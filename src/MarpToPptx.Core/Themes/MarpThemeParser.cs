namespace MarpToPptx.Core.Themes;

public static partial class MarpThemeParser
{
    public static ThemeDefinition Parse(string? css, string? themeName = null)
    {
        var theme = ThemeDefinition.Default with { Name = themeName ?? "default" };
        if (string.IsNullOrWhiteSpace(css))
        {
            return theme;
        }

        css = StripComments(css);

        var bodyStyle = theme.Body;
        var codeStyle = theme.Code;
        var inlineCodeStyle = theme.InlineCode;
        var preExplicitlySet = false;
        var headingStyles = theme.Headings.ToDictionary(static pair => pair.Key, static pair => pair.Value);
        var background = theme.BackgroundColor;
        var backgroundImage = theme.BackgroundImage;
        var backgroundSize = theme.BackgroundSize;
        var backgroundPosition = theme.BackgroundPosition;
        var textColor = theme.TextColor;
        var fontFamily = theme.FontFamily;
        var monospace = theme.MonospaceFontFamily;
        var slidePadding = theme.SlidePadding;
        var classVariants = new Dictionary<string, ClassVariant>(StringComparer.OrdinalIgnoreCase);

        foreach (var match in RuleRegex().Matches(css).Cast<System.Text.RegularExpressions.Match>())
        {
            var selectors = match.Groups[1].Value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var declarations = ParseDeclarations(match.Groups[2].Value);

            foreach (var selector in selectors)
            {
                switch (selector)
                {
                    case ":root":
                    case "body":
                    case "section":
                        fontFamily = declarations.TryGetValue("font-family", out var bodyFont) ? NormalizeFontFamily(bodyFont) : fontFamily;
                        textColor = declarations.TryGetValue("color", out var bodyColor) ? bodyColor : textColor;
                        background = declarations.TryGetValue("background", out var bg) ? ExtractColor(bg) : background;
                        background = declarations.TryGetValue("background-color", out var bgColor) ? bgColor : background;
                        backgroundImage = declarations.TryGetValue("background-image", out var bgImage) ? ExtractUrl(bgImage) : backgroundImage;
                        if (declarations.TryGetValue("background", out var bgShorthand))
                        {
                            var urlFromShorthand = ExtractUrl(bgShorthand);
                            if (urlFromShorthand is not null)
                            {
                                backgroundImage = urlFromShorthand;
                            }
                        }

                        backgroundSize = declarations.TryGetValue("background-size", out var bgSize) ? bgSize.Trim() : backgroundSize;
                        backgroundPosition = declarations.TryGetValue("background-position", out var bgPos) ? bgPos.Trim() : backgroundPosition;
                        slidePadding = declarations.TryGetValue("padding", out var padding) ? ParseSpacing(padding) : slidePadding;
                        bodyStyle = bodyStyle with
                        {
                            Color = textColor,
                            FontFamily = fontFamily,
                            FontSize = declarations.TryGetValue("font-size", out var size) ? ParseFontSize(size, bodyStyle.FontSize) : bodyStyle.FontSize,
                            Bold = declarations.TryGetValue("font-weight", out var bodyWeight) ? ParseFontWeight(bodyWeight) ?? bodyStyle.Bold : bodyStyle.Bold,
                            LineHeight = declarations.TryGetValue("line-height", out var bodyLineHeight) ? ParseLineHeight(bodyLineHeight) : bodyStyle.LineHeight,
                            LetterSpacing = declarations.TryGetValue("letter-spacing", out var bodyLetterSpacing) ? ParseFontSize(bodyLetterSpacing, 0) : bodyStyle.LetterSpacing,
                            TextTransform = declarations.TryGetValue("text-transform", out var bodyTextTransform) ? bodyTextTransform.Trim().ToLowerInvariant() : bodyStyle.TextTransform,
                        };
                        break;
                    case "pre":
                        preExplicitlySet = true;
                        monospace = declarations.TryGetValue("font-family", out var preFont) ? NormalizeFontFamily(preFont) : monospace;
                        codeStyle = codeStyle with
                        {
                            FontFamily = monospace,
                            FontSize = declarations.TryGetValue("font-size", out var preSize) ? ParseFontSize(preSize, codeStyle.FontSize) : codeStyle.FontSize,
                            Color = declarations.TryGetValue("color", out var preColor) ? preColor : codeStyle.Color,
                            BackgroundColor = declarations.TryGetValue("background-color", out var preBgColor) ? preBgColor : codeStyle.BackgroundColor,
                            LineHeight = declarations.TryGetValue("line-height", out var preLineHeight) ? ParseLineHeight(preLineHeight) : codeStyle.LineHeight,
                            LetterSpacing = declarations.TryGetValue("letter-spacing", out var preLetterSpacing) ? ParseFontSize(preLetterSpacing, 0) : codeStyle.LetterSpacing,
                        };
                        if (declarations.TryGetValue("background", out var preBg))
                        {
                            var extractedColor = ExtractColor(preBg);
                            if (!string.IsNullOrWhiteSpace(extractedColor))
                            {
                                codeStyle = codeStyle with { BackgroundColor = extractedColor };
                            }
                        }

                        break;
                    case "code":
                        monospace = declarations.TryGetValue("font-family", out var codeFont) ? NormalizeFontFamily(codeFont) : monospace;
                        inlineCodeStyle = inlineCodeStyle with
                        {
                            FontFamily = monospace,
                            FontSize = declarations.TryGetValue("font-size", out var codeSize) ? ParseFontSize(codeSize, inlineCodeStyle.FontSize) : inlineCodeStyle.FontSize,
                            Color = declarations.TryGetValue("color", out var codeColor) ? codeColor : inlineCodeStyle.Color,
                            BackgroundColor = declarations.TryGetValue("background-color", out var codeBgColor) ? codeBgColor : inlineCodeStyle.BackgroundColor,
                            LineHeight = declarations.TryGetValue("line-height", out var codeLineHeight) ? ParseLineHeight(codeLineHeight) : inlineCodeStyle.LineHeight,
                            LetterSpacing = declarations.TryGetValue("letter-spacing", out var codeLetterSpacing) ? ParseFontSize(codeLetterSpacing, 0) : inlineCodeStyle.LetterSpacing,
                        };
                        if (declarations.TryGetValue("background", out var codeBg))
                        {
                            var extractedColor = ExtractColor(codeBg);
                            if (!string.IsNullOrWhiteSpace(extractedColor))
                            {
                                inlineCodeStyle = inlineCodeStyle with { BackgroundColor = extractedColor };
                            }
                        }

                        // When no explicit `pre` rule has been seen, `code` sets both
                        // inline and block code styles (the universal code selector).
                        // Apply declarations independently to codeStyle from its own defaults
                        // so block-code colours are preserved (dark bg/light text).
                        if (!preExplicitlySet)
                        {
                            codeStyle = codeStyle with
                            {
                                FontFamily = monospace,
                                FontSize = declarations.TryGetValue("font-size", out var blockCodeSize) ? ParseFontSize(blockCodeSize, codeStyle.FontSize) : codeStyle.FontSize,
                                Color = declarations.TryGetValue("color", out var blockCodeColor) ? blockCodeColor : codeStyle.Color,
                                BackgroundColor = declarations.TryGetValue("background-color", out var blockCodeBgColor) ? blockCodeBgColor : codeStyle.BackgroundColor,
                                LineHeight = declarations.TryGetValue("line-height", out var blockCodeLineHeight) ? ParseLineHeight(blockCodeLineHeight) : codeStyle.LineHeight,
                                LetterSpacing = declarations.TryGetValue("letter-spacing", out var blockCodeLetterSpacing) ? ParseFontSize(blockCodeLetterSpacing, 0) : codeStyle.LetterSpacing,
                            };

                            if (declarations.TryGetValue("background", out var blockCodeBg))
                            {
                                var extractedBlockColor = ExtractColor(blockCodeBg);
                                if (!string.IsNullOrWhiteSpace(extractedBlockColor))
                                {
                                    codeStyle = codeStyle with { BackgroundColor = extractedBlockColor };
                                }
                            }
                        }

                        break;
                    default:
                        if (selector.StartsWith("h", StringComparison.OrdinalIgnoreCase) && selector.Length == 2 && char.IsDigit(selector[1]))
                        {
                            var level = selector[1] - '0';
                            var current = headingStyles[level];
                            headingStyles[level] = current with
                            {
                                FontFamily = declarations.TryGetValue("font-family", out var headingFont) ? NormalizeFontFamily(headingFont) : current.FontFamily,
                                FontSize = declarations.TryGetValue("font-size", out var headingSize) ? ParseFontSize(headingSize, current.FontSize) : current.FontSize,
                                Color = declarations.TryGetValue("color", out var headingColor) ? headingColor : current.Color,
                                Bold = declarations.TryGetValue("font-weight", out var headingWeight) ? ParseFontWeight(headingWeight) ?? current.Bold : current.Bold,
                                LineHeight = declarations.TryGetValue("line-height", out var headingLineHeight) ? ParseLineHeight(headingLineHeight) : current.LineHeight,
                                LetterSpacing = declarations.TryGetValue("letter-spacing", out var headingLetterSpacing) ? ParseFontSize(headingLetterSpacing, 0) : current.LetterSpacing,
                                TextTransform = declarations.TryGetValue("text-transform", out var headingTextTransform) ? headingTextTransform.Trim().ToLowerInvariant() : current.TextTransform,
                            };
                        }
                        else if (TryParseClassSelector(selector, out var className, out var subElement))
                        {
                            ApplyClassVariant(classVariants, className!, subElement, declarations, bodyStyle, headingStyles, inlineCodeStyle, codeStyle);
                        }
                        break;
                }
            }
        }

        return new ThemeDefinition
        {
            Name = themeName ?? theme.Name,
            FontFamily = fontFamily,
            MonospaceFontFamily = monospace,
            TextColor = textColor,
            BackgroundColor = background,
            BackgroundImage = backgroundImage,
            BackgroundSize = backgroundSize,
            BackgroundPosition = backgroundPosition,
            AccentColor = theme.AccentColor,
            SlidePadding = slidePadding,
            Body = bodyStyle,
            Code = codeStyle,
            InlineCode = inlineCodeStyle,
            Headings = headingStyles,
            ClassVariants = classVariants,
        };
    }

    private static string StripComments(string css)
        => CommentRegex().Replace(css, " ");

    private static Dictionary<string, string> ParseDeclarations(string block)
    {
        var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in block.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var separator = part.IndexOf(':');
            if (separator <= 0)
            {
                continue;
            }

            declarations[part[..separator].Trim()] = part[(separator + 1)..].Trim();
        }

        return declarations;
    }

    private static Spacing ParseSpacing(string value)
    {
        var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var numbers = parts.Select(part => ParseFontSize(part, 0)).ToArray();
        return numbers.Length switch
        {
            1 => Spacing.Uniform(numbers[0]),
            2 => new Spacing(numbers[0], numbers[1], numbers[0], numbers[1]),
            3 => new Spacing(numbers[0], numbers[1], numbers[2], numbers[1]),
            4 => new Spacing(numbers[0], numbers[1], numbers[2], numbers[3]),
            _ => ThemeDefinition.Default.SlidePadding,
        };
    }

    private static string NormalizeFontFamily(string value)
        => value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)[0].Trim('"', '\'');

    private static double ParseFontSize(string value, double fallback)
    {
        var normalized = value.Trim().ToLowerInvariant();
        if (!double.TryParse(new string(normalized.TakeWhile(ch => char.IsDigit(ch) || ch == '.').ToArray()), out var number))
        {
            return fallback;
        }

        if (normalized.EndsWith("rem", StringComparison.Ordinal))
        {
            return number * 16 * 0.75;
        }

        if (normalized.EndsWith("px", StringComparison.Ordinal))
        {
            return number * 0.75;
        }

        return number;
    }

    private static string ExtractColor(string value)
    {
        var colorMatch = ColorRegex().Match(value);
        return colorMatch.Success ? colorMatch.Value : value;
    }

    private static string? ExtractUrl(string value)
    {
        var urlMatch = UrlRegex().Match(value);
        if (!urlMatch.Success)
        {
            return null;
        }

        return urlMatch.Groups[1].Value.Trim('"', '\'');
    }

    private static bool? ParseFontWeight(string value)
    {
        var normalized = value.Trim().ToLowerInvariant();
        if (normalized == "bold" || normalized == "bolder")
        {
            return true;
        }

        if (normalized == "normal" || normalized == "lighter")
        {
            return false;
        }

        if (int.TryParse(normalized, out var weight))
        {
            return weight >= 600;
        }

        return null;
    }

    private static double? ParseLineHeight(string value)
    {
        var normalized = value.Trim().ToLowerInvariant();
        if (normalized == "normal")
        {
            return null;
        }

        if (normalized.EndsWith('%'))
        {
            if (double.TryParse(normalized[..^1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pct))
            {
                return pct / 100.0;
            }
        }

        if (!double.TryParse(new string(normalized.TakeWhile(ch => char.IsDigit(ch) || ch == '.').ToArray()), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number))
        {
            return null;
        }

        return number;
    }

    /// <summary>
    /// Tries to match selectors like <c>section.lead</c> or <c>section.lead h1</c>.
    /// Returns the class name and optional sub-element (e.g. "h1").
    /// </summary>
    private static bool TryParseClassSelector(string selector, out string? className, out string? subElement)
    {
        className = null;
        subElement = null;

        // Match "section.name" or "section.name subSelector"
        if (!selector.StartsWith("section.", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var afterDot = selector[8..]; // skip "section."
        var spaceIndex = afterDot.IndexOf(' ');
        if (spaceIndex < 0)
        {
            className = afterDot.Trim();
            return className.Length > 0;
        }

        className = afterDot[..spaceIndex].Trim();
        subElement = afterDot[(spaceIndex + 1)..].Trim();
        return className.Length > 0;
    }

    /// <summary>
    /// Applies CSS declarations to a class variant, creating it if it doesn't exist yet.
    /// </summary>
    private static void ApplyClassVariant(
        Dictionary<string, ClassVariant> variants,
        string className,
        string? subElement,
        Dictionary<string, string> declarations,
        TextStyle baseBody,
        Dictionary<int, TextStyle> baseHeadings,
        TextStyle? baseInlineCode = null,
        TextStyle? baseCode = null)
    {
        baseInlineCode ??= ThemeDefinition.Default.InlineCode;
        baseCode ??= ThemeDefinition.Default.Code;

        if (!variants.TryGetValue(className, out var variant))
        {
            variant = new ClassVariant();
        }

        if (string.IsNullOrEmpty(subElement))
        {
            // section.classname { ... } — section-level overrides.
            // Apply background shorthand first, then background-color so that
            // background-color takes precedence (matching the main loop priority).
            var backgroundColor = variant.BackgroundColor;
            if (declarations.TryGetValue("background", out var bg))
            {
                var extractedColor = ExtractColor(bg);
                if (!string.IsNullOrWhiteSpace(extractedColor))
                {
                    backgroundColor = extractedColor;
                }
            }
            if (declarations.TryGetValue("background-color", out var bgColor))
            {
                backgroundColor = bgColor;
            }

            variant = variant with { BackgroundColor = backgroundColor };

            // Build or update body text style from the CSS declarations.
            var body = variant.Body ?? baseBody;
            variant = variant with
            {
                Body = body with
                {
                    Color = declarations.TryGetValue("color", out var bodyColor) ? bodyColor : body.Color,
                    FontFamily = declarations.TryGetValue("font-family", out var bodyFont) ? NormalizeFontFamily(bodyFont) : body.FontFamily,
                    FontSize = declarations.TryGetValue("font-size", out var bodySize) ? ParseFontSize(bodySize, body.FontSize) : body.FontSize,
                    Bold = declarations.TryGetValue("font-weight", out var bodyWeight) ? ParseFontWeight(bodyWeight) ?? body.Bold : body.Bold,
                },
            };
        }
        else if (string.Equals(subElement, "code", StringComparison.OrdinalIgnoreCase))
        {
            var inlineCode = variant.InlineCode ?? baseInlineCode;
            var code = variant.Code ?? baseCode;

            inlineCode = inlineCode with
            {
                FontFamily = declarations.TryGetValue("font-family", out var inlineCodeFont) ? NormalizeFontFamily(inlineCodeFont) : inlineCode.FontFamily,
                FontSize = declarations.TryGetValue("font-size", out var inlineCodeSize) ? ParseFontSize(inlineCodeSize, inlineCode.FontSize) : inlineCode.FontSize,
                Color = declarations.TryGetValue("color", out var inlineCodeColor) ? inlineCodeColor : inlineCode.Color,
                BackgroundColor = declarations.TryGetValue("background-color", out var inlineCodeBgColor) ? inlineCodeBgColor : inlineCode.BackgroundColor,
                LineHeight = declarations.TryGetValue("line-height", out var inlineCodeLineHeight) ? ParseLineHeight(inlineCodeLineHeight) : inlineCode.LineHeight,
                LetterSpacing = declarations.TryGetValue("letter-spacing", out var inlineCodeLetterSpacing) ? ParseFontSize(inlineCodeLetterSpacing, 0) : inlineCode.LetterSpacing,
            };

            if (declarations.TryGetValue("background", out var inlineCodeBg))
            {
                var extractedColor = ExtractColor(inlineCodeBg);
                if (!string.IsNullOrWhiteSpace(extractedColor))
                {
                    inlineCode = inlineCode with { BackgroundColor = extractedColor };
                }
            }

            code = code with
            {
                FontFamily = declarations.TryGetValue("font-family", out var codeFont) ? NormalizeFontFamily(codeFont) : code.FontFamily,
                FontSize = declarations.TryGetValue("font-size", out var codeSize) ? ParseFontSize(codeSize, code.FontSize) : code.FontSize,
                Color = declarations.TryGetValue("color", out var codeColor) ? codeColor : code.Color,
                BackgroundColor = declarations.TryGetValue("background-color", out var codeBgColor) ? codeBgColor : code.BackgroundColor,
                LineHeight = declarations.TryGetValue("line-height", out var codeLineHeight) ? ParseLineHeight(codeLineHeight) : code.LineHeight,
                LetterSpacing = declarations.TryGetValue("letter-spacing", out var codeLetterSpacing) ? ParseFontSize(codeLetterSpacing, 0) : code.LetterSpacing,
            };

            if (declarations.TryGetValue("background", out var codeBg))
            {
                var extractedColor = ExtractColor(codeBg);
                if (!string.IsNullOrWhiteSpace(extractedColor))
                {
                    code = code with { BackgroundColor = extractedColor };
                }
            }

            variant = variant with { InlineCode = inlineCode, Code = code };
        }
        else if (string.Equals(subElement, "pre", StringComparison.OrdinalIgnoreCase))
        {
            var code = variant.Code ?? baseCode;
            code = code with
            {
                FontFamily = declarations.TryGetValue("font-family", out var codeFont) ? NormalizeFontFamily(codeFont) : code.FontFamily,
                FontSize = declarations.TryGetValue("font-size", out var codeSize) ? ParseFontSize(codeSize, code.FontSize) : code.FontSize,
                Color = declarations.TryGetValue("color", out var codeColor) ? codeColor : code.Color,
                BackgroundColor = declarations.TryGetValue("background-color", out var codeBgColor) ? codeBgColor : code.BackgroundColor,
                LineHeight = declarations.TryGetValue("line-height", out var codeLineHeight) ? ParseLineHeight(codeLineHeight) : code.LineHeight,
                LetterSpacing = declarations.TryGetValue("letter-spacing", out var codeLetterSpacing) ? ParseFontSize(codeLetterSpacing, 0) : code.LetterSpacing,
            };

            if (declarations.TryGetValue("background", out var codeBg))
            {
                var extractedColor = ExtractColor(codeBg);
                if (!string.IsNullOrWhiteSpace(extractedColor))
                {
                    code = code with { BackgroundColor = extractedColor };
                }
            }

            variant = variant with { Code = code };
        }
        else if (subElement.StartsWith("h", StringComparison.OrdinalIgnoreCase) && subElement.Length == 2 && char.IsDigit(subElement[1]))
        {
            // section.classname h1 { ... } — heading override within class.
            var level = subElement[1] - '0';
            if (!baseHeadings.TryGetValue(level, out var baseHeading))
            {
                // Ignore invalid heading levels (e.g., h0, h7) that have no base heading style.
                variants[className] = variant;
                return;
            }

            var headings = variant.Headings is not null
                ? new Dictionary<int, TextStyle>(variant.Headings)
                : new Dictionary<int, TextStyle>();
            var current = headings.TryGetValue(level, out var existing) ? existing : baseHeading;
            headings[level] = current with
            {
                FontFamily = declarations.TryGetValue("font-family", out var hFont) ? NormalizeFontFamily(hFont) : current.FontFamily,
                FontSize = declarations.TryGetValue("font-size", out var hSize) ? ParseFontSize(hSize, current.FontSize) : current.FontSize,
                Color = declarations.TryGetValue("color", out var hColor) ? hColor : current.Color,
                Bold = declarations.TryGetValue("font-weight", out var hWeight) ? ParseFontWeight(hWeight) ?? current.Bold : current.Bold,
            };
            variant = variant with { Headings = headings };
        }

        variants[className] = variant;
    }

    [System.Text.RegularExpressions.GeneratedRegex(@"([^{}]+)\{([^}]*)\}")]
    private static partial System.Text.RegularExpressions.Regex RuleRegex();

    [System.Text.RegularExpressions.GeneratedRegex(@"/\*.*?\*/", System.Text.RegularExpressions.RegexOptions.Singleline)]
    private static partial System.Text.RegularExpressions.Regex CommentRegex();

    [System.Text.RegularExpressions.GeneratedRegex(@"#(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{3})|rgba?\([^)]*\)")]
    private static partial System.Text.RegularExpressions.Regex ColorRegex();

    [System.Text.RegularExpressions.GeneratedRegex(@"url\(\s*([^)]*)\s*\)")]
    private static partial System.Text.RegularExpressions.Regex UrlRegex();
}
