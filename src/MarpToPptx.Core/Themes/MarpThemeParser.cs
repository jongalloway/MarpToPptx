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

        var bodyStyle = theme.Body;
        var codeStyle = theme.Code;
        var headingStyles = theme.Headings.ToDictionary(static pair => pair.Key, static pair => pair.Value);
        var background = theme.BackgroundColor;
        var textColor = theme.TextColor;
        var fontFamily = theme.FontFamily;
        var monospace = theme.MonospaceFontFamily;
        var slidePadding = theme.SlidePadding;

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
                        slidePadding = declarations.TryGetValue("padding", out var padding) ? ParseSpacing(padding) : slidePadding;
                        bodyStyle = bodyStyle with
                        {
                            Color = textColor,
                            FontFamily = fontFamily,
                            FontSize = declarations.TryGetValue("font-size", out var size) ? ParseFontSize(size, bodyStyle.FontSize) : bodyStyle.FontSize,
                        };
                        break;
                    case "pre":
                    case "code":
                        monospace = declarations.TryGetValue("font-family", out var codeFont) ? NormalizeFontFamily(codeFont) : monospace;
                        codeStyle = codeStyle with
                        {
                            FontFamily = monospace,
                            FontSize = declarations.TryGetValue("font-size", out var codeSize) ? ParseFontSize(codeSize, codeStyle.FontSize) : codeStyle.FontSize,
                            Color = declarations.TryGetValue("color", out var codeColor) ? codeColor : codeStyle.Color,
                            BackgroundColor = declarations.TryGetValue("background", out var codeBg) ? ExtractColor(codeBg) : codeStyle.BackgroundColor,
                        };
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
                            };
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
            AccentColor = theme.AccentColor,
            SlidePadding = slidePadding,
            Body = bodyStyle,
            Code = codeStyle,
            Headings = headingStyles,
        };
    }

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

    [System.Text.RegularExpressions.GeneratedRegex(@"([^{}]+)\{([^}]*)\}")]
    private static partial System.Text.RegularExpressions.Regex RuleRegex();

    [System.Text.RegularExpressions.GeneratedRegex(@"#(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{3})|rgba?\([^)]*\)")]
    private static partial System.Text.RegularExpressions.Regex ColorRegex();
}
