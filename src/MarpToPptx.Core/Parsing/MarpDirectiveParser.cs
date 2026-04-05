using MarpToPptx.Core.Models;

namespace MarpToPptx.Core.Parsing;

public static partial class MarpDirectiveParser
{
    public static SlideStyle ApplyDirective(SlideStyle style, string key, string value)
    {
        var updatedStyle = Clone(style);
        updatedStyle.Directives[key] = value;
        return ApplyKnownDirective(updatedStyle, key, value);
    }

    /// <summary>
    /// Parses HTML-comment directives from a single slide chunk.
    /// Returns the effective style (local + spot directives applied),
    /// the carry-forward style (local directives only, for propagation to subsequent slides),
    /// the cleaned markdown, and any presenter notes.
    /// Spot directives use a <c>_</c> prefix (e.g. <c>_class</c>) and apply only to the current slide.
    /// </summary>
    public static (SlideStyle EffectiveStyle, SlideStyle CarryForwardStyle, string MarkdownWithoutDirectives, string? Notes) Parse(string markdown, SlideStyle? inheritedStyle = null)
    {
        var localStyle = Clone(inheritedStyle ?? new SlideStyle());
        var spotOverrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var noteLines = new List<string>();

        var anyComment = AnyHtmlCommentRegex();
        var directive = DirectiveRegex();

        // Single pass over all HTML comments: process directives and collect notes.
        foreach (var match in anyComment.Matches(markdown).Cast<System.Text.RegularExpressions.Match>())
        {
            var directiveMatch = directive.Match(match.Value);
            if (directiveMatch.Success)
            {
                var rawKey = directiveMatch.Groups[1].Value.Trim();
                var value = directiveMatch.Groups[2].Value.Trim();

                // Spot directives use a _ prefix and apply only to the current slide.
                var isSpot = rawKey.StartsWith('_');
                var key = isSpot ? rawKey[1..] : rawKey;

                // Skip directives with an empty key (e.g. <!-- _: value -->).
                if (string.IsNullOrEmpty(key))
                {
                    continue;
                }

                if (string.Equals(key, "slideid", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(key, "slide-id", StringComparison.OrdinalIgnoreCase))
                {
                    spotOverrides[key] = value;
                    continue;
                }

                if (isSpot)
                {
                    spotOverrides[key] = value;
                }
                else
                {
                    localStyle = ApplyDirective(localStyle, key, value);
                }
            }
            else
            {
                var noteText = match.Groups[1].Value.Trim();
                if (!string.IsNullOrWhiteSpace(noteText))
                {
                    noteLines.Add(noteText);
                }
            }
        }

        // Build the carry-forward style (local directives only).
        var carryForwardStyle = Clone(localStyle);

        // Layer spot directives on top for the effective style.
        var effectiveStyle = Clone(localStyle);
        foreach (var (key, value) in spotOverrides)
        {
            effectiveStyle = ApplyDirective(effectiveStyle, key, value);
        }

        // Strip all HTML comments (directives and notes) from the cleaned output.
        var cleaned = anyComment.Replace(markdown, string.Empty).Trim();
        var notes = noteLines.Count > 0 ? string.Join("\n", noteLines) : null;
        return (effectiveStyle, carryForwardStyle, cleaned, notes);
    }

    private static SlideStyle ApplyKnownDirective(SlideStyle style, string key, string value)
    {
        return key.ToLowerInvariant() switch
        {
            "slideid" or "slide-id" => Clone(style, slideId: value),
            "theme" => Clone(style, themeName: value),
            "paginate" => Clone(style, paginate: bool.TryParse(value, out var p) ? p : null),
            "layout" => Clone(style, layout: value),
            "class" => Clone(style, className: value),
            "backgroundimage" => Clone(style, backgroundImage: UnwrapUrl(value)),
            "backgroundsize" => Clone(style, backgroundSize: value),
            "backgroundposition" => Clone(style, backgroundPosition: value),
            "backgroundcolor" => Clone(style, backgroundColor: value),
            "color" => Clone(style, color: value),
            "header" => Clone(style, header: value),
            "footer" => Clone(style, footer: value),
            "transition" => Clone(style, transition: ParseTransition(value)),
            "fontsize" or "font-size" => Clone(style, fontSize: ParseFontSizeDirective(value)),
            "shrinktofit" or "shrink-to-fit" => Clone(style, shrinkToFit: value),
            _ => style,
        };
    }

    /// <summary>
    /// Parses a font-size directive value into hundredths of a point.
    /// Accepts <c>20pt</c>, <c>20</c> (assumed pt), or <c>2000</c> (already in hundredths of a point).
    /// Returns <see langword="null"/> if the value cannot be parsed.
    /// </summary>
    public static int? ParseFontSizeDirective(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var trimmed = value.Trim();

        // Explicit "pt" suffix: "20pt" → 2000
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var numPart = trimmed[..^2].Trim();
            if (double.TryParse(numPart, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var pts)
                && double.IsFinite(pts) && pts > 0)
            {
                return (int)Math.Round(pts * 100);
            }
            return null;
        }

        // Bare number: treat as pt when <= 999, otherwise as hundredths of a point already.
        if (double.TryParse(trimmed, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var bare)
            && double.IsFinite(bare) && bare > 0)
        {
            // Heuristic: values <= 999 are treated as point values (e.g. "20" → 2000).
            // Values >= 1000 are treated as already in hundredths of a point (e.g. "2000" → 2000).
            return bare <= 999
                ? (int)Math.Round(bare * 100)
                : (int)Math.Round(bare);
        }

        return null;
    }

    private static SlideTransition? ParseTransition(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var tokens = value.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var type = tokens[0];
        string? direction = null;
        int? durationMs = null;

        foreach (var token in tokens.Skip(1))
        {
            if (token.StartsWith("dir:", StringComparison.OrdinalIgnoreCase))
            {
                var dir = token[4..];
                if (!string.IsNullOrEmpty(dir))
                {
                    direction = dir;
                }
            }
            else if (token.StartsWith("dur:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(token[4..], out var dur) &&
                     dur > 0)
            {
                durationMs = dur;
            }
        }

        return new SlideTransition(type, direction, durationMs);
    }

    private static SlideStyle Clone(
        SlideStyle source,
        string? slideId = null,
        string? themeName = null,
        bool? paginate = null,
        string? layout = null,
        string? className = null,
        string? backgroundImage = null,
        string? backgroundSize = null,
        string? backgroundPosition = null,
        string? backgroundColor = null,
        string? color = null,
        string? header = null,
        string? footer = null,
        SlideTransition? transition = null,
        int? fontSize = null,
        string? shrinkToFit = null)
    {
        var clone = new SlideStyle
        {
            SlideId = slideId ?? source.SlideId,
            ThemeName = themeName ?? source.ThemeName,
            Paginate = paginate ?? source.Paginate,
            Layout = layout ?? source.Layout,
            ClassName = className ?? source.ClassName,
            BackgroundImage = backgroundImage ?? source.BackgroundImage,
            BackgroundSize = backgroundSize ?? source.BackgroundSize,
            BackgroundPosition = backgroundPosition ?? source.BackgroundPosition,
            BackgroundColor = backgroundColor ?? source.BackgroundColor,
            Color = color ?? source.Color,
            Header = header ?? source.Header,
            Footer = footer ?? source.Footer,
            Transition = transition ?? source.Transition,
            FontSize = fontSize ?? source.FontSize,
            ShrinkToFit = shrinkToFit ?? source.ShrinkToFit,
        };

        foreach (var pair in source.Directives)
        {
            clone.Directives[pair.Key] = pair.Value;
        }

        return clone;
    }

    private static string UnwrapUrl(string value)
    {
        if (!value.StartsWith("url(", StringComparison.OrdinalIgnoreCase) || !value.EndsWith(")", StringComparison.Ordinal))
        {
            return value.Trim('"', '\'');
        }

        return value[4..^1].Trim().Trim('"', '\'');
    }

    [System.Text.RegularExpressions.GeneratedRegex(@"<!--\s*([\w-]+)\s*:\s*(.*?)\s*-->", System.Text.RegularExpressions.RegexOptions.Singleline)]
    private static partial System.Text.RegularExpressions.Regex DirectiveRegex();

    [System.Text.RegularExpressions.GeneratedRegex(@"<!--(.*?)-->", System.Text.RegularExpressions.RegexOptions.Singleline)]
    private static partial System.Text.RegularExpressions.Regex AnyHtmlCommentRegex();
}
