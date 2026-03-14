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
            "theme" => Clone(style, themeName: value),
            "paginate" => Clone(style, paginate: bool.TryParse(value, out var p) ? p : null),
            "layout" => Clone(style, layout: value),
            "class" => Clone(style, className: value),
            "backgroundimage" => Clone(style, backgroundImage: UnwrapUrl(value)),
            "backgroundsize" => Clone(style, backgroundSize: value),
            "backgroundposition" => Clone(style, backgroundPosition: value),
            "backgroundcolor" => Clone(style, backgroundColor: value),
            "header" => Clone(style, header: value),
            "footer" => Clone(style, footer: value),
            "transition" => Clone(style, transition: ParseTransition(value)),
            _ => style,
        };
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
        string? themeName = null,
        bool? paginate = null,
        string? layout = null,
        string? className = null,
        string? backgroundImage = null,
        string? backgroundSize = null,
        string? backgroundPosition = null,
        string? backgroundColor = null,
        string? header = null,
        string? footer = null,
        SlideTransition? transition = null)
    {
        var clone = new SlideStyle
        {
            ThemeName = themeName ?? source.ThemeName,
            Paginate = paginate ?? source.Paginate,
            Layout = layout ?? source.Layout,
            ClassName = className ?? source.ClassName,
            BackgroundImage = backgroundImage ?? source.BackgroundImage,
            BackgroundSize = backgroundSize ?? source.BackgroundSize,
            BackgroundPosition = backgroundPosition ?? source.BackgroundPosition,
            BackgroundColor = backgroundColor ?? source.BackgroundColor,
            Header = header ?? source.Header,
            Footer = footer ?? source.Footer,
            Transition = transition ?? source.Transition,
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
