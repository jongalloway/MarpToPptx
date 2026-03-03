using MarpToPptx.Core.Models;

namespace MarpToPptx.Core.Parsing;

public static partial class MarpDirectiveParser
{
    public static (SlideStyle Style, string MarkdownWithoutDirectives) Parse(string markdown, SlideStyle? inheritedStyle = null)
    {
        var style = Clone(inheritedStyle ?? new SlideStyle());

        foreach (var match in DirectiveRegex().Matches(markdown).Cast<System.Text.RegularExpressions.Match>())
        {
            var key = match.Groups[1].Value.Trim();
            var value = match.Groups[2].Value.Trim();
            style.Directives[key] = value;

            switch (key.ToLowerInvariant())
            {
                case "theme":
                    style = Clone(style, themeName: value);
                    break;
                case "paginate":
                    style = Clone(style, paginate: bool.TryParse(value, out var paginate) ? paginate : null);
                    break;
                case "class":
                    style = Clone(style, className: value);
                    break;
                case "backgroundimage":
                    style = Clone(style, backgroundImage: UnwrapUrl(value));
                    break;
                case "backgroundcolor":
                    style = Clone(style, backgroundColor: value);
                    break;
            }
        }

        var cleaned = DirectiveRegex().Replace(markdown, string.Empty).Trim();
        return (style, cleaned);
    }

    private static SlideStyle Clone(
        SlideStyle source,
        string? themeName = null,
        bool? paginate = null,
        string? className = null,
        string? backgroundImage = null,
        string? backgroundColor = null)
    {
        var clone = new SlideStyle
        {
            ThemeName = themeName ?? source.ThemeName,
            Paginate = paginate ?? source.Paginate,
            ClassName = className ?? source.ClassName,
            BackgroundImage = backgroundImage ?? source.BackgroundImage,
            BackgroundColor = backgroundColor ?? source.BackgroundColor,
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
}
