using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MarpToPptx.Core.Models;
using MarpToPptx.Core.Themes;
using System.Text.RegularExpressions;

namespace MarpToPptx.Core.Parsing;

public sealed class MarpMarkdownParser
{
    // Matches the opening or self-closing <video> tag and captures the src attribute value.
    private static readonly Regex VideoTagRegex = new(
        @"<video\b[^>]*?\bsrc\s*=\s*(?:""([^""]*)""|'([^']*)'|(\S+?)(?:\s|/?>))",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    private readonly MarkdownPipeline _pipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .Build();

    public SlideDeck Parse(string markdown, string? sourcePath = null, string? themeCss = null)
    {
        var (frontMatter, body) = FrontMatterParser.Parse(markdown);
        var deck = new SlideDeck { SourcePath = sourcePath };
        foreach (var pair in frontMatter)
        {
            deck.FrontMatter[pair.Key] = pair.Value;
        }

        var defaultStyle = new SlideStyle();
        if (frontMatter.TryGetValue("theme", out var themeName))
        {
            defaultStyle = new SlideStyle { ThemeName = themeName };
        }

        if (frontMatter.TryGetValue("paginate", out var paginate) && bool.TryParse(paginate, out var paginateValue))
        {
            defaultStyle = new SlideStyle
            {
                ThemeName = defaultStyle.ThemeName,
                Paginate = paginateValue,
                ClassName = defaultStyle.ClassName,
                BackgroundColor = defaultStyle.BackgroundColor,
                BackgroundImage = defaultStyle.BackgroundImage,
            };
        }

        if (frontMatter.TryGetValue("class", out var className))
        {
            defaultStyle = new SlideStyle
            {
                ThemeName = defaultStyle.ThemeName,
                Paginate = defaultStyle.Paginate,
                ClassName = className,
                BackgroundColor = defaultStyle.BackgroundColor,
                BackgroundImage = defaultStyle.BackgroundImage,
            };
        }

        if (frontMatter.TryGetValue("backgroundImage", out var backgroundImage))
        {
            defaultStyle = new SlideStyle
            {
                ThemeName = defaultStyle.ThemeName,
                Paginate = defaultStyle.Paginate,
                ClassName = defaultStyle.ClassName,
                BackgroundImage = backgroundImage,
                BackgroundColor = defaultStyle.BackgroundColor,
            };
        }

        if (frontMatter.TryGetValue("backgroundColor", out var backgroundColor))
        {
            defaultStyle = new SlideStyle
            {
                ThemeName = defaultStyle.ThemeName,
                Paginate = defaultStyle.Paginate,
                ClassName = defaultStyle.ClassName,
                BackgroundImage = defaultStyle.BackgroundImage,
                BackgroundColor = backgroundColor,
            };
        }

        deck.Theme = MarpThemeParser.Parse(themeCss, themeName: defaultStyle.ThemeName ?? "default");

        foreach (var chunk in SlideTokenizer.SplitSlides(body))
        {
            var (style, cleaned, notes) = MarpDirectiveParser.Parse(chunk, defaultStyle);
            var slide = new Slide { Style = style, Notes = notes };
            foreach (var element in ParseElements(cleaned))
            {
                slide.Elements.Add(element);
            }

            deck.Slides.Add(slide);
        }

        return deck;
    }

    private IReadOnlyList<ISlideElement> ParseElements(string markdown)
    {
        var document = Markdown.Parse(markdown, _pipeline);
        var elements = new List<ISlideElement>();

        foreach (var block in document)
        {
            switch (block)
            {
                case HeadingBlock heading:
                    elements.Add(new HeadingElement(heading.Level, TrimSpans(ExtractInlineSpans(heading.Inline))));
                    break;
                case ParagraphBlock paragraph:
                    AppendParagraphElements(paragraph, elements);
                    break;
                case ListBlock list:
                    elements.Add(new BulletListElement(FlattenListItems(list).ToArray(), list.IsOrdered));
                    break;
                case FencedCodeBlock fencedCode:
                    elements.Add(new CodeBlockElement(fencedCode.Info ?? string.Empty, fencedCode.Lines.ToString() ?? string.Empty));
                    break;
                case CodeBlock codeBlock:
                    elements.Add(new CodeBlockElement(string.Empty, codeBlock.Lines.ToString() ?? string.Empty));
                    break;
                case Table table:
                    elements.Add(ParseTable(table));
                    break;
                case HtmlBlock htmlBlock:
                    AppendHtmlBlockElements(htmlBlock, elements);
                    break;
            }
        }

        return elements;
    }

    private static void AppendHtmlBlockElements(HtmlBlock htmlBlock, ICollection<ISlideElement> elements)
    {
        var html = htmlBlock.Lines.ToString();
        var matches = VideoTagRegex.Matches(html);

        if (matches.Count == 0)
        {
            return;
        }

        foreach (Match match in matches)
        {
            var src = match.Groups[1].Success ? match.Groups[1].Value
                : match.Groups[2].Success ? match.Groups[2].Value
                : match.Groups[3].Value;
            elements.Add(new VideoElement(src, string.Empty));
        }
    }

    private static void AppendParagraphElements(ParagraphBlock paragraph, ICollection<ISlideElement> elements)
    {
        var media = ExtractMediaElements(paragraph.Inline).ToArray();
        var spans = TrimSpans(ExtractInlineSpans(paragraph.Inline));
        var hasText = spans.Any(s => s.Text.Length > 0);

        if (media.Length > 0 && !hasText)
        {
            foreach (var item in media)
            {
                elements.Add(item);
            }

            return;
        }

        if (hasText)
        {
            elements.Add(new ParagraphElement(spans));
        }

        foreach (var item in media)
        {
            elements.Add(item);
        }
    }

    private static IEnumerable<BulletListItem> FlattenListItems(ListBlock list, int depth = 0)
    {
        foreach (var item in list.OfType<ListItemBlock>())
        {
            var spanFragments = new List<InlineSpan>();
            foreach (var child in item)
            {
                switch (child)
                {
                    case ParagraphBlock paragraph:
                        var childSpans = TrimSpans(ExtractInlineSpans(paragraph.Inline));
                        spanFragments.AddRange(childSpans);
                        break;
                    case FencedCodeBlock code:
                        spanFragments.Add(new InlineSpan(code.Lines.ToString() ?? string.Empty, Code: true));
                        break;
                    case ListBlock nested:
                        foreach (var nestedItem in FlattenListItems(nested, depth + 1))
                        {
                            yield return nestedItem;
                        }
                        break;
                }
            }

            if (spanFragments.Any(s => s.Text.Length > 0))
            {
                yield return new BulletListItem(spanFragments, depth);
            }
        }
    }

    private static TableElement ParseTable(Table table)
    {
        var rows = new List<TableRowModel>();
        foreach (var rowObj in table)
        {
            if (rowObj is not TableRow row)
            {
                continue;
            }

            var cells = new List<IReadOnlyList<InlineSpan>>();
            foreach (var cellObj in row)
            {
                if (cellObj is not TableCell cell)
                {
                    continue;
                }

                var cellSpans = new List<InlineSpan>();
                foreach (var block in cell)
                {
                    if (block is ParagraphBlock paragraph)
                    {
                        var spans = TrimSpans(ExtractInlineSpans(paragraph.Inline));
                        cellSpans.AddRange(spans);
                    }
                    else
                    {
                        var text = block.ToString() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            cellSpans.Add(new InlineSpan(text));
                        }
                    }
                }

                cells.Add(cellSpans);
            }

            rows.Add(new TableRowModel(cells, row.IsHeader));
        }

        var alignments = table.ColumnDefinitions
            .Select(col => col.Alignment switch
            {
                TableColumnAlign.Left => (TableColumnAlignment?)TableColumnAlignment.Left,
                TableColumnAlign.Center => TableColumnAlignment.Center,
                TableColumnAlign.Right => TableColumnAlignment.Right,
                _ => null,
            })
            .ToArray();

        return new TableElement(rows, alignments);
    }

    private static IEnumerable<ISlideElement> ExtractMediaElements(ContainerInline? inline)
    {
        if (inline is null)
        {
            yield break;
        }

        foreach (var child in inline)
        {
            switch (child)
            {
                case LinkInline link when link.IsImage:
                    var url = link.Url ?? string.Empty;
                    var altText = ExtractInlineText(link).Trim();
                    if (url.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase))
                    {
                        yield return new VideoElement(url, altText);
                    }
                    else
                    {
                        yield return new ImageElement(url, altText);
                    }

                    break;
                case HtmlInline htmlInline:
                    var htmlMatch = VideoTagRegex.Match(htmlInline.Tag);
                    if (htmlMatch.Success)
                    {
                        var src = htmlMatch.Groups[1].Success ? htmlMatch.Groups[1].Value
                            : htmlMatch.Groups[2].Success ? htmlMatch.Groups[2].Value
                            : htmlMatch.Groups[3].Value;
                        yield return new VideoElement(src, string.Empty);
                    }

                    break;
                case ContainerInline nested:
                    foreach (var nestedItem in ExtractMediaElements(nested))
                    {
                        yield return nestedItem;
                    }

                    break;
            }
        }
    }

    /// <summary>
    /// Extracts a list of <see cref="InlineSpan"/> from a Markdig inline tree, preserving
    /// bold, italic, inline-code, strikethrough formatting and non-image hyperlink URLs.
    /// A <see cref="LineBreakInline"/> node is emitted as a span with <c>Text = "\n"</c>.
    /// </summary>
    private static IReadOnlyList<InlineSpan> ExtractInlineSpans(
        ContainerInline? inline,
        bool bold = false,
        bool italic = false,
        bool strikethrough = false,
        bool code = false,
        string? hyperlinkUrl = null)
    {
        if (inline is null)
        {
            return [];
        }

        var spans = new List<InlineSpan>();
        foreach (var child in inline)
        {
            switch (child)
            {
                case LiteralInline literal:
                    spans.Add(new InlineSpan(literal.Content.ToString(), bold, italic, code, strikethrough, hyperlinkUrl));
                    break;
                case LineBreakInline:
                    spans.Add(new InlineSpan("\n"));
                    break;
                case CodeInline codeInline:
                    spans.Add(new InlineSpan(codeInline.Content, bold, italic, Code: true, strikethrough, hyperlinkUrl));
                    break;
                case EmphasisInline emphasis:
                {
                    var isBold = (emphasis.DelimiterChar is '*' or '_') && emphasis.DelimiterCount >= 2;
                    var isItalic = (emphasis.DelimiterChar is '*' or '_') && emphasis.DelimiterCount == 1;
                    var isStrikethrough = emphasis.DelimiterChar == '~';
                    spans.AddRange(ExtractInlineSpans(
                        emphasis,
                        bold || isBold,
                        italic || isItalic,
                        strikethrough || isStrikethrough,
                        code,
                        hyperlinkUrl));
                    break;
                }

                case LinkInline link when !link.IsImage:
                    spans.AddRange(ExtractInlineSpans(link, bold, italic, strikethrough, code, link.Url));
                    break;
                case ContainerInline nested:
                    spans.AddRange(ExtractInlineSpans(nested, bold, italic, strikethrough, code, hyperlinkUrl));
                    break;
            }
        }

        return spans;
    }

    /// <summary>
    /// Trims leading and trailing whitespace-only spans (including newline markers)
    /// from the span list, matching the trim behavior of the previous plain-text extraction.
    /// </summary>
    private static IReadOnlyList<InlineSpan> TrimSpans(IReadOnlyList<InlineSpan> spans)
    {
        var list = spans.ToList();

        while (list.Count > 0 && string.IsNullOrWhiteSpace(list[0].Text))
        {
            list.RemoveAt(0);
        }

        while (list.Count > 0 && string.IsNullOrWhiteSpace(list[^1].Text))
        {
            list.RemoveAt(list.Count - 1);
        }

        return list;
    }

    private static string ExtractInlineText(ContainerInline? inline)
    {
        if (inline is null)
        {
            return string.Empty;
        }

        var parts = new List<string>();
        foreach (var child in inline)
        {
            switch (child)
            {
                case LiteralInline literal:
                    parts.Add(literal.Content.ToString());
                    break;
                case LineBreakInline:
                    parts.Add("\n");
                    break;
                case CodeInline code:
                    parts.Add(code.Content);
                    break;
                case LinkInline link when !link.IsImage:
                    parts.Add(ExtractInlineText(link));
                    break;
                case EmphasisInline emphasis:
                    parts.Add(ExtractInlineText(emphasis));
                    break;
                case ContainerInline nested:
                    parts.Add(ExtractInlineText(nested));
                    break;
            }
        }

        return string.Concat(parts);
    }
}
