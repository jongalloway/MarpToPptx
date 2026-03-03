using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MarpToPptx.Core.Models;
using MarpToPptx.Core.Themes;

namespace MarpToPptx.Core.Parsing;

public sealed class MarpMarkdownParser
{
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
            var (style, cleaned) = MarpDirectiveParser.Parse(chunk, defaultStyle);
            var slide = new Slide { Style = style };
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
                    elements.Add(new HeadingElement(heading.Level, ExtractInlineText(heading.Inline).Trim()));
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
            }
        }

        return elements;
    }

    private static void AppendParagraphElements(ParagraphBlock paragraph, ICollection<ISlideElement> elements)
    {
        var images = ExtractImages(paragraph.Inline).ToArray();
        var text = ExtractInlineText(paragraph.Inline).Trim();

        if (images.Length > 0 && text.Length == 0)
        {
            foreach (var image in images)
            {
                elements.Add(image);
            }

            return;
        }

        if (text.Length > 0)
        {
            elements.Add(new ParagraphElement(text));
        }

        foreach (var image in images)
        {
            elements.Add(image);
        }
    }

    private static IEnumerable<BulletListItem> FlattenListItems(ListBlock list, int depth = 0)
    {
        foreach (var item in list.OfType<ListItemBlock>())
        {
            var fragments = new List<string>();
            foreach (var child in item)
            {
                switch (child)
                {
                    case ParagraphBlock paragraph:
                        fragments.Add(ExtractInlineText(paragraph.Inline).Trim());
                        break;
                    case FencedCodeBlock code:
                        fragments.Add(code.Lines.ToString() ?? string.Empty);
                        break;
                    case ListBlock nested:
                        foreach (var nestedItem in FlattenListItems(nested, depth + 1))
                        {
                            yield return nestedItem;
                        }
                        break;
                }
            }

            if (fragments.Count > 0)
            {
                yield return new BulletListItem(string.Join(" ", fragments.Where(fragment => fragment.Length > 0)), depth);
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

            var cells = new List<string>();
            foreach (var cellObj in row)
            {
                if (cellObj is not TableCell cell)
                {
                    continue;
                }

                cells.Add(string.Join(" ", cell.Select(block => block switch
                {
                    ParagraphBlock paragraph => ExtractInlineText(paragraph.Inline).Trim(),
                    _ => block.ToString() ?? string.Empty,
                }).Where(text => !string.IsNullOrWhiteSpace(text))));
            }

            rows.Add(new TableRowModel(cells, row.IsHeader));
        }

        return new TableElement(rows);
    }

    private static IEnumerable<ImageElement> ExtractImages(ContainerInline? inline)
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
                    var altText = ExtractInlineText(link).Trim();
                    yield return new ImageElement(link.Url ?? string.Empty, altText);
                    break;
                case ContainerInline nested:
                    foreach (var nestedImage in ExtractImages(nested))
                    {
                        yield return nestedImage;
                    }
                    break;
            }
        }
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
