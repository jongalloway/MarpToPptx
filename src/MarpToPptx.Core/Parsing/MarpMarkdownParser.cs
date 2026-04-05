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

    // Matches the opening or self-closing <audio> tag and captures the src attribute value.
    private static readonly Regex AudioTagRegex = new(
        @"<audio\b[^>]*?\bsrc\s*=\s*(?:""([^""]*)""|'([^']*)'|(\S+?)(?:\s|/?>))",
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

        if (frontMatter.TryGetValue("layout", out var defaultContentLayout))
        {
            defaultContentLayout = defaultContentLayout?.Trim();
            if (!string.IsNullOrWhiteSpace(defaultContentLayout))
            {
                deck.DefaultContentLayout = defaultContentLayout;
            }
        }

        var defaultStyle = new SlideStyle();
        foreach (var pair in frontMatter)
        {
            switch (pair.Key.ToLowerInvariant())
            {
                case "theme":
                case "paginate":
                case "class":
                case "backgroundimage":
                case "backgroundsize":
                case "backgroundposition":
                case "backgroundcolor":
                case "color":
                case "header":
                case "footer":
                case "transition":
                    defaultStyle = MarpDirectiveParser.ApplyDirective(defaultStyle, pair.Key, pair.Value);
                    break;
            }
        }

        // lang: set deck language from front-matter.
        if (frontMatter.TryGetValue("lang", out var lang))
        {
            lang = lang?.Trim();
            if (!string.IsNullOrWhiteSpace(lang))
            {
                deck.Language = lang;
            }
        }

        // diagram-theme: set deck-level preferred DiagramForge theme from front-matter.
        if (frontMatter.TryGetValue("diagram-theme", out var diagramTheme))
        {
            diagramTheme = diagramTheme?.Trim();
            if (!string.IsNullOrWhiteSpace(diagramTheme))
            {
                deck.DiagramTheme = diagramTheme;
            }
        }

        // style: merge inline CSS from front-matter with the external theme CSS.
        string? mergedCss = themeCss;
        if (frontMatter.TryGetValue("style", out var inlineStyle) && !string.IsNullOrWhiteSpace(inlineStyle))
        {
            mergedCss = string.IsNullOrWhiteSpace(mergedCss)
                ? inlineStyle
                : mergedCss + "\n" + inlineStyle;
        }

        deck.Theme = MarpThemeParser.Parse(mergedCss, themeName: defaultStyle.ThemeName ?? "default");

        // headingDivider: parse from front-matter.
        int? headingDivider = null;
        if (frontMatter.TryGetValue("headingDivider", out var hdValue) &&
            int.TryParse(hdValue, out var hdLevel) &&
            hdLevel is >= 1 and <= 6)
        {
            headingDivider = hdLevel;
        }

        // Track carry-forward style: starts from front-matter defaults,
        // updated by local directives on each slide, unaffected by spot directives.
        var carryForwardStyle = defaultStyle;
        foreach (var chunk in SlideTokenizer.SplitSlides(body, headingDivider))
        {
            var (effectiveStyle, newCarryForward, cleaned, notes) = MarpDirectiveParser.Parse(chunk, carryForwardStyle);
            carryForwardStyle = newCarryForward;

            var allElements = ParseElements(cleaned);

            // Promote any ![bg](url) images to slide background image.
            // Only exact "bg" alt text (case-insensitive) is recognized in this slice.
            // A directive-specified backgroundImage always takes precedence: if a directive
            // (including an empty-value clear) has set BackgroundImage, bg syntax is ignored.
            var bgImages = allElements
                .OfType<ImageElement>()
                .Where(img => IsBgAltText(img.AltText))
                .ToList();

            if (bgImages.Count > 0 && effectiveStyle.BackgroundImage is null)
            {
                effectiveStyle = MarpDirectiveParser.ApplyDirective(effectiveStyle, "backgroundimage", bgImages[0].Source);
            }

            var slide = new Slide { Style = effectiveStyle, Notes = notes, NoteSpans = ParseNoteSpans(notes) };
            foreach (var element in allElements)
            {
                if (element is ImageElement img && IsBgAltText(img.AltText))
                {
                    continue;
                }

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
                case FencedCodeBlock fencedCode when string.Equals(fencedCode.Info, "mermaid", StringComparison.OrdinalIgnoreCase):
                    elements.Add(new MermaidDiagramElement(NormalizeDiagramSource(fencedCode.Lines.ToString() ?? string.Empty)));
                    break;
                case FencedCodeBlock fencedCode when string.Equals(fencedCode.Info, "diagram", StringComparison.OrdinalIgnoreCase):
                    elements.Add(new DiagramElement(NormalizeDiagramSource(fencedCode.Lines.ToString() ?? string.Empty)));
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
                case QuoteBlock quoteBlock:
                    elements.Add(ParseBlockquote(quoteBlock));
                    break;
                case HtmlBlock htmlBlock:
                    AppendHtmlBlockElements(htmlBlock, elements);
                    break;
            }
        }

        return elements;
    }

    private BlockquoteElement ParseBlockquote(QuoteBlock quoteBlock)
    {
        var spans = new List<InlineSpan>();
        var firstParagraph = true;
        foreach (var inner in quoteBlock)
        {
            if (inner is ParagraphBlock para && para.Inline is not null)
            {
                if (!firstParagraph)
                {
                    spans.Add(new InlineSpan("\n"));
                }

                spans.AddRange(ExtractInlineSpans(para.Inline));
                firstParagraph = false;
            }
        }

        return new BlockquoteElement(TrimSpans(spans));
    }

    private static string NormalizeDiagramSource(string source)
    {
        var lines = source.Replace("\r\n", "\n").Split('\n');

        var start = 0;
        while (start < lines.Length && string.IsNullOrWhiteSpace(lines[start]))
        {
            start++;
        }

        var end = lines.Length - 1;
        while (end >= start && string.IsNullOrWhiteSpace(lines[end]))
        {
            end--;
        }

        if (start > end)
        {
            return string.Empty;
        }

        var trimmed = lines[start..(end + 1)];

        var minIndent = trimmed
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => line.TakeWhile(char.IsWhiteSpace).Count())
            .DefaultIfEmpty(0)
            .Min();

        if (minIndent <= 0)
        {
            return string.Join('\n', trimmed);
        }

        return string.Join('\n', trimmed.Select(line =>
            string.IsNullOrWhiteSpace(line)
                ? string.Empty
                : line.Length >= minIndent
                    ? line[minIndent..]
                    : line.TrimStart()));
    }

    private IReadOnlyList<InlineSpan> ParseNoteSpans(string? notes)
    {
        if (string.IsNullOrWhiteSpace(notes))
        {
            return [];
        }

        var document = Markdown.Parse(notes, _pipeline);
        var spans = new List<InlineSpan>();
        foreach (var block in document)
        {
            var blockSpans = block switch
            {
                HeadingBlock heading => TrimSpans(ExtractInlineSpans(heading.Inline)),
                ParagraphBlock paragraph => TrimSpans(ExtractInlineSpans(paragraph.Inline)),
                ListBlock list => FlattenNoteListSpans(list),
                FencedCodeBlock fencedCode => FlattenCodeBlockSpans(fencedCode),
                CodeBlock codeBlock => FlattenCodeBlockSpans(codeBlock),
                _ => [],
            };

            if (blockSpans.Count == 0)
            {
                continue;
            }

            if (spans.Count > 0)
            {
                spans.Add(new InlineSpan("\n"));
            }

            spans.AddRange(blockSpans);
        }

        return spans.Count > 0 ? spans : CreateLiteralNoteSpans(notes);
    }

    private static IReadOnlyList<InlineSpan> FlattenNoteListSpans(ListBlock list)
    {
        var spans = new List<InlineSpan>();
        var items = FlattenListItems(list).ToArray();
        for (var index = 0; index < items.Length; index++)
        {
            if (index > 0)
            {
                spans.Add(new InlineSpan("\n"));
            }

            var item = items[index];
            var indent = item.Depth > 0 ? new string(' ', item.Depth * 2) : string.Empty;
            var prefix = list.IsOrdered ? $"{indent}{index + 1}. " : $"{indent}• ";
            spans.Add(new InlineSpan(prefix));
            spans.AddRange(item.Spans);
        }

        return spans;
    }

    private static IReadOnlyList<InlineSpan> FlattenCodeBlockSpans(CodeBlock codeBlock)
    {
        var text = codeBlock.Lines.ToString() ?? string.Empty;
        if (text.Length == 0)
        {
            return [];
        }

        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n', StringSplitOptions.None);
        var spans = new List<InlineSpan>();
        for (var index = 0; index < lines.Length; index++)
        {
            if (index > 0)
            {
                spans.Add(new InlineSpan("\n"));
            }

            if (lines[index].Length > 0)
            {
                spans.Add(new InlineSpan(lines[index], Code: true));
            }
        }

        return spans;
    }

    private static IReadOnlyList<InlineSpan> CreateLiteralNoteSpans(string notes)
    {
        var spans = new List<InlineSpan>();
        var lines = notes.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n', StringSplitOptions.None);
        for (var index = 0; index < lines.Length; index++)
        {
            if (index > 0)
            {
                spans.Add(new InlineSpan("\n"));
            }

            if (lines[index].Length > 0)
            {
                spans.Add(new InlineSpan(lines[index]));
            }
        }

        return spans;
    }

    private static void AppendHtmlBlockElements(HtmlBlock htmlBlock, ICollection<ISlideElement> elements)
    {
        var html = htmlBlock.Lines.ToString();

        var videoMatches = VideoTagRegex.Matches(html);
        foreach (Match match in videoMatches)
        {
            var src = match.Groups[1].Success ? match.Groups[1].Value
                : match.Groups[2].Success ? match.Groups[2].Value
                : match.Groups[3].Value;
            elements.Add(new VideoElement(src, string.Empty));
        }

        var audioMatches = AudioTagRegex.Matches(html);
        foreach (Match match in audioMatches)
        {
            var src = match.Groups[1].Success ? match.Groups[1].Value
                : match.Groups[2].Success ? match.Groups[2].Value
                : match.Groups[3].Value;
            elements.Add(new AudioElement(src, string.Empty));
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
                        if (spanFragments.Count > 0 && childSpans.Count > 0)
                        {
                            spanFragments.Add(new InlineSpan(" "));
                        }

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
                        if (cellSpans.Count > 0 && spans.Count > 0)
                        {
                            cellSpans.Add(new InlineSpan(" "));
                        }

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
                    var caption = string.IsNullOrWhiteSpace(link.Title) ? null : link.Title.Trim();
                    if (url.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase))
                    {
                        yield return new VideoElement(url, altText);
                    }
                    else if (url.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) ||
                             url.EndsWith(".wav", StringComparison.OrdinalIgnoreCase) ||
                             url.EndsWith(".m4a", StringComparison.OrdinalIgnoreCase))
                    {
                        yield return new AudioElement(url, altText);
                    }
                    else
                    {
                        // Skip Marpit sizing parsing for alt texts that start with the "bg" keyword
                        // (e.g. "bg 50%", "bg left"). Background-image sizing is handled separately
                        // and is not yet implemented; stripping tokens from those alt texts would
                        // incorrectly promote "bg 50%" to a background image via the IsBgAltText check.
                        var trimmedAlt = altText.Trim();
                        var isBgAlt = string.Equals(trimmedAlt, "bg", StringComparison.OrdinalIgnoreCase)
                            || trimmedAlt.StartsWith("bg ", StringComparison.OrdinalIgnoreCase);

                        if (!isBgAlt)
                        {
                            var (explicitWidth, explicitHeight, sizePercent, cleanAltText) = MarpitImageSizingParser.Parse(altText);
                            yield return new ImageElement(url, cleanAltText, caption, explicitWidth, explicitHeight, sizePercent);
                        }
                        else
                        {
                            yield return new ImageElement(url, altText, caption);
                        }
                    }

                    break;
                case HtmlInline htmlInline:
                    var videoMatch = VideoTagRegex.Match(htmlInline.Tag);
                    if (videoMatch.Success)
                    {
                        var src = videoMatch.Groups[1].Success ? videoMatch.Groups[1].Value
                            : videoMatch.Groups[2].Success ? videoMatch.Groups[2].Value
                            : videoMatch.Groups[3].Value;
                        yield return new VideoElement(src, string.Empty);
                    }

                    var audioMatch = AudioTagRegex.Match(htmlInline.Tag);
                    if (audioMatch.Success)
                    {
                        var src = audioMatch.Groups[1].Success ? audioMatch.Groups[1].Value
                            : audioMatch.Groups[2].Success ? audioMatch.Groups[2].Value
                            : audioMatch.Groups[3].Value;
                        yield return new AudioElement(src, string.Empty);
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
                    spans.Add(new InlineSpan(codeInline.Content, bold, italic, Code: true, Strikethrough: strikethrough, HyperlinkUrl: hyperlinkUrl));
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

                case LinkInline link when link.IsImage:
                    // Alt text is accessibility metadata stored on the image shape, not visible slide text.
                    break;
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

    /// <summary>
    /// Returns <see langword="true"/> when <paramref name="altText"/> is exactly <c>bg</c>
    /// (case-insensitive), indicating a Marpit background image marker.
    /// </summary>
    private static bool IsBgAltText(string altText) =>
        string.Equals(altText.Trim(), "bg", StringComparison.OrdinalIgnoreCase);
}
