using MarpToPptx.Core.Themes;

namespace MarpToPptx.Core.Models;

public sealed class SlideDeck
{
    public string? SourcePath { get; init; }

    public Dictionary<string, string> FrontMatter { get; } = new(StringComparer.OrdinalIgnoreCase);

    public List<Slide> Slides { get; } = [];

    /// <summary>
    /// Deck-level default template layout name for content slides when rendering with a PPTX template.
    /// Per-slide <c>layout</c> or <c>_layout</c> directives override this value.
    /// </summary>
    public string? DefaultContentLayout { get; set; }

    public ThemeDefinition Theme { get; set; } = ThemeDefinition.Default;

    /// <summary>
    /// BCP-47 language tag from the <c>lang</c> global directive.
    /// When set, this is used for document metadata and run-level language properties.
    /// </summary>
    public string? Language { get; set; }

    /// <summary>
    /// Deck-level preferred DiagramForge theme name from the <c>diagram-theme</c> global directive.
    /// Applied to all Mermaid and <c>diagram</c> fences that do not specify their own <c>theme:</c>
    /// in embedded front matter. A fence-level <c>theme:</c> always takes precedence.
    /// </summary>
    public string? DiagramTheme { get; set; }
}

public sealed class Slide
{
    public List<ISlideElement> Elements { get; } = [];

    public SlideStyle Style { get; init; } = new();

    public string? Notes { get; init; }

    public IReadOnlyList<InlineSpan> NoteSpans { get; init; } = [];
}

public sealed class SlideStyle
{
    public string? SlideId { get; init; }

    public string? ThemeName { get; init; }

    public bool? Paginate { get; init; }

    public string? Layout { get; init; }

    public string? ClassName { get; init; }

    public string? BackgroundImage { get; init; }

    public string? BackgroundSize { get; init; }

    public string? BackgroundPosition { get; init; }

    public string? BackgroundColor { get; init; }

    public string? Color { get; init; }

    public string? Header { get; init; }

    public string? Footer { get; init; }

    public SlideTransition? Transition { get; init; }

    /// <summary>
    /// Override body text font size in hundredths of a point (e.g. 2000 = 20pt).
    /// When set, overrides the theme body font size for body text on this slide.
    /// Title placeholder text is not affected.
    /// </summary>
    public int? FontSize { get; init; }

    public Dictionary<string, string> Directives { get; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>
/// Describes the slide transition animation to apply when advancing to this slide in PowerPoint.
/// </summary>
/// <param name="Type">
/// Transition type identifier (e.g. <c>fade</c>, <c>push</c>, <c>wipe</c>, <c>cut</c>, <c>cover</c>, <c>pull</c>, <c>random-bar</c>, <c>morph</c>).
/// Unknown values are preserved in the model but produce no transition element in the rendered PPTX.
/// </param>
/// <param name="Direction">
/// Optional direction: <c>left</c>, <c>right</c>, <c>up</c>, <c>down</c>. Ignored for transitions that do not support direction.
/// </param>
/// <param name="DurationMs">Optional transition duration in milliseconds.</param>
public sealed record SlideTransition(string Type, string? Direction = null, int? DurationMs = null);

public interface ISlideElement
{
}

/// <summary>
/// Represents a run of text with optional inline formatting and an optional hyperlink.
/// A span with <see cref="Text"/> equal to <c>"\n"</c> acts as a paragraph-break marker.
/// </summary>
public sealed record InlineSpan(
    string Text,
    bool Bold = false,
    bool Italic = false,
    bool Code = false,
    bool Strikethrough = false,
    string? HyperlinkUrl = null);

public sealed record HeadingElement(int Level, IReadOnlyList<InlineSpan> Spans) : ISlideElement
{
    public string Text => string.Concat(Spans.Select(s => s.Text));
}

public sealed record ParagraphElement(IReadOnlyList<InlineSpan> Spans) : ISlideElement
{
    public string Text => string.Concat(Spans.Select(s => s.Text));
}

public sealed record BulletListElement(IReadOnlyList<BulletListItem> Items, bool Ordered) : ISlideElement;

public sealed record BulletListItem(IReadOnlyList<InlineSpan> Spans, int Depth = 0)
{
    public string Text => string.Concat(Spans.Select(s => s.Text));
}

/// <summary>
/// Represents an embedded image on a slide.
/// </summary>
/// <param name="Source">File path or URL of the image.</param>
/// <param name="AltText">Accessibility alt text stored on the image shape; not rendered as visible slide text.</param>
/// <param name="Caption">
/// Optional visible caption rendered below the image.
/// Specify via the Markdown image title attribute: <c>![alt](url "Caption text")</c>.
/// </param>
/// <param name="ExplicitWidth">
/// Explicit image width in layout units (points), parsed from a Marpit <c>w:</c> sizing directive
/// (e.g. <c>![w:200px](img.png)</c>). When set, the renderer scales the image to this width;
/// height is preserved from the image's aspect ratio unless <see cref="ExplicitHeight"/> is also set.
/// </param>
/// <param name="ExplicitHeight">
/// Explicit image height in layout units (points), parsed from a Marpit <c>h:</c> sizing directive
/// (e.g. <c>![h:150px](img.png)</c>). When set, the renderer scales the image to this height;
/// width is preserved from the image's aspect ratio unless <see cref="ExplicitWidth"/> is also set.
/// </param>
/// <param name="SizePercent">
/// Percentage of the slide width to use for the image, parsed from a Marpit percentage sizing
/// directive (e.g. <c>![50%](img.png)</c>). Values are in the range 0–100. The renderer scales the
/// image width to this percentage of the full slide width; height is preserved from aspect ratio.
/// Ignored when <see cref="ExplicitWidth"/> or <see cref="ExplicitHeight"/> is set.
/// </param>
public sealed record ImageElement(
    string Source,
    string AltText,
    string? Caption = null,
    double? ExplicitWidth = null,
    double? ExplicitHeight = null,
    double? SizePercent = null) : ISlideElement
{
    /// <summary>
    /// Backward-compatible constructor matching the original 3-parameter signature.
    /// </summary>
    public ImageElement(string Source, string AltText, string? Caption = null)
        : this(Source, AltText, Caption, null, null, null)
    {
    }
}

public sealed record VideoElement(string Source, string AltText) : ISlideElement;

public sealed record AudioElement(string Source, string AltText) : ISlideElement;

public sealed record CodeBlockElement(string Language, string Code) : ISlideElement;

/// <summary>
/// A Mermaid diagram fenced code block that should be rendered to SVG and placed on the slide.
/// </summary>
public sealed record MermaidDiagramElement(string Source) : ISlideElement;

/// <summary>
/// A DiagramForge conceptual diagram fenced code block that should be rendered to SVG and placed on the slide.
/// </summary>
public sealed record DiagramElement(string Source) : ISlideElement;

public sealed record TableElement(IReadOnlyList<TableRowModel> Rows, IReadOnlyList<TableColumnAlignment?> ColumnAlignments) : ISlideElement
{
    public TableElement(IReadOnlyList<TableRowModel> Rows) : this(Rows, []) { }
}

/// <summary>
/// Represents a Markdown blockquote (<c>&gt; …</c>) on a slide.
/// </summary>
public sealed record BlockquoteElement(IReadOnlyList<InlineSpan> Spans) : ISlideElement
{
    public string Text => string.Concat(Spans.Select(s => s.Text));
}

public sealed record TableRowModel(IReadOnlyList<IReadOnlyList<InlineSpan>> Cells, bool IsHeader = false);

public enum TableColumnAlignment { Left, Center, Right }
