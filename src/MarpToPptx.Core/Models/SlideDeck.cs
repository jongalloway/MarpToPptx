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
    public string? ThemeName { get; init; }

    public bool? Paginate { get; init; }

    public string? Layout { get; init; }

    public string? ClassName { get; init; }

    public string? BackgroundImage { get; init; }

    public string? BackgroundSize { get; init; }

    public string? BackgroundColor { get; init; }

    public string? Header { get; init; }

    public string? Footer { get; init; }

    public Dictionary<string, string> Directives { get; } = new(StringComparer.OrdinalIgnoreCase);
}

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

public sealed record ImageElement(string Source, string AltText) : ISlideElement;

public sealed record VideoElement(string Source, string AltText) : ISlideElement;

public sealed record AudioElement(string Source, string AltText) : ISlideElement;

public sealed record CodeBlockElement(string Language, string Code) : ISlideElement;

/// <summary>
/// A Mermaid diagram fenced code block that should be rendered to SVG and placed on the slide.
/// </summary>
public sealed record MermaidDiagramElement(string Source) : ISlideElement;

public sealed record TableElement(IReadOnlyList<TableRowModel> Rows, IReadOnlyList<TableColumnAlignment?> ColumnAlignments) : ISlideElement
{
    public TableElement(IReadOnlyList<TableRowModel> Rows) : this(Rows, []) { }
}

public sealed record TableRowModel(IReadOnlyList<IReadOnlyList<InlineSpan>> Cells, bool IsHeader = false);

public enum TableColumnAlignment { Left, Center, Right }
