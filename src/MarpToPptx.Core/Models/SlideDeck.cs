using MarpToPptx.Core.Themes;

namespace MarpToPptx.Core.Models;

public sealed class SlideDeck
{
    public string? SourcePath { get; init; }

    public Dictionary<string, string> FrontMatter { get; } = new(StringComparer.OrdinalIgnoreCase);

    public List<Slide> Slides { get; } = [];

    public ThemeDefinition Theme { get; set; } = ThemeDefinition.Default;
}

public sealed class Slide
{
    public List<ISlideElement> Elements { get; } = [];

    public SlideStyle Style { get; init; } = new();
}

public sealed class SlideStyle
{
    public string? ThemeName { get; init; }

    public bool? Paginate { get; init; }

    public string? ClassName { get; init; }

    public string? BackgroundImage { get; init; }

    public string? BackgroundColor { get; init; }

    public Dictionary<string, string> Directives { get; } = new(StringComparer.OrdinalIgnoreCase);
}

public interface ISlideElement
{
}

public sealed record HeadingElement(int Level, string Text) : ISlideElement;

public sealed record ParagraphElement(string Text) : ISlideElement;

public sealed record BulletListElement(IReadOnlyList<BulletListItem> Items, bool Ordered) : ISlideElement;

public sealed record BulletListItem(string Text, int Depth = 0);

public sealed record ImageElement(string Source, string AltText) : ISlideElement;

public sealed record CodeBlockElement(string Language, string Code) : ISlideElement;

public sealed record TableElement(IReadOnlyList<TableRowModel> Rows, IReadOnlyList<TableColumnAlignment?> ColumnAlignments) : ISlideElement
{
    public TableElement(IReadOnlyList<TableRowModel> Rows) : this(Rows, []) { }
}

public sealed record TableRowModel(IReadOnlyList<string> Cells, bool IsHeader = false);

public enum TableColumnAlignment { Left, Center, Right }
