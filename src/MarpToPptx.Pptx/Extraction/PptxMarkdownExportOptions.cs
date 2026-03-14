namespace MarpToPptx.Pptx.Extraction;

public sealed class PptxMarkdownExportOptions
{
    public string? AssetsDirectory { get; init; }

    public string AssetPathPrefix { get; init; } = "assets";

    public bool IncludeNotes { get; init; } = true;

    public bool FilterNoise { get; init; } = true;
}