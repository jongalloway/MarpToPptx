using DocumentFormat.OpenXml.Packaging;

namespace MarpToPptx.Pptx.Extraction;

public sealed partial class PptxMarkdownExporter
{
    private const long SlideWidthEmu = 12192000L;
    private const long SlideHeightEmu = 6858000L;
    private const long BulletGroupMaxHorizontalDeltaEmu = 700000L;
    private const long BulletGroupMaxVerticalGapEmu = 1200000L;
    private const int MinimumInferredBulletCount = 3;
    private static readonly string[] CodeFontNames = ["consolas", "cascadia code", "courier new", "fira code", "jetbrains mono", "source code pro", "menlo", "monaco"];

    public void Export(string pptxPath, string outputMarkdownPath, PptxMarkdownExportOptions? options = null)
    {
        var markdown = ExportToMarkdown(pptxPath, options);
        var outputDirectory = Path.GetDirectoryName(outputMarkdownPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        File.WriteAllText(outputMarkdownPath, markdown);
    }

    public string ExportToMarkdown(string pptxPath, PptxMarkdownExportOptions? options = null)
    {
        options ??= new PptxMarkdownExportOptions();
        if (!string.IsNullOrWhiteSpace(options.AssetsDirectory))
        {
            Directory.CreateDirectory(options.AssetsDirectory);
        }

        using var document = PresentationDocument.Open(pptxPath, false);
        var presentationPart = document.PresentationPart ?? throw new InvalidOperationException("The presentation is missing a presentation part.");
        var slides = GetSlidesInPresentationOrder(presentationPart);
        var signatureCache = new Dictionary<string, string?>(StringComparer.Ordinal);
        var imageUseCounts = options.FilterNoise ? CountImageUses(slides, signatureCache) : new Dictionary<string, int>(StringComparer.Ordinal);
        var markdown = new List<string>();
        var usedAssetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < slides.Count; index++)
        {
            if (index > 0)
            {
                markdown.Add(string.Empty);
                markdown.Add("---");
                markdown.Add(string.Empty);
            }

            var slideLines = ExtractSlide(slides[index], index + 1, options, usedAssetNames, imageUseCounts, signatureCache);
            markdown.AddRange(slideLines);
        }

        return string.Join(Environment.NewLine, markdown).TrimEnd() + Environment.NewLine;
    }

    private static List<string> ExtractSlide(SlidePart slidePart, int slideNumber, PptxMarkdownExportOptions options, HashSet<string> usedAssetNames, IReadOnlyDictionary<string, int> imageUseCounts, Dictionary<string, string?> signatureCache)
    {
        var blocks = new List<MarkdownBlock>();
        var textShapes = GetTextShapes(slidePart);
        var titleShape = textShapes.FirstOrDefault(shape => shape.IsTitle) ?? textShapes.FirstOrDefault();

        blocks.AddRange(GetTextBlocks(textShapes, titleShape));

        blocks.AddRange(GetImageBlocks(slidePart, options, usedAssetNames, imageUseCounts, signatureCache));
        blocks.AddRange(GetTableBlocks(slidePart));

        var markdown = new List<string>();
        foreach (var block in blocks.OrderBy(block => block.Y).ThenBy(block => block.X))
        {
            if (markdown.Count > 0 && markdown[^1].Length > 0)
            {
                markdown.Add(string.Empty);
            }

            markdown.AddRange(block.Lines);
        }

        if (options.IncludeNotes)
        {
            var notes = GetNotesMarkdown(slidePart);
            if (notes.Count > 0)
            {
                if (markdown.Count > 0 && markdown[^1].Length > 0)
                {
                    markdown.Add(string.Empty);
                }

                markdown.AddRange(notes);
            }
        }

        if (markdown.Count == 0)
        {
            markdown.Add($"<!-- Slide {slideNumber} had no recoverable content -->");
        }

        return markdown;
    }

    private sealed record MarkdownBlock(long Y, long X, IReadOnlyList<string> Lines);

    private sealed record TextShapeInfo(string Name, bool IsTitle, long Y, long X, IReadOnlyList<ParagraphInfo> Paragraphs);

    private sealed record ParagraphInfo(string Text, bool IsBullet, bool IsOrdered, int Level, bool UsesCodeFont);

    private enum TableAlignment
    {
        Left,
        Center,
        Right,
    }
}
