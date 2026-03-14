using DocumentFormat.OpenXml.Packaging;
using System.Security.Cryptography;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Extraction;

public sealed partial class PptxMarkdownExporter
{
    private static IReadOnlyList<MarkdownBlock> GetImageBlocks(SlidePart slidePart, PptxMarkdownExportOptions options, HashSet<string> usedAssetNames, IReadOnlyDictionary<string, int> imageUseCounts, Dictionary<string, string?> signatureCache)
    {
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree is null || string.IsNullOrWhiteSpace(options.AssetsDirectory))
        {
            return [];
        }

        var blocks = new List<MarkdownBlock>();
        foreach (var picture in shapeTree.Elements<P.Picture>())
        {
            if (picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<A.VideoFromFile>() is not null ||
                picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<A.AudioFromFile>() is not null)
            {
                continue;
            }

            var relationshipId = GetImageRelationshipId(picture.BlipFill?.Blip);
            if (string.IsNullOrWhiteSpace(relationshipId) || !slidePart.TryGetPartById(relationshipId, out var part) || part is not ImagePart imagePart)
            {
                continue;
            }

            if (options.FilterNoise && IsLikelyDecorativeImage(picture, imagePart, imageUseCounts, signatureCache))
            {
                continue;
            }

            var drawingProperties = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
            var altText = drawingProperties?.Description?.Value;
            var preferredName = drawingProperties?.Name?.Value;
            var assetPath = ExportImagePart(imagePart, preferredName, options, usedAssetNames);
            if (assetPath is null)
            {
                continue;
            }

            TryGetPictureBounds(picture, out var x, out var y, out _, out _);
            var alt = string.IsNullOrWhiteSpace(altText) ? Path.GetFileNameWithoutExtension(assetPath) : altText;
            blocks.Add(new MarkdownBlock(y, x, [$"![{alt}]({assetPath})"]));
        }

        return blocks;
    }

    private static string? ExportImagePart(ImagePart imagePart, string? preferredName, PptxMarkdownExportOptions options, HashSet<string> usedAssetNames)
    {
        var sourceName = string.IsNullOrWhiteSpace(preferredName) ? Path.GetFileName(imagePart.Uri.ToString()) : Path.GetFileName(preferredName);
        var extension = Path.GetExtension(sourceName);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = GetExtensionForContentType(imagePart.ContentType);
        }

        var fileName = EnsureUniqueFileName(SanitizeFileName(Path.GetFileNameWithoutExtension(sourceName)), extension, usedAssetNames);
        var outputPath = Path.Combine(options.AssetsDirectory!, fileName);
        using var sourceStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var fileStream = File.Create(outputPath);
        sourceStream.CopyTo(fileStream);

        var prefix = options.AssetPathPrefix.Trim().TrimEnd('/').Replace('\\', '/');
        return string.IsNullOrWhiteSpace(prefix) ? fileName : $"{prefix}/{fileName}";
    }

    private static string EnsureUniqueFileName(string baseName, string extension, HashSet<string> usedAssetNames)
    {
        var normalizedBaseName = string.IsNullOrWhiteSpace(baseName) ? "asset" : baseName;
        var normalizedExtension = extension.StartsWith('.') ? extension : "." + extension;
        var candidate = normalizedBaseName + normalizedExtension;
        var suffix = 1;
        while (!usedAssetNames.Add(candidate))
        {
            candidate = $"{normalizedBaseName}-{suffix}{normalizedExtension}";
            suffix++;
        }

        return candidate;
    }

    private static string SanitizeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = new string(fileName.Select(ch => invalid.Contains(ch) ? '-' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(sanitized) ? "asset" : sanitized;
    }

    private static string GetExtensionForContentType(string contentType)
        => contentType.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tif",
            "image/svg+xml" => ".svg",
            _ => ".bin",
        };

    private static string? GetImageRelationshipId(A.Blip? blip)
    {
        if (blip?.Embed?.Value is { Length: > 0 } embed)
        {
            return embed;
        }

        return blip?
            .Descendants<DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip>()
            .Select(svgBlip => svgBlip.Embed?.Value)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    private static IReadOnlyList<MarkdownBlock> GetTableBlocks(SlidePart slidePart)
    {
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
        {
            return [];
        }

        var blocks = new List<MarkdownBlock>();
        foreach (var graphicFrame in shapeTree.Elements<P.GraphicFrame>())
        {
            var table = graphicFrame.Descendants<A.Table>().FirstOrDefault();
            if (table is null)
            {
                continue;
            }

            var rows = table.Elements<A.TableRow>()
                .Select(row => row.Elements<A.TableCell>().Select(GetTableCellText).ToArray())
                .Where(row => row.Length > 0)
                .ToArray();
            if (rows.Length == 0)
            {
                continue;
            }

            var columnCount = rows.Max(row => row.Length);
            var header = NormalizeRowLength(rows[0], columnCount);
            var alignments = Enumerable.Range(0, columnCount).Select(index => GetColumnAlignment(table, index)).ToArray();
            var lines = new List<string>
            {
                "| " + string.Join(" | ", header.Select(EscapeTableCell)) + " |",
                "| " + string.Join(" | ", alignments.Select(GetAlignmentMarker)) + " |",
            };

            foreach (var row in rows.Skip(1))
            {
                lines.Add("| " + string.Join(" | ", NormalizeRowLength(row, columnCount).Select(EscapeTableCell)) + " |");
            }

            TryGetGraphicFrameBounds(graphicFrame, out var x, out var y);
            blocks.Add(new MarkdownBlock(y, x, lines));
        }

        return blocks;
    }

    private static string GetTableCellText(A.TableCell cell)
        => string.Join("<br>", cell.TextBody?.Elements<A.Paragraph>().Select(paragraph => string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text)).Trim()) ?? []);

    private static string[] NormalizeRowLength(string[] row, int columnCount)
        => Enumerable.Range(0, columnCount).Select(index => index < row.Length ? row[index] : string.Empty).ToArray();

    private static string EscapeTableCell(string value)
        => value.Replace("|", "\\|", StringComparison.Ordinal);

    private static string GetAlignmentMarker(TableAlignment alignment)
        => alignment switch
        {
            TableAlignment.Center => ":---:",
            TableAlignment.Right => "---:",
            _ => "---",
        };

    private static TableAlignment GetColumnAlignment(A.Table table, int columnIndex)
    {
        foreach (var row in table.Elements<A.TableRow>())
        {
            var alignment = row.Elements<A.TableCell>().ElementAtOrDefault(columnIndex)?
                .TextBody?
                .Elements<A.Paragraph>()
                .FirstOrDefault()?
                .ParagraphProperties?
                .Alignment?
                .Value;
            if (alignment is null)
            {
                continue;
            }

            if (alignment == A.TextAlignmentTypeValues.Center)
            {
                return TableAlignment.Center;
            }

            if (alignment == A.TextAlignmentTypeValues.Right)
            {
                return TableAlignment.Right;
            }

            return TableAlignment.Left;
        }

        return TableAlignment.Left;
    }

    private static IReadOnlyList<string> GetNotesMarkdown(SlidePart slidePart)
    {
        var notesSlide = slidePart.NotesSlidePart?.NotesSlide;
        if (notesSlide?.CommonSlideData?.ShapeTree is null)
        {
            return [];
        }

        var noteLines = notesSlide.CommonSlideData.ShapeTree
            .Elements<P.Shape>()
            .Where(shape => shape.TextBody is not null)
            .SelectMany(shape => shape.TextBody!.Elements<A.Paragraph>())
            .Select(paragraph => BuildParagraphText(paragraph, slidePart))
            .Where(text => !IsLikelyNoiseText(text, y: SlideHeightEmu, isNote: true))
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray();
        if (noteLines.Length == 0)
        {
            return [];
        }

        var markdown = new List<string> { "<!--" };
        markdown.AddRange(noteLines);
        markdown.Add("-->");
        return markdown;
    }

    private static IReadOnlyList<SlidePart> GetSlidesInPresentationOrder(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>()
            .Where(slideId => !string.IsNullOrWhiteSpace(slideId.RelationshipId))
            .ToArray() ?? [];
        return slideIds.Select(slideId => (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!)).ToArray();
    }

    private static bool TryGetShapeBounds(P.Shape shape, out long x, out long y)
    {
        x = 0L;
        y = 0L;
        var transform = shape.ShapeProperties?.Transform2D;
        if (transform?.Offset is null)
        {
            return false;
        }

        x = transform.Offset.X?.Value ?? 0L;
        y = transform.Offset.Y?.Value ?? 0L;
        return true;
    }

    private static bool TryGetPictureBounds(P.Picture picture, out long x, out long y, out long cx, out long cy)
    {
        x = 0L;
        y = 0L;
        cx = 0L;
        cy = 0L;
        var transform = picture.ShapeProperties?.Transform2D;
        if (transform?.Offset is null || transform.Extents is null)
        {
            return false;
        }

        x = transform.Offset.X?.Value ?? 0L;
        y = transform.Offset.Y?.Value ?? 0L;
        cx = transform.Extents.Cx?.Value ?? 0L;
        cy = transform.Extents.Cy?.Value ?? 0L;
        return true;
    }

    private static bool TryGetGraphicFrameBounds(P.GraphicFrame graphicFrame, out long x, out long y)
    {
        x = 0L;
        y = 0L;
        if (graphicFrame.Transform?.Offset is null)
        {
            return false;
        }

        x = graphicFrame.Transform.Offset.X?.Value ?? 0L;
        y = graphicFrame.Transform.Offset.Y?.Value ?? 0L;
        return true;
    }

    private static Dictionary<string, int> CountImageUses(IReadOnlyList<SlidePart> slides, Dictionary<string, string?> signatureCache)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var slidePart in slides)
        {
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (shapeTree is null)
            {
                continue;
            }

            foreach (var picture in shapeTree.Elements<P.Picture>())
            {
                if (picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<A.VideoFromFile>() is not null ||
                    picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<A.AudioFromFile>() is not null)
                {
                    continue;
                }

                var relationshipId = GetImageRelationshipId(picture.BlipFill?.Blip);
                if (string.IsNullOrWhiteSpace(relationshipId) || !slidePart.TryGetPartById(relationshipId, out var part) || part is not ImagePart imagePart)
                {
                    continue;
                }

                var signature = GetOrComputeImageSignature(imagePart, signatureCache);
                if (signature is null)
                {
                    continue;
                }

                counts[signature] = counts.TryGetValue(signature, out var count) ? count + 1 : 1;
            }
        }

        return counts;
    }

    private static bool IsLikelyDecorativeImage(P.Picture picture, ImagePart imagePart, IReadOnlyDictionary<string, int> imageUseCounts, Dictionary<string, string?> signatureCache)
    {
        var altText = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
        if (!string.IsNullOrWhiteSpace(altText) &&
            altText.Contains("AI-generated content may be incorrect", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!TryGetPictureBounds(picture, out var x, out var y, out var cx, out var cy))
        {
            return false;
        }

        var signature = GetOrComputeImageSignature(imagePart, signatureCache);
        var repeated = signature is not null && imageUseCounts.TryGetValue(signature, out var count) && count > 1;
        if (!repeated)
        {
            return false;
        }

        var relativeArea = (double)(cx * cy) / (SlideWidthEmu * (double)SlideHeightEmu);
        var nearCorner = x < SlideWidthEmu * 0.2 && y < SlideHeightEmu * 0.25;
        return relativeArea < 0.05 && nearCorner;
    }

    private static string? GetOrComputeImageSignature(ImagePart imagePart, Dictionary<string, string?> signatureCache)
    {
        var key = imagePart.Uri.ToString();
        if (!signatureCache.TryGetValue(key, out var signature))
        {
            signatureCache[key] = signature = ComputeImageSignature(imagePart);
        }

        return signature;
    }

    private static string? ComputeImageSignature(ImagePart imagePart)
    {
        using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static bool IsLikelyNoiseText(string text, long y, bool isNote)
    {
        var normalized = string.Join(' ', text.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries)).Trim();
        if (normalized.Length == 0)
        {
            return true;
        }

        if (normalized.StartsWith("© Microsoft Corporation. All rights reserved.", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("MICROSOFT MAKES NO WARRANTIES", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var nearBottom = y >= SlideHeightEmu * 0.82;
        if ((nearBottom || isNote) && int.TryParse(normalized, out _))
        {
            return true;
        }

        if (isNote && DateTime.TryParse(normalized, out _))
        {
            return true;
        }

        return false;
    }
}