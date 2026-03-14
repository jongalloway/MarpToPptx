using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Extraction;

public sealed partial class PptxMarkdownExporter
{
    private static IReadOnlyList<MarkdownBlock> GetTextBlocks(IReadOnlyList<TextShapeInfo> textShapes, TextShapeInfo? titleShape)
    {
        var blocks = new List<MarkdownBlock>();
        var consumedIndexes = new HashSet<int>();

        if (TryBuildSlideWideLabelGroup(textShapes, titleShape, out var labelGroup, out var groupedIndexes))
        {
            blocks.Add(labelGroup);
            foreach (var index in groupedIndexes)
            {
                consumedIndexes.Add(index);
            }
        }

        for (var index = 0; index < textShapes.Count; index++)
        {
            if (consumedIndexes.Contains(index))
            {
                continue;
            }

            var textShape = textShapes[index];
            if (!ReferenceEquals(textShape, titleShape) &&
                TryBuildGroupedBulletBlock(textShapes, index, out var groupedBlock, out var nextIndex))
            {
                blocks.Add(groupedBlock);
                index = nextIndex - 1;
                continue;
            }

            var lines = BuildTextBlockLines(textShape, ReferenceEquals(textShape, titleShape));
            if (lines.Count > 0)
            {
                blocks.Add(new MarkdownBlock(textShape.Y, textShape.X, lines));
            }
        }

        return blocks;
    }

    private static bool TryBuildSlideWideLabelGroup(IReadOnlyList<TextShapeInfo> textShapes, TextShapeInfo? titleShape, out MarkdownBlock block, out IReadOnlyList<int> groupedIndexes)
    {
        block = default!;
        groupedIndexes = [];

        var candidates = textShapes
            .Select((shape, index) => (shape, index))
            .Where(item => !ReferenceEquals(item.shape, titleShape))
            .Where(item => item.shape.Paragraphs.Count == 1 && IsLikelyLabelParagraph(item.shape.Paragraphs[0]))
            .ToArray();

        if (candidates.Length < 5)
        {
            return false;
        }

        var rows = candidates.Select(item => item.shape.Y / 600000L).Distinct().Count();
        var columns = candidates.Select(item => item.shape.X / 900000L).Distinct().Count();
        var isGrid = rows >= 2 && columns >= 2;
        var isSingleRowStrip = rows == 1 && columns >= 5;
        if (!isGrid && !isSingleRowStrip)
        {
            return false;
        }

        var ordered = candidates
            .OrderBy(item => item.shape.Y)
            .ThenBy(item => item.shape.X)
            .ToArray();

        block = new MarkdownBlock(
            ordered[0].shape.Y,
            ordered[0].shape.X,
            ordered.Select((item, index) => FormatListItem(item.shape.Paragraphs[0], index + 1)).ToArray());
        groupedIndexes = ordered.Select(item => item.index).ToArray();
        return true;
    }

    private static bool TryBuildGroupedBulletBlock(IReadOnlyList<TextShapeInfo> textShapes, int startIndex, out MarkdownBlock block, out int nextIndex)
    {
        block = default!;
        nextIndex = startIndex;

        var firstShape = textShapes[startIndex];
        if (firstShape.IsTitle || firstShape.Paragraphs.Count != 1 || !IsLikelyLabelParagraph(firstShape.Paragraphs[0]))
        {
            return false;
        }

        var groupedShapes = new List<TextShapeInfo> { firstShape };
        var anchorX = firstShape.X;
        var previousY = firstShape.Y;

        for (var index = startIndex + 1; index < textShapes.Count; index++)
        {
            var candidate = textShapes[index];
            if (candidate.IsTitle || candidate.Paragraphs.Count != 1 || !IsLikelyLabelParagraph(candidate.Paragraphs[0]))
            {
                break;
            }

            if (Math.Abs(candidate.X - anchorX) > BulletGroupMaxHorizontalDeltaEmu)
            {
                break;
            }

            var verticalGap = candidate.Y - previousY;
            if (verticalGap <= 0 || verticalGap > BulletGroupMaxVerticalGapEmu)
            {
                break;
            }

            groupedShapes.Add(candidate);
            previousY = candidate.Y;
        }

        if (groupedShapes.Count < MinimumInferredBulletCount)
        {
            return false;
        }

        block = new MarkdownBlock(
            groupedShapes[0].Y,
            groupedShapes[0].X,
            groupedShapes.Select((shape, index) => FormatListItem(shape.Paragraphs[0], index + 1)).ToArray());
        nextIndex = startIndex + groupedShapes.Count;
        return true;
    }

    private static IReadOnlyList<TextShapeInfo> GetTextShapes(SlidePart slidePart)
    {
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
        {
            return [];
        }

        var shapes = new List<TextShapeInfo>();
        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            if (shape.TextBody is null)
            {
                continue;
            }

            var placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();
            var placeholderType = placeholder?.Type?.Value;
            if (placeholderType == P.PlaceholderValues.Footer ||
                placeholderType == P.PlaceholderValues.DateAndTime ||
                placeholderType == P.PlaceholderValues.SlideNumber)
            {
                continue;
            }

            var hasBounds = TryGetShapeBounds(shape, out var x, out var y);
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            var isTitle = placeholderType == P.PlaceholderValues.Title ||
                placeholderType == P.PlaceholderValues.CenteredTitle ||
                name.Contains("title", StringComparison.OrdinalIgnoreCase);
            if (!hasBounds)
            {
                x = 0L;
                y = isTitle ? 0L : 1_000_000L;
            }

            var paragraphs = shape.TextBody.Elements<A.Paragraph>()
                .Select(paragraph => CreateParagraphInfo(paragraph, slidePart))
                .Where(info => !IsLikelyNoiseText(info.Text, y, false))
                .Where(info => !string.IsNullOrWhiteSpace(info.Text))
                .ToArray();
            if (paragraphs.Length == 0)
            {
                continue;
            }

            shapes.Add(new TextShapeInfo(name, isTitle, y, x, paragraphs));
        }

        return shapes.OrderBy(shape => shape.Y).ThenBy(shape => shape.X).ToArray();
    }

    private static ParagraphInfo CreateParagraphInfo(A.Paragraph paragraph, SlidePart slidePart)
    {
        var properties = paragraph.ParagraphProperties;
        var isOrdered = properties?.Elements<A.AutoNumberedBullet>().Any() == true;
        var hasCharacterBullet = properties?.Elements<A.CharacterBullet>().Any() == true;
        var hasLevel = properties?.Level is not null;
        var hasNoBullet = properties?.Elements<A.NoBullet>().Any() == true;
        var isBullet = isOrdered || hasCharacterBullet || (hasLevel && !hasNoBullet);
        var level = properties?.Level?.Value ?? 0;
        return new ParagraphInfo(BuildParagraphText(paragraph, slidePart), isBullet, isOrdered, level, UsesCodeFont(paragraph));
    }

    private static string BuildParagraphText(A.Paragraph paragraph, SlidePart slidePart)
    {
        var pieces = new List<string>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case A.Run run:
                    pieces.Add(BuildRunText(run, slidePart));
                    break;
                case A.Break:
                    pieces.Add("<br>");
                    break;
                case A.Field field:
                    pieces.Add(string.Concat(field.Descendants<A.Text>().Select(text => text.Text)));
                    break;
            }
        }

        return string.Concat(pieces).Trim();
    }

    private static string BuildRunText(A.Run run, SlidePart slidePart)
    {
        var text = run.Text?.Text ?? string.Empty;
        var hyperlinkId = run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value;
        if (string.IsNullOrWhiteSpace(hyperlinkId))
        {
            return text;
        }

        var hyperlink = slidePart.HyperlinkRelationships.FirstOrDefault(relationship => relationship.Id == hyperlinkId);
        return hyperlink?.Uri is null ? text : $"[{text}]({hyperlink.Uri})";
    }

    private static IReadOnlyList<string> BuildTextBlockLines(TextShapeInfo textShape, bool isTitle)
    {
        if (textShape.Paragraphs.Count == 0)
        {
            return [];
        }

        if (isTitle)
        {
            var title = textShape.Paragraphs[0].Text.Replace("<br>", " ", StringComparison.Ordinal).Trim();
            var lines = new List<string> { $"# {title}" };
            foreach (var paragraph in textShape.Paragraphs.Skip(1))
            {
                if (!string.IsNullOrWhiteSpace(paragraph.Text))
                {
                    lines.Add(string.Empty);
                    lines.Add(paragraph.Text.Replace("<br>", "  " + Environment.NewLine, StringComparison.Ordinal));
                }
            }

            return lines;
        }

        if (textShape.Name.StartsWith("Code", StringComparison.OrdinalIgnoreCase) || IsLikelyCodeBlock(textShape))
        {
            var language = ExtractCodeLanguage(textShape.Name, textShape.Paragraphs);
            var lines = new List<string> { $"```{language}" };
            lines.AddRange(textShape.Paragraphs.Select(paragraph => paragraph.Text.Replace("<br>", Environment.NewLine, StringComparison.Ordinal)));
            lines.Add("```");
            return lines;
        }

        if (TryBuildLabelDescriptionLines(textShape, out var labelDescriptionLines))
        {
            return labelDescriptionLines;
        }

        if (textShape.Paragraphs.All(paragraph => paragraph.IsBullet) || IsLikelyBulletList(textShape))
        {
            return textShape.Paragraphs.Select((paragraph, index) => FormatListItem(paragraph, index + 1)).ToArray();
        }

        return textShape.Paragraphs
            .Select(paragraph => paragraph.Text.Replace("<br>", "  " + Environment.NewLine, StringComparison.Ordinal))
            .ToArray();
    }

    private static string ExtractCodeLanguage(string shapeName, IReadOnlyList<ParagraphInfo> paragraphs)
    {
        var start = shapeName.IndexOf('(');
        var end = shapeName.LastIndexOf(')');
        if (start >= 0 && end > start)
        {
            return shapeName[(start + 1)..end].Trim();
        }

        var text = string.Join("\n", paragraphs.Select(paragraph => paragraph.Text));
        return DetectFallbackCodeLanguage(text);
    }

    private static string DetectFallbackCodeLanguage(string text)
    {
        if (LooksLikeCSharp(text))
        {
            return "csharp";
        }

        if (LooksLikeTypeScript(text))
        {
            return "typescript";
        }

        if (LooksLikeJavaScript(text))
        {
            return "javascript";
        }

        if (LooksLikePython(text))
        {
            return "python";
        }

        if (LooksLikeJava(text))
        {
            return "java";
        }

        return string.Empty;
    }

    private static string FormatListItem(ParagraphInfo paragraph, int orderNumber)
    {
        var indent = new string(' ', paragraph.Level * 2);
        var marker = paragraph.IsOrdered ? $"{orderNumber}." : "-";
        return $"{indent}{marker} {paragraph.Text.Replace("<br>", " ", StringComparison.Ordinal)}";
    }

    private static bool IsLikelyBulletList(TextShapeInfo textShape)
    {
        if (textShape.Paragraphs.Count < MinimumInferredBulletCount)
        {
            return false;
        }

        if (textShape.Paragraphs.Any(paragraph => paragraph.IsBullet) || IsLikelyCodeBlock(textShape))
        {
            return false;
        }

        var labelLikeCount = textShape.Paragraphs.Count(IsLikelyLabelParagraph);
        return labelLikeCount >= MinimumInferredBulletCount && labelLikeCount >= (int)Math.Ceiling(textShape.Paragraphs.Count * 0.8);
    }

    private static bool TryBuildLabelDescriptionLines(TextShapeInfo textShape, out IReadOnlyList<string> lines)
    {
        lines = [];
        if (textShape.Paragraphs.Count < 2 || IsLikelyCodeBlock(textShape))
        {
            return false;
        }

        var builtLines = new List<string>();
        var bulletCount = 0;

        for (var index = 0; index < textShape.Paragraphs.Count; index++)
        {
            var paragraph = textShape.Paragraphs[index];
            var nextParagraph = index + 1 < textShape.Paragraphs.Count ? textShape.Paragraphs[index + 1] : null;

            if (IsLikelyLabelParagraph(paragraph))
            {
                builtLines.Add(FormatListItem(paragraph, bulletCount + 1));
                bulletCount++;

                if (nextParagraph is not null && !IsLikelyLabelParagraph(nextParagraph) && IsLikelyDescriptionParagraph(nextParagraph))
                {
                    builtLines.Add("  " + nextParagraph.Text.Replace("<br>", "  " + Environment.NewLine + "  ", StringComparison.Ordinal));
                    index++;
                }

                continue;
            }

            if (bulletCount > 0)
            {
                builtLines.Add(string.Empty);
            }

            builtLines.Add(paragraph.Text.Replace("<br>", "  " + Environment.NewLine, StringComparison.Ordinal));
        }

        var minimumBulletCount = textShape.Paragraphs.Count <= 2 ? 1 : 2;
        if (bulletCount < minimumBulletCount)
        {
            return false;
        }

        lines = builtLines;
        return true;
    }

    private static bool IsLikelyDescriptionParagraph(ParagraphInfo paragraph)
    {
        var text = paragraph.Text.Replace("<br>", " ", StringComparison.Ordinal).Trim();
        if (text.Length < 20)
        {
            return false;
        }

        var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return words.Length >= 4;
    }

    private static bool IsLikelyLabelParagraph(ParagraphInfo paragraph)
    {
        var text = paragraph.Text.Replace("<br>", " ", StringComparison.Ordinal).Trim();
        if (text.Length is < 2 or > 42)
        {
            return false;
        }

        if (text.Contains("http://", StringComparison.OrdinalIgnoreCase) || text.Contains("https://", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (text.IndexOfAny(['=', '{', '}', ';', '(', ')', '[', ']']) >= 0)
        {
            return false;
        }

        if (text.EndsWith(".", StringComparison.Ordinal) ||
            text.EndsWith(":", StringComparison.Ordinal) ||
            text.EndsWith("?", StringComparison.Ordinal) ||
            text.EndsWith("!", StringComparison.Ordinal))
        {
            return false;
        }

        var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return words.Length is > 0 and <= 5;
    }

    private static bool IsLikelyCodeBlock(TextShapeInfo textShape)
    {
        if (textShape.Paragraphs.Count < 3)
        {
            return false;
        }

        var codeLikeLines = textShape.Paragraphs.Count(paragraph => IsCodeLikeLine(paragraph.Text));
        var codeFontLines = textShape.Paragraphs.Count(paragraph => paragraph.UsesCodeFont);

        if (codeFontLines >= 2 && codeFontLines >= (int)Math.Ceiling(textShape.Paragraphs.Count * 0.6))
        {
            return true;
        }

        return codeLikeLines >= 3 && codeLikeLines >= (int)Math.Ceiling(textShape.Paragraphs.Count * 0.6);
    }

    private static bool UsesCodeFont(A.Paragraph paragraph)
    {
        foreach (var run in paragraph.Elements<A.Run>())
        {
            var typeface = run.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(typeface))
            {
                continue;
            }

            if (CodeFontNames.Any(font => typeface.Contains(font, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsCodeLikeLine(string text)
    {
        var normalized = text.Trim();
        if (normalized.Length == 0)
        {
            return false;
        }

        if (StartsWithOrdinal(normalized, "//") ||
            StartsWithOrdinal(normalized, "using ") ||
            StartsWithOrdinal(normalized, "var ") ||
            StartsWithOrdinal(normalized, "new ") ||
            StartsWithOrdinal(normalized, "await ") ||
            StartsWithOrdinal(normalized, "return "))
        {
            return true;
        }

        if (ContainsOrdinal(normalized, "=>") ||
            ContainsOrdinal(normalized, "();") ||
            ContainsOrdinal(normalized, " = ") ||
            ContainsOrdinal(normalized, "{") ||
            ContainsOrdinal(normalized, "}") ||
            EndsWithOrdinal(normalized, ";"))
        {
            return true;
        }

        return false;
    }

    private static bool LooksLikeCSharp(string text)
    {
        return ContainsOrdinal(text, "using System") ||
            ContainsOrdinal(text, "Console.") ||
            ContainsOrdinal(text, "namespace ") ||
            ContainsOrdinal(text, "async Task") ||
            ContainsOrdinal(text, "public record ") ||
            ContainsOrdinal(text, "public sealed class ") ||
            ContainsOrdinal(text, "get; set;") ||
            ContainsOrdinal(text, "using var ") ||
            ContainsOrdinal(text, "nameof(") ||
            ContainsOrdinal(text, "IEnumerable<");
    }

    private static bool LooksLikeTypeScript(string text)
    {
        return ContainsOrdinal(text, "interface ") ||
            ContainsOrdinal(text, "type ") ||
            ContainsOrdinal(text, "implements ") ||
            ContainsOrdinal(text, "readonly ") ||
            ContainsOrdinal(text, ": string") ||
            ContainsOrdinal(text, ": number") ||
            ContainsOrdinal(text, ": boolean") ||
            ContainsOrdinal(text, " as const");
    }

    private static bool LooksLikeJavaScript(string text)
    {
        return ContainsOrdinal(text, "console.log") ||
            ContainsOrdinal(text, "function ") ||
            ContainsOrdinal(text, "const ") ||
            ContainsOrdinal(text, "let ") ||
            ContainsOrdinal(text, "export default") ||
            ContainsOrdinal(text, "module.exports") ||
            ContainsOrdinal(text, "require(");
    }

    private static bool LooksLikePython(string text)
    {
        return ContainsOrdinal(text, "def ") ||
            ContainsOrdinal(text, "print(") ||
            ContainsOrdinal(text, "if __name__ == \"__main__\":") ||
            ContainsOrdinal(text, "if __name__ == '__main__':") ||
            ContainsOrdinal(text, "from ") && ContainsOrdinal(text, " import ");
    }

    private static bool LooksLikeJava(string text)
    {
        return ContainsOrdinal(text, "public class ") ||
            ContainsOrdinal(text, "public static void main") ||
            ContainsOrdinal(text, "System.out.") ||
            ContainsOrdinal(text, "import java.") ||
            ContainsOrdinal(text, "package ") ||
            ContainsOrdinal(text, "private static final");
    }

    private static bool ContainsOrdinal(string text, string value)
        => text.Contains(value, StringComparison.Ordinal);

    private static bool StartsWithOrdinal(string text, string value)
        => text.StartsWith(value, StringComparison.Ordinal);

    private static bool EndsWithOrdinal(string text, string value)
        => text.EndsWith(value, StringComparison.Ordinal);
}
