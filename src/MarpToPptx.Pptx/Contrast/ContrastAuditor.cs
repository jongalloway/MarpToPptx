using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Contrast;

/// <summary>
/// Audits a PPTX file for WCAG 2.1 text/background contrast failures
/// by inspecting solid-fill pairs on slides, shapes, and table cells.
/// </summary>
public sealed class ContrastAuditor
{
    /// <summary>
    /// Default fallback background color (white) used when no explicit fill can be resolved.
    /// </summary>
    private const string DefaultBackgroundColor = "FFFFFF";

    /// <summary>
    /// Opens and audits all slides in the PPTX at the given path.
    /// </summary>
    public IReadOnlyList<ContrastAuditResult> Audit(string pptxPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pptxPath);

        using var document = PresentationDocument.Open(pptxPath, false);
        return Audit(document);
    }

    /// <summary>
    /// Audits all slides in an already-open <see cref="PresentationDocument"/>.
    /// </summary>
    public IReadOnlyList<ContrastAuditResult> Audit(PresentationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var results = new List<ContrastAuditResult>();

        var presentationPart = document.PresentationPart
            ?? throw new InvalidOperationException("PresentationDocument has no PresentationPart.");

        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<SlideId>().ToArray()
            ?? [];

        for (var i = 0; i < slideIds.Length; i++)
        {
            var relId = slideIds[i].RelationshipId?.Value;
            if (relId is null)
                continue;

            var slidePart = (SlidePart)presentationPart.GetPartById(relId);
            AuditSlide(slidePart, slideNumber: i + 1, results);
        }

        return results;
    }

    // ─────────────────────────────────────────────────────────────
    // Slide-level traversal
    // ─────────────────────────────────────────────────────────────

    private static void AuditSlide(SlidePart slidePart, int slideNumber, List<ContrastAuditResult> results)
    {
        var slide = slidePart.Slide;
        if (slide is null)
            return;

        var slideBackground = ResolveSlideBackground(slide);

        // Audit regular (non-table) shapes
        foreach (var shape in slide.Descendants<P.Shape>())
        {
            var shapeName = shape.NonVisualShapeProperties?
                .NonVisualDrawingProperties?.Name?.Value ?? "Shape";

            // The background rectangle itself has no foreground text we care about
            if (shapeName == "Background")
                continue;

            var shapeFill = GetSolidFillFromElement(shape.ShapeProperties);
            var effectiveBg = shapeFill ?? slideBackground ?? DefaultBackgroundColor;

            var textBody = shape.TextBody;
            if (textBody is null)
                continue;

            AuditTextBody(textBody, effectiveBg, slideNumber, $"Shape \"{shapeName}\"", results);
        }

        // Audit table cells
        var tableIndex = 0;
        foreach (var table in slide.Descendants<A.Table>())
        {
            tableIndex++;
            var rowIndex = 0;
            foreach (var row in table.Elements<A.TableRow>())
            {
                rowIndex++;
                var cellIndex = 0;
                foreach (var cell in row.Elements<A.TableCell>())
                {
                    cellIndex++;
                    var cellFill = GetTableCellFill(cell);
                    var effectiveBg = cellFill ?? slideBackground ?? DefaultBackgroundColor;
                    var context = $"Table {tableIndex}, Row {rowIndex}, Cell {cellIndex}";

                    var textBody = cell.TextBody;
                    if (textBody is null)
                        continue;

                    AuditTextBody(textBody, effectiveBg, slideNumber, context, results);
                }
            }
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Text body / paragraph / run traversal
    // ─────────────────────────────────────────────────────────────

    private static void AuditTextBody(
        DocumentFormat.OpenXml.OpenXmlElement textBody,
        string backgroundHex,
        int slideNumber,
        string context,
        List<ContrastAuditResult> results)
    {
        foreach (var paragraph in textBody.Elements<A.Paragraph>())
        {
            // WCAG 2.1 large text: 18pt+ (any weight) or 14pt+ bold.
            // The paragraph-level default run properties may carry a font size.
            var paraFontSize = GetParagraphFontSizePt(paragraph);

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var runColor = GetRunColor(run);
                if (runColor is null)
                    continue;

                var runFontSize = GetRunFontSizePt(run) ?? paraFontSize;
                var isBold = run.RunProperties?.Bold?.Value == true;
                var isLargeText = runFontSize.HasValue
                    && (runFontSize.Value >= 18.0 || (isBold && runFontSize.Value >= 14.0));

                var ratio = ContrastCalculator.ContrastRatio(runColor, backgroundHex);

                results.Add(new ContrastAuditResult(
                    SlideNumber: slideNumber,
                    ShapeContext: context,
                    ForegroundColor: runColor,
                    BackgroundColor: backgroundHex,
                    ContrastRatio: ratio,
                    IsLargeText: isLargeText));
            }
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Color resolution helpers
    // ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Resolves the slide-level background color.
    /// First checks for the renderer-emitted "Background" rectangle shape,
    /// then falls back to the CSld.Background element.
    /// </summary>
    private static string? ResolveSlideBackground(Slide slide)
    {
        // The renderer places a filled rectangle named "Background" in the shape tree.
        foreach (var shape in slide.Descendants<P.Shape>())
        {
            if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Background")
            {
                var color = GetSolidFillFromElement(shape.ShapeProperties);
                if (color is not null)
                    return color;
            }
        }

        // Fall back to the OOXML Background element on the common slide data.
        var cSld = slide.CommonSlideData;
        if (cSld?.Background?.BackgroundProperties is { } bgPr)
        {
            var color = GetSolidFillFromElement(bgPr);
            if (color is not null)
                return color;
        }

        return null;
    }

    /// <summary>
    /// Returns the hex color string (no '#') from the first SolidFill > RgbColorModelHex
    /// child of the given element, or null if none exists.
    /// </summary>
    private static string? GetSolidFillFromElement(DocumentFormat.OpenXml.OpenXmlElement? element)
    {
        if (element is null)
            return null;

        return element
            .Descendants<A.SolidFill>()
            .Select(sf => sf.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value)
            .FirstOrDefault(v => v is not null);
    }

    /// <summary>
    /// Returns the hex color from a table cell's TableCellProperties fill, or null.
    /// </summary>
    private static string? GetTableCellFill(A.TableCell cell)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr is null)
            return null;

        return tcPr
            .Descendants<A.SolidFill>()
            .Select(sf => sf.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value)
            .FirstOrDefault(v => v is not null);
    }

    /// <summary>
    /// Returns the hex color from a run's RunProperties SolidFill, or null if not set.
    /// </summary>
    private static string? GetRunColor(A.Run run)
    {
        var rPr = run.RunProperties;
        if (rPr is null)
            return null;

        return rPr
            .Descendants<A.SolidFill>()
            .Select(sf => sf.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value)
            .FirstOrDefault(v => v is not null);
    }

    // ─────────────────────────────────────────────────────────────
    // Font-size helpers (returns points as double)
    // ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Returns the font size in points from a run's RunProperties, or null if not set.
    /// Open XML stores font sizes as hundredths of a point.
    /// </summary>
    private static double? GetRunFontSizePt(A.Run run)
    {
        var sz = run.RunProperties?.FontSize;
        if (sz is null)
            return null;
        return sz.Value / 100.0;
    }

    /// <summary>
    /// Returns the default font size in points from paragraph-level default run properties, or null.
    /// </summary>
    private static double? GetParagraphFontSizePt(A.Paragraph paragraph)
    {
        var sz = paragraph.ParagraphProperties
            ?.GetFirstChild<A.DefaultRunProperties>()?.FontSize;
        if (sz is null)
            return null;
        return sz.Value / 100.0;
    }
}
