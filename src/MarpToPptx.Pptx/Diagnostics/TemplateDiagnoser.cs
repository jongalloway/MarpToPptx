using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Pptx.Rendering;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Analyzes a PPTX template and produces a <see cref="TemplateDiagnosticReport"/> that summarizes
/// layout structure, inferred semantic roles, placeholder coverage, and Markdown-directive
/// recommendations for use with MarpToPptx.
/// </summary>
public sealed class TemplateDiagnoser
{
    /// <summary>
    /// Opens the PPTX file at <paramref name="templatePath"/> and returns a diagnostic report.
    /// </summary>
    public TemplateDiagnosticReport Diagnose(string templatePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);

        using var document = PresentationDocument.Open(templatePath, false);
        return Diagnose(document, templatePath);
    }

    /// <summary>
    /// Returns a diagnostic report for an already-open <see cref="PresentationDocument"/>.
    /// <paramref name="templatePath"/> is recorded in the report for display purposes only.
    /// </summary>
    public TemplateDiagnosticReport Diagnose(PresentationDocument document, string templatePath)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(templatePath);

        var presentationPart = document.PresentationPart
            ?? throw new InvalidOperationException("PresentationDocument has no PresentationPart.");

        var masterParts = presentationPart.SlideMasterParts.ToList();

        // Collect all layouts in document order across all masters.
        var allLayouts = new List<SlideLayoutPart>();
        foreach (var master in masterParts)
        {
            allLayouts.AddRange(master.SlideLayoutParts);
        }

        // Build per-layout diagnostics (redundancy flag set after grouping).
        var preDiagnostics = allLayouts
            .Select((layoutPart, i) => BuildLayoutDiagnostic(layoutPart, layoutIndex: i + 1, likelyVisuallyRedundant: false))
            .ToList();

        // A layout is "likely visually redundant" when it has no distinct non-placeholder shapes
        // and shares its semantic role with at least one other zero-shape layout.
        var redundantIndices = preDiagnostics
            .Select((d, i) => (d, i))
            .Where(x => x.d.NonPlaceholderShapeCount == 0)
            .GroupBy(x => x.d.SemanticRole)
            .Where(g => g.Count() > 1)
            .SelectMany(g => g.Select(x => x.i))
            .ToHashSet();

        var layouts = preDiagnostics
            .Select((d, i) => redundantIndices.Contains(i) ? d with { LikelyVisuallyRedundant = true } : d)
            .ToArray();

        return new TemplateDiagnosticReport(
            templatePath,
            SlideMasterCount: masterParts.Count,
            Layouts: layouts,
            RecommendedDefaultContentLayout: PickDefaultContentLayout(layouts),
            RecommendedTitleLayout: layouts.FirstOrDefault(l => l.SemanticRole == LayoutSemanticRole.Title)?.Name,
            RecommendedSectionLayout: layouts.FirstOrDefault(l => l.SemanticRole == LayoutSemanticRole.SectionHeader)?.Name,
            RecommendedPictureCaptionLayout: PickPictureCaptionLayout(layouts),
            Warnings: BuildWarnings(layouts));
    }

    private static LayoutDiagnostic BuildLayoutDiagnostic(SlideLayoutPart layoutPart, int layoutIndex, bool likelyVisuallyRedundant)
    {
        var slideLayout = layoutPart.SlideLayout;
        var name = GetLayoutName(slideLayout, layoutIndex);
        var typeCode = slideLayout?.Type?.InnerText;
        var semanticRole = MapSemanticRole(typeCode);
        var hasTitle = SlideTemplateSelector.GetTitlePlaceholder(layoutPart) is not null;
        var hasBody = SlideTemplateSelector.GetBodyPlaceholder(layoutPart) is not null;
        var hasPicture = SlideTemplateSelector.GetPicturePlaceholder(layoutPart) is not null;
        var nonPlaceholderCount = CountNonPlaceholderShapes(layoutPart);

        return new LayoutDiagnostic(
            name,
            typeCode,
            semanticRole,
            hasTitle,
            hasBody,
            hasPicture,
            nonPlaceholderCount,
            likelyVisuallyRedundant);
    }

    private static string GetLayoutName(P.SlideLayout? slideLayout, int layoutIndex)
    {
        if (slideLayout?.MatchingName?.Value is { Length: > 0 } matchingName)
        {
            return matchingName;
        }

        if (slideLayout?.CommonSlideData?.Name?.Value is { Length: > 0 } csdName)
        {
            return csdName;
        }

        return $"Layout {layoutIndex}";
    }

    private static int CountNonPlaceholderShapes(SlideLayoutPart layoutPart)
    {
        var shapeTree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
        {
            return 0;
        }

        var count = 0;

        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            var ph = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();
            if (ph is null)
            {
                count++;
            }
        }

        // Pictures, connectors, and graphic frames (tables/charts) are also distinct design elements.
        count += shapeTree.Elements<P.Picture>().Count();
        count += shapeTree.Elements<P.GraphicFrame>().Count();
        count += shapeTree.Elements<P.ConnectionShape>().Count();

        return count;
    }

    private static LayoutSemanticRole MapSemanticRole(string? typeCode)
    {
        // SlideLayoutValues in OpenXML SDK 3.x is a struct with static properties, not an enum,
        // so we compare the raw XML type-code string from the <p:sldLayout type="..."> attribute.
        return typeCode switch
        {
            "title"               => LayoutSemanticRole.Title,
            "titleOnly"           => LayoutSemanticRole.TitleOnly,
            "secHead"             => LayoutSemanticRole.SectionHeader,
            "blank"               => LayoutSemanticRole.Blank,
            // Picture / media + caption layouts
            "picTx"               => LayoutSemanticRole.PictureCaption,
            "clipArtAndTx"        => LayoutSemanticRole.PictureCaption,
            "txAndClipArt"        => LayoutSemanticRole.PictureCaption,
            "clipArtAndVertTx"    => LayoutSemanticRole.PictureCaption,
            "txAndMedia"          => LayoutSemanticRole.PictureCaption,
            "mediaAndTx"          => LayoutSemanticRole.PictureCaption,
            // Multi-column / comparison layouts
            "twoColTx"            => LayoutSemanticRole.Comparison,
            "twoTxTwoObj"         => LayoutSemanticRole.Comparison,
            "twoObj"              => LayoutSemanticRole.Comparison,
            "txAndTwoObj"         => LayoutSemanticRole.Comparison,
            "twoObjAndTx"         => LayoutSemanticRole.Comparison,
            "twoObjOverTx"        => LayoutSemanticRole.Comparison,
            "objAndTwoObj"        => LayoutSemanticRole.Comparison,
            "twoObjAndObj"        => LayoutSemanticRole.Comparison,
            // Custom / unrecognised
            "cust"                => LayoutSemanticRole.Other,
            // Everything else (tx, obj, fourObj, tbl, chart, vert*, …) is a content layout.
            null                  => LayoutSemanticRole.Other,
            _                     => LayoutSemanticRole.Content,
        };
    }

    private static string? PickDefaultContentLayout(IReadOnlyList<LayoutDiagnostic> layouts)
    {
        // Prefer a content layout that has both title and body placeholders.
        var withBoth = layouts.FirstOrDefault(l =>
            l.SemanticRole == LayoutSemanticRole.Content &&
            l.HasTitlePlaceholder &&
            l.HasBodyPlaceholder);

        if (withBoth is not null)
        {
            return withBoth.Name;
        }

        return layouts.FirstOrDefault(l => l.SemanticRole == LayoutSemanticRole.Content)?.Name;
    }

    private static string? PickPictureCaptionLayout(IReadOnlyList<LayoutDiagnostic> layouts)
    {
        // Prefer a layout explicitly typed as PictureCaption.
        var pictureCaptionLayout = layouts.FirstOrDefault(l => l.SemanticRole == LayoutSemanticRole.PictureCaption);
        if (pictureCaptionLayout is not null)
        {
            return pictureCaptionLayout.Name;
        }

        // Fall back to any layout that declares a picture placeholder.
        return layouts.FirstOrDefault(l => l.HasPicturePlaceholder)?.Name;
    }

    private static IReadOnlyList<string> BuildWarnings(IReadOnlyList<LayoutDiagnostic> layouts)
    {
        var warnings = new List<string>();

        // Warn about each group of visually redundant layouts.
        var redundantGroups = layouts
            .Where(l => l.LikelyVisuallyRedundant)
            .GroupBy(l => l.SemanticRole)
            .ToList();

        foreach (var group in redundantGroups)
        {
            var names = string.Join(", ", group.Select(l => $"\"{l.Name}\""));
            warnings.Add(
                $"Layouts {names} share the {group.Key} role and have no distinct non-placeholder shapes. " +
                "They will appear visually identical with the current renderer. " +
                "Test each before using in Markdown directives.");
        }

        // Warn if no content layout has both a title and body placeholder.
        var contentLayouts = layouts.Where(l => l.SemanticRole == LayoutSemanticRole.Content).ToList();
        if (contentLayouts.Count > 0 && !contentLayouts.Any(l => l.HasTitlePlaceholder && l.HasBodyPlaceholder))
        {
            warnings.Add(
                "No Content layout was found with both a title and a body placeholder. " +
                "Normal content slides may not inherit placeholder geometry from the template.");
        }

        // Warn if there is no title layout.
        if (!layouts.Any(l => l.SemanticRole == LayoutSemanticRole.Title))
        {
            warnings.Add(
                "No title-type layout was found. " +
                "The first slide may not use an appropriate layout automatically.");
        }

        return warnings;
    }
}
