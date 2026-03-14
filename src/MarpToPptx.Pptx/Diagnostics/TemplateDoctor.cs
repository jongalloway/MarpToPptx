using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Pptx.Rendering;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Analyzes a PPTX template for structural issues that are likely to degrade MarpToPptx output
/// and optionally writes a repaired copy with safe, automatable fixups applied.
/// </summary>
public sealed class TemplateDoctor
{
    /// <summary>
    /// Analyzes the template at <paramref name="templatePath"/> and returns a report of all
    /// structural issues found.  This is a read-only, dry-run operation — no files are written.
    /// </summary>
    public TemplateDoctorReport Analyze(string templatePath)
        => Run(templatePath, outputPath: null);

    /// <summary>
    /// Analyzes the template at <paramref name="templatePath"/> and, if <paramref name="outputPath"/>
    /// is non-null, writes a repaired copy to that path with all safe fixups applied.
    /// </summary>
    /// <param name="templatePath">Source template to inspect.</param>
    /// <param name="outputPath">
    /// Destination for the fixed template, or <c>null</c> to perform a dry run only.
    /// The original file is never modified.
    /// </param>
    public TemplateDoctorReport Run(string templatePath, string? outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);

        // ── Phase 1: read-only analysis ─────────────────────────────────────
        var issues = new List<TemplateDoctorIssue>();
        List<SlideLayoutPart> analysisLayouts;

        using (var document = PresentationDocument.Open(templatePath, false))
        {
            var presentationPart = document.PresentationPart
                ?? throw new InvalidOperationException("PresentationDocument has no PresentationPart.");

            analysisLayouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts)
                .ToList();

            CollectIssues(analysisLayouts, issues);
        }

        // ── Phase 2: optional write-back with fixups ─────────────────────────
        var appliedFixes = new List<string>();
        var wroteFix = false;

        if (outputPath is not null)
        {
            var fixableIssues = issues.Where(i => i.Severity == IssueSeverity.Fixable).ToList();

            File.Copy(templatePath, outputPath, overwrite: true);

            if (fixableIssues.Count > 0)
            {
                using var editDoc = PresentationDocument.Open(outputPath, true);
                ApplyGeometryFixes(editDoc, appliedFixes);
            }

            wroteFix = true;
        }

        return new TemplateDoctorReport(templatePath, issues, wroteFix, outputPath, appliedFixes);
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Analysis
    // ──────────────────────────────────────────────────────────────────────────

    private static void CollectIssues(List<SlideLayoutPart> layouts, List<TemplateDoctorIssue> issues)
    {
        CheckDuplicateLayoutNames(layouts, issues);
        CheckEmptyLayoutNames(layouts, issues);

        for (var i = 0; i < layouts.Count; i++)
        {
            var layoutPart = layouts[i];
            var name = GetLayoutName(layoutPart.SlideLayout, i + 1);
            var typeCode = layoutPart.SlideLayout?.Type?.InnerText;
            var role = MapSemanticRole(typeCode);

            if (role == LayoutSemanticRole.Content)
            {
                CheckContentLayoutPlaceholders(layoutPart, name, issues);
                CheckPlaceholderGeometryInherited(layoutPart, name, issues);
                CheckTypelessIndexedBody(layoutPart, name, issues);
            }

            if (role is LayoutSemanticRole.PictureCaption or LayoutSemanticRole.Comparison)
            {
                CheckUnmappableLayout(layoutPart, name, role, issues);
            }
        }

        CheckVisuallyRedundantLayouts(layouts, issues);
    }

    private static void CheckDuplicateLayoutNames(List<SlideLayoutPart> layouts, List<TemplateDoctorIssue> issues)
    {
        var nameCounts = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < layouts.Count; i++)
        {
            var name = GetLayoutName(layouts[i].SlideLayout, i + 1);
            if (!nameCounts.TryGetValue(name, out var indices))
            {
                indices = [];
                nameCounts[name] = indices;
            }

            indices.Add(i + 1);
        }

        foreach (var (name, indices) in nameCounts)
        {
            if (indices.Count <= 1)
            {
                continue;
            }

            issues.Add(new TemplateDoctorIssue(
                LayoutName: name,
                Severity: IssueSeverity.Warning,
                Code: "DuplicateLayoutName",
                Description: $"Layout name \"{name}\" is used by {indices.Count} layouts " +
                    $"(at positions {string.Join(", ", indices)}). " +
                    "Duplicate names make it impossible to target layouts by name in Markdown directives."));
        }
    }

    private static void CheckEmptyLayoutNames(List<SlideLayoutPart> layouts, List<TemplateDoctorIssue> issues)
    {
        for (var i = 0; i < layouts.Count; i++)
        {
            var layout = layouts[i].SlideLayout;
            var hasMatchingName = layout?.MatchingName?.Value is { Length: > 0 };
            var hasCsdName = layout?.CommonSlideData?.Name?.Value is { Length: > 0 };

            if (!hasMatchingName && !hasCsdName)
            {
                issues.Add(new TemplateDoctorIssue(
                    LayoutName: null,
                    Severity: IssueSeverity.Warning,
                    Code: "EmptyLayoutName",
                    Description: $"Layout at position {i + 1} has no name and will be referenced as \"Layout {i + 1}\". " +
                        "Unnamed layouts are hard to target by name in Markdown directives."));
            }
        }
    }

    private static void CheckContentLayoutPlaceholders(
        SlideLayoutPart layoutPart,
        string name,
        List<TemplateDoctorIssue> issues)
    {
        var hasTitle = SlideTemplateSelector.GetTitlePlaceholder(layoutPart) is not null;
        var hasBody = SlideTemplateSelector.GetBodyPlaceholder(layoutPart) is not null;

        if (!hasTitle)
        {
            issues.Add(new TemplateDoctorIssue(
                LayoutName: name,
                Severity: IssueSeverity.Warning,
                Code: "ContentLayoutMissingTitlePlaceholder",
                Description: $"Content layout \"{name}\" has no title placeholder. " +
                    "Slide headings will not be placed using the template's title geometry."));
        }

        if (!hasBody)
        {
            issues.Add(new TemplateDoctorIssue(
                LayoutName: name,
                Severity: IssueSeverity.Warning,
                Code: "ContentLayoutMissingBodyPlaceholder",
                Description: $"Content layout \"{name}\" has no body placeholder. " +
                    "Slide body content will not be placed using the template's body geometry."));
        }
    }

    private static void CheckPlaceholderGeometryInherited(
        SlideLayoutPart layoutPart,
        string name,
        List<TemplateDoctorIssue> issues)
    {
        var titlePlaceholder = SlideTemplateSelector.GetTitlePlaceholder(layoutPart);
        var bodyPlaceholder = SlideTemplateSelector.GetBodyPlaceholder(layoutPart);

        if (titlePlaceholder is not null && !HasOwnTransform(layoutPart, titlePlaceholder,
                P.PlaceholderValues.Title, P.PlaceholderValues.CenteredTitle))
        {
            issues.Add(new TemplateDoctorIssue(
                LayoutName: name,
                Severity: IssueSeverity.Fixable,
                Code: "PlaceholderGeometryInherited",
                Description: $"Content layout \"{name}\" title placeholder has no explicit geometry — " +
                    "position and size are inherited from the slide master. " +
                    "The renderer can recover via master fallback, but materializing the geometry " +
                    "improves template portability and predictability.",
                ProposedFix: $"Copy the title placeholder transform from the slide master onto layout \"{name}\"."));
        }

        if (bodyPlaceholder is not null && !HasOwnTransform(layoutPart, bodyPlaceholder,
                P.PlaceholderValues.Body, P.PlaceholderValues.SubTitle))
        {
            issues.Add(new TemplateDoctorIssue(
                LayoutName: name,
                Severity: IssueSeverity.Fixable,
                Code: "PlaceholderGeometryInherited",
                Description: $"Content layout \"{name}\" body placeholder has no explicit geometry — " +
                    "position and size are inherited from the slide master. " +
                    "The renderer can recover via master fallback, but materializing the geometry " +
                    "improves template portability and predictability.",
                ProposedFix: $"Copy the body placeholder transform from the slide master onto layout \"{name}\"."));
        }
    }

    private static void CheckTypelessIndexedBody(
        SlideLayoutPart layoutPart,
        string name,
        List<TemplateDoctorIssue> issues)
    {
        // Body placeholder is resolved as a typeless indexed placeholder (no type attribute,
        // idx only). This works, but is a weaker signal than an explicit body type.
        var hasTypedBody = SlideTemplateSelector.GetBodyPlaceholder(layoutPart) is { Type: not null };
        var hasTypelessBody = !hasTypedBody &&
            SlideTemplateSelector.GetBodyPlaceholder(layoutPart) is { Type: null };

        if (hasTypelessBody)
        {
            issues.Add(new TemplateDoctorIssue(
                LayoutName: name,
                Severity: IssueSeverity.Info,
                Code: "TypelessIndexedBodyPlaceholder",
                Description: $"Content layout \"{name}\" exposes body content only via a typeless " +
                    "indexed placeholder (<p:ph idx=\"...\"/>). " +
                    "This is the standard form for object/content placeholders and is fully supported."));
        }
    }

    private static void CheckUnmappableLayout(
        SlideLayoutPart layoutPart,
        string name,
        LayoutSemanticRole role,
        List<TemplateDoctorIssue> issues)
    {
        issues.Add(new TemplateDoctorIssue(
            LayoutName: name,
            Severity: IssueSeverity.Info,
            Code: "UnmappableLayoutRole",
            Description: $"Layout \"{name}\" has semantic role {role}, which MarpToPptx " +
                "cannot yet map to a Markdown content type. " +
                "It will not be auto-selected during rendering."));
    }

    private static void CheckVisuallyRedundantLayouts(
        List<SlideLayoutPart> layouts,
        List<TemplateDoctorIssue> issues)
    {
        // Build the same redundancy groups that TemplateDiagnoser uses.
        var noShapeLayouts = layouts
            .Select((lp, i) => (lp, i, name: GetLayoutName(lp.SlideLayout, i + 1), role: MapSemanticRole(lp.SlideLayout?.Type?.InnerText)))
            .Where(x => CountNonPlaceholderShapes(x.lp) == 0)
            .ToList();

        var redundantGroups = noShapeLayouts
            .GroupBy(x => x.role)
            .Where(g => g.Count() > 1);

        foreach (var group in redundantGroups)
        {
            var names = string.Join(", ", group.Select(x => $"\"{x.name}\""));
            issues.Add(new TemplateDoctorIssue(
                LayoutName: null,
                Severity: IssueSeverity.Info,
                Code: "VisuallyRedundantLayouts",
                Description: $"Layouts {names} share the {group.Key} role and have no distinct " +
                    "non-placeholder shapes — they will appear visually identical with the current renderer. " +
                    "Test each before using in Markdown directives."));
        }
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Fixup application
    // ──────────────────────────────────────────────────────────────────────────

    private static void ApplyGeometryFixes(PresentationDocument document, List<string> appliedFixes)
    {
        var masterParts = document.PresentationPart!.SlideMasterParts.ToList();
        var allLayouts = masterParts.SelectMany(m => m.SlideLayoutParts).ToList();

        for (var i = 0; i < allLayouts.Count; i++)
        {
            var layoutPart = allLayouts[i];
            var name = GetLayoutName(layoutPart.SlideLayout, i + 1);
            var typeCode = layoutPart.SlideLayout?.Type?.InnerText;
            var role = MapSemanticRole(typeCode);

            if (role != LayoutSemanticRole.Content)
            {
                continue;
            }

            var titlePlaceholder = SlideTemplateSelector.GetTitlePlaceholder(layoutPart);
            if (titlePlaceholder is not null && !HasOwnTransform(layoutPart, titlePlaceholder,
                    P.PlaceholderValues.Title, P.PlaceholderValues.CenteredTitle))
            {
                if (TryMaterializeTransform(layoutPart, titlePlaceholder,
                        [P.PlaceholderValues.Title, P.PlaceholderValues.CenteredTitle],
                        [P.PlaceholderValues.Title, P.PlaceholderValues.CenteredTitle]))
                {
                    appliedFixes.Add(
                        $"Materialized title placeholder geometry from slide master onto layout \"{name}\".");
                }
            }

            var bodyPlaceholder = SlideTemplateSelector.GetBodyPlaceholder(layoutPart);
            if (bodyPlaceholder is not null && !HasOwnTransform(layoutPart, bodyPlaceholder,
                    P.PlaceholderValues.Body, P.PlaceholderValues.SubTitle))
            {
                if (TryMaterializeTransform(layoutPart, bodyPlaceholder,
                        [P.PlaceholderValues.Body, P.PlaceholderValues.SubTitle],
                        [P.PlaceholderValues.Body, P.PlaceholderValues.SubTitle]))
                {
                    appliedFixes.Add(
                        $"Materialized body placeholder geometry from slide master onto layout \"{name}\".");
                }
            }
        }
    }

    /// <summary>
    /// Finds the matching placeholder shape in the layout's own shape tree, then finds the
    /// corresponding shape in the slide master, and copies the master's transform to the layout
    /// shape.  Returns <see langword="true"/> if the fix was successfully applied.
    /// </summary>
    private static bool TryMaterializeTransform(
        SlideLayoutPart layoutPart,
        TemplatePlaceholder placeholder,
        P.PlaceholderValues[] layoutMatchTypes,
        P.PlaceholderValues[] masterMatchTypes)
    {
        var layoutShape = FindPlaceholderShape(
            layoutPart.SlideLayout?.CommonSlideData?.ShapeTree,
            placeholder,
            layoutMatchTypes);

        if (layoutShape is null)
        {
            return false;
        }

        var masterShape = FindMasterPlaceholderShape(
            layoutPart.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree,
            placeholder,
            masterMatchTypes);

        if (masterShape is null)
        {
            return false;
        }

        var masterXfrm = masterShape.ShapeProperties?.Transform2D;
        if (masterXfrm is null ||
            masterXfrm.Offset is null ||
            masterXfrm.Extents is null ||
            (masterXfrm.Extents.Cx?.Value ?? 0) <= 0 ||
            (masterXfrm.Extents.Cy?.Value ?? 0) <= 0)
        {
            return false;
        }

        var clonedXfrm = (A.Transform2D)masterXfrm.CloneNode(true);

        var shapeProps = layoutShape.ShapeProperties;
        if (shapeProps is null)
        {
            layoutShape.ShapeProperties = new P.ShapeProperties();
            shapeProps = layoutShape.ShapeProperties;
        }

        // Replace any existing transform.
        shapeProps.Transform2D = clonedXfrm;

        return true;
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Shape-tree helpers
    // ──────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Returns <see langword="true"/> when the layout's own shape tree contains a shape matching
    /// <paramref name="placeholder"/> that carries a valid, non-zero transform.
    /// Does not fall back to the slide master.
    /// </summary>
    private static bool HasOwnTransform(
        SlideLayoutPart layoutPart,
        TemplatePlaceholder placeholder,
        params P.PlaceholderValues[] masterFallbackTypes)
    {
        var shape = FindPlaceholderShape(
            layoutPart.SlideLayout?.CommonSlideData?.ShapeTree,
            placeholder,
            masterFallbackTypes);

        if (shape is null)
        {
            return false;
        }

        var xfrm = shape.ShapeProperties?.Transform2D;
        return xfrm?.Offset is not null
            && xfrm.Extents is not null
            && (xfrm.Extents.Cx?.Value ?? 0) > 0
            && (xfrm.Extents.Cy?.Value ?? 0) > 0;
    }

    private static P.Shape? FindPlaceholderShape(
        P.ShapeTree? shapeTree,
        TemplatePlaceholder placeholder,
        P.PlaceholderValues[] fallbackTypes)
    {
        if (shapeTree is null)
        {
            return null;
        }

        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            var ph = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();

            if (ph is null || !MatchesPlaceholder(ph, placeholder, fallbackTypes))
            {
                continue;
            }

            return shape;
        }

        return null;
    }

    private static P.Shape? FindMasterPlaceholderShape(
        P.ShapeTree? masterShapeTree,
        TemplatePlaceholder placeholder,
        P.PlaceholderValues[] masterTypes)
    {
        if (masterShapeTree is null)
        {
            return null;
        }

        foreach (var shape in masterShapeTree.Elements<P.Shape>())
        {
            var ph = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();

            if (ph is null)
            {
                continue;
            }

            // For typed placeholders, match by type.
            if (placeholder.Type is { } phType)
            {
                if (masterTypes.Any(t => ph.Type?.Value == t))
                {
                    return shape;
                }

                continue;
            }

            // For typeless placeholders, look for body/content type on master.
            if (masterTypes.Any(t => ph.Type?.Value == t) || ph.Type?.Value is null)
            {
                return shape;
            }
        }

        return null;
    }

    private static bool MatchesPlaceholder(
        P.PlaceholderShape ph,
        TemplatePlaceholder placeholder,
        P.PlaceholderValues[] fallbackTypes)
    {
        // Index must match when specified.
        if (placeholder.Index is not null && ph.Index?.Value != placeholder.Index)
        {
            return false;
        }

        if (placeholder.Type is { } type)
        {
            return ph.Type?.Value == type;
        }

        // Typeless placeholder: match typeless or body-like shapes.
        var actualType = ph.Type?.Value;
        return actualType is null || fallbackTypes.Any(t => actualType == t);
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

        count += shapeTree.Elements<P.Picture>().Count();
        count += shapeTree.Elements<P.GraphicFrame>().Count();
        count += shapeTree.Elements<P.ConnectionShape>().Count();

        return count;
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Shared helpers (mirrored from TemplateDiagnoser)
    // ──────────────────────────────────────────────────────────────────────────

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

    private static LayoutSemanticRole MapSemanticRole(string? typeCode)
        => typeCode switch
        {
            "title"               => LayoutSemanticRole.Title,
            "titleOnly"           => LayoutSemanticRole.TitleOnly,
            "secHead"             => LayoutSemanticRole.SectionHeader,
            "blank"               => LayoutSemanticRole.Blank,
            "picTx"               => LayoutSemanticRole.PictureCaption,
            "clipArtAndTx"        => LayoutSemanticRole.PictureCaption,
            "txAndClipArt"        => LayoutSemanticRole.PictureCaption,
            "clipArtAndVertTx"    => LayoutSemanticRole.PictureCaption,
            "txAndMedia"          => LayoutSemanticRole.PictureCaption,
            "mediaAndTx"          => LayoutSemanticRole.PictureCaption,
            "twoColTx"            => LayoutSemanticRole.Comparison,
            "twoTxTwoObj"         => LayoutSemanticRole.Comparison,
            "twoObj"              => LayoutSemanticRole.Comparison,
            "txAndTwoObj"         => LayoutSemanticRole.Comparison,
            "twoObjAndTx"         => LayoutSemanticRole.Comparison,
            "twoObjOverTx"        => LayoutSemanticRole.Comparison,
            "objAndTwoObj"        => LayoutSemanticRole.Comparison,
            "twoObjAndObj"        => LayoutSemanticRole.Comparison,
            "cust"                => LayoutSemanticRole.Other,
            null                  => LayoutSemanticRole.Other,
            _                     => LayoutSemanticRole.Content,
        };
}
