using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core.Layout;
using MarpToPptx.Core.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Rendering;

/// <summary>
/// Represents the semantic content type of a slide, used to select the most
/// appropriate layout from a template.
/// </summary>
internal enum SlideKind
{
    /// <summary>A title slide: first element is an H1 heading with at most one supporting paragraph.</summary>
    Title,

    /// <summary>A standard content slide with headings and body elements.</summary>
    Content,

    /// <summary>A slide where images make up the majority of the non-heading content.</summary>
    ImageFocused,
}

/// <summary>
/// Classifies slides by semantic content and selects the best matching
/// <see cref="SlideLayoutPart"/> from the available template layouts.
/// Also exposes helpers for reading placeholder bounds from a layout.
/// </summary>
internal sealed class SlideTemplateSelector
{
    private const int LayoutScale = 12700;

    private readonly IReadOnlyList<SlideLayoutPart> _layouts;

    public SlideTemplateSelector(IReadOnlyList<SlideLayoutPart> layouts)
    {
        _layouts = layouts;
    }

    /// <summary>
    /// Classifies a slide based on its content elements.
    /// </summary>
    public static SlideKind Classify(Slide slide)
    {
        if (slide.Elements.Count == 0)
        {
            return SlideKind.Content;
        }

        // Title slide: starts with H1 and has at most one additional paragraph element.
        var isH1First = slide.Elements[0] is HeadingElement { Level: 1 };
        if (isH1First && slide.Elements.Count <= 2 &&
            (slide.Elements.Count == 1 || slide.Elements[1] is ParagraphElement))
        {
            return SlideKind.Title;
        }

        // Image-focused slide: images constitute at least half of non-heading content.
        var imageCount = slide.Elements.OfType<ImageElement>().Count();
        var nonHeadingCount = slide.Elements.Count(e => e is not HeadingElement);
        if (nonHeadingCount > 0 && imageCount > 0 &&
            (double)imageCount / nonHeadingCount >= 0.5)
        {
            return SlideKind.ImageFocused;
        }

        return SlideKind.Content;
    }

    /// <summary>
    /// Returns the most appropriate <see cref="SlideLayoutPart"/> for the given <paramref name="kind"/>.
    /// Falls back to the first available layout when no exact match is found.
    /// </summary>
    public SlideLayoutPart SelectLayout(SlideKind kind)
        => kind switch
        {
            SlideKind.Title =>
                FindLayout(P.SlideLayoutValues.Title) ??
                FindLayout(P.SlideLayoutValues.Text) ??
                _layouts[0],

            SlideKind.ImageFocused =>
                FindLayout(P.SlideLayoutValues.Blank) ??
                _layouts[0],

            _ =>
                FindLayout(P.SlideLayoutValues.Text) ??
                _layouts[0],
        };

    /// <summary>
    /// Returns the bounding rectangle (in layout units) of the title placeholder in the
    /// given layout, or <c>null</c> if the placeholder has no explicit transform.
    /// </summary>
    public static Rect? GetTitlePlaceholderRect(SlideLayoutPart layoutPart)
        => GetPlaceholderRect(layoutPart, P.PlaceholderValues.Title, P.PlaceholderValues.CenteredTitle);

    /// <summary>
    /// Returns the bounding rectangle (in layout units) of the body placeholder in the
    /// given layout, or <c>null</c> if the placeholder has no explicit transform.
    /// </summary>
    public static Rect? GetBodyPlaceholderRect(SlideLayoutPart layoutPart)
        => GetPlaceholderRect(layoutPart, P.PlaceholderValues.Body, P.PlaceholderValues.SubTitle);

    private SlideLayoutPart? FindLayout(P.SlideLayoutValues targetType)
    {
        foreach (var layoutPart in _layouts)
        {
            if (layoutPart.SlideLayout?.Type?.Value == targetType)
            {
                return layoutPart;
            }
        }

        return null;
    }

    private static Rect? GetPlaceholderRect(SlideLayoutPart layoutPart, params P.PlaceholderValues[] types)
    {
        var shapeTree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree;
        if (shapeTree is null)
        {
            return null;
        }

        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            var ph = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();
            if (ph is null)
            {
                continue;
            }

            var phType = ph.Type?.Value;
            if (phType is null || !types.Any(t => phType == t))
            {
                continue;
            }

            var xfrm = shape.ShapeProperties?.Transform2D;
            if (xfrm?.Offset is null || xfrm.Extents is null)
            {
                continue;
            }

            var x = xfrm.Offset.X?.Value ?? 0L;
            var y = xfrm.Offset.Y?.Value ?? 0L;
            var cx = xfrm.Extents.Cx?.Value ?? 0L;
            var cy = xfrm.Extents.Cy?.Value ?? 0L;

            if (cx <= 0 || cy <= 0)
            {
                continue;
            }

            return new Rect(
                (double)x / LayoutScale,
                (double)y / LayoutScale,
                (double)cx / LayoutScale,
                (double)cy / LayoutScale);
        }

        return null;
    }
}
