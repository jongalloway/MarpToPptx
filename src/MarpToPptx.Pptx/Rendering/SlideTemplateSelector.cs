using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core.Layout;
using MarpToPptx.Core.Models;
using System.Xml.Linq;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Rendering;

internal sealed record SelectedSlideLayout(SlideLayoutPart LayoutPart, bool UseTemplateStyle);

/// <summary>
/// Identity of a placeholder shape on a template layout, captured by the
/// <c>&lt;p:ph&gt;</c> element's <c>type</c> and optional <c>idx</c>. A slide-level
/// placeholder shape with the same type+index inherits geometry and text
/// styling from the layout and master.
///
/// A <c>null</c> <see cref="Type"/> represents a typeless placeholder
/// (<c>&lt;p:ph idx="..."/&gt;</c>), which is the conventional form of an object/content
/// placeholder on <c>obj</c>-type layouts such as "Title and Content". The slide-level
/// echo must also omit the <c>type</c> attribute for inheritance to match.
/// </summary>
internal sealed record TemplatePlaceholder(P.PlaceholderValues? Type, uint? Index);

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

    /// <summary>A slide where images make up at least half of the non-heading content.</summary>
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

        // Image-focused slide: images, videos, and audio elements constitute at least half of non-heading content.
        var mediaCount = slide.Elements.OfType<ImageElement>().Count() +
                         slide.Elements.OfType<VideoElement>().Count() +
                         slide.Elements.OfType<AudioElement>().Count();
        var nonHeadingCount = slide.Elements.Count(e => e is not HeadingElement);
        if (nonHeadingCount > 0 && mediaCount > 0 &&
            (double)mediaCount / nonHeadingCount >= 0.5)
        {
            return SlideKind.ImageFocused;
        }

        return SlideKind.Content;
    }

    /// <summary>
    /// Returns the most appropriate <see cref="SlideLayoutPart"/> for the given slide.
    /// A matching named layout enables template-first visual styling for that slide.
    /// </summary>
    public SelectedSlideLayout SelectLayout(Slide slide, SlideKind kind, string? defaultContentLayout)
    {
        var requestedLayout = ResolveRequestedLayout(slide, kind, defaultContentLayout);
        if (!string.IsNullOrWhiteSpace(requestedLayout))
        {
            var namedLayout = FindLayoutByName(requestedLayout);
            if (namedLayout is not null)
            {
                return new SelectedSlideLayout(namedLayout, UseTemplateStyle: true);
            }
        }

        return new SelectedSlideLayout(SelectAutoLayout(kind), UseTemplateStyle: false);
    }

    private SlideLayoutPart SelectAutoLayout(SlideKind kind)
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

    /// <summary>
    /// Returns the title placeholder identity (<c>type</c>+<c>idx</c>) declared on the
    /// given layout, or <c>null</c> if no title-like placeholder exists. Unlike
    /// <see cref="GetTitlePlaceholderRect"/>, this does not require an explicit
    /// transform; layouts commonly inherit placeholder geometry from the master.
    /// </summary>
    public static TemplatePlaceholder? GetTitlePlaceholder(SlideLayoutPart layoutPart)
        => FindPlaceholder(layoutPart, P.PlaceholderValues.Title, P.PlaceholderValues.CenteredTitle);

    /// <summary>
    /// Returns the body placeholder identity (<c>type</c>+<c>idx</c>) declared on the
    /// given layout, or <c>null</c> if no body-like placeholder exists.
    ///
    /// Falls back to the first typeless indexed placeholder (<c>&lt;p:ph idx="..."/&gt;</c>),
    /// which is the conventional content placeholder on <c>obj</c> layouts such as
    /// "Title and Content" and "Two Content".
    /// </summary>
    public static TemplatePlaceholder? GetBodyPlaceholder(SlideLayoutPart layoutPart)
        => FindPlaceholder(layoutPart, P.PlaceholderValues.Body, P.PlaceholderValues.SubTitle)
            ?? FindTypelessIndexedPlaceholder(layoutPart);

    private static TemplatePlaceholder? FindPlaceholder(SlideLayoutPart layoutPart, params P.PlaceholderValues[] types)
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
            if (ph?.Type?.Value is not { } phType || !types.Any(t => phType == t))
            {
                continue;
            }

            return new TemplatePlaceholder(phType, ph.Index?.Value);
        }

        return null;
    }

    private static TemplatePlaceholder? FindTypelessIndexedPlaceholder(SlideLayoutPart layoutPart)
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
            // A <p:ph/> with no type attribute and a non-null idx is the generic
            // object/content placeholder. Typed placeholders (dt, ftr, sldNum, title, ...)
            // are excluded here because their Type is non-null.
            if (ph is null || ph.Type is not null || ph.Index?.Value is not { } idx)
            {
                continue;
            }

            return new TemplatePlaceholder(Type: null, Index: idx);
        }

        return null;
    }

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

    private SlideLayoutPart? FindLayoutByName(string requestedLayout)
    {
        foreach (var layoutPart in _layouts)
        {
            foreach (var candidate in GetLayoutNames(layoutPart))
            {
                if (string.Equals(candidate, requestedLayout.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return layoutPart;
                }
            }
        }

        return null;
    }

    private static string? ResolveRequestedLayout(Slide slide, SlideKind kind, string? defaultContentLayout)
    {
        if (!string.IsNullOrWhiteSpace(slide.Style.Layout))
        {
            return slide.Style.Layout;
        }

        return kind == SlideKind.Content && !string.IsNullOrWhiteSpace(defaultContentLayout)
            ? defaultContentLayout
            : null;
    }

    private static IEnumerable<string> GetLayoutNames(SlideLayoutPart layoutPart)
    {
        var slideLayout = layoutPart.SlideLayout;
        if (slideLayout is null)
        {
            yield break;
        }

        var document = XDocument.Parse(slideLayout.OuterXml);
        var root = document.Root;
        if (root is null)
        {
            yield break;
        }

        var matchingName = root.Attribute("matchingName")?.Value;
        if (!string.IsNullOrWhiteSpace(matchingName))
        {
            yield return matchingName;
        }

        XNamespace presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var commonSlideDataName = root.Element(presentationNamespace + "cSld")?.Attribute("name")?.Value;
        if (!string.IsNullOrWhiteSpace(commonSlideDataName))
        {
            yield return commonSlideDataName;
        }
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
