using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core.Layout;
using MarpToPptx.Core.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Rendering;

/// <summary>
/// A reference to a pre-cloned template slide part, created before <c>ClearSlides</c>
/// removes the original slide's <c>&lt;p:sldId&gt;</c> XML reference. The pre-clone
/// is a proper package member whose sub-parts (images, etc.) remain alive even after
/// the original template slide and its parts are destroyed by the SDK.
/// </summary>
/// <param name="SlideNumber">1-based position of the template slide in the original presentation.</param>
/// <param name="SlidePart">The pre-cloned <see cref="SlidePart"/>; alive for the lifetime of the rendering operation.</param>
/// <param name="LayoutPart">The <see cref="SlideLayoutPart"/> captured from the pre-clone; safe to access at any time.</param>
internal sealed record TemplateSlideReference(int SlideNumber, SlidePart SlidePart, SlideLayoutPart? LayoutPart);

internal sealed record SelectedSlideLayout(SlideLayoutPart LayoutPart, bool UseTemplateStyle, TemplateSlideReference? TemplateSlide = null);

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
    private readonly IReadOnlyList<TemplateSlideReference> _templateSlides;

    public SlideTemplateSelector(IReadOnlyList<SlideLayoutPart> layouts, IReadOnlyList<SlidePart>? templateSlides = null)
    {
        _layouts = layouts;
        // The caller (OpenXmlPptxRenderer.Render) passes pre-cloned SlidePart objects that
        // were created before ClearSlides ran. Eagerly capture LayoutPart here so
        // SelectLayout never needs to access it lazily through the SlidePart reference.
        _templateSlides = templateSlides?.Select(
            (sp, i) => new TemplateSlideReference(i + 1, sp, sp.SlideLayoutPart)).ToArray() ?? [];
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
            if (TryResolveTemplateSlide(requestedLayout, out var templateSlide) &&
                templateSlide.LayoutPart is { } templateSlideLayout)
            {
                return new SelectedSlideLayout(templateSlideLayout, UseTemplateStyle: true, templateSlide);
            }

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
    /// given layout, or <c>null</c> if neither the layout nor its slide master exposes
    /// a usable transform for the inherited placeholder.
    /// </summary>
    public static Rect? GetTitlePlaceholderRect(SlideLayoutPart layoutPart)
        => GetPlaceholderRect(
            layoutPart,
            GetTitlePlaceholder(layoutPart),
            P.PlaceholderValues.Title,
            P.PlaceholderValues.CenteredTitle);

    /// <summary>
    /// Returns the bounding rectangle (in layout units) of the body placeholder in the
    /// given layout, or <c>null</c> if neither the layout nor its slide master exposes
    /// a usable transform for the inherited placeholder.
    /// </summary>
    public static Rect? GetBodyPlaceholderRect(SlideLayoutPart layoutPart)
        => GetBodyPlaceholderRect(layoutPart, GetBodyPlaceholder(layoutPart));

    /// <summary>
    /// Returns the bounding rectangle (in layout units) of the body placeholder in the
    /// given layout using an already-resolved <see cref="TemplatePlaceholder"/>, avoiding
    /// a second traversal of the shape tree. Returns <c>null</c> if the placeholder is
    /// <c>null</c> or no usable transform is found on the layout or its slide master.
    /// </summary>
    public static Rect? GetBodyPlaceholderRect(SlideLayoutPart layoutPart, TemplatePlaceholder? bodyPlaceholder)
        => GetPlaceholderRect(
            layoutPart,
            bodyPlaceholder,
            P.PlaceholderValues.Body,
            P.PlaceholderValues.SubTitle);

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

    /// <summary>
    /// Returns the picture placeholder identity (<c>type</c>+<c>idx</c>) declared on the
    /// given layout, or <c>null</c> if no picture placeholder exists.
    /// </summary>
    public static TemplatePlaceholder? GetPicturePlaceholder(SlideLayoutPart layoutPart)
        => FindPlaceholder(layoutPart, P.PlaceholderValues.Picture);

    /// <summary>
    /// Returns the bounding rectangle (in layout units) of the picture placeholder in the
    /// given layout, or <c>null</c> if neither the layout nor its slide master exposes
    /// a usable transform for the inherited placeholder.
    /// </summary>
    public static Rect? GetPicturePlaceholderRect(SlideLayoutPart layoutPart)
        => GetPicturePlaceholderRect(layoutPart, GetPicturePlaceholder(layoutPart));

    /// <summary>
    /// Returns the bounding rectangle (in layout units) of the picture placeholder in the
    /// given layout using an already-resolved <see cref="TemplatePlaceholder"/>, avoiding
    /// a second traversal of the shape tree. Returns <c>null</c> if the placeholder is
    /// <c>null</c> or no usable transform is found on the layout or its slide master.
    /// </summary>
    public static Rect? GetPicturePlaceholderRect(SlideLayoutPart layoutPart, TemplatePlaceholder? picturePlaceholder)
        => GetPlaceholderRect(
            layoutPart,
            picturePlaceholder,
            P.PlaceholderValues.Picture);

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

    private bool TryResolveTemplateSlide(string requestedLayout, out TemplateSlideReference templateSlide)
    {
        templateSlide = null!;

        if (!TryParseTemplateSlideNumber(requestedLayout, out var slideNumber) ||
            slideNumber < 1 || slideNumber > _templateSlides.Count)
        {
            return false;
        }

        templateSlide = _templateSlides[slideNumber - 1];
        return true;
    }

    private static bool TryParseTemplateSlideNumber(string requestedLayout, out int slideNumber)
    {
        slideNumber = 0;
        var trimmed = requestedLayout.Trim();

        if (trimmed.StartsWith("Template[", StringComparison.OrdinalIgnoreCase) &&
            trimmed.EndsWith("]", StringComparison.Ordinal))
        {
            return int.TryParse(trimmed[9..^1].Trim(), out slideNumber);
        }

        const string templateSlidePrefix = "Template Slide ";
        if (trimmed.StartsWith(templateSlidePrefix, StringComparison.OrdinalIgnoreCase))
        {
            return int.TryParse(trimmed[templateSlidePrefix.Length..].Trim(), out slideNumber);
        }

        return false;
    }

    private static IEnumerable<string> GetLayoutNames(SlideLayoutPart layoutPart)
    {
        var slideLayout = layoutPart.SlideLayout;
        if (slideLayout is null)
        {
            yield break;
        }

        var matchingName = slideLayout.MatchingName?.Value;
        if (!string.IsNullOrWhiteSpace(matchingName))
        {
            yield return matchingName;
        }

        var commonSlideDataName = slideLayout.CommonSlideData?.Name?.Value;
        if (!string.IsNullOrWhiteSpace(commonSlideDataName))
        {
            yield return commonSlideDataName;
        }
    }

    private static Rect? GetPlaceholderRect(
        SlideLayoutPart layoutPart,
        TemplatePlaceholder? placeholder,
        params P.PlaceholderValues[] masterFallbackTypes)
    {
        if (placeholder is null)
        {
            return null;
        }

        return GetPlaceholderRect(
                layoutPart.SlideLayout?.CommonSlideData?.ShapeTree,
                placeholder,
                masterFallbackTypes)
            ?? GetPlaceholderRect(
                layoutPart.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree,
                placeholder,
                masterFallbackTypes);
    }

    private static Rect? GetPlaceholderRect(
        P.ShapeTree? shapeTree,
        TemplatePlaceholder placeholder,
        params P.PlaceholderValues[] masterFallbackTypes)
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
            if (ph is null || !PlaceholderMatches(ph, placeholder, masterFallbackTypes))
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

    private static bool PlaceholderMatches(
        P.PlaceholderShape placeholderShape,
        TemplatePlaceholder placeholder,
        params P.PlaceholderValues[] masterFallbackTypes)
    {
        if (placeholder.Index is not null && placeholderShape.Index?.Value != placeholder.Index)
        {
            return false;
        }

        if (placeholder.Type is { } type)
        {
            return placeholderShape.Type?.Value == type;
        }

        var actualType = placeholderShape.Type?.Value;
        return actualType is null || masterFallbackTypes.Any(candidate => candidate == actualType);
    }
}
