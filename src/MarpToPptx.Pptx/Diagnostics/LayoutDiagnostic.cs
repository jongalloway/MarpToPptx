namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Diagnostic information for a single slide layout in a PPTX template.
/// </summary>
/// <param name="Name">Display name of the layout (from <c>MatchingName</c> or <c>CommonSlideData.Name</c>).</param>
/// <param name="TypeCode">
/// The raw OpenXML layout type value (e.g. <c>tx</c>, <c>blank</c>, <c>title</c>),
/// or <c>null</c> for custom or unset layouts.
/// </param>
/// <param name="SemanticRole">Semantic role inferred from the layout type and placeholder structure.</param>
/// <param name="HasTitlePlaceholder">Whether the layout declares a title-type placeholder.</param>
/// <param name="HasBodyPlaceholder">Whether the layout declares a body-type (content) placeholder.</param>
/// <param name="HasPicturePlaceholder">Whether the layout declares a picture placeholder.</param>
/// <param name="NonPlaceholderShapeCount">
/// Number of non-placeholder shapes in the layout's own shape tree (master shapes are not counted).
/// A count of zero means the layout's visual design is inherited entirely from its slide master.
/// </param>
/// <param name="LikelyVisuallyRedundant">
/// <see langword="true"/> when this layout has no distinct non-placeholder shapes and shares its
/// semantic role with at least one other layout, meaning the layouts will likely appear visually
/// identical with the current MarpToPptx renderer until placeholder-based rendering is expanded.
/// </param>
public sealed record LayoutDiagnostic(
    string Name,
    string? TypeCode,
    LayoutSemanticRole SemanticRole,
    bool HasTitlePlaceholder,
    bool HasBodyPlaceholder,
    bool HasPicturePlaceholder,
    int NonPlaceholderShapeCount,
    bool LikelyVisuallyRedundant);
