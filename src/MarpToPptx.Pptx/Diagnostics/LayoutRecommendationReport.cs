namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Layout recommendation for a single slide in a deck.
/// </summary>
/// <param name="SlideNumber">1-based slide number.</param>
/// <param name="SlideTitle">Title text extracted from the slide's H1, or <c>null</c> if absent.</param>
/// <param name="ContentKind">Heuristic content classification of the slide.</param>
/// <param name="RecommendedLayout">Name of the suggested template layout to use for this slide.</param>
/// <param name="Reason">Short human-readable hint explaining why the layout was chosen, or <c>null</c>.</param>
/// <param name="IsExplicitLayout">
/// <see langword="true"/> when the slide already carries an explicit <c>_layout:</c> directive;
/// the <see cref="RecommendedLayout"/> value is the directive's value, not a computed suggestion.
/// </param>
public sealed record SlideRecommendation(
    int SlideNumber,
    string? SlideTitle,
    SlideContentKind ContentKind,
    string RecommendedLayout,
    string? Reason,
    bool IsExplicitLayout = false);

/// <summary>
/// Full layout recommendation report produced by <see cref="LayoutRecommender"/>
/// for a Marp deck analysed against a PPTX template.
/// </summary>
/// <param name="DeckPath">Path of the analysed Markdown deck, or a display label when not file-backed.</param>
/// <param name="TemplatePath">Path of the PPTX template used for recommendations.</param>
/// <param name="Recommendations">Per-slide recommendation records.</param>
/// <param name="SuggestedFrontMatterLayout">
/// The layout name that is recommended as the global <c>layout:</c> front-matter directive
/// (the most-used content layout across all slides), or <c>null</c> if no clear winner exists.
/// </param>
/// <param name="PhotoLayoutRotation">
/// Ordered list of all photo/picture-focused layout names available in the template.
/// When there are multiple photo layouts, slides with images will rotate through these variants.
/// Empty when the template has no picture-focused layouts.
/// </param>
public sealed record LayoutRecommendationReport(
    string DeckPath,
    string TemplatePath,
    IReadOnlyList<SlideRecommendation> Recommendations,
    string? SuggestedFrontMatterLayout,
    IReadOnlyList<string> PhotoLayoutRotation);
