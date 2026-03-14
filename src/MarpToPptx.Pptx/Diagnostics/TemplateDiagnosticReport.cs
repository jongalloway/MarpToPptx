namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Top-level diagnostic report produced by <see cref="TemplateDiagnoser"/> for a PPTX template.
/// </summary>
/// <param name="TemplatePath">Path of the inspected template file as supplied to the diagnoser.</param>
/// <param name="SlideMasterCount">Number of slide masters in the template.</param>
/// <param name="Layouts">
/// Per-layout diagnostic records, ordered as they appear across masters in the template.
/// </param>
/// <param name="RecommendedDefaultContentLayout">
/// Name of the layout best suited for the front-matter <c>layout:</c> directive on normal content slides,
/// or <c>null</c> if no suitable layout was found.
/// </param>
/// <param name="RecommendedTitleLayout">
/// Name of the layout best suited for title/cover slides (<c>_layout:</c> on the first slide),
/// or <c>null</c> if no title layout was found.
/// </param>
/// <param name="RecommendedSectionLayout">
/// Name of the layout best suited for section-header slides (<c>_layout:</c> on section dividers),
/// or <c>null</c> if no section-header layout was found.
/// </param>
/// <param name="RecommendedPictureCaptionLayout">
/// Name of the layout best suited for image/picture slides (<c>_layout:</c>),
/// or <c>null</c> if the template has no picture-focused layout.
/// </param>
/// <param name="Warnings">Human-readable warning strings describing potential issues.</param>
public sealed record TemplateDiagnosticReport(
    string TemplatePath,
    int SlideMasterCount,
    IReadOnlyList<LayoutDiagnostic> Layouts,
    string? RecommendedDefaultContentLayout,
    string? RecommendedTitleLayout,
    string? RecommendedSectionLayout,
    string? RecommendedPictureCaptionLayout,
    IReadOnlyList<string> Warnings);
