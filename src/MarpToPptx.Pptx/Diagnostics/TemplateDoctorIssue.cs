namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Represents a single structural issue or observation found by <see cref="TemplateDoctor"/>
/// for a specific layout (or the template as a whole).
/// </summary>
/// <param name="LayoutName">
/// The name of the layout that triggered the issue, or <c>null</c> for template-wide issues.
/// </param>
/// <param name="Severity">
/// How serious the issue is: <see cref="IssueSeverity.Info"/>,
/// <see cref="IssueSeverity.Warning"/>, or <see cref="IssueSeverity.Fixable"/>.
/// </param>
/// <param name="Code">
/// A short machine-readable identifier for the issue type (e.g. <c>DuplicateLayoutName</c>).
/// </param>
/// <param name="Description">Human-readable description of the issue.</param>
/// <param name="ProposedFix">
/// Description of the automated repair that would be applied in write mode,
/// or <c>null</c> if the issue is informational only.
/// </param>
public sealed record TemplateDoctorIssue(
    string? LayoutName,
    IssueSeverity Severity,
    string Code,
    string Description,
    string? ProposedFix = null);
