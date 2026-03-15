namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Top-level report produced by <see cref="TemplateDoctor"/> after analyzing (and optionally
/// repairing) a PPTX template.
/// </summary>
/// <param name="TemplatePath">Path of the template that was inspected.</param>
/// <param name="Issues">
/// All structural issues and observations found, in the order the checks were run
/// (template-wide checks first, then per-layout checks in layout order).
/// </param>
/// <param name="WroteFixedTemplate">
/// <see langword="true"/> when the doctor wrote a repaired copy of the template.
/// </param>
/// <param name="FixedTemplatePath">
/// Path of the written repaired template, or <c>null</c> if no file was written.
/// </param>
/// <param name="AppliedFixes">
/// Human-readable descriptions of each repair that was applied to the fixed template,
/// or an empty list when <paramref name="WroteFixedTemplate"/> is <see langword="false"/>.
/// </param>
public sealed record TemplateDoctorReport(
    string TemplatePath,
    IReadOnlyList<TemplateDoctorIssue> Issues,
    bool WroteFixedTemplate,
    string? FixedTemplatePath,
    IReadOnlyList<string> AppliedFixes);
