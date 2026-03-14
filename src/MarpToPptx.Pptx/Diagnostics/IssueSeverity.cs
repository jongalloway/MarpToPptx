namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Severity level of a <see cref="TemplateDoctorIssue"/> reported by <see cref="TemplateDoctor"/>.
/// </summary>
public enum IssueSeverity
{
    /// <summary>Informational finding that does not degrade MarpToPptx output.</summary>
    Info,

    /// <summary>Potential problem that may produce unexpected or degraded output.</summary>
    Warning,

    /// <summary>Issue that can be automatically repaired by the template doctor write-back.</summary>
    Fixable,
}
