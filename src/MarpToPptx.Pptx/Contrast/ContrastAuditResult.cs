namespace MarpToPptx.Pptx.Contrast;

/// <summary>
/// Represents the result of a contrast audit for a single text/background color pair found in a PPTX slide.
/// </summary>
public sealed record ContrastAuditResult(
    int SlideNumber,
    string ShapeContext,
    string ForegroundColor,
    string BackgroundColor,
    double ContrastRatio,
    bool IsLargeText)
{
    /// <summary>
    /// The minimum WCAG 2.1 contrast ratio required for this text size.
    /// Large text requires 3:1; normal text requires 4.5:1.
    /// </summary>
    public double MinimumRequiredRatio => IsLargeText ? 3.0 : 4.5;

    /// <summary>
    /// Returns true if the contrast ratio is below the minimum required for the text size.
    /// </summary>
    public bool IsFailing => ContrastRatio < MinimumRequiredRatio;
}
