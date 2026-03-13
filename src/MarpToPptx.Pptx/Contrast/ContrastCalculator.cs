namespace MarpToPptx.Pptx.Contrast;

/// <summary>
/// Computes WCAG 2.1 relative luminance and contrast ratios for RGB colors.
/// </summary>
internal static class ContrastCalculator
{
    /// <summary>
    /// Computes the WCAG 2.1 relative luminance of a 6-digit hex color (e.g. "FFFFFF").
    /// </summary>
    public static double RelativeLuminance(string hexColor)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hexColor);

        var hex = hexColor.TrimStart('#');
        if (hex.Length != 6)
            throw new ArgumentException($"Expected a 6-digit hex color, got: '{hexColor}'.", nameof(hexColor));

        var r = Convert.ToInt32(hex[0..2], 16) / 255.0;
        var g = Convert.ToInt32(hex[2..4], 16) / 255.0;
        var b = Convert.ToInt32(hex[4..6], 16) / 255.0;

        return 0.2126 * Linearize(r) + 0.7152 * Linearize(g) + 0.0722 * Linearize(b);
    }

    /// <summary>
    /// Computes the WCAG 2.1 contrast ratio between two 6-digit hex colors.
    /// Returns a value between 1 (no contrast) and 21 (maximum contrast).
    /// </summary>
    public static double ContrastRatio(string foreground, string background)
    {
        var l1 = RelativeLuminance(foreground);
        var l2 = RelativeLuminance(background);
        var lighter = Math.Max(l1, l2);
        var darker = Math.Min(l1, l2);
        return (lighter + 0.05) / (darker + 0.05);
    }

    private static double Linearize(double c)
        => c <= 0.04045 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
}
