using System.Text.RegularExpressions;

namespace MarpToPptx.Core.Parsing;

/// <summary>
/// Parses Marpit-compatible image sizing tokens from an image alt-text string.
/// </summary>
/// <remarks>
/// Supported Marpit sizing forms:
/// <list type="bullet">
///   <item><c>w:200px</c> — explicit width in CSS pixels</item>
///   <item><c>h:150px</c> — explicit height in CSS pixels</item>
///   <item><c>50%</c> — percentage of slide width</item>
/// </list>
/// Sizing tokens are stripped from the alt text; any remaining text becomes the
/// accessible description. Units other than <c>px</c> are not currently supported
/// and are treated as plain alt-text tokens.
/// </remarks>
internal static partial class MarpitImageSizingParser
{
    // Conversion: 1 CSS px = 0.75 pt. The renderer's layout coordinate space uses
    // points as its unit (1 layout unit = 1/72 inch = 12700 EMU).
    private const double PxToLayoutUnits = 0.75;

    // Matches w:Npx or h:Npx tokens (case-insensitive, whole word).
    [GeneratedRegex(@"(?<!\S)([wh]):(\d+(?:\.\d+)?)px(?!\S)", RegexOptions.IgnoreCase)]
    private static partial Regex DimensionPattern();

    // Matches a standalone N% token (e.g. "50%" or "100%").
    [GeneratedRegex(@"(?<!\S)(\d+(?:\.\d+)?)%(?!\S)")]
    private static partial Regex PercentPattern();

    /// <summary>
    /// Parses Marpit image sizing tokens from <paramref name="altText"/> and returns the
    /// parsed sizing values along with the alt text with sizing tokens removed.
    /// </summary>
    /// <param name="altText">Raw alt-text string from the Markdown image (e.g. <c>"w:200px My photo"</c>).</param>
    /// <returns>
    /// A tuple of:
    /// <list type="bullet">
    ///   <item><c>ExplicitWidth</c> — layout units, or <c>null</c> if not specified.</item>
    ///   <item><c>ExplicitHeight</c> — layout units, or <c>null</c> if not specified.</item>
    ///   <item><c>SizePercent</c> — 0–100 value, or <c>null</c> if not specified.</item>
    ///   <item><c>CleanAltText</c> — alt text with all recognised sizing tokens removed.</item>
    /// </list>
    /// </returns>
    public static (double? ExplicitWidth, double? ExplicitHeight, double? SizePercent, string CleanAltText) Parse(string altText)
    {
        if (string.IsNullOrWhiteSpace(altText))
        {
            return (null, null, null, altText ?? string.Empty);
        }

        double? explicitWidth = null;
        double? explicitHeight = null;
        double? sizePercent = null;

        var cleaned = altText;

        // Process w:/h: dimension tokens.
        cleaned = DimensionPattern().Replace(cleaned, match =>
        {
            var axis = match.Groups[1].Value.ToLowerInvariant();
            if (double.TryParse(match.Groups[2].Value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var value))
            {
                var layoutUnits = value * PxToLayoutUnits;
                if (axis == "w")
                {
                    explicitWidth = layoutUnits;
                }
                else
                {
                    explicitHeight = layoutUnits;
                }
            }

            return string.Empty;
        });

        // Process percentage tokens:
        // - Always strip them from the alt text (even when w:/h: directives are present).
        // - Only apply SizePercent when no explicit px dimensions are present.
        cleaned = PercentPattern().Replace(cleaned, match =>
        {
            if (explicitWidth is null && explicitHeight is null &&
                double.TryParse(match.Groups[1].Value,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var pct))
            {
                sizePercent = Math.Clamp(pct, 0.0, 100.0);
            }

            return string.Empty;
        });

        // Normalise whitespace left by removed tokens.
        cleaned = string.Join(' ', cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries));

        return (explicitWidth, explicitHeight, sizePercent, cleaned);
    }
}
