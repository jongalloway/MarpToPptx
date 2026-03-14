using DiagramForge.Models;
using MarpToPptx.Core.Themes;

namespace MarpToPptx.Pptx.Rendering;

internal static class DiagramThemeFactory
{
    private const string DefaultAccentColor = "#0F766E";
    private const string DefaultBackgroundColor = "#FFFFFF";
    private const string DefaultLightTextColor = "#F8FAFC";
    private const string DefaultDarkTextColor = "#1F2937";

    internal static Theme Create(ThemeDefinition effectiveTheme)
    {
        var hasCodeBackground = !string.IsNullOrWhiteSpace(effectiveTheme.Code.BackgroundColor);
        var backgroundColor = NormalizeHexColor(hasCodeBackground
            ? effectiveTheme.Code.BackgroundColor
            : effectiveTheme.BackgroundColor, DefaultBackgroundColor);

        var primaryAccent = NormalizeHexColor(
            string.IsNullOrWhiteSpace(effectiveTheme.AccentColor)
                ? effectiveTheme.GetHeadingStyle(2).Color
                : effectiveTheme.AccentColor,
            DefaultAccentColor);

        var isLightBackground = IsLight(backgroundColor);
        var secondaryAccent = RotateHue(primaryAccent, isLightBackground ? 34 : 30, saturationDelta: 0.02, lightnessDelta: isLightBackground ? -0.02 : 0.06);
        var tertiaryAccent = RotateHue(primaryAccent, isLightBackground ? -42 : -38, saturationDelta: 0.05, lightnessDelta: isLightBackground ? 0.01 : 0.08);

        var textColor = NormalizeHexColor(
            hasCodeBackground ? effectiveTheme.Code.Color : effectiveTheme.Body.Color,
            isLightBackground ? DefaultDarkTextColor : DefaultLightTextColor);

        var theme = Theme.FromPalette(
            primaryColor: primaryAccent,
            secondaryColor: secondaryAccent,
            accentColor: tertiaryAccent,
            backgroundColor: backgroundColor);

        theme.TextColor = textColor;
        theme.FontFamily = effectiveTheme.FontFamily;
        theme.BorderRadius = 12;
        return theme;
    }

    private static string NormalizeHexColor(string? value, string fallback)
    {
        if (TryParseRgb(value, out var red, out var green, out var blue))
        {
            return FormatHex(red, green, blue);
        }

        return TryParseRgb(fallback, out red, out green, out blue)
            ? FormatHex(red, green, blue)
            : "#000000";
    }

    private static bool TryParseRgb(string? value, out int red, out int green, out int blue)
    {
        red = 0;
        green = 0;
        blue = 0;

        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var trimmed = value.Trim().Trim('"', '\'');
        if (trimmed.StartsWith('#'))
        {
            trimmed = trimmed[1..];
        }

        if (trimmed.Length == 3)
        {
            trimmed = string.Concat(trimmed.Select(ch => new string(ch, 2)));
        }

        if (trimmed.Length < 6)
        {
            return false;
        }

        trimmed = trimmed[..6];
        return int.TryParse(trimmed[..2], System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out red)
            && int.TryParse(trimmed[2..4], System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out green)
            && int.TryParse(trimmed[4..6], System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out blue);
    }

    private static string RotateHue(string hexColor, double degrees, double saturationDelta, double lightnessDelta)
    {
        if (!TryParseRgb(hexColor, out var red, out var green, out var blue))
        {
            return hexColor;
        }

        var (hue, saturation, lightness) = RgbToHsl(red, green, blue);
        hue = NormalizeHue(hue + degrees);
        saturation = Clamp01(saturation + saturationDelta);
        lightness = Clamp01(lightness + lightnessDelta);

        var (rotatedRed, rotatedGreen, rotatedBlue) = HslToRgb(hue, saturation, lightness);
        return FormatHex(rotatedRed, rotatedGreen, rotatedBlue);
    }

    private static bool IsLight(string hexColor)
    {
        if (!TryParseRgb(hexColor, out var red, out var green, out var blue))
        {
            return true;
        }

        var luminance = (0.2126 * red) + (0.7152 * green) + (0.0722 * blue);
        return luminance >= 150;
    }

    private static (double Hue, double Saturation, double Lightness) RgbToHsl(int red, int green, int blue)
    {
        var r = red / 255d;
        var g = green / 255d;
        var b = blue / 255d;

        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var lightness = (max + min) / 2d;

        if (Math.Abs(max - min) < double.Epsilon)
        {
            return (0, 0, lightness);
        }

        var delta = max - min;
        var saturation = lightness > 0.5
            ? delta / (2d - max - min)
            : delta / (max + min);

        double hue;
        if (Math.Abs(max - r) < double.Epsilon)
        {
            hue = ((g - b) / delta) + (g < b ? 6 : 0);
        }
        else if (Math.Abs(max - g) < double.Epsilon)
        {
            hue = ((b - r) / delta) + 2;
        }
        else
        {
            hue = ((r - g) / delta) + 4;
        }

        hue *= 60;
        return (hue, saturation, lightness);
    }

    private static (int Red, int Green, int Blue) HslToRgb(double hue, double saturation, double lightness)
    {
        if (saturation <= 0)
        {
            var gray = (int)Math.Round(lightness * 255);
            return (gray, gray, gray);
        }

        var q = lightness < 0.5
            ? lightness * (1 + saturation)
            : lightness + saturation - (lightness * saturation);
        var p = (2 * lightness) - q;
        var hk = hue / 360d;

        var red = HueToRgb(p, q, hk + (1d / 3d));
        var green = HueToRgb(p, q, hk);
        var blue = HueToRgb(p, q, hk - (1d / 3d));

        return ((int)Math.Round(red * 255), (int)Math.Round(green * 255), (int)Math.Round(blue * 255));
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0)
        {
            t += 1;
        }

        if (t > 1)
        {
            t -= 1;
        }

        if (t < 1d / 6d)
        {
            return p + ((q - p) * 6 * t);
        }

        if (t < 0.5)
        {
            return q;
        }

        if (t < 2d / 3d)
        {
            return p + ((q - p) * ((2d / 3d) - t) * 6);
        }

        return p;
    }

    private static double NormalizeHue(double hue)
    {
        var normalized = hue % 360;
        return normalized < 0 ? normalized + 360 : normalized;
    }

    private static double Clamp01(double value) => Math.Clamp(value, 0d, 1d);

    private static string FormatHex(int red, int green, int blue)
        => $"#{red:X2}{green:X2}{blue:X2}";
}