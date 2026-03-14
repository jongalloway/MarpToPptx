using MarpToPptx.Core.Themes;
using MarpToPptx.Pptx.Rendering;

namespace MarpToPptx.Tests;

public class DiagramThemeFactoryTests
{
    [Fact]
    public void Create_ProvidesDistinctNodePaletteEntries()
    {
        var theme = DiagramThemeFactory.Create(ThemeDefinition.Default);

        Assert.NotNull(theme.NodePalette);
        Assert.True(theme.NodePalette!.Count >= 6);
        Assert.True(theme.NodePalette.Distinct(StringComparer.OrdinalIgnoreCase).Count() >= 6);
    }

    [Fact]
    public void Create_UsesCodeSurfaceTextColorWhenCodeBackgroundExists()
    {
        var effectiveTheme = ThemeDefinition.Default with
        {
            FontFamily = "IBM Plex Sans",
            AccentColor = "#2563EB",
            Code = ThemeDefinition.Default.Code with
            {
                BackgroundColor = "#111827",
                Color = "#F8FAFC",
            },
        };

        var theme = DiagramThemeFactory.Create(effectiveTheme);

        Assert.Equal("IBM Plex Sans", theme.FontFamily);
        Assert.Equal("#F8FAFC", theme.TextColor);
        Assert.NotNull(theme.NodePalette);
        Assert.True(theme.NodePalette!.Distinct(StringComparer.OrdinalIgnoreCase).Count() >= 6);
    }

    [Fact]
    public void Create_HandlesRgbAccentColor()
    {
        var effectiveTheme = ThemeDefinition.Default with
        {
            AccentColor = "rgb(37, 99, 235)",
        };

        var theme = DiagramThemeFactory.Create(effectiveTheme);

        Assert.NotNull(theme.NodePalette);
        Assert.NotNull(theme.PrimaryColor);
        Assert.Equal("#2563EB", theme.PrimaryColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void Create_HandlesRgbaAccentColor_IgnoresAlpha()
    {
        var effectiveTheme = ThemeDefinition.Default with
        {
            AccentColor = "rgba(37, 99, 235, 0.8)",
        };

        var theme = DiagramThemeFactory.Create(effectiveTheme);

        Assert.NotNull(theme.NodePalette);
        Assert.NotNull(theme.PrimaryColor);
        Assert.Equal("#2563EB", theme.PrimaryColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void Create_HandlesRgbPercentageAccentColor()
    {
        // rgb(100%, 0%, 0%) = #FF0000
        var effectiveTheme = ThemeDefinition.Default with
        {
            AccentColor = "rgb(100%, 0%, 0%)",
        };

        var theme = DiagramThemeFactory.Create(effectiveTheme);

        Assert.NotNull(theme.PrimaryColor);
        Assert.Equal("#FF0000", theme.PrimaryColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void Create_HandlesRgbBackgroundColor()
    {
        var effectiveTheme = ThemeDefinition.Default with
        {
            BackgroundColor = "rgb(31, 41, 55)",
            Code = ThemeDefinition.Default.Code with { BackgroundColor = null },
        };

        var theme = DiagramThemeFactory.Create(effectiveTheme);

        Assert.NotNull(theme.BackgroundColor);
        Assert.Equal("#1F2937", theme.BackgroundColor, StringComparer.OrdinalIgnoreCase);
    }
}