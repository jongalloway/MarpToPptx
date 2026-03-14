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
}