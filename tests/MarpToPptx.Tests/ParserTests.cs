using MarpToPptx.Core;
using MarpToPptx.Core.Models;
using MarpToPptx.Core.Parsing;
using MarpToPptx.Core.Themes;

namespace MarpToPptx.Tests;

public class ParserTests
{
    [Fact]
    public void SplitSlides_IgnoresSeparatorsInsideCodeFences()
    {
        const string markdown = """
        # Slide 1

        ```md
        ---
        ```

        ---

        # Slide 2
        """;

        var slides = SlideTokenizer.SplitSlides(markdown);

        Assert.Equal(2, slides.Count);
    }

    [Fact]
    public void MarpCompiler_ParsesFrontMatterDirectivesAndElements()
    {
        const string markdown = """
        ---
        theme: gaia
        paginate: true
        ---

        <!-- class: lead -->
        # Hello

        Welcome to the deck.

        - One
        - Two
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal("gaia", deck.Theme.Name);
        Assert.Single(deck.Slides);
        Assert.True(deck.Slides[0].Style.Paginate);
        Assert.Equal("lead", deck.Slides[0].Style.ClassName);
        Assert.Collection(
            deck.Slides[0].Elements,
            element => Assert.IsType<HeadingElement>(element),
            element => Assert.IsType<ParagraphElement>(element),
            element => Assert.IsType<BulletListElement>(element));
    }

    [Fact]
    public void MarpThemeParser_ExtractsCoreThemeValues()
    {
        const string css = """
        section {
          font-family: "IBM Plex Sans", sans-serif;
          font-size: 32px;
          color: #112233;
          background-color: #faf8f2;
          padding: 40px 56px;
        }

        h1 {
          font-size: 56px;
          color: #334455;
        }

        code {
          font-family: "IBM Plex Mono", monospace;
          color: #f5f5f5;
          background: #101820;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "editorial");

        Assert.Equal("editorial", theme.Name);
        Assert.Equal("IBM Plex Sans", theme.FontFamily);
        Assert.Equal("#112233", theme.TextColor);
        Assert.Equal("#faf8f2", theme.BackgroundColor);
        Assert.Equal(40 * 0.75, theme.SlidePadding.Top);
        Assert.Equal(56 * 0.75, theme.SlidePadding.Right);
        Assert.Equal("IBM Plex Mono", theme.Code.FontFamily);
        Assert.Equal("#101820", theme.Code.BackgroundColor);
        Assert.Equal(56 * 0.75, theme.Headings[1].FontSize);
    }
}
