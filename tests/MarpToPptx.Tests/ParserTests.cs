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

    [Fact]
    public void MarpThemeParser_ParsesBackgroundImageFromSection()
    {
        const string css = """
        section {
          background-image: url("slides-bg.jpg");
          background-size: cover;
          background-position: center;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "bgtest");

        Assert.Equal("slides-bg.jpg", theme.BackgroundImage);
        Assert.Equal("cover", theme.BackgroundSize);
        Assert.Equal("center", theme.BackgroundPosition);
    }

    [Fact]
    public void MarpThemeParser_ParsesBackgroundImageFromBackgroundShorthand()
    {
        const string css = """
        section {
          background: url('hero.png') center/cover no-repeat #1a1a2e;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "bgshorthand");

        Assert.Equal("hero.png", theme.BackgroundImage);
        Assert.Equal("#1a1a2e", theme.BackgroundColor);
    }

    [Fact]
    public void MarpThemeParser_ParsesLineHeightAndLetterSpacingForBody()
    {
        const string css = """
        section {
          font-size: 24px;
          line-height: 1.6;
          letter-spacing: 1px;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "spacing");

        Assert.Equal(1.6, theme.Body.LineHeight);
        Assert.Equal(1 * 0.75, theme.Body.LetterSpacing);
    }

    [Fact]
    public void MarpThemeParser_ParsesLineHeightAsPercentage()
    {
        const string css = """
        section {
          line-height: 150%;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "lhpct");

        Assert.Equal(1.5, theme.Body.LineHeight);
    }

    [Fact]
    public void MarpThemeParser_ParsesTextTransformAndFontWeightForBody()
    {
        const string css = """
        section {
          font-weight: normal;
          text-transform: uppercase;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "transform");

        Assert.False(theme.Body.Bold);
        Assert.Equal("uppercase", theme.Body.TextTransform);
    }

    [Fact]
    public void MarpThemeParser_ParsesHeadingFontWeightAndTypography()
    {
        const string css = """
        h1 {
          font-weight: normal;
          line-height: 1.2;
          letter-spacing: 2px;
          text-transform: uppercase;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "headingtype");

        Assert.False(theme.Headings[1].Bold);
        Assert.Equal(1.2, theme.Headings[1].LineHeight);
        Assert.Equal(2 * 0.75, theme.Headings[1].LetterSpacing);
        Assert.Equal("uppercase", theme.Headings[1].TextTransform);
    }

    [Fact]
    public void MarpThemeParser_ParsesCodeLineHeightAndLetterSpacing()
    {
        const string css = """
        code {
          font-family: "Fira Code", monospace;
          font-size: 16px;
          line-height: 1.5;
          letter-spacing: 0px;
          background-color: #282c34;
          color: #abb2bf;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "codeblock");

        Assert.Equal("Fira Code", theme.Code.FontFamily);
        Assert.Equal(1.5, theme.Code.LineHeight);
        Assert.Equal(0.0, theme.Code.LetterSpacing);
        Assert.Equal("#282c34", theme.Code.BackgroundColor);
        Assert.Equal("#abb2bf", theme.Code.Color);
    }

    [Fact]
    public void MarpThemeParser_ParsesFontWeightNumericBold()
    {
        const string css = """
        h2 {
          font-weight: 700;
        }

        h3 {
          font-weight: 400;
        }
        """;

        var theme = MarpThemeParser.Parse(css, "numericweight");

        Assert.True(theme.Headings[2].Bold);
        Assert.False(theme.Headings[3].Bold);
    }

    [Fact]
    public void Parser_ParsesHtmlVideoInlineTag_AsVideoElement()
    {
        const string markdown = """
        # Slide

        <video src="clip.mp4"></video>
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        Assert.Collection(
            slide.Elements,
            e => Assert.IsType<HeadingElement>(e),
            e =>
            {
                var video = Assert.IsType<VideoElement>(e);
                Assert.Equal("clip.mp4", video.Source);
            });
    }

    [Fact]
    public void Parser_ParsesHtmlVideoSelfClosingTag_AsVideoElement()
    {
        const string markdown = "# Slide\n\n<video src=\"clip.mp4\" />\n\ntext";

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var videoElement = slide.Elements.OfType<VideoElement>().SingleOrDefault();
        Assert.NotNull(videoElement);
        Assert.Equal("clip.mp4", videoElement!.Source);
    }

    [Fact]
    public void Parser_ParsesHtmlVideoTag_WithExtraAttributes()
    {
        const string markdown = """
        # Slide

        <video controls width="640" src="demo.mp4" height="360"></video>
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var video = slide.Elements.OfType<VideoElement>().SingleOrDefault();
        Assert.NotNull(video);
        Assert.Equal("demo.mp4", video!.Source);
    }

    [Fact]
    public void Parser_ParsesMarkdownImageWithMp4Extension_AsVideoElement()
    {
        const string markdown = """
        # Slide

        ![My clip](video.mp4)
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var video = slide.Elements.OfType<VideoElement>().SingleOrDefault();
        Assert.NotNull(video);
        Assert.Equal("video.mp4", video!.Source);
        Assert.Equal("My clip", video.AltText);
    }

    [Fact]
    public void Parser_DoesNotTreatNonMp4ImageAsVideo()
    {
        const string markdown = """
        # Slide

        ![Photo](photo.png)
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        Assert.Empty(slide.Elements.OfType<VideoElement>());
        Assert.Single(slide.Elements.OfType<ImageElement>());
    }

    [Fact]
    public void Parser_ExtractsNotesFromNonDirectiveHtmlComment()
    {
        const string markdown = """
        # Slide

        Content here.

        <!-- This is a presenter note. -->
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        Assert.Equal("This is a presenter note.", slide.Notes);
    }

    [Fact]
    public void Parser_ExcludesDirectiveCommentsFromNotes()
    {
        const string markdown = """
        <!-- class: lead -->
        # Slide

        Content here.

        <!-- This is a note. -->
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        Assert.Equal("lead", slide.Style.ClassName);
        Assert.Equal("This is a note.", slide.Notes);
    }

    [Fact]
    public void Parser_CombinesMultipleNoteCommentsWithNewline()
    {
        const string markdown = """
        # Slide

        <!-- First note. -->

        Content.

        <!-- Second note. -->
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        Assert.Equal("First note.\nSecond note.", slide.Notes);
    }

    [Fact]
    public void Parser_NoteCommentIsNotEmittedAsSlideElement()
    {
        const string markdown = """
        # Slide

        <!-- Presenter note text. -->
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        Assert.Single(slide.Elements);
        Assert.IsType<HeadingElement>(slide.Elements[0]);
        Assert.Equal("Presenter note text.", slide.Notes);
    }

    [Fact]
    public void Parser_AssignsNotesToCorrectSlides_InMultiSlideDeck()
    {
        const string markdown = """
        # Slide One

        <!-- Note for slide one. -->

        ---

        # Slide Two

        No notes here.

        ---

        # Slide Three

        <!-- Note for slide three. -->
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Equal("Note for slide one.", deck.Slides[0].Notes);
        Assert.Null(deck.Slides[1].Notes);
        Assert.Equal("Note for slide three.", deck.Slides[2].Notes);
    }
}
