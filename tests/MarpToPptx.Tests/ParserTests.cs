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

    [Fact]
    public void Parser_ExtractsBoldSpan_FromInlineBoldText()
    {
        const string markdown = """
        # Heading

        Normal **bold** text.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);

        var boldSpan = paragraph.Spans.FirstOrDefault(s => s.Bold);
        Assert.NotNull(boldSpan);
        Assert.Equal("bold", boldSpan!.Text);
    }

    [Fact]
    public void Parser_ExtractsItalicSpan_FromInlineItalicText()
    {
        const string markdown = """
        # Heading

        Normal *italic* text.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);

        var italicSpan = paragraph.Spans.FirstOrDefault(s => s.Italic);
        Assert.NotNull(italicSpan);
        Assert.Equal("italic", italicSpan!.Text);
    }

    [Fact]
    public void Parser_ExtractsStrikethroughSpan_FromInlineStrikethroughText()
    {
        const string markdown = """
        # Heading

        Normal ~~struck~~ text.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);

        var struckSpan = paragraph.Spans.FirstOrDefault(s => s.Strikethrough);
        Assert.NotNull(struckSpan);
        Assert.Equal("struck", struckSpan!.Text);
    }

    [Fact]
    public void Parser_ExtractsCodeSpan_FromInlineCode()
    {
        const string markdown = """
        # Heading

        Use `printf()` here.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);

        var codeSpan = paragraph.Spans.FirstOrDefault(s => s.Code);
        Assert.NotNull(codeSpan);
        Assert.Equal("printf()", codeSpan!.Text);
    }

    [Fact]
    public void Parser_ExtractsHyperlinkUrl_FromInlineLink()
    {
        const string markdown = """
        # Heading

        Visit [the site](https://example.com) for details.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);

        var linkSpan = paragraph.Spans.FirstOrDefault(s => s.HyperlinkUrl is not null);
        Assert.NotNull(linkSpan);
        Assert.Equal("the site", linkSpan!.Text);
        Assert.Equal("https://example.com", linkSpan.HyperlinkUrl);
    }

    [Fact]
    public void Parser_ExtractsNestedInlineFormatting_BoldItalic()
    {
        const string markdown = """
        # Heading

        ***bold and italic***
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);

        var span = paragraph.Spans.FirstOrDefault(s => s.Bold && s.Italic);
        Assert.NotNull(span);
        Assert.Equal("bold and italic", span!.Text);
    }

    [Fact]
    public void Parser_PreservesTextPropertyForBackwardCompat()
    {
        const string markdown = """
        # My **Title**

        Plain and **bold** text.

        - Item with *italic*
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var heading = Assert.IsType<HeadingElement>(slide.Elements[0]);
        Assert.Equal("My Title", heading.Text);

        var paragraph = Assert.IsType<ParagraphElement>(slide.Elements[1]);
        Assert.Equal("Plain and bold text.", paragraph.Text);

        var list = Assert.IsType<BulletListElement>(slide.Elements[2]);
        Assert.Equal("Item with italic", list.Items[0].Text);
    }

    [Fact]
    public void Parser_ExtractsInlineFormattingInBulletListItems()
    {
        const string markdown = """
        # Heading

        - Plain item
        - **Bold item**
        - *Italic item*
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var list = Assert.IsType<BulletListElement>(slide.Elements[1]);

        Assert.Equal(3, list.Items.Count);
        Assert.DoesNotContain(list.Items[0].Spans, s => s.Bold || s.Italic);
        Assert.Contains(list.Items[1].Spans, s => s.Bold);
        Assert.Contains(list.Items[2].Spans, s => s.Italic);
    }

    [Fact]
    public void Parser_ExtractsInlineFormattingInTableCells()
    {
        const string markdown = """
        # Table

        | Feature       | Status   |
        |---------------|----------|
        | **Bold cell** | *Italic* |
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var table = Assert.IsType<TableElement>(slide.Elements[1]);

        // Header row (row index 0) has plain text.
        var bodyRow = table.Rows[1];
        Assert.Contains(bodyRow.Cells[0], s => s.Bold);
        Assert.Contains(bodyRow.Cells[1], s => s.Italic);
    }

    [Fact]
    public void Parser_ParsesHeaderAndFooter_FromFrontMatter()
    {
        const string markdown = """
        ---
        header: My Header
        footer: My Footer
        ---

        # Slide
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Single(deck.Slides);
        Assert.Equal("My Header", deck.Slides[0].Style.Header);
        Assert.Equal("My Footer", deck.Slides[0].Style.Footer);
    }

    [Fact]
    public void Parser_ParsesHeaderAndFooter_FromSlideDirective()
    {
        const string markdown = """
        # Slide One

        ---

        <!-- header: Slide Two Header -->
        <!-- footer: Slide Two Footer -->
        # Slide Two
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(2, deck.Slides.Count);
        Assert.Null(deck.Slides[0].Style.Header);
        Assert.Null(deck.Slides[0].Style.Footer);
        Assert.Equal("Slide Two Header", deck.Slides[1].Style.Header);
        Assert.Equal("Slide Two Footer", deck.Slides[1].Style.Footer);
    }

    [Fact]
    public void Parser_ParsesAudioElement_FromImageSyntaxWithMp3Extension()
    {
        const string markdown = """
        # Slide

        ![Background music](intro.mp3)
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var audio = slide.Elements.OfType<AudioElement>().SingleOrDefault();
        Assert.NotNull(audio);
        Assert.Equal("intro.mp3", audio!.Source);
        Assert.Equal("Background music", audio.AltText);
    }

    [Fact]
    public void Parser_ParsesAudioElement_FromImageSyntaxWithWavExtension()
    {
        const string markdown = """
        # Slide

        ![Sound effect](effect.wav)
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var audio = slide.Elements.OfType<AudioElement>().SingleOrDefault();
        Assert.NotNull(audio);
        Assert.Equal("effect.wav", audio!.Source);
    }

    [Fact]
    public void Parser_ParsesAudioElement_FromHtmlAudioBlock()
    {
        const string markdown = """
        # Slide

        <audio src="background.mp3" controls></audio>
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        var slide = Assert.Single(deck.Slides);
        var audio = Assert.IsType<AudioElement>(slide.Elements[1]);
        Assert.Equal("background.mp3", audio.Source);
    }

    // ── Directive scoping and spot-directive tests ──────────────────────

    [Fact]
    public void Parser_LocalDirective_CarriesForwardToSubsequentSlides()
    {
        const string markdown = """
        ---
        paginate: true
        ---

        # Slide One

        ---

        <!-- class: lead -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Null(deck.Slides[0].Style.ClassName);
        Assert.Equal("lead", deck.Slides[1].Style.ClassName);
        // Local directive carries forward to slide 3.
        Assert.Equal("lead", deck.Slides[2].Style.ClassName);
        // Global front-matter directive persists through all slides.
        Assert.True(deck.Slides[2].Style.Paginate);
    }

    [Fact]
    public void Parser_SpotDirective_AppliesOnlyToCurrentSlide()
    {
        const string markdown = """
        # Slide One

        ---

        <!-- _class: lead -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Null(deck.Slides[0].Style.ClassName);
        // Spot directive applies to slide 2.
        Assert.Equal("lead", deck.Slides[1].Style.ClassName);
        // Spot directive does NOT carry forward.
        Assert.Null(deck.Slides[2].Style.ClassName);
    }

    [Fact]
    public void Parser_SpotDirective_OverridesLocalDirectiveForOneSlide()
    {
        const string markdown = """
        ---
        class: default-class
        ---

        # Slide One

        ---

        <!-- _class: override -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Equal("default-class", deck.Slides[0].Style.ClassName);
        // Spot overrides for this slide only.
        Assert.Equal("override", deck.Slides[1].Style.ClassName);
        // Reverts to carry-forward (front-matter default).
        Assert.Equal("default-class", deck.Slides[2].Style.ClassName);
    }

    [Fact]
    public void Parser_LocalHeaderFooter_CarryForward()
    {
        const string markdown = """
        # Slide One

        ---

        <!-- header: Page Header -->
        <!-- footer: Page Footer -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Null(deck.Slides[0].Style.Header);
        Assert.Null(deck.Slides[0].Style.Footer);
        Assert.Equal("Page Header", deck.Slides[1].Style.Header);
        Assert.Equal("Page Footer", deck.Slides[1].Style.Footer);
        // Carries forward.
        Assert.Equal("Page Header", deck.Slides[2].Style.Header);
        Assert.Equal("Page Footer", deck.Slides[2].Style.Footer);
    }

    [Fact]
    public void Parser_SpotPaginate_AppliesToSingleSlide()
    {
        const string markdown = """
        ---
        paginate: true
        ---

        # Slide One

        ---

        <!-- _paginate: false -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.True(deck.Slides[0].Style.Paginate);
        // Spot override turns off pagination for this slide only.
        Assert.False(deck.Slides[1].Style.Paginate);
        // Reverts to carry-forward (front-matter default).
        Assert.True(deck.Slides[2].Style.Paginate);
    }

    [Fact]
    public void Parser_MultipleLocalDirectives_AllCarryForward()
    {
        const string markdown = """
        # Slide One

        ---

        <!-- class: dark -->
        <!-- backgroundColor: #000 -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Equal("dark", deck.Slides[1].Style.ClassName);
        Assert.Equal("#000", deck.Slides[1].Style.BackgroundColor);
        // Both carry forward.
        Assert.Equal("dark", deck.Slides[2].Style.ClassName);
        Assert.Equal("#000", deck.Slides[2].Style.BackgroundColor);
    }

    [Fact]
    public void Parser_LocalThenSpotOnSameKey_CarriesLocalOnly()
    {
        const string markdown = """
        # Slide One

        ---

        <!-- class: base -->
        <!-- _class: spot-override -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        // Spot override wins for effective style on this slide.
        Assert.Equal("spot-override", deck.Slides[1].Style.ClassName);
        // Carry-forward uses the local directive value.
        Assert.Equal("base", deck.Slides[2].Style.ClassName);
    }

    [Fact]
    public void Parser_SpotBackgroundColor_DoesNotCarryForward()
    {
        const string markdown = """
        # Slide One

        ---

        <!-- _backgroundColor: red -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        Assert.Equal("red", deck.Slides[1].Style.BackgroundColor);
        Assert.Null(deck.Slides[2].Style.BackgroundColor);
    }

    [Fact]
    public void Parser_SpotDirective_WithEmptyKeyAfterStrip_IsIgnored()
    {
        // A directive like <!-- _: value --> has an empty key after stripping the '_' prefix
        // and should be silently ignored rather than inserted with an empty key.
        const string markdown = """
        # Slide One

        ---

        <!-- _: ignored-value -->
        # Slide Two

        ---

        # Slide Three
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal(3, deck.Slides.Count);
        // No empty-string key should appear in any slide's directive map.
        Assert.DoesNotContain("", deck.Slides[0].Style.Directives.Keys);
        Assert.DoesNotContain("", deck.Slides[1].Style.Directives.Keys);
        Assert.DoesNotContain("", deck.Slides[2].Style.Directives.Keys);
    }

    // ────────────────────────────────────────────────────────
    // Issue #38 – Global directives: headingDivider, lang, style
    // ────────────────────────────────────────────────────────

    [Fact]
    public void Lang_SetsLanguageOnSlideDeck()
    {
        const string markdown = """
        ---
        marp: true
        lang: ja
        ---

        # Slide 1
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Equal("ja", deck.Language);
    }

    [Fact]
    public void Lang_DefaultsToNull_WhenNotSpecified()
    {
        const string markdown = """
        ---
        marp: true
        ---

        # Slide 1
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Null(deck.Language);
    }

    [Fact]
    public void Lang_IsNull_WhenEmptyValue()
    {
        const string markdown = """
        ---
        marp: true
        lang:
        ---

        # Slide 1
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        Assert.Null(deck.Language);
    }

    [Fact]
    public void Style_MergesInlineCssWithThemeCss()
    {
        const string themeCss = "section { background: red; }";
        const string markdown = """
        ---
        marp: true
        style: |
          h1 { color: blue; }
        ---

        # Slide 1
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown, themeCss: themeCss);

        // Both the theme CSS (background: red on section) and the inline CSS
        // (h1 { color: blue }) should have been applied to the ThemeDefinition.
        Assert.Equal("red", deck.Theme.BackgroundColor);
        Assert.Equal("blue", deck.Theme.Headings[1].Color);
    }

    [Fact]
    public void Style_UsedAlone_WhenNoThemeCss()
    {
        const string markdown = """
        ---
        marp: true
        style: |
          section { font-family: "Comic Sans MS"; }
        ---

        # Slide 1
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        // Inline CSS alone should override the default font family.
        Assert.Equal("Comic Sans MS", deck.Theme.FontFamily);
    }

    [Fact]
    public void HeadingDivider_SplitsSlidesOnHeading()
    {
        const string markdown = """
        ---
        marp: true
        headingDivider: 2
        ---

        # Part One

        Intro text.

        ## Section A

        Content A.

        ## Section B

        Content B.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        // headingDivider: 2 should split on ## headings, producing 3 slides:
        // 1. "# Part One" + "Intro text."
        // 2. "## Section A" + "Content A."
        // 3. "## Section B" + "Content B."
        Assert.Equal(3, deck.Slides.Count);
    }

    [Fact]
    public void HeadingDivider_IgnoresLowerLevelHeadings()
    {
        const string markdown = """
        ---
        marp: true
        headingDivider: 1
        ---

        # First Slide

        ### This is H3 (below divider level)

        More content.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        // headingDivider: 1 splits only on # headings; ### should not split.
        Assert.Single(deck.Slides);
    }

    [Fact]
    public void HeadingDivider_WorksWithExplicitSeparators()
    {
        const string markdown = """
        ---
        marp: true
        headingDivider: 2
        ---

        # Slide One

        ---

        # Slide Two

        ## Sub-section

        Content here.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        // Explicit --- creates slide 1→2 break.
        // headingDivider: 2 creates slide 2→3 break at ## Sub-section.
        Assert.Equal(3, deck.Slides.Count);
    }

    [Fact]
    public void HeadingDivider_InvalidValue_IsIgnored()
    {
        const string markdown = """
        ---
        marp: true
        headingDivider: abc
        ---

        # Slide One

        ## Section

        Content.
        """;

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown);

        // Invalid value should be ignored — no heading-based splitting.
        // Only the initial slide exists (no --- separators).
        Assert.Single(deck.Slides);
    }

    // ────────────────────────────────────────────────────────
    // Issue #39 – Theme class variant parsing
    // ────────────────────────────────────────────────────────

    [Fact]
    public void ThemeParser_ParsesSectionClassSelector()
    {
        const string css = """
        section { color: #000000; }
        section.lead { background-color: #102A43; color: #FFFFFF; }
        """;

        var theme = MarpThemeParser.Parse(css);

        Assert.True(theme.ClassVariants.ContainsKey("lead"));
        var variant = theme.ClassVariants["lead"];
        Assert.Equal("#102A43", variant.BackgroundColor);
        Assert.Equal("#FFFFFF", variant.Body?.Color);
    }

    [Fact]
    public void ThemeParser_ParsesSectionClassHeadingSelector()
    {
        const string css = """
        section.invert h1 { color: #FFDD57; font-size: 40px; }
        """;

        var theme = MarpThemeParser.Parse(css);

        Assert.True(theme.ClassVariants.ContainsKey("invert"));
        var variant = theme.ClassVariants["invert"];
        Assert.NotNull(variant.Headings);
        Assert.True(variant.Headings!.ContainsKey(1));
        Assert.Equal("#FFDD57", variant.Headings[1].Color);
        Assert.Equal(30, variant.Headings[1].FontSize); // 40px * 0.75 = 30pt
    }

    [Fact]
    public void ThemeParser_MultipleClassVariants_ParsedIndependently()
    {
        const string css = """
        section.lead { background-color: #1A1A2E; }
        section.invert { background-color: #16213E; color: #E94560; }
        """;

        var theme = MarpThemeParser.Parse(css);

        Assert.Equal(2, theme.ClassVariants.Count);
        Assert.Equal("#1A1A2E", theme.ClassVariants["lead"].BackgroundColor);
        Assert.Equal("#16213E", theme.ClassVariants["invert"].BackgroundColor);
        Assert.Equal("#E94560", theme.ClassVariants["invert"].Body?.Color);
        Assert.Equal("#1F2937", theme.ClassVariants["lead"].Body?.Color); // no color override; inherits base theme body color
    }

    [Fact]
    public void ThemeParser_NoClassVariants_WhenNone()
    {
        const string css = "section { color: #333; }";

        var theme = MarpThemeParser.Parse(css);

        Assert.Empty(theme.ClassVariants);
    }
}
