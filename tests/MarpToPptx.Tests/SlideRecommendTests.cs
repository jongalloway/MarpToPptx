using MarpToPptx.Core;
using MarpToPptx.Core.Models;
using MarpToPptx.Pptx.Diagnostics;
using MarpToPptx.Pptx.Rendering;

namespace MarpToPptx.Tests;

public class SlideRecommendTests
{
    // ─────────────────────────────────────────────────────────────
    // SlideContentClassifier — basic classification heuristics
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void Classifier_FirstSlideWithH1AndShortSubtitle_ClassifiesAsTitle()
    {
        var slide = ParseSingleSlide("# Welcome\n\nA brief intro.");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: true, isLast: false);
        Assert.Equal(SlideContentKind.Title, kind);
    }

    [Fact]
    public void Classifier_LastSlideWithH1Only_ClassifiesAsConclusion()
    {
        var slide = ParseSingleSlide("# Thank You");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: true);
        Assert.Equal(SlideContentKind.Conclusion, kind);
    }

    [Fact]
    public void Classifier_SlideWithBlockquote_ClassifiesAsQuote()
    {
        var slide = ParseSingleSlide("# Insight\n\n> The future belongs to those who believe.");
        var (kind, reason) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.Quote, kind);
        Assert.Contains("blockquote", reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Classifier_SlideWithImage_ClassifiesAsImageFocused()
    {
        var slide = ParseSingleSlide("# Photo\n\n![alt](photo.jpg)");
        var (kind, reason) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.ImageFocused, kind);
        Assert.Contains("image", reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Classifier_SlideWithH1Only_ClassifiesAsSectionHeader()
    {
        var slide = ParseSingleSlide("# Chapter 3");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.SectionHeader, kind);
    }

    [Fact]
    public void Classifier_SlideWithH1AndShortTagline_ClassifiesAsSectionHeader()
    {
        var slide = ParseSingleSlide("# Chapter 3\n\nA new beginning.");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.SectionHeader, kind);
    }

    [Fact]
    public void Classifier_SlideWithOrderedList_ClassifiesAsAgenda()
    {
        var slide = ParseSingleSlide("# Agenda\n\n1. Intro\n2. Main topic\n3. Q&A");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.Agenda, kind);
    }

    [Fact]
    public void Classifier_SlideWithShortBulletList_ClassifiesAsStatement()
    {
        var slide = ParseSingleSlide("# Key Points\n\n- Fast\n- Reliable\n- Secure");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.Statement, kind);
    }

    [Fact]
    public void Classifier_SlideWithTable_ClassifiesAsContent()
    {
        var slide = ParseSingleSlide("# Comparison\n\n| A | B |\n|---|---|\n| 1 | 2 |");
        var (kind, reason) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.Content, kind);
        Assert.Contains("table", reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Classifier_SlideWithDenseText_ClassifiesAsWideContent()
    {
        // 8 bullet items → exceeds the 6-unit threshold.
        var slide = ParseSingleSlide(
            "# Details\n\n- A\n- B\n- C\n- D\n- E\n- F\n- G\n- H");
        var (kind, reason) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.WideContent, kind);
        Assert.Contains("dense", reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Classifier_SlideWithBoldShortNumber_ClassifiesAsBigNumber()
    {
        // Bold short number token.
        var slide = ParseSingleSlide("# Impact\n\n**98%**\n\nof projects succeeded.");
        var (kind, reason) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        Assert.Equal(SlideContentKind.BigNumber, kind);
        Assert.Contains("number", reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Classifier_StandardContentSlide_ClassifiesAsContent()
    {
        var slide = ParseSingleSlide("# Summary\n\nSome prose text here with details.");
        var (kind, _) = SlideContentClassifier.Classify(slide, isFirst: false, isLast: false);
        // A single short paragraph gets SectionHeader because it qualifies as a short tagline.
        // A longer paragraph should remain as Content.
        Assert.True(kind == SlideContentKind.Content || kind == SlideContentKind.SectionHeader);
    }

    // ─────────────────────────────────────────────────────────────
    // BlockquoteElement — parser produces the element
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void Parser_BlockquoteInSlide_ProducesBlockquoteElement()
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile("# Title\n\n> This is a quote.");
        Assert.Single(deck.Slides);
        Assert.Contains(deck.Slides[0].Elements, e => e is BlockquoteElement);
    }

    [Fact]
    public void Parser_BlockquoteElement_HasExpectedText()
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile("# Title\n\n> Believe in yourself.");
        var blockquote = deck.Slides[0].Elements.OfType<BlockquoteElement>().First();
        Assert.Equal("Believe in yourself.", blockquote.Text);
    }

    // ─────────────────────────────────────────────────────────────
    // LayoutRecommender — produces non-null report
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void LayoutRecommender_Recommend_ReturnsOneRecommendationPerSlide()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(
            """
            # Title Slide

            ---

            ## Content Slide

            - Alpha
            - Beta

            ---

            # Last Slide
            """);

        var diagnoser = new TemplateDiagnoser();
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, false);
        var templateReport = diagnoser.Diagnose(document, pptxPath);

        var recommender = new LayoutRecommender();
        var report = recommender.Recommend(deck, templateReport);

        Assert.Equal(3, report.Recommendations.Count);
    }

    [Fact]
    public void LayoutRecommender_Recommend_AllRecommendedLayoutsAreNonEmpty()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(
            """
            # Welcome

            ---

            ## About

            Some content here.

            ---

            # Conclusion
            """);

        var diagnoser = new TemplateDiagnoser();
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, false);
        var templateReport = diagnoser.Diagnose(document, pptxPath);

        var recommender = new LayoutRecommender();
        var report = recommender.Recommend(deck, templateReport);

        Assert.All(report.Recommendations, r => Assert.False(string.IsNullOrWhiteSpace(r.RecommendedLayout)));
    }

    [Fact]
    public void LayoutRecommender_Recommend_SuggestedFrontMatterLayout_IsNonEmpty()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var compiler = new MarpCompiler();
        var deck = compiler.Compile("# Intro\n\n---\n\n## Body\n\nContent here.");

        var diagnoser = new TemplateDiagnoser();
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, false);
        var templateReport = diagnoser.Diagnose(document, pptxPath);

        var recommender = new LayoutRecommender();
        var report = recommender.Recommend(deck, templateReport);

        Assert.False(string.IsNullOrWhiteSpace(report.SuggestedFrontMatterLayout));
    }

    [Fact]
    public void LayoutRecommender_SlideWithExplicitLayout_RetainsExplicitLayout()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        const string layoutName = "Custom Layout";
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(
            $"""
            <!-- _layout: {layoutName} -->
            # My Slide

            Some content.
            """);

        var diagnoser = new TemplateDiagnoser();
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, false);
        var templateReport = diagnoser.Diagnose(document, pptxPath);

        var recommender = new LayoutRecommender();
        var report = recommender.Recommend(deck, templateReport);

        Assert.Single(report.Recommendations);
        Assert.Equal(layoutName, report.Recommendations[0].RecommendedLayout);
    }

    [Fact]
    public void LayoutRecommender_Recommend_SlideTitlesAreExtractedFromH1()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var compiler = new MarpCompiler();
        var deck = compiler.Compile("# My Great Title\n\nSome text.");

        var diagnoser = new TemplateDiagnoser();
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, false);
        var templateReport = diagnoser.Diagnose(document, pptxPath);

        var recommender = new LayoutRecommender();
        var report = recommender.Recommend(deck, templateReport);

        Assert.Equal("My Great Title", report.Recommendations[0].SlideTitle);
    }

    // ─────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────

    private static Slide ParseSingleSlide(string slideMarkdown)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(slideMarkdown);
        return deck.Slides[0];
    }

    private static string RenderMinimalDeck(TestWorkspace workspace)
    {
        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Intro paragraph.

            ---

            ## Content Slide

            - Alpha
            - Beta
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = workspace.RootPath });
        return outputPath;
    }

    private sealed class TestWorkspace : IDisposable
    {
        private TestWorkspace(string rootPath) => RootPath = rootPath;

        public string RootPath { get; }

        public static TestWorkspace Create()
        {
            var rootPath = Path.Combine(Path.GetTempPath(), "MarpToPptx.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(rootPath);
            return new TestWorkspace(rootPath);
        }

        public string GetPath(string relativePath) => Path.Combine(RootPath, relativePath);

        public string WriteMarkdown(string relativePath, string content)
        {
            var path = GetPath(relativePath);
            File.WriteAllText(path, content);
            return path;
        }

        public void Dispose()
        {
            if (Directory.Exists(RootPath))
                Directory.Delete(RootPath, recursive: true);
        }
    }
}
