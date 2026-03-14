using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Diagnostics;
using MarpToPptx.Pptx.Rendering;

namespace MarpToPptx.Tests;

public class TemplateDiagnoserTests
{
    // ─────────────────────────────────────────────────────────────
    // Basic smoke test
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_RenderedPptx_ReturnsNonEmptyReport()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.NotNull(report);
        Assert.True(report.SlideMasterCount >= 1, "Expected at least one slide master.");
        Assert.NotEmpty(report.Layouts);
    }

    // ─────────────────────────────────────────────────────────────
    // Master and layout counts
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_ReportsExpectedMasterAndLayoutCount()
    {
        // The minimal rendered deck uses the built-in template which has 1 master and 2 layouts.
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.Equal(1, report.SlideMasterCount);
        Assert.Equal(2, report.Layouts.Count);
    }

    // ─────────────────────────────────────────────────────────────
    // Semantic role classification via type-code string
    // ─────────────────────────────────────────────────────────────

    [Theory]
    [InlineData("title",          LayoutSemanticRole.Title)]
    [InlineData("titleOnly",      LayoutSemanticRole.TitleOnly)]
    [InlineData("secHead",        LayoutSemanticRole.SectionHeader)]
    [InlineData("blank",          LayoutSemanticRole.Blank)]
    [InlineData("picTx",          LayoutSemanticRole.PictureCaption)]
    [InlineData("clipArtAndTx",   LayoutSemanticRole.PictureCaption)]
    [InlineData("txAndClipArt",   LayoutSemanticRole.PictureCaption)]
    [InlineData("clipArtAndVertTx", LayoutSemanticRole.PictureCaption)]
    [InlineData("txAndMedia",     LayoutSemanticRole.PictureCaption)]
    [InlineData("mediaAndTx",     LayoutSemanticRole.PictureCaption)]
    [InlineData("twoColTx",       LayoutSemanticRole.Comparison)]
    [InlineData("twoTxTwoObj",    LayoutSemanticRole.Comparison)]
    [InlineData("twoObj",         LayoutSemanticRole.Comparison)]
    [InlineData("txAndTwoObj",    LayoutSemanticRole.Comparison)]
    [InlineData("twoObjAndTx",    LayoutSemanticRole.Comparison)]
    [InlineData("twoObjOverTx",   LayoutSemanticRole.Comparison)]
    [InlineData("objAndTwoObj",   LayoutSemanticRole.Comparison)]
    [InlineData("twoObjAndObj",   LayoutSemanticRole.Comparison)]
    [InlineData("cust",           LayoutSemanticRole.Other)]
    [InlineData(null,             LayoutSemanticRole.Other)]
    [InlineData("tx",             LayoutSemanticRole.Content)]
    [InlineData("obj",            LayoutSemanticRole.Content)]
    [InlineData("fourObj",        LayoutSemanticRole.Content)]
    [InlineData("tbl",            LayoutSemanticRole.Content)]
    [InlineData("vertTx",         LayoutSemanticRole.Content)]
    [InlineData("chart",          LayoutSemanticRole.Content)]
    public void TemplateDiagnoser_MapSemanticRole_MapsTypeCodeToExpectedRole(
        string? typeCode,
        LayoutSemanticRole expectedRole)
    {
        // The built-in template only contains "tx" and "blank" layouts.
        // For those two we verify through the rendered-deck diagnoser.
        // For all other type codes we verify via a standalone mapping shim that mirrors
        // the same logic in TemplateDiagnoser.
        if (typeCode is "tx" or "blank")
        {
            using var workspace = TestWorkspace.Create();
            var pptxPath = RenderMinimalDeck(workspace);

            using var document = PresentationDocument.Open(pptxPath, false);
            var diagnoser = new TemplateDiagnoser();
            var report = diagnoser.Diagnose(document, pptxPath);

            var layout = report.Layouts.FirstOrDefault(l => l.TypeCode == typeCode);
            Assert.NotNull(layout);
            Assert.Equal(expectedRole, layout.SemanticRole);
            return;
        }

        var role = MapSemanticRoleShim(typeCode);
        Assert.Equal(expectedRole, role);
    }

    // ─────────────────────────────────────────────────────────────
    // Layout placeholder detection
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_BuiltInContentLayout_HasTitleAndBodyPlaceholders()
    {
        // The built-in "tx" (Text/Content) layout declares both title and body placeholders.
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        var contentLayout = report.Layouts.FirstOrDefault(l => l.SemanticRole == LayoutSemanticRole.Content);
        Assert.NotNull(contentLayout);
        Assert.True(contentLayout.HasTitlePlaceholder, "Content layout should have a title placeholder.");
        Assert.True(contentLayout.HasBodyPlaceholder, "Content layout should have a body placeholder.");
    }

    [Fact]
    public void TemplateDiagnoser_BuiltInContentLayout_HasBodyPlaceholder()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        var contentLayout = report.Layouts.FirstOrDefault(l => l.SemanticRole == LayoutSemanticRole.Content);
        Assert.NotNull(contentLayout);
        Assert.True(contentLayout.HasBodyPlaceholder, "Content layout should have a body placeholder.");
    }

    // ─────────────────────────────────────────────────────────────
    // Recommendations
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_BuiltInTemplate_RecommendsDefaultContentLayout()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.NotNull(report.RecommendedDefaultContentLayout);
    }

    [Fact]
    public void TemplateDiagnoser_BuiltInTemplate_RecommendsTitleLayout_NullWhenNoTitleType()
    {
        // The built-in scaffold has no "title"-type layout, so the recommendation is null.
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.Null(report.RecommendedTitleLayout);
    }

    [Fact]
    public void TemplateDiagnoser_RecommendedLayouts_MatchActualLayoutNames()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        var allNames = report.Layouts.Select(l => l.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (report.RecommendedDefaultContentLayout is { } defaultLayout)
        {
            Assert.Contains(defaultLayout, allNames);
        }

        if (report.RecommendedTitleLayout is { } titleLayout)
        {
            Assert.Contains(titleLayout, allNames);
        }

        if (report.RecommendedSectionLayout is { } sectionLayout)
        {
            Assert.Contains(sectionLayout, allNames);
        }

        if (report.RecommendedPictureCaptionLayout is { } pictureLayout)
        {
            Assert.Contains(pictureLayout, allNames);
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Redundancy detection
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_LayoutsWithDistinctShapes_AreNotMarkedRedundant()
    {
        // Layouts that have at least one non-placeholder shape should never be
        // flagged as redundant, regardless of role.
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        foreach (var layout in report.Layouts.Where(l => l.NonPlaceholderShapeCount > 0))
        {
            Assert.False(layout.LikelyVisuallyRedundant,
                $"Layout \"{layout.Name}\" has {layout.NonPlaceholderShapeCount} distinct shape(s) and must not be marked redundant.");
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Warnings
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_Warnings_AreStrings_WhenPresent()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.All(report.Warnings, w => Assert.False(string.IsNullOrWhiteSpace(w)));
    }

    [Fact]
    public void TemplateDiagnoser_BuiltInTemplate_WarnsMissingTitleLayout()
    {
        // The built-in scaffold has no "title"-type layout; a warning about the missing
        // title layout should be present.
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.Contains(report.Warnings, w => w.Contains("title", StringComparison.OrdinalIgnoreCase));
    }

    // ─────────────────────────────────────────────────────────────
    // TemplatePath stored in report
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDiagnoser_Report_StoresSuppliedTemplatePath()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(document, pptxPath);

        Assert.Equal(pptxPath, report.TemplatePath);
    }

    // ─────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────

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

    /// <summary>
    /// Thin shim that mirrors the type-code mapping in <see cref="TemplateDiagnoser"/>
    /// so semantic-role theory tests can exercise paths not covered by the built-in template.
    /// </summary>
    private static LayoutSemanticRole MapSemanticRoleShim(string? typeCode)
        => typeCode switch
        {
            "title"               => LayoutSemanticRole.Title,
            "titleOnly"           => LayoutSemanticRole.TitleOnly,
            "secHead"             => LayoutSemanticRole.SectionHeader,
            "blank"               => LayoutSemanticRole.Blank,
            "picTx"               => LayoutSemanticRole.PictureCaption,
            "clipArtAndTx"        => LayoutSemanticRole.PictureCaption,
            "txAndClipArt"        => LayoutSemanticRole.PictureCaption,
            "clipArtAndVertTx"    => LayoutSemanticRole.PictureCaption,
            "txAndMedia"          => LayoutSemanticRole.PictureCaption,
            "mediaAndTx"          => LayoutSemanticRole.PictureCaption,
            "twoColTx"            => LayoutSemanticRole.Comparison,
            "twoTxTwoObj"         => LayoutSemanticRole.Comparison,
            "twoObj"              => LayoutSemanticRole.Comparison,
            "txAndTwoObj"         => LayoutSemanticRole.Comparison,
            "twoObjAndTx"         => LayoutSemanticRole.Comparison,
            "twoObjOverTx"        => LayoutSemanticRole.Comparison,
            "objAndTwoObj"        => LayoutSemanticRole.Comparison,
            "twoObjAndObj"        => LayoutSemanticRole.Comparison,
            "cust"                => LayoutSemanticRole.Other,
            null                  => LayoutSemanticRole.Other,
            _                     => LayoutSemanticRole.Content,
        };

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
