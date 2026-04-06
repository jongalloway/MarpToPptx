using System.Text.Json;
using MarpToPptx.Cli.Mcp;

namespace MarpToPptx.Tests;

public class McpToolsTests
{
    private static readonly string TwoSlideDeck = """
        # Intro Slide

        ---

        # Details Slide
        """;

    // ── marp_render ────────────────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_Render_GeneratesOutputFile()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var outputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        var result = await tools.marp_render(markdownPath, outputPath);

        Assert.Contains(outputPath, result);
        Assert.Contains("2 slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task McpTools_Render_DefaultsOutputPathToSameDirectoryAsInput()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var expectedOutputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        var result = await tools.marp_render(markdownPath);

        Assert.Contains(expectedOutputPath, result);
        Assert.True(File.Exists(expectedOutputPath));
    }

    [Fact]
    public async Task McpTools_Render_WithWriteSlideIds_AddsSlideIdDirectives()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var outputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        var result = await tools.marp_render(markdownPath, outputPath, writeSlideIds: true);

        Assert.Contains("Added 2 slideId directive(s)", result);
        Assert.True(File.Exists(outputPath));

        var rewrittenMarkdown = File.ReadAllText(markdownPath);
        Assert.Contains("<!-- slideId:", rewrittenMarkdown);
    }

    [Fact]
    public async Task McpTools_Render_ThrowsFileNotFound_ForMissingInput()
    {
        using var workspace = TestWorkspace.Create();
        var missingPath = workspace.GetPath("nonexistent.md");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<FileNotFoundException>(() => tools.marp_render(missingPath));
    }

    // ── marp_inspect ───────────────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_Inspect_ReturnsStructuredMetadata()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);

        var tools = new MarpToPptxTools();
        var json = await tools.marp_inspect(markdownPath);

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(2, root.GetProperty("slideCount").GetInt32());
        Assert.Equal(2, root.GetProperty("slides").GetArrayLength());
        Assert.Equal("Intro Slide", root.GetProperty("slides")[0].GetProperty("title").GetString());
    }

    [Fact]
    public async Task McpTools_Inspect_ThrowsFileNotFound_ForMissingInput()
    {
        using var workspace = TestWorkspace.Create();
        var missingPath = workspace.GetPath("nonexistent.md");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<FileNotFoundException>(() => tools.marp_inspect(missingPath));
    }

    // ── marp_write_slide_ids ───────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_WriteSlideIds_AddsMissingDirectives()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);

        var tools = new MarpToPptxTools();
        var result = await tools.marp_write_slide_ids(markdownPath);

        Assert.Contains("Added 2 slideId directive(s)", result);

        var updatedMarkdown = File.ReadAllText(markdownPath);
        Assert.Contains("<!-- slideId: intro-slide -->", updatedMarkdown);
        Assert.Contains("<!-- slideId: details-slide -->", updatedMarkdown);
    }

    [Fact]
    public async Task McpTools_WriteSlideIds_ReportsNoChanges_WhenAllSlidesAlreadyHaveIds()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", """
            <!-- slideId: intro-slide -->
            # Intro Slide

            ---

            <!-- slideId: details-slide -->
            # Details Slide
            """);

        var tools = new MarpToPptxTools();
        var result = await tools.marp_write_slide_ids(markdownPath);

        Assert.Contains("All slides already have slideId directives", result);
    }

    [Fact]
    public async Task McpTools_WriteSlideIds_ThrowsFileNotFound_ForMissingInput()
    {
        using var workspace = TestWorkspace.Create();
        var missingPath = workspace.GetPath("nonexistent.md");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<FileNotFoundException>(() => tools.marp_write_slide_ids(missingPath));
    }

    // ── marp_update_deck ───────────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_UpdateDeck_UpdatesExistingPptx()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var outputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        await tools.marp_render(markdownPath, outputPath, writeSlideIds: true);

        var updatedMarkdown = File.ReadAllText(markdownPath);
        File.WriteAllText(markdownPath, updatedMarkdown + "\n\n---\n\n# Third Slide\n");

        var result = await tools.marp_update_deck(markdownPath, outputPath, outputPath);

        Assert.Contains(outputPath, result);
        Assert.Contains("Updated '", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task McpTools_UpdateDeck_ThrowsFileNotFound_ForMissingExistingDeck()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var missingDeckPath = workspace.GetPath("nonexistent.pptx");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            tools.marp_update_deck(markdownPath, missingDeckPath));
    }

    // ── marp_render_string ─────────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_RenderString_GeneratesOutputFile()
    {
        using var workspace = TestWorkspace.Create();
        var outputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        var result = await tools.marp_render_string(TwoSlideDeck, outputPath);

        Assert.Contains(outputPath, result);
        Assert.Contains("2 slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task McpTools_RenderString_ThrowsArgumentException_ForEmptyMarkdown()
    {
        using var workspace = TestWorkspace.Create();
        var outputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<ArgumentException>(() =>
            tools.marp_render_string("", outputPath));
    }

    [Fact]
    public async Task McpTools_RenderString_ThrowsArgumentException_ForEmptyOutputPath()
    {
        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<ArgumentException>(() =>
            tools.marp_render_string(TwoSlideDeck, ""));
    }

    [Fact]
    public async Task McpTools_RenderString_WithSourceDirectory_DoesNotThrow()
    {
        using var workspace = TestWorkspace.Create();
        var outputPath = workspace.GetPath("deck.pptx");

        var tools = new MarpToPptxTools();
        var result = await tools.marp_render_string(TwoSlideDeck, outputPath, sourceDirectory: workspace.RootPath);

        Assert.Contains(outputPath, result);
        Assert.True(File.Exists(outputPath));
    }

    // ── marp_diagnose_template ────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_DiagnoseTemplate_ReturnsJsonWithLayoutCatalog()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var templatePath = workspace.GetPath("template.pptx");

        var tools = new MarpToPptxTools();
        await tools.marp_render(markdownPath, templatePath);

        var json = await tools.marp_diagnose_template(templatePath);

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(templatePath, root.GetProperty("templatePath").GetString());
        Assert.True(root.GetProperty("slideMasterCount").GetInt32() >= 1);
        Assert.True(root.GetProperty("layoutCount").GetInt32() >= 1);
        Assert.True(root.GetProperty("layouts").GetArrayLength() >= 1);

        var firstLayout = root.GetProperty("layouts")[0];
        Assert.True(firstLayout.TryGetProperty("index", out _));
        Assert.True(firstLayout.TryGetProperty("name", out _));
        Assert.True(firstLayout.TryGetProperty("role", out _));
        Assert.True(firstLayout.TryGetProperty("hasTitle", out _));
        Assert.True(firstLayout.TryGetProperty("hasBody", out _));
        Assert.True(firstLayout.TryGetProperty("hasPicture", out _));

        var recommendations = root.GetProperty("recommendations");
        Assert.True(recommendations.TryGetProperty("defaultContentLayout", out _));
        Assert.True(recommendations.TryGetProperty("titleLayout", out _));
        Assert.True(recommendations.TryGetProperty("sectionHeaderLayout", out _));
        Assert.True(recommendations.TryGetProperty("pictureLayout", out _));
        // The built-in template always yields at least one content layout recommendation.
        Assert.False(string.IsNullOrEmpty(recommendations.GetProperty("defaultContentLayout").GetString()));
    }

    [Fact]
    public async Task McpTools_DiagnoseTemplate_ThrowsFileNotFound_ForMissingTemplate()
    {
        using var workspace = TestWorkspace.Create();
        var missingPath = workspace.GetPath("nonexistent.pptx");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<FileNotFoundException>(() => tools.marp_diagnose_template(missingPath));
    }

    // ── marp_recommend_layouts ────────────────────────────────────────────────

    [Fact]
    public async Task McpTools_RecommendLayouts_ReturnsJsonWithPerSlideRecommendations()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var templatePath = workspace.GetPath("template.pptx");

        var tools = new MarpToPptxTools();
        await tools.marp_render(markdownPath, templatePath);

        var json = await tools.marp_recommend_layouts(markdownPath, templatePath);

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(2, root.GetProperty("slideCount").GetInt32());
        Assert.Equal(2, root.GetProperty("slides").GetArrayLength());
        Assert.True(root.TryGetProperty("suggestedFrontMatterLayout", out _));
        Assert.True(root.TryGetProperty("photoLayoutRotation", out _));

        var firstSlide = root.GetProperty("slides")[0];
        Assert.Equal(0, firstSlide.GetProperty("index").GetInt32());
        Assert.Equal("Intro Slide", firstSlide.GetProperty("title").GetString());
        Assert.True(firstSlide.TryGetProperty("contentKind", out _));
        Assert.True(firstSlide.TryGetProperty("recommendedLayout", out _));
        Assert.True(firstSlide.TryGetProperty("isExplicitLayout", out _));
        Assert.True(firstSlide.TryGetProperty("explicitLayout", out _));
    }

    [Fact]
    public async Task McpTools_RecommendLayouts_ThrowsFileNotFound_ForMissingMarkdown()
    {
        using var workspace = TestWorkspace.Create();
        var missingPath = workspace.GetPath("nonexistent.md");
        // Create and render a real deck so the template file exists; the missing-path check
        // should fail on the markdown argument, not the template argument.
        var deckPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var templatePath = workspace.GetPath("template.pptx");

        var tools = new MarpToPptxTools();
        await tools.marp_render(deckPath, templatePath);

        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            tools.marp_recommend_layouts(missingPath, templatePath));
    }

    [Fact]
    public async Task McpTools_RecommendLayouts_ThrowsFileNotFound_ForMissingTemplate()
    {
        using var workspace = TestWorkspace.Create();
        var markdownPath = workspace.WriteMarkdown("deck.md", TwoSlideDeck);
        var missingTemplatePath = workspace.GetPath("nonexistent.pptx");

        var tools = new MarpToPptxTools();
        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            tools.marp_recommend_layouts(markdownPath, missingTemplatePath));
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

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
