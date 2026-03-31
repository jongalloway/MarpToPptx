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
