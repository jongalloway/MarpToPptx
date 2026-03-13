using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Contrast;
using MarpToPptx.Pptx.Rendering;

namespace MarpToPptx.Tests;

public class ContrastAuditorTests
{
    // ─────────────────────────────────────────────────────────────
    // ContrastCalculator unit tests (tested through the auditor)
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void ContrastAuditor_DefaultTheme_AllColorPairsPass()
    {
        // Default theme uses dark gray (#1F2937) text on a white (#FFFFFF) background,
        // which is well above the 4.5:1 WCAG 2.1 normal-text threshold.
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderGoodContrastDeck(workspace);

        using var document = PresentationDocument.Open(pptxPath, false);
        var auditor = new ContrastAuditor();
        var results = auditor.Audit(document);

        // There must be at least one result from the deck.
        Assert.NotEmpty(results);

        // All results from the default dark-on-light theme must pass.
        Assert.All(results, r => Assert.False(r.IsFailing,
            $"Slide {r.SlideNumber} \"{r.ShapeContext}\": #{r.ForegroundColor} on #{r.BackgroundColor} " +
            $"= {r.ContrastRatio:F2}:1, requires {r.MinimumRequiredRatio:F1}:1"));
    }

    [Fact]
    public void ContrastAuditor_LowContrastNormalText_ReportsFailure()
    {
        // Very light gray (#EEEEEE) on white (#FFFFFF) has a contrast ratio of ~1.19:1,
        // well below the 4.5:1 WCAG 2.1 threshold for normal text.
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
            section { color: #EEEEEE; background-color: #FFFFFF; }
            """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Low Contrast Deck

            This body paragraph text has extremely low contrast.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var auditor = new ContrastAuditor();
        var results = auditor.Audit(document);

        // At least the body text run should fail.
        var failures = results.Where(r => r.IsFailing).ToList();
        Assert.NotEmpty(failures);

        var failure = failures.First();
        Assert.Equal(1, failure.SlideNumber);
        Assert.True(
            failure.ContrastRatio < 4.5,
            $"Expected contrast ratio below 4.5:1 but got {failure.ContrastRatio:F2}:1");
    }

    [Fact]
    public void ContrastAuditor_TableCellFills_AreInspected()
    {
        // A rendered deck with a Markdown table must produce audit results whose
        // ShapeContext identifies the table/row/cell location.
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Table Slide

            | Column A | Column B |
            | -------- | -------- |
            | Alpha    | Beta     |
            | Gamma    | Delta    |
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var auditor = new ContrastAuditor();
        var results = auditor.Audit(document);

        // Must have results from table cells.
        Assert.Contains(results, r => r.ShapeContext.Contains("Table", StringComparison.Ordinal));
    }

    [Fact]
    public void ContrastAuditor_DefaultTableColors_PassContrast()
    {
        // The renderer emits a fixed body-row fill (#FFFFFF / #F8FAFC) with dark text (#1F2937).
        // This combination must satisfy the 4.5:1 WCAG normal-text threshold.
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Table Slide

            | Header A | Header B |
            | -------- | -------- |
            | Body 1   | Body 2   |
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var auditor = new ContrastAuditor();
        var results = auditor.Audit(document);

        var tableResults = results.Where(r => r.ShapeContext.Contains("Table", StringComparison.Ordinal)).ToList();
        Assert.NotEmpty(tableResults);

        Assert.All(tableResults, r => Assert.False(r.IsFailing,
            $"Table cell contrast failure: #{r.ForegroundColor} on #{r.BackgroundColor} " +
            $"= {r.ContrastRatio:F2}:1 at {r.ShapeContext}"));
    }

    [Fact]
    public void ContrastAuditor_MultiSlide_ReportsCorrectSlideNumbers()
    {
        // The auditor must correctly attribute results to the right slide number.
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
            section.bad { color: #EEEEEE; background-color: #FFFFFF; }
            """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: custom
            ---

            # Slide One (good contrast)

            Normal body text here.

            ---

            <!-- _class: bad -->

            # Slide Two (bad contrast body)

            Low contrast body text.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var auditor = new ContrastAuditor();
        var results = auditor.Audit(document);

        // There should be failures only on slide 2 (or at least including slide 2).
        var failingSlides = results.Where(r => r.IsFailing).Select(r => r.SlideNumber).Distinct().ToList();
        Assert.Contains(2, failingSlides);
    }

    [Fact]
    public void ContrastAuditor_FileNotFound_Throws()
    {
        var auditor = new ContrastAuditor();
        var missingPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        Assert.Throws<FileNotFoundException>(() => auditor.Audit(missingPath));
    }

    // ─────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────

    private static string RenderGoodContrastDeck(TestWorkspace workspace)
    {
        // Default theme: #1F2937 text on #FFFFFF background — excellent contrast.
        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title

            A body paragraph with good contrast.

            ---

            ## Slide Two

            - Bullet point
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        return outputPath;
    }

    private static void RenderDeck(string markdownPath, string outputPath, string sourceDirectory)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);

        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = sourceDirectory });
    }

    private static void RenderDeckWithTheme(string markdownPath, string outputPath, string sourceDirectory, string themeCss)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath, themeCss);

        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = sourceDirectory });
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
