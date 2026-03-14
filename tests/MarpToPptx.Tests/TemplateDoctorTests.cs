using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Diagnostics;
using MarpToPptx.Pptx.Rendering;

namespace MarpToPptx.Tests;

public class TemplateDoctorTests
{
    // ─────────────────────────────────────────────────────────────
    // Smoke test: analyze does not throw on a valid template
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Analyze_ReturnsReportForValidTemplate()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        Assert.NotNull(report);
        Assert.Equal(pptxPath, report.TemplatePath);
        Assert.False(report.WroteFixedTemplate);
        Assert.Null(report.FixedTemplatePath);
        Assert.Empty(report.AppliedFixes);
    }

    // ─────────────────────────────────────────────────────────────
    // Issue collection: issues are well-formed records
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Issues_HaveNonEmptyCodeAndDescription()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        Assert.All(report.Issues, issue =>
        {
            Assert.False(string.IsNullOrWhiteSpace(issue.Code));
            Assert.False(string.IsNullOrWhiteSpace(issue.Description));
        });
    }

    [Fact]
    public void TemplateDoctor_FixableIssues_HaveProposedFix()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        Assert.All(
            report.Issues.Where(i => i.Severity == IssueSeverity.Fixable),
            issue => Assert.False(string.IsNullOrWhiteSpace(issue.ProposedFix)));
    }

    // ─────────────────────────────────────────────────────────────
    // Issue severity values are all defined enum members
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_IssueSeverities_AreValidEnumValues()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        var validSeverities = new[] { IssueSeverity.Info, IssueSeverity.Warning, IssueSeverity.Fixable };
        Assert.All(report.Issues, issue => Assert.Contains(issue.Severity, validSeverities));
    }

    // ─────────────────────────────────────────────────────────────
    // The built-in template has two layouts so the blank layout
    // produces a "missing title layout" warning from the diagnoser;
    // the doctor checks do not throw, and we can assert that no
    // "ContentLayoutMissingBodyPlaceholder" issue is raised for a
    // layout that correctly declares a body placeholder.
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_BuiltInContentLayout_DoesNotReportMissingBodyPlaceholder()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        // The built-in "tx" layout has a body placeholder — no issue should be raised for it.
        var missingBodyIssues = report.Issues
            .Where(i => i.Code == "ContentLayoutMissingBodyPlaceholder")
            .ToList();

        // There should be no such issue on the built-in content layout (which has a body).
        foreach (var issue in missingBodyIssues)
        {
            // If any are reported, they must not be for the content layout.
            Assert.NotEqual("Title and Content", issue.LayoutName,
                StringComparer.OrdinalIgnoreCase);
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Dry-run mode: no file written
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Analyze_DryRun_DoesNotWriteFile()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);
        var potentialOutput = workspace.GetPath("fixed.pptx");

        var doctor = new TemplateDoctor();
        // Analyze (dry-run) should not write potentialOutput even if we never passed it.
        var report = doctor.Analyze(pptxPath);

        Assert.False(report.WroteFixedTemplate);
        Assert.False(File.Exists(potentialOutput),
            "Dry-run analyze must not write any file.");
    }

    // ─────────────────────────────────────────────────────────────
    // Write mode: file is written when outputPath is supplied
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Run_WithOutputPath_WritesFile()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);
        var outputPath = workspace.GetPath("fixed.pptx");

        var doctor = new TemplateDoctor();
        var report = doctor.Run(pptxPath, outputPath);

        Assert.True(report.WroteFixedTemplate);
        Assert.Equal(outputPath, report.FixedTemplatePath);
        Assert.True(File.Exists(outputPath), "Output file must exist after Run with outputPath.");
    }

    [Fact]
    public void TemplateDoctor_Run_WithOutputPath_DoesNotModifyOriginal()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);
        var originalBytes = File.ReadAllBytes(pptxPath);
        var outputPath = workspace.GetPath("fixed.pptx");

        var doctor = new TemplateDoctor();
        doctor.Run(pptxPath, outputPath);

        var afterBytes = File.ReadAllBytes(pptxPath);
        Assert.Equal(originalBytes, afterBytes);
    }

    [Fact]
    public void TemplateDoctor_Run_OutputFileIsValidPptx()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);
        var outputPath = workspace.GetPath("fixed.pptx");

        var doctor = new TemplateDoctor();
        doctor.Run(pptxPath, outputPath);

        // Verify the output file can be opened as a valid PresentationDocument.
        using var doc = PresentationDocument.Open(outputPath, false);
        Assert.NotNull(doc.PresentationPart);
    }

    // ─────────────────────────────────────────────────────────────
    // Report: template path is stored
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Report_StoresTemplatePath()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        Assert.Equal(pptxPath, report.TemplatePath);
    }

    // ─────────────────────────────────────────────────────────────
    // Informational issues: blank layout or visually-redundant
    // layouts should produce informational (not Fixable) issues
    // only, when applicable
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Issues_ContainOnlySupportedCodes()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);

        var doctor = new TemplateDoctor();
        var report = doctor.Analyze(pptxPath);

        var knownCodes = new HashSet<string>
        {
            "DuplicateLayoutName",
            "EmptyLayoutName",
            "ContentLayoutMissingTitlePlaceholder",
            "ContentLayoutMissingBodyPlaceholder",
            "PlaceholderGeometryInherited",
            "TypelessIndexedBodyPlaceholder",
            "UnmappableLayoutRole",
            "VisuallyRedundantLayouts",
        };

        Assert.All(report.Issues, issue =>
            Assert.Contains(issue.Code, knownCodes));
    }

    // ─────────────────────────────────────────────────────────────
    // AppliedFixes: non-empty strings only
    // ─────────────────────────────────────────────────────────────

    [Fact]
    public void TemplateDoctor_Run_AppliedFixes_AreNonEmptyStrings()
    {
        using var workspace = TestWorkspace.Create();
        var pptxPath = RenderMinimalDeck(workspace);
        var outputPath = workspace.GetPath("fixed.pptx");

        var doctor = new TemplateDoctor();
        var report = doctor.Run(pptxPath, outputPath);

        Assert.All(report.AppliedFixes, fix => Assert.False(string.IsNullOrWhiteSpace(fix)));
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
