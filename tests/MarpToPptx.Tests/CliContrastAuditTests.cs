namespace MarpToPptx.Tests;

public class CliContrastAuditTests
{
    [Fact]
    public async Task Cli_ContrastWarningsSummary_PrintsSlideSummaryWithoutFailingGeneration()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Low Contrast Deck

            This body paragraph text has extremely low contrast.
            """);

        var cssPath = workspace.WriteText(
            "theme.css",
            """
            section { color: #EEEEEE; background-color: #FFFFFF; }
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var (exitCode, stdout, stderr) = await RunCliAsync(markdownPath, "-o", outputPath, "--theme-css", cssPath, "--contrast-warnings", "summary");

        Assert.Equal(0, exitCode);
        Assert.Contains("Generated '", stdout);
        Assert.Contains("Warning: Slides 1 may have low-contrast accessibility issues.", stderr);
        Assert.DoesNotContain("Shape \"Text\"", stderr);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Cli_ContrastWarningsDetailed_PrintsPerShapeFindings()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Low Contrast Deck

            This body paragraph text has extremely low contrast.
            """);

        var cssPath = workspace.WriteText(
            "theme.css",
            """
            section { color: #EEEEEE; background-color: #FFFFFF; }
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var (exitCode, stdout, stderr) = await RunCliAsync(markdownPath, "-o", outputPath, "--theme-css", cssPath, "--contrast-warnings", "detailed");

        Assert.Equal(0, exitCode);
        Assert.Contains("Generated '", stdout);
        Assert.Contains("Warning: Contrast audit found", stderr);
        Assert.Contains("Slide 1 - Shape \"Text\"", stderr);
    }

    [Fact]
    public async Task Cli_ContrastReport_WritesAuditSummaryToRequestedFile()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Good Contrast Deck

            This body paragraph text should pass.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var reportPath = workspace.GetPath("contrast-report.txt");
        var (exitCode, stdout, stderr) = await RunCliAsync(markdownPath, "-o", outputPath, "--contrast-report", reportPath);

        Assert.Equal(0, exitCode);
        Assert.Contains("Contrast audit report written to", stdout);
        Assert.DoesNotContain("Warning: Contrast audit found", stderr);
        Assert.True(File.Exists(reportPath));
        Assert.Contains("Contrast audit passed", File.ReadAllText(reportPath));
    }

    [Fact]
    public async Task Cli_WarnLowContrast_Alias_MapsToDetailedMode()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Low Contrast Deck

            This body paragraph text has extremely low contrast.
            """);

        var cssPath = workspace.WriteText(
            "theme.css",
            """
            section { color: #EEEEEE; background-color: #FFFFFF; }
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var (exitCode, _, stderr) = await RunCliAsync(markdownPath, "-o", outputPath, "--theme-css", cssPath, "--warn-low-contrast");

        Assert.Equal(0, exitCode);
        Assert.Contains("Slide 1 - Shape \"Text\"", stderr);
    }

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunCliAsync(params string[] args)
    {
        var stdoutWriter = new StringWriter();
        var stderrWriter = new StringWriter();
        var originalOut = Console.Out;
        var originalError = Console.Error;

        Console.SetOut(stdoutWriter);
        Console.SetError(stderrWriter);

        try
        {
            var exitCode = await ProgramEntry.RunAsync(args);
            return (exitCode, stdoutWriter.ToString(), stderrWriter.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }
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

        public string WriteText(string relativePath, string content)
        {
            var path = GetPath(relativePath);
            File.WriteAllText(path, content);
            return path;
        }

        public void Dispose()
        {
            if (Directory.Exists(RootPath))
            {
                Directory.Delete(RootPath, recursive: true);
            }
        }
    }
}