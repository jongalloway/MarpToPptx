using System.Text.RegularExpressions;

namespace MarpToPptx.Tests;

public class CliWriteSlideIdsTests
{
    [Fact]
    public async Task Cli_WriteSlideIds_AddsMissingDirectives_AndRendersDeck()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Intro Slide

            ---

            # Details Slide
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var (exitCode, stdout, stderr) = await RunCliAsync(markdownPath, "-o", outputPath, "--write-slide-ids");

        Assert.Equal(0, exitCode);
        Assert.Contains("Added 2 slideId directive(s)", stdout);
        Assert.Contains("Generated '", stdout);
        Assert.True(string.IsNullOrWhiteSpace(stderr));
        Assert.True(File.Exists(outputPath));

        var rewrittenMarkdown = File.ReadAllText(markdownPath);
        Assert.Contains("<!-- slideId: intro-slide -->", rewrittenMarkdown);
        Assert.Contains("<!-- slideId: details-slide -->", rewrittenMarkdown);
    }

    [Fact]
    public async Task Cli_WriteSlideIds_PreservesExistingSlideIds_AndOnlyAddsMissingOnes()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- slideId: custom-intro -->
            # Intro Slide

            ---

            # Details Slide
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var (exitCode, stdout, _) = await RunCliAsync(markdownPath, "-o", outputPath, "--write-slide-ids");

        Assert.Equal(0, exitCode);
        Assert.Contains("Added 1 slideId directive(s)", stdout);

        var rewrittenMarkdown = File.ReadAllText(markdownPath);
        Assert.Contains("<!-- slideId: custom-intro -->", rewrittenMarkdown);
        Assert.Contains("<!-- slideId: details-slide -->", rewrittenMarkdown);
        Assert.DoesNotContain("<!-- slideId: intro-slide -->", rewrittenMarkdown);
    }

    [Fact]
    public async Task Cli_WriteSlideIds_DoesNotDuplicateDirectives_OnSecondRun()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Intro Slide

            ---

            # Details Slide
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var firstRun = await RunCliAsync(markdownPath, "-o", outputPath, "--write-slide-ids");
        var secondRun = await RunCliAsync(markdownPath, "-o", outputPath, "--write-slide-ids");

        Assert.Equal(0, firstRun.ExitCode);
        Assert.Equal(0, secondRun.ExitCode);

        var rewrittenMarkdown = File.ReadAllText(markdownPath);
        Assert.Equal(2, Regex.Matches(rewrittenMarkdown, @"<!--\s*slideId:").Count);
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

        public void Dispose()
        {
            if (Directory.Exists(RootPath))
            {
                Directory.Delete(RootPath, recursive: true);
            }
        }
    }
}