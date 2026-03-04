using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;
using MarpToPptx.Pptx.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Tests;

public class SyntaxHighlighterTests
{
    [Theory]
    [InlineData("csharp")]
    [InlineData("javascript")]
    [InlineData("typescript")]
    [InlineData("json")]
    [InlineData("html")]
    [InlineData("css")]
    [InlineData("xml")]
    [InlineData("powershell")]
    [InlineData("python")]
    [InlineData("sql")]
    public void IsSupported_ReturnsTrueForKnownLanguages(string language)
    {
        Assert.True(SyntaxHighlighter.IsSupported(language));
    }

    [Theory]
    [InlineData("cs")]
    [InlineData("js")]
    [InlineData("ts")]
    [InlineData("py")]
    [InlineData("ps1")]
    [InlineData("bash")]
    [InlineData("sh")]
    [InlineData("yml")]
    public void IsSupported_ReturnsTrueForCommonAliases(string alias)
    {
        Assert.True(SyntaxHighlighter.IsSupported(alias));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("unknown-lang-xyz")]
    [InlineData("nonexistent-lang-xyz")]
    public void IsSupported_ReturnsFalseForUnknownLanguages(string? language)
    {
        Assert.False(SyntaxHighlighter.IsSupported(language));
    }

    [Fact]
    public void Tokenize_CSharp_ProducesColoredRunsForKeywords()
    {
        var lines = SyntaxHighlighter.Tokenize("csharp", "public class Foo { }");

        Assert.Single(lines);
        var runs = lines[0];
        Assert.True(runs.Count > 1, "Expected multiple runs for syntax-highlighted C#");

        // At least one run should have a non-null color (keyword highlight)
        Assert.Contains(runs, r => r.Color is not null);
    }

    [Fact]
    public void Tokenize_CSharp_KeywordRunHasDifferentColorFromDefaultText()
    {
        var lines = SyntaxHighlighter.Tokenize("csharp", "public class Foo { }");
        var runs = lines[0];

        // Collect distinct colors
        var colors = runs.Where(r => r.Color is not null).Select(r => r.Color).Distinct().ToList();

        // Should have at least one colored token (keyword "public", "class", type "Foo")
        Assert.NotEmpty(colors);
    }

    [Fact]
    public void Tokenize_CSharp_PreservesFullCodeText()
    {
        const string code = "public class Foo { }";
        var lines = SyntaxHighlighter.Tokenize("csharp", code);

        Assert.Single(lines);
        var reassembled = string.Concat(lines[0].Select(r => r.Text));
        Assert.Equal(code, reassembled);
    }

    [Fact]
    public void Tokenize_CSharp_MultilinePreservesAllLines()
    {
        const string code = "public class Foo {\n    int x = 0;\n}";
        var lines = SyntaxHighlighter.Tokenize("csharp", code);

        Assert.Equal(3, lines.Count);
        Assert.Equal("public class Foo {", string.Concat(lines[0].Select(r => r.Text)));
        Assert.Equal("    int x = 0;", string.Concat(lines[1].Select(r => r.Text)));
        Assert.Equal("}", string.Concat(lines[2].Select(r => r.Text)));
    }

    [Fact]
    public void Tokenize_UnsupportedLanguage_ReturnsOnePlainRunPerLine()
    {
        const string code = "hello world\nfoo bar";
        var lines = SyntaxHighlighter.Tokenize("nonexistent-lang-xyz", code);

        Assert.Equal(2, lines.Count);
        Assert.Single(lines[0]);
        Assert.Single(lines[1]);
        Assert.Null(lines[0][0].Color);
        Assert.Null(lines[1][0].Color);
        Assert.Equal("hello world", lines[0][0].Text);
        Assert.Equal("foo bar", lines[1][0].Text);
    }

    [Fact]
    public void Tokenize_EmptyCode_ReturnsOneEmptyLine()
    {
        var lines = SyntaxHighlighter.Tokenize("csharp", string.Empty);

        Assert.Single(lines);
    }

    [Fact]
    public void Tokenize_ColorValues_AreValidSixDigitHex()
    {
        var lines = SyntaxHighlighter.Tokenize("csharp", "public int x = 42;");

        foreach (var run in lines.SelectMany(l => l))
        {
            if (run.Color is not null)
            {
                Assert.Equal(6, run.Color.Length);
                Assert.All(run.Color, ch => Assert.True(
                    (ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F'),
                    $"Expected uppercase hex digit, got '{ch}' in color '{run.Color}'"));
            }
        }
    }

    [Fact]
    public void Renderer_CreatesHighlightedCodeBlock_ForKnownLanguage()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Code Slide

            ```csharp
            public class Foo
            {
                public void Bar() { }
            }
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // Find the code-block shape by its name ("Code (csharp)")
        var codeShape = slidePart.Slide!
            .Descendants<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?
                .NonVisualDrawingProperties?.Name?.Value?.StartsWith("Code (", StringComparison.Ordinal) == true);
        Assert.NotNull(codeShape);

        // The first content paragraph of the highlighted code block should have >1 run
        // (because keywords, whitespace, and identifiers get distinct colored runs)
        var firstCodeParagraph = codeShape.Descendants<A.Paragraph>()
            .FirstOrDefault(p => p.Descendants<A.Run>().Count() > 1);
        Assert.NotNull(firstCodeParagraph);

        // There must be at least 2 distinct RgbColorModelHex values within that paragraph,
        // proving that syntax highlighting produced different token colors
        var distinctColors = firstCodeParagraph!
            .Descendants<A.RgbColorModelHex>()
            .Select(rgb => rgb.Val?.Value)
            .Where(v => v is not null)
            .Distinct()
            .ToList();
        Assert.True(distinctColors.Count >= 2, $"Expected at least 2 distinct token colors in highlighted code paragraph, got {distinctColors.Count}: [{string.Join(", ", distinctColors)}]");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_RendersPlainCodeBlock_ForUnknownLanguage()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Code Slide

            ```nonexistent-lang-xyz
            DISPLAY "Hello".
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    private static void RenderDeck(string markdownPath, string outputPath, string sourceDirectory)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);

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
            {
                Directory.Delete(RootPath, recursive: true);
            }
        }
    }
}

