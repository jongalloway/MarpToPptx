using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;
using MarpToPptx.Pptx.Validation;
using System.IO.Compression;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Tests;

public class PptxRendererTests
{
    [Fact]
    public void Renderer_MatchesGoldenPackageBaseline_ForMinimalDeck()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Intro paragraph.

            ---

            ## Second Slide

            - Alpha
            - Beta

            ![Pixel](pixel.png)
            """);

        workspace.WriteFile(
            "pixel.png",
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII="));

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        PptxGoldenPackage.AssertMatchesFixture(outputPath, "minimal-deck.package.json");
    }

    [Fact]
    public void Renderer_MatchesGoldenPackageBaseline_ForSvgBackgroundDeck()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile(
            "accent-wave.svg",
            """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
              <rect width="100" height="100" fill="#102A43" />
              <path d="M0 70 C 20 40, 40 40, 60 70 S 100 100, 100 60 L100 100 L0 100 Z" fill="#F7C948" />
            </svg>
            """);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: gaia
            backgroundColor: "#F7F3E8"
            ---

            # Quoted Color

            ---

            <!-- backgroundImage: url(accent-wave.svg) -->
            # Svg Background
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        PptxGoldenPackage.AssertMatchesFixture(outputPath, "svg-background.package.json");
    }

    [Fact]
    public void Renderer_CreatesPresentationWithSlidesTextAndImage()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Intro paragraph.

            ---

            ## Second Slide

            - Alpha
            - Beta

            ![Pixel](pixel.png)
            """);

        workspace.WriteFile(
            "pixel.png",
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII="));

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var presentationPart = document.PresentationPart;
        Assert.NotNull(presentationPart);
        Assert.Equal(2, presentationPart!.Presentation!.SlideIdList!.Count());
        Assert.NotNull(presentationPart.PresentationPropertiesPart);
        Assert.NotNull(presentationPart.ViewPropertiesPart);
        Assert.NotNull(presentationPart.TableStylesPart);
        Assert.NotNull(presentationPart.ThemePart);
        Assert.NotNull(document.ExtendedFilePropertiesPart);
        Assert.Equal(2, presentationPart.SlideMasterParts.Single().SlideLayoutParts.Count());

        var slideParts = presentationPart.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);
        Assert.All(slideParts, slidePart => Assert.Equal("/ppt/slideLayouts/slideLayout2.xml", slidePart.SlideLayoutPart?.Uri.ToString()));
        Assert.NotNull(slideParts[0].Slide);
        Assert.NotNull(slideParts[1].Slide);
        Assert.Contains("Title Slide", slideParts[0].Slide!.Descendants<A.Text>().Select(text => text.Text));
        Assert.Contains("Intro paragraph.", slideParts[0].Slide!.Descendants<A.Text>().Select(text => text.Text));
        Assert.Single(slideParts[1].Slide!.Descendants<P.Picture>());

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);

        using var archive = ZipFile.OpenRead(outputPath);
        Assert.NotNull(archive.GetEntry("docProps/core.xml"));
        Assert.NotNull(archive.GetEntry("docProps/app.xml"));
        Assert.NotNull(archive.GetEntry("ppt/slideLayouts/_rels/slideLayout1.xml.rels"));
        Assert.NotNull(archive.GetEntry("ppt/slideLayouts/_rels/slideLayout2.xml.rels"));

        using var contentTypesReader = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
        var contentTypes = contentTypesReader.ReadToEnd();
        Assert.Contains("Extension=\"xml\" ContentType=\"application/xml\"", contentTypes);
        Assert.Contains("PartName=\"/ppt/presentation.xml\"", contentTypes);

        using var slideRelationshipsReader = new StreamReader(archive.GetEntry("ppt/slides/_rels/slide1.xml.rels")!.Open());
        var slideRelationships = slideRelationshipsReader.ReadToEnd();
        Assert.Contains("../slideLayouts/slideLayout2.xml", slideRelationships);
    }

    [Fact]
    public void Renderer_AcceptsQuotedFrontMatterColorsAndSvgBackgrounds()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile(
            "accent-wave.svg",
            """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
              <rect width="100" height="100" fill="#102A43" />
              <path d="M0 70 C 20 40, 40 40, 60 70 S 100 100, 100 60 L100 100 L0 100 Z" fill="#F7C948" />
            </svg>
            """);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: gaia
            backgroundColor: "#F7F3E8"
            ---

            # Quoted Color

            ---

            <!-- backgroundImage: url(accent-wave.svg) -->
            # Svg Background
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);

        using var archive = ZipFile.OpenRead(outputPath);
        using var contentTypesReader = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
        var contentTypes = contentTypesReader.ReadToEnd();
        Assert.Contains("image/svg+xml", contentTypes);
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
        private TestWorkspace(string rootPath)
        {
            RootPath = rootPath;
        }

        public string RootPath { get; }

        public static TestWorkspace Create()
        {
            var rootPath = Path.Combine(Path.GetTempPath(), "MarpToPptx.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(rootPath);
            return new TestWorkspace(rootPath);
        }

        public string GetPath(string relativePath)
            => Path.Combine(RootPath, relativePath);

        public string WriteMarkdown(string relativePath, string content)
        {
            var path = GetPath(relativePath);
            File.WriteAllText(path, content);
            return path;
        }

        public void WriteFile(string relativePath, string content)
        {
            File.WriteAllText(GetPath(relativePath), content);
        }

        public void WriteFile(string relativePath, byte[] content)
        {
            File.WriteAllBytes(GetPath(relativePath), content);
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
