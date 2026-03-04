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

    [Fact]
    public void Renderer_CreatesNativePptxTable_ForMarkdownTable()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Table Slide

            | Name  | Score | Rank |
            |-------|------:|:----:|
            | Alice | 95    | 1    |
            | Bob   | 87    | 2    |
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var graphicFrames = slidePart.Slide!.Descendants<P.GraphicFrame>().ToArray();
        Assert.NotEmpty(graphicFrames);

        var table = graphicFrames[0].Descendants<A.Table>().Single();
        var rows = table.Elements<A.TableRow>().ToArray();
        Assert.Equal(3, rows.Length);

        var headerCells = rows[0].Elements<A.TableCell>().ToArray();
        Assert.Equal(3, headerCells.Length);
        Assert.Contains("Name", headerCells[0].Descendants<A.Text>().Select(t => t.Text));

        var headerRunProps = headerCells[0].Descendants<A.RunProperties>().First();
        Assert.Equal(true, headerRunProps.Bold?.Value);

        var scoreColProps = rows[1].Elements<A.TableCell>().ToArray()[1]
            .Descendants<A.ParagraphProperties>().FirstOrDefault();
        Assert.Equal(A.TextAlignmentTypeValues.Right, scoreColProps?.Alignment?.Value);

        var rankColProps = rows[1].Elements<A.TableCell>().ToArray()[2]
            .Descendants<A.ParagraphProperties>().FirstOrDefault();
        Assert.Equal(A.TextAlignmentTypeValues.Center, rankColProps?.Alignment?.Value);

        var tableProperties = table.Elements<A.TableProperties>().Single();
        Assert.Equal(true, tableProperties.FirstRow?.Value);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_ResolvesRemoteImage_WhenHttpHandlerReturnsImage()
    {
        using var workspace = TestWorkspace.Create();

        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");

        var handler = new StubHttpMessageHandler(req =>
        {
            var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new System.Net.Http.ByteArrayContent(pngBytes),
            };
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/png");
            return response;
        });

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide With Remote Image

            ![Remote](https://example.com/photo.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            SourceDirectory = workspace.RootPath,
            AllowRemoteAssets = true,
            RemoteAssetHandler = handler,
        });

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        Assert.DoesNotContain("Missing image", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));
    }

    [Fact]
    public void Renderer_ShowsActionableError_WhenRemoteImageReturnsHttpError()
    {
        using var workspace = TestWorkspace.Create();

        var handler = new StubHttpMessageHandler(_ =>
            new HttpResponseMessage(System.Net.HttpStatusCode.NotFound));

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Broken](https://example.com/missing.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            SourceDirectory = workspace.RootPath,
            AllowRemoteAssets = true,
            RemoteAssetHandler = handler,
        });

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains(texts, t => t.Contains("Missing image") && t.Contains("https://example.com/missing.png"));
        Assert.Contains(texts, t => t.Contains("404") || t.Contains("Not Found"));
    }

    [Fact]
    public void Renderer_ShowsActionableError_WhenRemoteAssetsDisabled()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Remote](https://example.com/photo.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            SourceDirectory = workspace.RootPath,
            AllowRemoteAssets = false,
        });

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains(texts, t => t.Contains("Missing image") && t.Contains("https://example.com/photo.png"));
    }

    [Fact]
    public void Renderer_AcceptsBmpImage()
    {
        using var workspace = TestWorkspace.Create();

        // Minimal valid 16x16 BMP with a BITMAPINFOHEADER
        var bmpBytes = new byte[]
        {
            0x42, 0x4D,             // BM signature
            0x36, 0x03, 0x00, 0x00, // file size = 822 (doesn't matter for test)
            0x00, 0x00, 0x00, 0x00, // reserved
            0x36, 0x00, 0x00, 0x00, // pixel data offset = 54
            0x28, 0x00, 0x00, 0x00, // BITMAPINFOHEADER size = 40
            0x10, 0x00, 0x00, 0x00, // width = 16
            0x10, 0x00, 0x00, 0x00, // height = 16
            0x01, 0x00,             // color planes = 1
            0x18, 0x00,             // bits per pixel = 24
            0x00, 0x00, 0x00, 0x00, // compression = none
            0x00, 0x03, 0x00, 0x00, // image size (can be 0 for uncompressed)
            0x13, 0x0B, 0x00, 0x00, // x pixels per meter
            0x13, 0x0B, 0x00, 0x00, // y pixels per meter
            0x00, 0x00, 0x00, 0x00, // colors in color table
            0x00, 0x00, 0x00, 0x00, // important color count
        };

        workspace.WriteFile("image.bmp", bmpBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # BMP Slide

            ![Bitmap](image.bmp)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        Assert.DoesNotContain("Missing image", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));
        Assert.DoesNotContain("Unsupported image format", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));
    }

    [Fact]
    public void Renderer_AcceptsWebpImage()
    {
        using var workspace = TestWorkspace.Create();

        // Minimal VP8X extended WebP with 16x16 canvas
        // Structure: RIFF header (12) + VP8X chunk (20 bytes: 4 tag + 4 size + 4 flags + 3 cw + 3 ch + 2 padding)
        var webpBytes = new byte[]
        {
            0x52, 0x49, 0x46, 0x46, // RIFF
            0x1E, 0x00, 0x00, 0x00, // file size = 30 (just enough)
            0x57, 0x45, 0x42, 0x50, // WEBP
            0x56, 0x50, 0x38, 0x58, // VP8X
            0x0A, 0x00, 0x00, 0x00, // chunk size = 10
            0x00, 0x00, 0x00, 0x00, // flags
            0x0F, 0x00, 0x00,       // canvas_width_minus_one = 15 → width = 16
            0x0F, 0x00, 0x00,       // canvas_height_minus_one = 15 → height = 16
        };

        workspace.WriteFile("image.webp", webpBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # WebP Slide

            ![Webp](image.webp)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        Assert.DoesNotContain("Missing image", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));
        Assert.DoesNotContain("Unsupported image format", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));
    }

    [Fact]
    public void Renderer_ShowsActionableError_ForUnsupportedImageFormat()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile("icon.ico", new byte[] { 0x00, 0x00, 0x01, 0x00, 0x00, 0x00 });

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Icon](icon.ico)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains(texts, t => t.Contains("Unsupported image format"));
    }

    [Fact]
    public void Renderer_ResolvesRemoteImage_WhenUrlHasNoFileExtension()
    {
        using var workspace = TestWorkspace.Create();

        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");

        var handler = new StubHttpMessageHandler(req =>
        {
            var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new System.Net.Http.ByteArrayContent(pngBytes),
            };
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/png");
            return response;
        });

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Remote No Ext](https://cdn.example.com/images/thumbnail)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            SourceDirectory = workspace.RootPath,
            AllowRemoteAssets = true,
            RemoteAssetHandler = handler,
        });

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
    }

    private static void RenderDeck(string markdownPath, string outputPath, string sourceDirectory)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);

        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = sourceDirectory });
    }

    private sealed class StubHttpMessageHandler(Func<HttpRequestMessage, HttpResponseMessage> handler) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            => Task.FromResult(handler(request));
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
