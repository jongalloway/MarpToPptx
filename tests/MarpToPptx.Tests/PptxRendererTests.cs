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
        // Slide 1 (H1 + paragraph = Title kind) and slide 2 (H2 + bullets + image = Content kind)
        // both use the content layout (type="tx"); image-focused selection applies when images are at least 50% of non-heading elements.
        Assert.All(slideParts, slidePart => Assert.Equal("/ppt/slideLayouts/slideLayout1.xml", slidePart.SlideLayoutPart?.Uri.ToString()));
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
        Assert.Contains("../slideLayouts/slideLayout1.xml", slideRelationships);
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

    [Fact]
    public void Renderer_EmbedsMp4Video_WhenLocalFileExists()
    {
        using var workspace = TestWorkspace.Create();

        // Minimal ftyp box to produce a recognizable but tiny MP4-like file.
        var mp4Bytes = new byte[]
        {
            0x00, 0x00, 0x00, 0x14, // box size = 20
            0x66, 0x74, 0x79, 0x70, // 'ftyp'
            0x69, 0x73, 0x6F, 0x6D, // major brand = 'isom'
            0x00, 0x00, 0x02, 0x00, // minor version
            0x69, 0x73, 0x6F, 0x6D, // compatible brand = 'isom'
        };

        workspace.WriteFile("clip.mp4", mp4Bytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Video Slide

            ![Demo clip](clip.mp4)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // The picture element should be present with a video file reference.
        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);

        var videoFile = pictures[0].Descendants<A.VideoFromFile>().SingleOrDefault();
        Assert.NotNull(videoFile);
        Assert.NotNull(videoFile!.Link?.Value);

        // The slide should contain a video reference relationship pointing to an mp4 media part.
        var videoRels = slidePart.DataPartReferenceRelationships
            .OfType<VideoReferenceRelationship>()
            .ToArray();
        Assert.Single(videoRels);
        Assert.Equal("video/mp4", videoRels[0].DataPart.ContentType);

        // No error text should appear.
        Assert.DoesNotContain("Missing video", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));

        // Package structure should be valid.
        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);

        // Content-Types should declare the mp4 media type.
        using var archive = ZipFile.OpenRead(outputPath);
        using var contentTypesReader = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
        var contentTypes = contentTypesReader.ReadToEnd();
        Assert.Contains("video/mp4", contentTypes);
    }

    [Fact]
    public void Renderer_ShowsActionableError_WhenMp4FileIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Missing video](missing.mp4)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains(texts, t => t.Contains("Missing video") && t.Contains("missing.mp4"));
    }

    [Fact]
    public void Renderer_ShowsActionableError_ForUnsupportedVideoFormat()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile("video.avi", new byte[] { 0x52, 0x49, 0x46, 0x46 });

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![AVI](video.avi)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains(texts, t => t.Contains("Unsupported image format") && t.Contains("video.avi"));
    }

    [Fact]
    public void Renderer_SelectsBlankLayout_ForVideoFocusedSlide()
    {
        using var workspace = TestWorkspace.Create();

        var mp4Bytes = new byte[]
        {
            0x00, 0x00, 0x00, 0x14,
            0x66, 0x74, 0x79, 0x70,
            0x69, 0x73, 0x6F, 0x6D,
            0x00, 0x00, 0x02, 0x00,
            0x69, 0x73, 0x6F, 0x6D,
        };
        workspace.WriteFile("clip.mp4", mp4Bytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Video Slide

            ![Demo](clip.mp4)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // A slide with only a heading + video should select the blank layout (image-focused).
        Assert.Equal("/ppt/slideLayouts/slideLayout2.xml", slidePart.SlideLayoutPart?.Uri.ToString());
    }

    [Fact]
    public void Renderer_EmbedsMp4Video_WhenSpecifiedViaHtmlVideoTag()
    {
        using var workspace = TestWorkspace.Create();

        var mp4Bytes = new byte[]
        {
            0x00, 0x00, 0x00, 0x14,
            0x66, 0x74, 0x79, 0x70,
            0x69, 0x73, 0x6F, 0x6D,
            0x00, 0x00, 0x02, 0x00,
            0x69, 0x73, 0x6F, 0x6D,
        };
        workspace.WriteFile("clip.mp4", mp4Bytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Video Slide

            <video src="clip.mp4" controls></video>
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);
        Assert.NotNull(pictures[0].Descendants<A.VideoFromFile>().SingleOrDefault());

        var videoRels = slidePart.DataPartReferenceRelationships
            .OfType<VideoReferenceRelationship>()
            .ToArray();
        Assert.Single(videoRels);
        Assert.Equal("video/mp4", videoRels[0].DataPart.ContentType);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmbedsMp4Video_WhenSpecifiedViaSelfClosingHtmlVideoTag()
    {
        using var workspace = TestWorkspace.Create();

        var mp4Bytes = new byte[]
        {
            0x00, 0x00, 0x00, 0x14,
            0x66, 0x74, 0x79, 0x70,
            0x69, 0x73, 0x6F, 0x6D,
            0x00, 0x00, 0x02, 0x00,
            0x69, 0x73, 0x6F, 0x6D,
        };
        workspace.WriteFile("clip.mp4", mp4Bytes);

        // Self-closing <video /> becomes an HtmlBlock when surrounded by blank lines.
        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            "# Video Slide\n\n<video src=\"clip.mp4\" />\n\nCaption text.");

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);
        Assert.NotNull(pictures[0].Descendants<A.VideoFromFile>().SingleOrDefault());

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_SelectsDifferentLayouts_ForDifferentSlideKinds()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile(
            "pixel.png",
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII="));

        // Slide 1: H1 + paragraph  → Title kind     → content layout (type="tx")
        // Slide 2: H2 + bullets    → Content kind   → content layout (type="tx")
        // Slide 3: H1 + image      → ImageFocused   → blank layout   (type="blank")
        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Welcome message.

            ---

            ## Content Slide

            - Bullet 1
            - Bullet 2

            ---

            # Image Slide

            ![Pixel](pixel.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(3, slideParts.Length);

        // Slide 1 (Title kind): no dedicated title layout in scaffold → falls back to content layout.
        Assert.Equal("/ppt/slideLayouts/slideLayout1.xml", slideParts[0].SlideLayoutPart?.Uri.ToString());
        // Slide 2 (Content kind): content layout (type="tx").
        Assert.Equal("/ppt/slideLayouts/slideLayout1.xml", slideParts[1].SlideLayoutPart?.Uri.ToString());
        // Slide 3 (ImageFocused kind): blank layout (type="blank").
        Assert.Equal("/ppt/slideLayouts/slideLayout2.xml", slideParts[2].SlideLayoutPart?.Uri.ToString());

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_SelectsTitleLayoutFromTemplate_ForTitleSlide()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithMultipleLayouts(templatePath);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Subtitle text.

            ---

            ## Content Slide

            Body text here.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            TemplatePath = templatePath,
            SourceDirectory = workspace.RootPath,
        });

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);

        // Slide 1 (Title kind) should use the title layout from the template.
        Assert.Equal(P.SlideLayoutValues.Title, slideParts[0].SlideLayoutPart?.SlideLayout?.Type?.Value);

        // Slide 2 (Content kind) should use the content (text) layout from the template.
        Assert.Equal(P.SlideLayoutValues.Text, slideParts[1].SlideLayoutPart?.SlideLayout?.Type?.Value);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_UsesTemplatePlaceholderRects_WhenBothTitleAndBodyAreExplicit()
    {
        using var workspace = TestWorkspace.Create();

        // Template with title layout that has explicit title (x=50,y=80 w=860,h=120) and
        // body (x=50,y=230 w=860,h=200) placeholder transforms (in layout units = EMU/12700).
        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithExplicitPlaceholders(
            templatePath,
            titleBounds: new PlaceholderBounds(X: 50, Y: 80, W: 860, H: 120),
            bodyBounds: new PlaceholderBounds(X: 50, Y: 230, W: 860, H: 200));

        // A title slide: H1 + one paragraph (both placeholders will fire).
        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Heading

            Subtitle paragraph.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            TemplatePath = templatePath,
            SourceDirectory = workspace.RootPath,
        });

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // Verify the title shape is positioned according to the template placeholder.
        var titleShape = slidePart.Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Title");
        var titleOff = titleShape.ShapeProperties!.Transform2D!.Offset!;
        Assert.Equal(50L * 12700, titleOff.X?.Value);
        Assert.Equal(80L * 12700, titleOff.Y?.Value);

        // Verify the body shape uses the body placeholder rect.
        var bodyShape = slidePart.Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Text");
        var bodyOff = bodyShape.ShapeProperties!.Transform2D!.Offset!;
        Assert.Equal(50L * 12700, bodyOff.X?.Value);
        Assert.Equal(230L * 12700, bodyOff.Y?.Value);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    /// <summary>
    /// Creates a minimal template PPTX with three layouts: Title (type="title"),
    /// Text/Content (type="tx"), and Blank (type="blank").
    /// </summary>
    private static void CreateTemplateWithMultipleLayouts(string path)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        var titleLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        titleLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        { Type = P.SlideLayoutValues.Title };
        titleLayoutPart.AddPart(slideMasterPart, "rId1");
        titleLayoutPart.SlideLayout.Save();

        var contentLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId2");
        contentLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        { Type = P.SlideLayoutValues.Text };
        contentLayoutPart.AddPart(slideMasterPart, "rId1");
        contentLayoutPart.SlideLayout.Save();

        var blankLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId3");
        blankLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Blank,
            Preserve = true,
        };
        blankLayoutPart.AddPart(slideMasterPart, "rId1");
        blankLayoutPart.SlideLayout.Save();

        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId4");
        themePart.Theme = new A.Theme
        {
            Name = "Test",
            ThemeElements = new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "1F2937" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "F8FAFC" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "0F766E" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "2563EB" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "F59E0B" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "DC2626" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "7C3AED" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "0891B2" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "2563EB" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "7C3AED" }))
                { Name = "Test" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "Test" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 6350 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 12700 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 19050 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })))
                { Name = "Test" }),
            ObjectDefaults = new A.ObjectDefaults(),
            ExtraColorSchemeList = new A.ExtraColorSchemeList(),
        };
        themePart.Theme.Save();

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new P.SlideLayoutIdList(
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(titleLayoutPart) },
                new P.SlideLayoutId { Id = 2147483650U, RelationshipId = slideMasterPart.GetIdOfPart(contentLayoutPart) },
                new P.SlideLayoutId { Id = 2147483651U, RelationshipId = slideMasterPart.GetIdOfPart(blankLayoutPart) }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }),
            new P.SlideIdList(),
            new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 },
            new P.DefaultTextStyle());
        presentationPart.Presentation.Save();

        doc.Save();
    }

    /// <summary>
    /// Creates a minimal template PPTX with a single title-type layout that has
    /// explicit transforms on both the title and body (subtitle) placeholders.
    /// </summary>
    private static void CreateTemplateWithExplicitPlaceholders(
        string path, PlaceholderBounds titleBounds, PlaceholderBounds bodyBounds)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        static P.Shape MakePlaceholder(uint id, string name, P.PlaceholderValues phType, PlaceholderBounds bounds)
            => new(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = phType })),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = bounds.X * 12700L, Y = bounds.Y * 12700L },
                        new A.Extents { Cx = bounds.W * 12700L, Cy = bounds.H * 12700L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));

        var titleLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        titleLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })),
                MakePlaceholder(2U, "Title Placeholder", P.PlaceholderValues.Title, titleBounds),
                MakePlaceholder(3U, "Body Placeholder", P.PlaceholderValues.SubTitle, bodyBounds))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        { Type = P.SlideLayoutValues.Title };
        titleLayoutPart.AddPart(slideMasterPart, "rId1");
        titleLayoutPart.SlideLayout.Save();

        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId2");
        themePart.Theme = new A.Theme
        {
            Name = "Test",
            ThemeElements = new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "1F2937" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "F8FAFC" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "0F766E" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "2563EB" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "F59E0B" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "DC2626" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "7C3AED" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "0891B2" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "2563EB" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "7C3AED" }))
                { Name = "Test" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "Test" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 6350 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 12700 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 19050 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })))
                { Name = "Test" }),
            ObjectDefaults = new A.ObjectDefaults(),
            ExtraColorSchemeList = new A.ExtraColorSchemeList(),
        };
        themePart.Theme.Save();

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new P.SlideLayoutIdList(
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(titleLayoutPart) }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }),
            new P.SlideIdList(),
            new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 },
            new P.DefaultTextStyle());
        presentationPart.Presentation.Save();

        doc.Save();
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

    private sealed record PlaceholderBounds(long X, long Y, long W, long H);
}
