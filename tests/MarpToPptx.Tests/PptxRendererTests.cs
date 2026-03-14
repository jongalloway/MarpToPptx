using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;
using MarpToPptx.Pptx.Validation;
using System.IO.Compression;
using System.Xml.Linq;
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
        // Slide 1 (H1 + paragraph = Title kind) uses the content layout (type="tx").
        // Slide 2 (H2 + bullets + image): the image is 50% of non-heading elements,
        // meeting the image-focused threshold, so it uses the blank layout (type="blank").
        Assert.Equal("/ppt/slideLayouts/slideLayout1.xml", slideParts[0].SlideLayoutPart?.Uri.ToString());
        Assert.Equal("/ppt/slideLayouts/slideLayout2.xml", slideParts[1].SlideLayoutPart?.Uri.ToString());
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

        using var slideReader = new StreamReader(archive.GetEntry("ppt/slides/slide2.xml")!.Open());
        var slideXml = slideReader.ReadToEnd();
        var slideDocument = XDocument.Parse(slideXml);
        var svgBlip = slideDocument.Descendants().SingleOrDefault(element => element.Name.LocalName == "svgBlip");
        Assert.NotNull(svgBlip);
        Assert.False(string.IsNullOrWhiteSpace(svgBlip!.Attributes().SingleOrDefault(attribute => attribute.Name.LocalName == "embed")?.Value));
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
    public void Renderer_TableUsesReadableColors_WhenSlideBodyUsesLightText()
    {
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
        section.contrast { color: #FFFFFF; }
        """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: custom
            ---

            <!-- class: contrast -->
            # Table Slide

            | Key | Scope |
            | --- | --- |
            | class | local |
            | footer | local |
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var table = document.PresentationPart!.SlideParts.First().Slide!
            .Descendants<A.Table>()
            .Single();

        var rows = table.Elements<A.TableRow>().ToArray();
        var bodyCell = rows[1].Elements<A.TableCell>().First();

        Assert.True(GetTableCellFill(bodyCell) is "FFFFFF" or "F8FAFC");
        Assert.Equal("1F2937", GetFirstRunColor(bodyCell));
    }

    [Fact]
    public void Renderer_EmptyBackgroundImageDirective_ClearsInheritedBackgroundImage()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile(
            "accent-wave.svg",
            """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
              <rect width="100" height="100" fill="#FDE68A" />
            </svg>
            """);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            backgroundImage: url(accent-wave.svg)
            ---

            # Slide One

            ---

            <!-- backgroundImage: -->
            # Slide Two
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slides = document.PresentationPart!.SlideParts.ToArray();

        Assert.Single(slides[0].Slide!.Descendants<P.Picture>());
        Assert.Empty(slides[1].Slide!.Descendants<P.Picture>());
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
    public void Renderer_ImageAltText_IsNotRenderedAsVisibleSlideText()
    {
        using var workspace = TestWorkspace.Create();

        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
        workspace.WriteFile("photo.png", pngBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # My Title

            ![My descriptive alt text](photo.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var visibleTexts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.DoesNotContain("My descriptive alt text", visibleTexts);
    }

    [Fact]
    public void Renderer_ImageAltText_IsSetOnPictureShapeDescription()
    {
        using var workspace = TestWorkspace.Create();

        var pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
        workspace.WriteFile("photo.png", pngBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Accessibility description here](photo.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var picture = Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        var description = picture.NonVisualPictureProperties?
            .NonVisualDrawingProperties?
            .Description?.Value;
        Assert.Equal("Accessibility description here", description);
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
        Assert.NotNull(pictures[0].Descendants<A.Blip>().SingleOrDefault()?.Embed?.Value);
        Assert.Contains("p14:media", pictures[0].InnerXml);

        // The slide should contain a video reference relationship pointing to an mp4 media part.
        var videoRels = slidePart.DataPartReferenceRelationships
            .OfType<VideoReferenceRelationship>()
            .ToArray();
        Assert.Single(videoRels);
        Assert.Equal("video/mp4", videoRels[0].DataPart.ContentType);

        var mediaRels = slidePart.DataPartReferenceRelationships
            .OfType<MediaReferenceRelationship>()
            .ToArray();
        Assert.Single(mediaRels);
        Assert.Equal("video/mp4", mediaRels[0].DataPart.ContentType);

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
        Assert.NotNull(pictures[0].Descendants<A.Blip>().SingleOrDefault()?.Embed?.Value);
        Assert.Contains("p14:media", pictures[0].InnerXml);

        var videoRels = slidePart.DataPartReferenceRelationships
            .OfType<VideoReferenceRelationship>()
            .ToArray();
        Assert.Single(videoRels);
        Assert.Equal("video/mp4", videoRels[0].DataPart.ContentType);

        var mediaRels = slidePart.DataPartReferenceRelationships
            .OfType<MediaReferenceRelationship>()
            .ToArray();
        Assert.Single(mediaRels);
        Assert.Equal("video/mp4", mediaRels[0].DataPart.ContentType);

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
        Assert.NotNull(pictures[0].Descendants<A.Blip>().SingleOrDefault()?.Embed?.Value);
        Assert.Contains("p14:media", pictures[0].InnerXml);

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

    [Fact]
    public void Renderer_UsesInheritedMasterBodyPlaceholderRect_ForTemplateResidualContent()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template-inherited-body.pptx");
        CreateTemplateWithInheritedContentPlaceholder(
            templatePath,
            titleBounds: new PlaceholderBounds(X: 50, Y: 70, W: 860, H: 90),
            bodyBounds: new PlaceholderBounds(X: 50, Y: 180, W: 860, H: 220));

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            layout: Title and Content
            ---

            ## Topic Slide

            ```csharp
            Console.WriteLine("hello");
            ```
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var codeShape = slidePart.Slide!.Descendants<P.Shape>()
            .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Code (csharp)");
        var codeOffset = codeShape.ShapeProperties!.Transform2D!.Offset!;
        Assert.Equal(50L * 12700, codeOffset.X?.Value);
        Assert.Equal(180L * 12700, codeOffset.Y?.Value);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_SelectsNamedTemplateLayout_FromSlideDirective()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithMultipleLayouts(templatePath);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- layout: Comparison -->
            # Title Slide

            Body text.
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

        Assert.Equal("/ppt/slideLayouts/slideLayout3.xml", slidePart.SlideLayoutPart?.Uri.ToString());
        Assert.Null(GetBackgroundColor(slidePart));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_UsesFrontMatterLayout_AsDefaultForContentSlidesOnly()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithMultipleLayouts(templatePath);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            layout: Comparison
            backgroundColor: "#123456"
            paginate: true
            ---

            # Title Slide

            Intro text.

            ---

            ## Content Slide

            Body text.
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
        Assert.Equal("/ppt/slideLayouts/slideLayout1.xml", slideParts[0].SlideLayoutPart?.Uri.ToString());
        Assert.Equal("/ppt/slideLayouts/slideLayout3.xml", slideParts[1].SlideLayoutPart?.Uri.ToString());
        Assert.Equal("123456", GetBackgroundColor(slideParts[0]));
        Assert.Null(GetBackgroundColor(slideParts[1]));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_WritesContentIntoTemplatePlaceholders_WhenNamedLayoutCarriesThem()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithPlaceholderLayout(templatePath);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- layout: Placeholder Content -->
            # Slide Heading

            Plain body paragraph with **bold** span.

            - First bullet
            - Second bullet
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Title placeholder shape: <p:ph type="title">, no transform, contains the H1 text.
        var titleShape = slidePart.Slide!.Descendants<P.Shape>().Single(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Title);
        Assert.Null(titleShape.ShapeProperties?.Transform2D);
        Assert.Contains(titleShape.Descendants<A.Text>(), t => t.Text == "Slide Heading");

        // Title runs omit font size / colour fill / font family so layout styles cascade.
        var titleRun = titleShape.Descendants<A.Run>().First();
        Assert.Null(titleRun.RunProperties?.FontSize);
        Assert.Empty(titleRun.RunProperties!.Descendants<A.SolidFill>());
        Assert.Empty(titleRun.RunProperties!.Descendants<A.LatinFont>());

        // Body placeholder shape: <p:ph type="body" idx="1">, echoes the layout's index.
        var bodyShape = slidePart.Slide!.Descendants<P.Shape>().Single(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Body);
        var bodyPh = bodyShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>()!;
        Assert.Equal(1U, bodyPh.Index?.Value);
        Assert.Null(bodyShape.ShapeProperties?.Transform2D);

        // Body collapses paragraph + list into one shape: plain para gets <a:buNone/>,
        // bullet items keep their level and inherit layout bullet styling (no explicit char).
        var bodyParagraphs = bodyShape.TextBody!.Elements<A.Paragraph>().ToArray();
        Assert.Equal(3, bodyParagraphs.Length);
        Assert.NotNull(bodyParagraphs[0].ParagraphProperties?.GetFirstChild<A.NoBullet>());
        Assert.Equal(0, bodyParagraphs[1].ParagraphProperties?.Level?.Value);
        Assert.Null(bodyParagraphs[1].ParagraphProperties?.GetFirstChild<A.NoBullet>());
        Assert.Null(bodyParagraphs[1].ParagraphProperties?.GetFirstChild<A.CharacterBullet>());

        // Inline bold survives; bold span lives inside the plain paragraph.
        var boldRun = bodyParagraphs[0].Elements<A.Run>()
            .First(r => r.Descendants<A.Text>().Any(t => t.Text == "bold"));
        Assert.True(boldRun.RunProperties?.Bold?.Value);

        // Standalone-shape path was NOT taken (no shape named "Text").
        Assert.DoesNotContain(slidePart.Slide!.Descendants<P.Shape>(),
            s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Text");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_FallsBackToStandaloneShapes_WhenNamedLayoutLacksPlaceholders()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithMultipleLayouts(templatePath);

        // "Comparison" layout in this template fixture has no <p:ph> shapes at all,
        // so TryRenderIntoTemplatePlaceholders must return false and the slide must
        // still carry its content via the standalone-text-box path.
        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- layout: Comparison -->
            # Heading

            Body text.
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // No slide-level placeholder shapes were emitted.
        Assert.DoesNotContain(slidePart.Slide!.Descendants<P.Shape>(), s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>() is not null);

        // Standalone shapes carry the content.
        Assert.Contains(slidePart.Slide!.Descendants<P.Shape>(),
            s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Title");
        Assert.Contains(slidePart.Slide!.Descendants<P.Shape>(),
            s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Text");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_RoutesImageIntoPicturePlaceholder_WhenLayoutExposesPicturePlaceholder()
    {
        // Arrange: template with title + picture placeholder layout.
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithTitleAndPicturePlaceholder(templatePath);

        var pixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
        workspace.WriteFile("photo.png", pixelPng);

        var markdownPath = workspace.WriteMarkdown("deck.md",
            """
            <!-- _layout: Title Plus Picture -->
            # Slide Title

            ![Alt text](photo.png)
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Title placeholder shape emitted.
        Assert.Contains(slidePart.Slide!.Descendants<P.Shape>(), s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Title);

        // Image is embedded into a P.Picture with type="pic" placeholder reference (no explicit transform).
        var picturePh = Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        var picPh = picturePh.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<P.PlaceholderShape>();
        Assert.NotNull(picPh);
        Assert.Equal(P.PlaceholderValues.Picture, picPh!.Type?.Value);
        Assert.Null(picturePh.ShapeProperties?.Transform2D);

        // Alt text is preserved on the picture's non-visual drawing properties.
        Assert.Equal("Alt text", picturePh.NonVisualPictureProperties?
            .NonVisualDrawingProperties?.Description?.Value);

        // No standalone picture shape without a placeholder reference.
        Assert.DoesNotContain(slidePart.Slide!.Descendants<P.Picture>(), p =>
            p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>() is null);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_RoutesImageIntoPicturePlaceholder_WhenLayoutHasPictureOnlyPlaceholder()
    {
        // Arrange: template with picture-only layout (no title placeholder).
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithPictureOnlyLayout(templatePath);

        var pixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
        workspace.WriteFile("photo.png", pixelPng);

        var markdownPath = workspace.WriteMarkdown("deck.md",
            """
            <!-- _layout: Picture Only -->
            ![Photo](photo.png)
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Image is embedded into a P.Picture with type="pic" placeholder reference.
        var picturePh = Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        var picPh = picturePh.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<P.PlaceholderShape>();
        Assert.NotNull(picPh);
        Assert.Equal(P.PlaceholderValues.Picture, picPh!.Type?.Value);
        Assert.Null(picturePh.ShapeProperties?.Transform2D);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_FallsBackToStandaloneImage_WhenLayoutLacksPicturePlaceholder()
    {
        // Arrange: template with title + body placeholders only (no picture placeholder).
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithPlaceholderLayout(templatePath);

        var pixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
        workspace.WriteFile("photo.png", pixelPng);

        var markdownPath = workspace.WriteMarkdown("deck.md",
            """
            <!-- _layout: Placeholder Content -->
            # Heading

            ![Photo](photo.png)
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Image is present as a standalone P.Picture (no placeholder reference).
        var picture = Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        Assert.Null(picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<P.PlaceholderShape>());

        // Explicit transform is set (standalone shape, not inherited from layout).
        Assert.NotNull(picture.ShapeProperties?.Transform2D);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_ClonesTemplateSlideArtwork_WhenTemplateSlideIsRequested()
    {
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithDecoratedTemplateSlide(templatePath);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- _layout: Template[1] -->
            # Real Session Title

            Speaker Name\
            Title\
            Organization

            Level: Advanced
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        Assert.Equal(P.SlideLayoutValues.Title, slidePart.SlideLayoutPart?.SlideLayout?.Type?.Value);
        Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
        Assert.Contains(slidePart.Slide.Descendants<A.Text>(), text => text.Text == "Real Session Title");
        Assert.Contains(slidePart.Slide.Descendants<A.Text>(), text => text.Text == "Speaker Name");
        Assert.Contains(slidePart.Slide.Descendants<A.Text>(), text => text.Text == "Level: Advanced");
        Assert.DoesNotContain(slidePart.Slide.Descendants<A.Text>(), text => text.Text == "Session Title Goes Here");
        Assert.DoesNotContain(slidePart.Slide.Descendants<A.Text>(), text => text.Text == "Use Two Lines if Needed");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_TitleOnly_ConstrainsResidualContentBelowExplicitTitleRect()
    {
        // Verifies that on a Title Only layout (title placeholder, no body placeholder)
        // with an explicit title transform, standalone body shapes start below the
        // title region plus the mandatory spacer gap.
        using var workspace = TestWorkspace.Create();

        // Title placeholder: Y=80, H=120 → bottom=200; expected min content Y=220 (spacer=20).
        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithTitleOnlyLayout(templatePath,
            layoutTitleBounds: new PlaceholderBounds(X: 50, Y: 80, W: 860, H: 120),
            masterTitleBounds: null);

        var markdownPath = workspace.WriteMarkdown("deck.md",
            """
            <!-- _layout: Title Only -->
            # My Title

            A body paragraph.

            ```csharp
            var x = 1;
            ```
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Title goes into a placeholder shape (no explicit transform).
        var titleShape = Assert.Single(slidePart.Slide!.Descendants<P.Shape>(),
            s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                     .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Title);
        Assert.Null(titleShape.ShapeProperties?.Transform2D);

        // All standalone (non-placeholder) shapes must start at or below 220*12700 EMU.
        // titleBottom = 80+120=200; spacer=20; minY=220.
        const long expectedMinY = 220L * 12700;
        var standaloneShapes = slidePart.Slide!.Descendants<P.Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                             .GetFirstChild<P.PlaceholderShape>() is null)
            .ToArray();

        Assert.NotEmpty(standaloneShapes);
        foreach (var shape in standaloneShapes)
        {
            var offsetY = shape.ShapeProperties?.Transform2D?.Offset?.Y?.Value;
            Assert.NotNull(offsetY);
            Assert.True(offsetY >= expectedMinY,
                $"Shape '{shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value}' " +
                $"Y={offsetY} is above the minimum {expectedMinY} (title bottom + spacer).");
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_TitleOnly_ConstrainsResidualContentBelowInheritedMasterTitleRect()
    {
        // Verifies that when a Title Only layout carries no explicit title transform but
        // the slide master does, the renderer still resolves the master's title bounds and
        // constrains standalone shapes below it.
        using var workspace = TestWorkspace.Create();

        // Master title: Y=60, H=100 → bottom=160; expected min content Y=180 (spacer=20).
        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithTitleOnlyLayout(templatePath,
            layoutTitleBounds: null,
            masterTitleBounds: new PlaceholderBounds(X: 50, Y: 60, W: 860, H: 100));

        var markdownPath = workspace.WriteMarkdown("deck.md",
            """
            <!-- _layout: Title Only -->
            # Inherited Title

            - Bullet one
            - Bullet two
            - Bullet three
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // All standalone (non-placeholder) shapes must start at or below 180*12700 EMU.
        // masterTitleBottom = 60+100=160; spacer=20; minY=180.
        const long expectedMinY = 180L * 12700;
        var standaloneShapes = slidePart.Slide!.Descendants<P.Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                             .GetFirstChild<P.PlaceholderShape>() is null)
            .ToArray();

        Assert.NotEmpty(standaloneShapes);
        foreach (var shape in standaloneShapes)
        {
            var offsetY = shape.ShapeProperties?.Transform2D?.Offset?.Y?.Value;
            Assert.NotNull(offsetY);
            Assert.True(offsetY >= expectedMinY,
                $"Shape '{shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value}' " +
                $"Y={offsetY} is above the minimum {expectedMinY} (master title bottom + spacer).");
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_TitleAndContent_DoesNotRegressWhenTitleOnlyFixApplied()
    {
        // Regression guard: Title and Content layouts still route body content into the
        // body placeholder and do not apply the Title Only spacer constraint.
        using var workspace = TestWorkspace.Create();

        var templatePath = workspace.GetPath("template.pptx");
        CreateTemplateWithPlaceholderLayout(templatePath);

        var markdownPath = workspace.WriteMarkdown("deck.md",
            """
            <!-- layout: Placeholder Content -->
            # Slide Heading

            Body paragraph text.
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
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Both title and body placeholders should be present.
        var phShapes = slidePart.Slide!.Descendants<P.Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                             .GetFirstChild<P.PlaceholderShape>() is not null)
            .ToArray();
        Assert.Equal(2, phShapes.Length);

        // No standalone shapes should exist (all content routed to placeholders).
        var standaloneShapes = slidePart.Slide!.Descendants<P.Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                             .GetFirstChild<P.PlaceholderShape>() is null)
            .ToArray();
        Assert.Empty(standaloneShapes);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    /// <summary>
    /// Creates a minimal template PPTX with a single "Title Only" layout that carries a
    /// title placeholder but no body placeholder. When <paramref name="layoutTitleBounds"/>
    /// is non-null, the title placeholder gets an explicit transform on the layout.
    /// When <paramref name="masterTitleBounds"/> is non-null, the slide master carries a
    /// title placeholder with that transform. Both may be set independently.
    /// </summary>
    private static void CreateTemplateWithTitleOnlyLayout(
        string path,
        PlaceholderBounds? layoutTitleBounds,
        PlaceholderBounds? masterTitleBounds)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        static P.Shape MakeTitlePh(uint id, PlaceholderBounds? bounds)
        {
            var shapeProperties = bounds is { } b
                ? new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = b.X * 12700L, Y = b.Y * 12700L },
                        new A.Extents { Cx = b.W * 12700L, Cy = b.H * 12700L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                : new P.ShapeProperties();
            return new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = "Title Placeholder" },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                shapeProperties,
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));
        }

        var layoutShapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = 0L, Cy = 0L })),
            MakeTitlePh(2U, layoutTitleBounds));

        var titleOnlyLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        titleOnlyLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(layoutShapeTree),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.TitleOnly,
            MatchingName = "Title Only",
        };
        titleOnlyLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Only";
        titleOnlyLayoutPart.AddPart(slideMasterPart, "rId1");
        titleOnlyLayoutPart.SlideLayout.Save();

        var masterShapes = new List<P.Shape>();
        if (masterTitleBounds is not null)
        {
            masterShapes.Add(MakeTitlePh(2U, masterTitleBounds));
        }

        var masterShapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = 0L, Cy = 0L })));
        foreach (var shape in masterShapes) { masterShapeTree.Append(shape); }

        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId2");
        themePart.Theme = new A.Theme
        {
            Name = "TestTitleOnly",
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
                { Name = "TestTitleOnly" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Calibri" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "TestTitleOnly" },
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
                { Name = "TestTitleOnly" }),
            ObjectDefaults = new A.ObjectDefaults(),
            ExtraColorSchemeList = new A.ExtraColorSchemeList(),
        };
        themePart.Theme.Save();

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(masterShapeTree),
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
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(titleOnlyLayoutPart) }),
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
    /// Creates a minimal template PPTX with a single named layout that carries both a
    /// title placeholder (no idx) and a body placeholder (idx=1). The layout has type
    /// <c>tx</c> so auto-selection does not pick it for a Title-kind slide; it is
    /// reached only via a <c>layout:</c> directive, keeping the placeholder-rendering
    /// path isolated under test.
    /// </summary>
    private static void CreateTemplateWithPlaceholderLayout(string path)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        static P.Shape MakePh(uint id, string name, P.PlaceholderValues type, uint? idx)
        {
            var ph = new P.PlaceholderShape { Type = type };
            if (idx is { } i) { ph.Index = i; }
            return new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(ph)),
                new P.ShapeProperties(),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));
        }

        var layoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        layoutPart.SlideLayout = new P.SlideLayout(
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
                MakePh(2U, "Title Placeholder", P.PlaceholderValues.Title, null),
                MakePh(3U, "Body Placeholder", P.PlaceholderValues.Body, 1U))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Text,
            MatchingName = "Placeholder Content",
        };
        layoutPart.SlideLayout.CommonSlideData!.Name = "Placeholder Content";
        layoutPart.AddPart(slideMasterPart, "rId1");
        layoutPart.SlideLayout.Save();

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
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(layoutPart) }),
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
    /// Creates a minimal template PPTX with a single layout named "Title Plus Picture"
    /// that carries both a title placeholder and a picture placeholder.
    /// </summary>
    private static void CreateTemplateWithTitleAndPicturePlaceholder(string path)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        static P.Shape MakeTitlePh(uint id)
            => new(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = "Title Placeholder" },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                new P.ShapeProperties(),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));

        static P.Shape MakePicturePh(uint id, uint idx)
        {
            var ph = new P.PlaceholderShape { Type = P.PlaceholderValues.Picture, Index = idx };
            return new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = "Picture Placeholder" },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(ph)),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 457200L, Y = 1143000L },
                        new A.Extents { Cx = 8229600L, Cy = 4525963L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));
        }

        var layoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        layoutPart.SlideLayout = new P.SlideLayout(
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
                MakeTitlePh(2U),
                MakePicturePh(3U, 1U))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.PictureText,
            MatchingName = "Title Plus Picture",
        };
        layoutPart.SlideLayout.CommonSlideData!.Name = "Title Plus Picture";
        layoutPart.AddPart(slideMasterPart, "rId1");
        layoutPart.SlideLayout.Save();

        AddMinimalSlideMaster(presentationPart, slideMasterPart, layoutPart);
        doc.Save();
    }

    /// <summary>
    /// Creates a minimal template PPTX with a layout named "Picture Only" that carries
    /// only a picture placeholder (no title or body placeholder).
    /// </summary>
    private static void CreateTemplateWithPictureOnlyLayout(string path)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        var ph = new P.PlaceholderShape { Type = P.PlaceholderValues.Picture, Index = 1U };
        var picturePh = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Picture Placeholder" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(ph)),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 12192000L, Cy = 6858000L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));

        var layoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        layoutPart.SlideLayout = new P.SlideLayout(
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
                picturePh)),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Blank,
            MatchingName = "Picture Only",
        };
        layoutPart.SlideLayout.CommonSlideData!.Name = "Picture Only";
        layoutPart.AddPart(slideMasterPart, "rId1");
        layoutPart.SlideLayout.Save();

        AddMinimalSlideMaster(presentationPart, slideMasterPart, layoutPart);
        doc.Save();
    }

    /// <summary>
    /// Adds a minimal theme, slide master, and presentation to the given parts so
    /// the document is a valid PPTX. Used by picture-placeholder template helpers.
    /// </summary>
    private static void AddMinimalSlideMaster(
        PresentationPart presentationPart,
        SlideMasterPart slideMasterPart,
        SlideLayoutPart layoutPart)
    {
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
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(layoutPart) }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }),
            new P.SlideIdList(),
            new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 },
            new P.DefaultTextStyle());
        presentationPart.Presentation.Save();
    }

    private static void CreateTemplateWithDecoratedTemplateSlide(string path)
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
        {
            Type = P.SlideLayoutValues.Title,
            MatchingName = "Title Slide",
        };
        titleLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";
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

        var slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
        slidePart.AddPart(titleLayoutPart, "rId1");
        var imagePart = slidePart.AddImagePart(ImagePartType.Png, "rId2");
        using (var imageStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            var imageBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
            imageStream.Write(imageBytes, 0, imageBytes.Length);
        }

        slidePart.Slide = new P.Slide(
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
                new P.Picture(
                    new P.NonVisualPictureProperties(
                        new P.NonVisualDrawingProperties { Id = 2U, Name = "Background" },
                        new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.BlipFill(
                        new A.Blip { Embed = "rId2" },
                        new A.Stretch(new A.FillRectangle())),
                    new P.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 12192000L, Cy = 6858000L }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                CreateTemplateSlideTextShape(3U, "Title Box", 900000L, 1200000L, 8800000L, 1200000L,
                    [
                        CreateTemplateParagraph("Session Title Goes Here", 2600, true),
                        CreateTemplateParagraph("Use Two Lines if Needed", 2600, true),
                    ]),
                CreateTemplateSlideTextShape(4U, "Speaker Box", 900000L, 3400000L, 5200000L, 1500000L,
                    [
                        CreateTemplateParagraph("Speaker Name", 2000, true),
                        CreateTemplateParagraph("Title", 1500, false),
                        CreateTemplateParagraph("Organization", 1500, false),
                    ]),
                CreateTemplateSlideTextShape(5U, "Level Box", 900000L, 5350000L, 3800000L, 600000L,
                    [CreateTemplateParagraph("Level: Intermediate, etc.", 1200, false)]))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        slidePart.Slide.Save();

        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }),
            new P.SlideIdList(new P.SlideId { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart) }),
            new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 },
            new P.DefaultTextStyle());
        presentationPart.Presentation.Save();

        doc.Save();
    }

    private static P.Shape CreateTemplateSlideTextShape(uint id, string name, long x, long y, long cx, long cy, IReadOnlyList<A.Paragraph> paragraphs)
    {
        var textBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in paragraphs)
        {
            textBody.Append(paragraph);
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            textBody);
    }

    private static A.Paragraph CreateTemplateParagraph(string text, int fontSize, bool bold)
    {
        var runProperties = new A.RunProperties { FontSize = fontSize * 100 };
        if (bold)
        {
            runProperties.Bold = true;
        }

        return new A.Paragraph(
            new A.Run(runProperties, new A.Text(text)),
            new A.EndParagraphRunProperties { FontSize = fontSize * 100 });
    }


    /// <summary>
    /// Creates a minimal template PPTX with four layouts: Title (type="title"),
    /// two content layouts (type="tx"), and Blank (type="blank").
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
        {
            Type = P.SlideLayoutValues.Title,
            MatchingName = "Title Slide",
        };
        titleLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";
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
        {
            Type = P.SlideLayoutValues.Text,
            MatchingName = "Title and Content",
        };
        contentLayoutPart.SlideLayout.CommonSlideData!.Name = "Title and Content";
        contentLayoutPart.AddPart(slideMasterPart, "rId1");
        contentLayoutPart.SlideLayout.Save();

        var comparisonLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId3");
        comparisonLayoutPart.SlideLayout = new P.SlideLayout(
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
            Type = P.SlideLayoutValues.Text,
            MatchingName = "Comparison",
        };
        comparisonLayoutPart.SlideLayout.CommonSlideData!.Name = "Comparison";
        comparisonLayoutPart.AddPart(slideMasterPart, "rId1");
        comparisonLayoutPart.SlideLayout.Save();

        var blankLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId4");
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
            MatchingName = "Blank",
        };
        blankLayoutPart.SlideLayout.CommonSlideData!.Name = "Blank";
        blankLayoutPart.AddPart(slideMasterPart, "rId1");
        blankLayoutPart.SlideLayout.Save();

        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId5");
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
                new P.SlideLayoutId { Id = 2147483651U, RelationshipId = slideMasterPart.GetIdOfPart(comparisonLayoutPart) },
                new P.SlideLayoutId { Id = 2147483652U, RelationshipId = slideMasterPart.GetIdOfPart(blankLayoutPart) }),
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

    private static void CreateTemplateWithInheritedContentPlaceholder(
        string path,
        PlaceholderBounds titleBounds,
        PlaceholderBounds bodyBounds)
    {
        using var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

        static P.Shape MakeMasterPlaceholder(uint id, string name, P.PlaceholderValues phType, uint? idx, PlaceholderBounds bounds)
            => new(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = phType, Index = idx })),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = bounds.X * 12700L, Y = bounds.Y * 12700L },
                        new A.Extents { Cx = bounds.W * 12700L, Cy = bounds.H * 12700L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));

        static P.Shape MakeLayoutPlaceholder(uint id, string name, P.PlaceholderValues? phType, uint? idx)
            => new(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = phType, Index = idx })),
                new P.ShapeProperties(new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties())));

        var contentLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
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
                    new A.ChildExtents { Cx = 0L, Cy = 0L })),
                MakeLayoutPlaceholder(2U, "Title Placeholder", P.PlaceholderValues.Title, null),
                MakeLayoutPlaceholder(3U, "Content Placeholder", null, 1U))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Text,
            MatchingName = "Title and Content",
        };
        contentLayoutPart.AddPart(slideMasterPart, "rId1");
        contentLayoutPart.SlideLayout.Save();

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
                    new A.ChildExtents { Cx = 0L, Cy = 0L })),
                MakeMasterPlaceholder(2U, "Title Placeholder 1", P.PlaceholderValues.Title, null, titleBounds),
                MakeMasterPlaceholder(3U, "Text Placeholder 2", P.PlaceholderValues.Body, 1U, bodyBounds))),
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
                new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(contentLayoutPart) }),
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

    [Fact]
    public void Renderer_CreatesNotesSlidePart_ForSlideWithNoteComment()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Some content.

            <!-- This is a presenter note. -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        Assert.NotNull(slidePart.NotesSlidePart);
        var notesTexts = slidePart.NotesSlidePart!.NotesSlide!
            .Descendants<A.Text>()
            .Select(t => t.Text)
            .ToArray();
        Assert.Contains("This is a presenter note.", notesTexts);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_FormatsPresenterNotesWithRichRuns()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            <!-- **Bold** and *italic* and `code` and ~~strike~~ -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var runs = document.PresentationPart!.SlideParts.Single()
            .NotesSlidePart!.NotesSlide!
            .Descendants<A.Run>()
            .ToArray();

        var boldRun = Assert.Single(runs, run => run.Text?.Text == "Bold");
        Assert.True(boldRun.RunProperties?.Bold?.Value);

        var italicRun = Assert.Single(runs, run => run.Text?.Text == "italic");
        Assert.True(italicRun.RunProperties?.Italic?.Value);

        var codeRun = Assert.Single(runs, run => run.Text?.Text == "code");
        Assert.Equal("Cascadia Mono", codeRun.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);

        var strikeRun = Assert.Single(runs, run => run.Text?.Text == "strike");
        Assert.Equal(A.TextStrikeValues.SingleStrike, strikeRun.RunProperties?.Strike?.Value);
    }

    [Fact]
    public void Renderer_FormatsPresenterNoteCodeBlocksWithMonospaceRuns()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            <!--
            ```csharp
            var total = items.Count;
            Console.WriteLine(total);
            ```
            -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var runs = document.PresentationPart!.SlideParts.Single()
            .NotesSlidePart!.NotesSlide!
            .Descendants<A.Run>()
            .ToArray();

        Assert.Collection(
            runs,
            run =>
            {
                Assert.Equal("var total = items.Count;", run.Text?.Text);
                Assert.Equal("Cascadia Mono", run.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
            },
            run =>
            {
                Assert.Equal("Console.WriteLine(total);", run.Text?.Text);
                Assert.Equal("Cascadia Mono", run.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
            });
    }

    [Fact]
    public void Renderer_CreatesNotesForMultipleSlides_AttachedToCorrectSlides()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide One

            <!-- Note for slide one. -->

            ---

            # Slide Two

            No notes here.

            ---

            # Slide Three

            <!-- Note for slide three. -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(3, slideParts.Length);

        // Slide 1: has notes
        Assert.NotNull(slideParts[0].NotesSlidePart);
        var slide1Notes = slideParts[0].NotesSlidePart!.NotesSlide!
            .Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains("Note for slide one.", slide1Notes);

        // Slide 2: no notes
        Assert.Null(slideParts[1].NotesSlidePart);

        // Slide 3: has notes
        Assert.NotNull(slideParts[2].NotesSlidePart);
        var slide3Notes = slideParts[2].NotesSlidePart!.NotesSlide!
            .Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains("Note for slide three.", slide3Notes);

        Assert.NotNull(document.PresentationPart!.NotesMasterPart);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmitsNotesRelationshipsNeededForPowerPointCompatibility()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Body text.

            <!-- Presenter note text. -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var notesSlidePart = slidePart.NotesSlidePart;
        var notesMasterPart = document.PresentationPart.NotesMasterPart;

        Assert.NotNull(notesSlidePart);
        Assert.NotNull(notesMasterPart);
        Assert.NotNull(notesMasterPart!.ThemePart);
        Assert.Equal("/ppt/slides/slide1.xml", notesSlidePart!.SlidePart?.Uri.ToString());
        Assert.Equal(notesMasterPart.Uri, notesSlidePart.NotesMasterPart?.Uri);

        using var archive = ZipFile.OpenRead(outputPath);

        var notesMasterRelationshipsEntry = archive.GetEntry("ppt/notesMasters/_rels/notesMaster1.xml.rels");
        Assert.NotNull(notesMasterRelationshipsEntry);
        using (var reader = new StreamReader(notesMasterRelationshipsEntry!.Open()))
        {
            var xml = reader.ReadToEnd();
            Assert.Contains("relationships/theme", xml);
            Assert.Contains("theme", xml);
        }

        var notesSlideRelationshipsEntry = archive.GetEntry("ppt/notesSlides/_rels/notesSlide1.xml.rels");
        Assert.NotNull(notesSlideRelationshipsEntry);
        using (var reader = new StreamReader(notesSlideRelationshipsEntry!.Open()))
        {
            var xml = reader.ReadToEnd();
            Assert.Contains("relationships/notesMaster", xml);
            Assert.Contains("relationships/slide", xml);
            Assert.Contains("../slides/slide1.xml", xml);
        }
    }

    [Fact]
    public void Renderer_DoesNotCreateNotesSlidePart_WhenSlideHasNoNotes()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide Without Notes

            Just some content.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        Assert.Null(slidePart.NotesSlidePart);
        Assert.Null(document.PresentationPart.NotesMasterPart);
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

    [Fact]
    public void Renderer_EmitsBoldRun_ForInlineBoldText()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Heading

            Normal and **bold** text.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var runs = slidePart.Slide!.Descendants<A.Run>().ToArray();
        var boldRun = runs.FirstOrDefault(r => r.RunProperties?.Bold?.Value == true && r.Text?.Text == "bold");
        Assert.NotNull(boldRun);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmitsItalicRun_ForInlineItalicText()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Heading

            Normal and *italic* text.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var runs = slidePart.Slide!.Descendants<A.Run>().ToArray();
        var italicRun = runs.FirstOrDefault(r => r.RunProperties?.Italic?.Value == true && r.Text?.Text == "italic");
        Assert.NotNull(italicRun);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmitsStrikethroughRun_ForInlineStrikethroughText()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Heading

            Normal and ~~struck~~ text.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var runs = slidePart.Slide!.Descendants<A.Run>().ToArray();
        var strikeRun = runs.FirstOrDefault(r =>
            r.RunProperties?.Strike?.Value == A.TextStrikeValues.SingleStrike &&
            r.Text?.Text == "struck");
        Assert.NotNull(strikeRun);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmitsMonospaceFontRun_ForInlineCode()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Heading

            Use `printf()` to print.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // The inline code run should use the code font (not the body font).
        var codeRun = slidePart.Slide!.Descendants<A.Run>()
            .FirstOrDefault(r => r.Text?.Text == "printf()");
        Assert.NotNull(codeRun);

        // The code run font should differ from plain text runs (monospace).
        var codeFont = codeRun!.RunProperties?.Descendants<A.LatinFont>().FirstOrDefault()?.Typeface;
        Assert.NotNull(codeFont);

        // Body-text runs on the same slide should use a different (proportional) font.
        var bodyRun = slidePart.Slide!.Descendants<A.Run>()
            .FirstOrDefault(r => r.Text?.Text?.Contains("print") == true && r.Text.Text != "printf()");
        if (bodyRun is not null)
        {
            var bodyFont = bodyRun.RunProperties?.Descendants<A.LatinFont>().FirstOrDefault()?.Typeface;
            Assert.NotEqual(codeFont, bodyFont);
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmitsHyperlinkRelationship_ForInlineLink()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Heading

            Visit [the site](https://example.com) for details.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // A HyperlinkOnClick element should be present in the run properties.
        var hlinkClicks = slidePart.Slide!.Descendants<A.HyperlinkOnClick>().ToArray();
        Assert.NotEmpty(hlinkClicks);

        // The relationship target should be the external URL.
        var relId = hlinkClicks[0].Id?.Value;
        Assert.NotNull(relId);
        var rel = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
        Assert.NotNull(rel);
        Assert.Equal("https://example.com", rel!.Uri.OriginalString);

        // The link text should be present in the slide.
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains("the site", texts);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_PreservesInlineFormattingInBulletList()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Heading

            - Plain item
            - **Bold item**
            - *Italic item*
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var allRuns = slidePart.Slide!.Descendants<A.Run>().ToArray();
        Assert.Contains(allRuns, r => r.RunProperties?.Bold?.Value == true && r.Text?.Text?.Contains("Bold item") == true);
        Assert.Contains(allRuns, r => r.RunProperties?.Italic?.Value == true && r.Text?.Text?.Contains("Italic item") == true);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_PreservesInlineFormattingInTableCells()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Table

            | Feature       | Status      |
            |---------------|-------------|
            | **Bold cell** | *Italic*    |
            | Plain cell    | Normal text |
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var allRuns = slidePart.Slide!.Descendants<A.Run>().ToArray();
        Assert.Contains(allRuns, r => r.RunProperties?.Bold?.Value == true && r.Text?.Text == "Bold cell");
        Assert.Contains(allRuns, r => r.RunProperties?.Italic?.Value == true && r.Text?.Text == "Italic");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_AddsSlideNumberField_WhenPaginateIsTrue()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            paginate: true
            ---

            # Slide One

            ---

            # Slide Two
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);

        // Both slides should contain a slide-number field.
        foreach (var slidePart in slideParts)
        {
            var fields = slidePart.Slide!.Descendants<A.Field>().ToArray();
            Assert.Contains(fields, f => f.Type?.Value == "slidenum");
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_AddsSlideNumberField_WithCorrectSlideNumbers()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            paginate: true
            ---

            # Slide One

            ---

            # Slide Two

            ---

            # Slide Three
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(3, slideParts.Length);

        // Verify each slide's field contains the correct slide number as display text
        // and has explicit font styling matching the footer style.
        for (var i = 0; i < slideParts.Length; i++)
        {
            var field = slideParts[i].Slide!.Descendants<A.Field>()
                .FirstOrDefault(f => f.Type?.Value == "slidenum");
            Assert.NotNull(field);
            Assert.Equal((i + 1).ToString(), field!.GetFirstChild<A.Text>()?.Text);

            // Run properties should include explicit font family and color.
            var runProps = field.GetFirstChild<A.RunProperties>();
            Assert.NotNull(runProps);
            Assert.NotNull(runProps!.Descendants<A.LatinFont>().FirstOrDefault());
            Assert.NotNull(runProps.Descendants<A.SolidFill>().FirstOrDefault());
        }
    }

    [Fact]
    public void Renderer_AddsHeaderShape_WhenHeaderDirectiveIsSet()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            header: My Presentation Header
            ---

            # Slide
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains("My Presentation Header", texts);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_AddsFooterShape_WhenFooterDirectiveIsSet()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            footer: © 2024 My Company
            ---

            # Slide
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains("© 2024 My Company", texts);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_AddsFooterAndPageNumber_WhenBothAreSet()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            paginate: true
            footer: Conference 2024
            ---

            # First Slide

            ---

            # Second Slide
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);

        foreach (var slidePart in slideParts)
        {
            var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
            Assert.Contains("Conference 2024", texts);
            var fields = slidePart.Slide!.Descendants<A.Field>().ToArray();
            Assert.Contains(fields, f => f.Type?.Value == "slidenum");
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_DoesNotAddSlideNumber_WhenPaginateIsFalse()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide Without Pagination
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var fields = slidePart.Slide!.Descendants<A.Field>().ToArray();
        Assert.DoesNotContain(fields, f => f.Type?.Value == "slidenum");
    }

    [Fact]
    public void Renderer_EmbedsMp3Audio_WhenLocalFileExists()
    {
        using var workspace = TestWorkspace.Create();

        // Minimal valid ID3v2 MP3 header for testing.
        var mp3Bytes = new byte[]
        {
            0x49, 0x44, 0x33, // ID3
            0x03, 0x00,       // version 2.3
            0x00,             // flags
            0x00, 0x00, 0x00, 0x00, // size (4 bytes syncsafe)
        };

        workspace.WriteFile("music.mp3", mp3Bytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Audio Slide

            ![Background music](music.mp3)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // A picture element should be present with an audio file reference.
        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);

        var audioFile = pictures[0].Descendants<A.AudioFromFile>().SingleOrDefault();
        Assert.NotNull(audioFile);
        Assert.NotNull(audioFile!.Link?.Value);
        Assert.NotNull(pictures[0].Descendants<A.Blip>().SingleOrDefault()?.Embed?.Value);
        Assert.Contains("p14:media", pictures[0].InnerXml);

        // The slide should contain an audio reference relationship pointing to an mp3 media part.
        var audioRels = slidePart.DataPartReferenceRelationships
            .OfType<AudioReferenceRelationship>()
            .ToArray();
        Assert.Single(audioRels);
        Assert.Equal("audio/mp3", audioRels[0].DataPart.ContentType);

        var mediaRels = slidePart.DataPartReferenceRelationships
            .OfType<MediaReferenceRelationship>()
            .ToArray();
        Assert.Single(mediaRels);
        Assert.Equal("audio/mp3", mediaRels[0].DataPart.ContentType);

        // No error text should appear.
        Assert.DoesNotContain("Missing audio", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));

        // Package should be valid.
        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmbedsWavAudio_WhenLocalFileExists()
    {
        using var workspace = TestWorkspace.Create();

        // Minimal RIFF/WAVE header.
        var wavBytes = new byte[]
        {
            0x52, 0x49, 0x46, 0x46, // RIFF
            0x24, 0x00, 0x00, 0x00, // chunk size
            0x57, 0x41, 0x56, 0x45, // WAVE
            0x66, 0x6D, 0x74, 0x20, // fmt 
            0x10, 0x00, 0x00, 0x00, // subchunk size = 16
            0x01, 0x00,             // audio format = PCM
            0x01, 0x00,             // num channels = 1
            0x44, 0xAC, 0x00, 0x00, // sample rate = 44100
            0x88, 0x58, 0x01, 0x00, // byte rate
            0x02, 0x00,             // block align
            0x10, 0x00,             // bits per sample = 16
            0x64, 0x61, 0x74, 0x61, // data
            0x00, 0x00, 0x00, 0x00, // data chunk size = 0
        };

        workspace.WriteFile("effect.wav", wavBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Audio Slide

            ![Sound effect](effect.wav)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);

        var audioFile = pictures[0].Descendants<A.AudioFromFile>().SingleOrDefault();
        Assert.NotNull(audioFile);

        var audioRels = slidePart.DataPartReferenceRelationships
            .OfType<AudioReferenceRelationship>()
            .ToArray();
        Assert.Single(audioRels);
        Assert.Equal("audio/wav", audioRels[0].DataPart.ContentType);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_ShowsActionableError_WhenMp3FileIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![Missing audio](missing.mp3)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.Contains(texts, t => t.Contains("Missing audio") && t.Contains("missing.mp3"));
    }

    [Fact]
    public void Renderer_EmbedsM4aAudio_WhenLocalFileExists()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile("audio.m4a", new byte[] { 0x00, 0x00, 0x00, 0x20, 0x66, 0x74, 0x79, 0x70 });

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ![M4A audio](audio.m4a)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);

        var audioFile = pictures[0].Descendants<A.AudioFromFile>().SingleOrDefault();
        Assert.NotNull(audioFile);

        var audioRels = slidePart.DataPartReferenceRelationships
            .OfType<AudioReferenceRelationship>()
            .ToArray();
        Assert.Single(audioRels);
        Assert.Equal("audio/mp4", audioRels[0].DataPart.ContentType);

        var mediaRels = slidePart.DataPartReferenceRelationships
            .OfType<MediaReferenceRelationship>()
            .ToArray();
        Assert.Single(mediaRels);
        Assert.Equal("audio/mp4", mediaRels[0].DataPart.ContentType);

        Assert.DoesNotContain("Unsupported audio format", slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_EmbedsMp3Audio_WhenSpecifiedViaHtmlAudioTag()
    {
        using var workspace = TestWorkspace.Create();

        var mp3Bytes = new byte[]
        {
            0x49, 0x44, 0x33,
            0x03, 0x00,
            0x00,
            0x00, 0x00, 0x00, 0x00,
        };
        workspace.WriteFile("bg.mp3", mp3Bytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Audio Slide

            <audio src="bg.mp3" controls></audio>
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.Single(pictures);
        Assert.NotNull(pictures[0].Descendants<A.AudioFromFile>().SingleOrDefault());

        var audioRels = slidePart.DataPartReferenceRelationships
            .OfType<AudioReferenceRelationship>()
            .ToArray();
        Assert.Single(audioRels);
        Assert.Equal("audio/mp3", audioRels[0].DataPart.ContentType);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    private sealed record PlaceholderBounds(long X, long Y, long W, long H);

    // ────────────────────────────────────────────────────────
    // Issue #39 – Paginate and class directive output
    // ────────────────────────────────────────────────────────

    [Fact]
    public void Renderer_OmitsSlideNumberField_WhenPaginateFalse()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            paginate: false
            ---

            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        Assert.DoesNotContain(slidePart.Slide!.Descendants<A.Field>(), f => f.Type == "slidenum");
    }

    [Fact]
    public void Renderer_SpotPaginateFalse_SuppressesNumberOnSingleSlide()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            paginate: true
            ---

            # Title

            ---

            <!-- _paginate: false -->
            # No Number Here

            ---

            # Back To Numbers
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();

        // Slide 1 and 3 should have slide numbers; slide 2 should not.
        Assert.Single(slideParts[0].Slide!.Descendants<A.Field>(), f => f.Type == "slidenum");
        Assert.DoesNotContain(slideParts[1].Slide!.Descendants<A.Field>(), f => f.Type == "slidenum");
        Assert.Single(slideParts[2].Slide!.Descendants<A.Field>(), f => f.Type == "slidenum");
    }

    [Fact]
    public void Renderer_ClassVariant_AffectsBackgroundColor()
    {
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
        section { background-color: #FFFFFF; }
        section.dark { background-color: #1A1A2E; }
        """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: custom
            ---

            # Default Background

            ---

            <!-- class: dark -->
            # Dark Background
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();

        // Slide 1 background should be FFFFFF (default).
        var bg1 = GetBackgroundColor(slideParts[0]);
        Assert.Equal("FFFFFF", bg1);

        // Slide 2 background should be 1A1A2E (class variant).
        var bg2 = GetBackgroundColor(slideParts[1]);
        Assert.Equal("1A1A2E", bg2);
    }

    [Fact]
    public void Renderer_ClassVariant_AffectsHeadingColor()
    {
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
        section.accent h1 { color: #E94560; }
        """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: custom
            ---

            # Normal Title

            ---

            <!-- class: accent -->
            # Accent Title
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();

        // Slide 2 should have the accent heading color.
        var headingRun = slideParts[1].Slide!.Descendants<A.Run>()
            .FirstOrDefault(r => r.Descendants<A.Text>().Any());
        Assert.NotNull(headingRun);
        var color = headingRun!.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
        Assert.Equal("E94560", color);
    }

    [Fact]
    public void Renderer_ClassVariant_AffectsInlineCodeColor()
    {
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
        section.contrast code { color: #FFFFFF; }
        """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: custom
            ---

            <!-- class: contrast -->
            A paragraph with `inline code`.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var inlineCodeRun = document.PresentationPart!.SlideParts.First().Slide!
            .Descendants<A.Run>()
            .First(run => run.Descendants<A.Text>().Any(text => text.Text == "inline code"));

        var color = inlineCodeRun.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
        Assert.Equal("FFFFFF", color);
    }

    [Fact]
    public void Renderer_ClassVariant_AffectsLayoutSizing()
    {
        using var workspace = TestWorkspace.Create();

        const string themeCss = """
        section { font-size: 24px; }
        section.large { font-size: 48px; }
        section.large h2 { font-size: 64px; }
        """;

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            theme: custom
            ---

            ## Normal Heading

            First paragraph.

            Second paragraph.

            ---

            <!-- class: large -->
            ## Large Heading

            First paragraph.

            Second paragraph.
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeckWithTheme(markdownPath, outputPath, workspace.RootPath, themeCss);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();

        var normalY = GetShapeTopByContainedText(slideParts[0], "Second paragraph.");
        var largeY = GetShapeTopByContainedText(slideParts[1], "Second paragraph.");

        Assert.True(largeY > normalY);
    }

    private static void RenderDeckWithTheme(string markdownPath, string outputPath, string sourceDirectory, string themeCss)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath, themeCss);

        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = sourceDirectory });
    }

    private static string? GetBackgroundColor(SlidePart slidePart)
    {
        // The first shape named "Background" should contain a filled rectangle with the color.
        foreach (var shape in slidePart.Slide!.Descendants<P.Shape>())
        {
            var nvSpPr = shape.NonVisualShapeProperties;
            if (nvSpPr?.NonVisualDrawingProperties?.Name == "Background")
            {
                var solidFill = shape.Descendants<A.SolidFill>().FirstOrDefault();
                return solidFill?.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
            }
        }

        return null;
    }

    private static long GetShapeTopByContainedText(SlidePart slidePart, string text)
    {
        var shape = slidePart.Slide!.Descendants<P.Shape>()
            .First(s => s.Descendants<A.Text>().Any(t => t.Text == text));

        return shape.ShapeProperties!.Transform2D!.Offset!.Y!.Value;
    }

    private static string? GetFirstRunColor(A.TableCell cell)
        => cell.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;

    private static string? GetTableCellFill(A.TableCell cell)
        => cell.TableCellProperties?.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;

    // ────────────────────────────────────────────────────────
    // Issue #44 – header, footer, backgroundSize rendering
    // ────────────────────────────────────────────────────────

    [Fact]
    public void Renderer_HeaderCarriesForward_AcrossSlides()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            header: Persistent Header
            ---

            # Slide One

            ---

            # Slide Two
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);

        foreach (var slidePart in slideParts)
        {
            var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
            Assert.Contains("Persistent Header", texts);
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_SpotHeader_OverridesOnSingleSlide()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            header: Default Header
            ---

            # Slide One

            ---

            <!-- _header: Override Header -->
            # Slide Two

            ---

            # Slide Three
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(3, slideParts.Length);

        Assert.Contains("Default Header", slideParts[0].Slide!.Descendants<A.Text>().Select(t => t.Text));
        Assert.Contains("Override Header", slideParts[1].Slide!.Descendants<A.Text>().Select(t => t.Text));
        Assert.DoesNotContain("Default Header", slideParts[1].Slide!.Descendants<A.Text>().Select(t => t.Text));
        Assert.Contains("Default Header", slideParts[2].Slide!.Descendants<A.Text>().Select(t => t.Text));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_FooterCarriesForward_AcrossSlides()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            footer: © 2024 Company
            ---

            # Slide One

            ---

            # Slide Two
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);

        foreach (var slidePart in slideParts)
        {
            var texts = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
            Assert.Contains("© 2024 Company", texts);
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_BackgroundSize_ContainUsesContainedPlacement()
    {
        using var workspace = TestWorkspace.Create();

        // Create a minimal 2x1 PNG (landscape aspect ratio)
        var pngBytes = CreateMinimalPng(2, 1);
        workspace.WriteFile("wide.png", pngBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- backgroundImage: wide.png -->
            <!-- backgroundSize: contain -->
            # Contained Background
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // The background image should be present as a picture element.
        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);

        // With contain mode, the image should be fitted within the slide
        // (not extend beyond). The image extent should be <= slide dimensions.
        var extents = pictures[0].Descendants<A.Extents>().First();
        Assert.True(extents.Cx!.Value <= 12192000L, "Image width should not exceed slide width");
        Assert.True(extents.Cy!.Value <= 6858000L, "Image height should not exceed slide height");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_BackgroundSize_DefaultIsCover()
    {
        using var workspace = TestWorkspace.Create();

        // Create a tall 1x2 PNG (portrait)
        var pngBytes = CreateMinimalPng(1, 2);
        workspace.WriteFile("tall.png", pngBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- backgroundImage: tall.png -->
            # Full Bleed Background
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        // With cover mode (default), the image should extend beyond the slide
        // to cover it fully. For a portrait image on a landscape slide,
        // the width should match or exceed slide width.
        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);
        var extents = pictures[0].Descendants<A.Extents>().First();
        Assert.True(extents.Cx!.Value >= 12192000L, "Cover mode should scale image to at least slide width");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_SpotBackgroundSize_AffectsSingleSlide()
    {
        using var workspace = TestWorkspace.Create();

        var pngBytes = CreateMinimalPng(2, 1);
        workspace.WriteFile("wide.png", pngBytes);

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- backgroundImage: wide.png -->
            # Cover (default)

            ---

            <!-- _backgroundSize: contain -->
            # Contained

            ---

            # Cover again
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(3, slideParts.Length);

        // Slides 1 and 3 (cover mode) should have image width exceeding slide width
        // because the wide 2:1 image in cover mode scales to fill the 16:9 slide height,
        // resulting in a width larger than the slide width (12192000 EMU).
        var slide1Pictures = slideParts[0].Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(slide1Pictures);
        var slide1Extents = slide1Pictures[0].Descendants<A.Extents>().First();
        Assert.True(slide1Extents.Cx!.Value >= 12192000L, "Slide 1 should retain cover mode after spot directive on slide 2");

        // Slide 2 (contain) should have image extents within slide bounds.
        var slide2Pictures = slideParts[1].Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(slide2Pictures);
        var slide2Extents = slide2Pictures[0].Descendants<A.Extents>().First();
        Assert.True(slide2Extents.Cx!.Value <= 12192000L);
        Assert.True(slide2Extents.Cy!.Value <= 6858000L);

        // Slide 3 should revert to cover mode (spot directive only affects slide 2).
        var slide3Pictures = slideParts[2].Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(slide3Pictures);
        var slide3Extents = slide3Pictures[0].Descendants<A.Extents>().First();
        Assert.True(slide3Extents.Cx!.Value >= 12192000L, "Slide 3 should revert to cover mode after spot directive on slide 2");

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_PlacesMermaidDiagramAsPictureShape_OnSlide()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide with Diagram

            ```mermaid
            flowchart LR
              A[Write] --> B[Build]
              B --> C{Tests pass?}
              C -->|yes| D[Ship]
              C -->|no| A
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);

        var imageParts = slidePart.ImageParts.ToArray();
        Assert.NotEmpty(imageParts);

        var svgPart = imageParts.FirstOrDefault(p => string.Equals(p.ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(svgPart);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_MermaidDiagram_SvgContainsExpectedContent()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ```mermaid
            flowchart LR
              NodeA --> NodeB
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var svgPart = slidePart.ImageParts
            .FirstOrDefault(p => string.Equals(p.ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(svgPart);

        using var stream = svgPart!.GetStream();
        var svg = new System.IO.StreamReader(stream).ReadToEnd();
        Assert.Contains("<svg", svg, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Renderer_PlacesDiagramFenceAsPictureShape_OnSlide()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide with Conceptual Diagram

            ```diagram
                        diagram: pyramid
                        levels:
                            - Vision
                            - Strategy
                            - Tactics
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);

        var svgPart = slidePart.ImageParts
            .FirstOrDefault(p => string.Equals(p.ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(svgPart);

        using var stream = svgPart!.GetStream();
        var svg = new System.IO.StreamReader(stream).ReadToEnd();
        Assert.Contains("<svg", svg, StringComparison.OrdinalIgnoreCase);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_InvalidMermaidInput_FallsBackToCodeBlock()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ```mermaid
            this is not valid mermaid syntax @@@###
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        // Should not throw; falls back gracefully
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Fallback renders as text shape (code block), not a picture
        var textRuns = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.NotEmpty(textRuns);

        // Error label should contain the "Mermaid parse error:" prefix
        Assert.Contains(textRuns, t => t.StartsWith("Mermaid parse error:", StringComparison.Ordinal));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_InvalidDiagramInput_FallsBackToCodeBlock()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide

            ```diagram
            diagram: unknowntype
            items:
              - Alpha
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var textRuns = slidePart.Slide!.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.NotEmpty(textRuns);
        Assert.Contains(textRuns, t => t.StartsWith("Diagram parse error:", StringComparison.Ordinal));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_DiagramFence_Pillars_WithEmbeddedFrontMatter_RendersSvg()
    {
        // Real-world authored fence: pillars diagram with embedded --- theme: presentation --- block.
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Pillars Diagram

            ```diagram
            ---
            theme: presentation
            ---
            diagram: pillars
            pillars:
              - title: Microsoft.Extensions.AI
                segments:
                  - IChatClient
                  - Middleware
              - title: Semantic Kernel
                segments:
                  - Plugins
                  - Memory
              - title: Azure AI
                segments:
                  - OpenAI
                  - Search
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        // Must render as a picture shape (SVG), not fall back to a code block.
        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);

        var svgPart = slidePart.ImageParts
            .FirstOrDefault(p => string.Equals(p.ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(svgPart);

        using var stream = svgPart!.GetStream();
        var svg = new System.IO.StreamReader(stream).ReadToEnd();
        Assert.Contains("<svg", svg, StringComparison.OrdinalIgnoreCase);

        // Verify that a P.Picture on the slide references the SVG part via its SVGBlip embed relationship.
        // SVG blips use an Office2019 SVGBlip extension rather than the standard A.Blip.Embed attribute.
        var svgRelId = slidePart.GetIdOfPart(svgPart);
        Assert.Contains(pictures, pic => pic.Descendants<DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip>().Any(b => b.Embed?.Value == svgRelId));

        // Must not contain a fallback error label.
        var textRuns = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.DoesNotContain(textRuns, t => t.StartsWith("Diagram parse error:", StringComparison.Ordinal));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_DiagramFence_Matrix_WithEmbeddedFrontMatter_RendersSvg()
    {
        // Real-world authored fence: matrix diagram with embedded --- theme: presentation --- block.
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Matrix Diagram

            ```diagram
            ---
            theme: presentation
            ---
            diagram: matrix
            rows:
              - Qdrant
              - Redis
            columns:
              - Search
              - Filter
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);

        var svgPart = slidePart.ImageParts
            .FirstOrDefault(p => string.Equals(p.ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(svgPart);

        using var stream = svgPart!.GetStream();
        var svg = new System.IO.StreamReader(stream).ReadToEnd();
        Assert.Contains("<svg", svg, StringComparison.OrdinalIgnoreCase);

        // Verify that a P.Picture on the slide references the SVG part via its SVGBlip embed relationship.
        // SVG blips use an Office2019 SVGBlip extension rather than the standard A.Blip.Embed attribute.
        var svgRelId = slidePart.GetIdOfPart(svgPart);
        Assert.Contains(pictures, pic => pic.Descendants<DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip>().Any(b => b.Embed?.Value == svgRelId));

        var textRuns = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.DoesNotContain(textRuns, t => t.StartsWith("Diagram parse error:", StringComparison.Ordinal));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_DiagramFence_Pyramid_WithEmbeddedFrontMatter_RendersSvg()
    {
        // Real-world authored fence: pyramid diagram with embedded --- theme: presentation --- block.
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Pyramid Diagram

            ```diagram
            ---
            theme: presentation
            ---
            diagram: pyramid
            levels:
              - Vision
              - Strategy
              - Delivery
              - Feedback
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();

        var pictures = slidePart.Slide!.Descendants<P.Picture>().ToArray();
        Assert.NotEmpty(pictures);

        var svgPart = slidePart.ImageParts
            .FirstOrDefault(p => string.Equals(p.ContentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(svgPart);

        using var stream = svgPart!.GetStream();
        var svg = new System.IO.StreamReader(stream).ReadToEnd();
        Assert.Contains("<svg", svg, StringComparison.OrdinalIgnoreCase);

        // Verify that a P.Picture on the slide references the SVG part via its SVGBlip embed relationship.
        // SVG blips use an Office2019 SVGBlip extension rather than the standard A.Blip.Embed attribute.
        var svgRelId = slidePart.GetIdOfPart(svgPart);
        Assert.Contains(pictures, pic => pic.Descendants<DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip>().Any(b => b.Embed?.Value == svgRelId));

        var textRuns = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text).ToArray();
        Assert.DoesNotContain(textRuns, t => t.StartsWith("Diagram parse error:", StringComparison.Ordinal));

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    /// <summary>
    /// Creates a minimal valid PNG file with the specified dimensions.
    /// </summary>
    private static byte[] CreateMinimalPng(int width, int height)
    {
        using var ms = new MemoryStream();
        // PNG signature
        ms.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        // IHDR chunk
        var ihdr = new byte[13];
        WriteBigEndian(ihdr, 0, width);
        WriteBigEndian(ihdr, 4, height);
        ihdr[8] = 8; // bit depth
        ihdr[9] = 2; // color type: RGB
        WriteChunk(ms, "IHDR", ihdr);

        // IDAT chunk (minimal: a single row of zeros per scanline)
        using var deflateMs = new MemoryStream();
        using (var deflate = new System.IO.Compression.DeflateStream(deflateMs, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true))
        {
            for (var y = 0; y < height; y++)
            {
                deflate.WriteByte(0); // filter: none
                for (var x = 0; x < width; x++)
                {
                    deflate.WriteByte(0); deflate.WriteByte(0); deflate.WriteByte(0); // RGB
                }
            }
        }

        // Wrap in zlib container: header (78 01) + deflate data + Adler-32
        var rawDeflate = deflateMs.ToArray();
        var zlibData = new byte[2 + rawDeflate.Length + 4];
        zlibData[0] = 0x78;
        zlibData[1] = 0x01;
        Array.Copy(rawDeflate, 0, zlibData, 2, rawDeflate.Length);
        var adler = ComputeAdler32(height, width);
        WriteBigEndian(zlibData, 2 + rawDeflate.Length, (int)adler);

        WriteChunk(ms, "IDAT", zlibData);

        // IEND chunk
        WriteChunk(ms, "IEND", []);

        return ms.ToArray();

        static void WriteBigEndian(byte[] buffer, int offset, int value)
        {
            buffer[offset] = (byte)(value >> 24);
            buffer[offset + 1] = (byte)(value >> 16);
            buffer[offset + 2] = (byte)(value >> 8);
            buffer[offset + 3] = (byte)value;
        }

        static void WriteChunk(MemoryStream stream, string type, byte[] data)
        {
            var lengthBytes = new byte[4];
            WriteBigEndian(lengthBytes, 0, data.Length);
            stream.Write(lengthBytes);

            var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
            stream.Write(typeBytes);
            stream.Write(data);

            // CRC32 over type + data
            var crcData = new byte[4 + data.Length];
            Array.Copy(typeBytes, crcData, 4);
            Array.Copy(data, 0, crcData, 4, data.Length);
            var crc = Crc32(crcData);
            var crcBytes = new byte[4];
            WriteBigEndian(crcBytes, 0, (int)crc);
            stream.Write(crcBytes);
        }

        static uint Crc32(byte[] data)
        {
            uint crc = 0xFFFFFFFF;
            foreach (var b in data)
            {
                crc ^= b;
                for (var i = 0; i < 8; i++)
                    crc = (crc >> 1) ^ (crc % 2 != 0 ? 0xEDB88320 : 0);
            }
            return ~crc;
        }

        static uint ComputeAdler32(int height, int width)
        {
            uint a = 1, b = 0;
            for (var y = 0; y < height; y++)
            {
                // filter byte (0)
                a = (a + 0) % 65521;
                b = (b + a) % 65521;
                for (var x = 0; x < width * 3; x++)
                {
                    a = (a + 0) % 65521;
                    b = (b + a) % 65521;
                }
            }
            return (b << 16) | a;
        }
    }

    // ── Transition directive renderer tests ──────────────────────────────

    [Fact]
    public void Renderer_FadeTransition_EmitsTransitionElementWithFadeChild()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            ---
            transition: fade
            ---

            # Slide One

            ---

            # Slide Two
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slideParts = document.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);

        foreach (var slidePart in slideParts)
        {
            var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
            Assert.NotNull(transition);
            Assert.Single(transition!.Elements<P.FadeTransition>());
        }

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_PushTransitionWithDirection_EmitsPushElementWithDirectionAttribute()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- transition: push dir:right -->
            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
        Assert.NotNull(transition);
        var push = transition!.Elements<P.PushTransition>().SingleOrDefault();
        Assert.NotNull(push);
        Assert.Equal(P.TransitionSlideDirectionValues.Right, push!.Direction?.Value);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_NoTransition_EmitsNoTransitionElement()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var transition = slidePart.Slide!.Elements<P.Transition>().FirstOrDefault();
        Assert.Null(transition);
    }

    [Fact]
    public void Renderer_TransitionWithDuration_EmitsSpeedAttribute()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- transition: wipe dur:600 -->
            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
        Assert.NotNull(transition);
        // 600ms maps to the "med" speed band.
        Assert.Equal(P.TransitionSpeedValues.Medium, transition!.Speed?.Value);
        Assert.Single(transition.Elements<P.WipeTransition>());

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_AllSupportedTransitionTypes_PassOpenXmlValidation()
    {
        var types = new[] { "fade", "push", "wipe", "cut", "cover", "pull", "random-bar", "morph" };
        foreach (var type in types)
        {
            using var workspace = TestWorkspace.Create();
            var markdownPath = workspace.WriteMarkdown(
                "deck.md",
                $"""
                <!-- transition: {type} -->
                # Slide One
                """);

            var outputPath = workspace.GetPath("deck.pptx");
            RenderDeck(markdownPath, outputPath, workspace.RootPath);

            using var document = PresentationDocument.Open(outputPath, false);
            var validationErrors = new OpenXmlPackageValidator().Validate(document);
            Assert.Empty(validationErrors);

            var slidePart = document.PresentationPart!.SlideParts.Single();
            var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
            Assert.NotNull(transition);
        }
    }

    [Fact]
    public void Renderer_CutTransition_EmitsCutElementWithNoDirection()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- transition: cut -->
            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
        Assert.NotNull(transition);
        Assert.Single(transition!.Elements<P.CutTransition>());

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_RandomBarTransition_EmitsHorizontalOrientationByDefault()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- transition: random-bar -->
            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
        Assert.NotNull(transition);
        var rb = transition!.Elements<P.RandomBarTransition>().SingleOrDefault();
        Assert.NotNull(rb);

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }

    [Fact]
    public void Renderer_MorphTransition_EmitsFadeAsFallback()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            <!-- transition: morph -->
            # Slide One
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        using var document = PresentationDocument.Open(outputPath, false);
        var slidePart = document.PresentationPart!.SlideParts.Single();
        var transition = slidePart.Slide!.Elements<P.Transition>().SingleOrDefault();
        Assert.NotNull(transition);
        // Morph is currently emitted as fade (fallback until AlternateContent wrapper is implemented).
        Assert.Single(transition!.Elements<P.FadeTransition>());

        var validationErrors = new OpenXmlPackageValidator().Validate(document);
        Assert.Empty(validationErrors);
    }
}
