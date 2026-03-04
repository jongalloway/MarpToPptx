using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;
using System.IO.Compression;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Tests;

public class PptxRendererTests
{
    [Fact]
    public void Renderer_CreatesPresentationWithSlidesTextAndImage()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), "MarpToPptx.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);

        var markdownPath = Path.Combine(tempRoot, "deck.md");
        var imagePath = Path.Combine(tempRoot, "pixel.png");
        var outputPath = Path.Combine(tempRoot, "deck.pptx");

        File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII="));
        File.WriteAllText(markdownPath, """
        # Title Slide

        Intro paragraph.

        ---

        ## Second Slide

        - Alpha
        - Beta

        ![Pixel](pixel.png)
        """);

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);

        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = tempRoot });

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
        Assert.All(slideParts, slidePart => Assert.Equal("/ppt/slideLayouts/slideLayout2.xml", slidePart.SlideLayoutPart?.Uri.ToString()));
        Assert.Contains("Title Slide", slideParts[0].Slide.Descendants<A.Text>().Select(text => text.Text));
        Assert.Contains("Intro paragraph.", slideParts[0].Slide.Descendants<A.Text>().Select(text => text.Text));
        Assert.Single(slideParts[1].Slide.Descendants<P.Picture>());

        var validationErrors = new OpenXmlValidator().Validate(document).ToList();
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
        var tempRoot = Path.Combine(Path.GetTempPath(), "MarpToPptx.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);

        var markdownPath = Path.Combine(tempRoot, "deck.md");
        var svgPath = Path.Combine(tempRoot, "accent-wave.svg");
        var outputPath = Path.Combine(tempRoot, "deck.pptx");

        File.WriteAllText(svgPath, """
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
          <rect width="100" height="100" fill="#102A43" />
          <path d="M0 70 C 20 40, 40 40, 60 70 S 100 100, 100 60 L100 100 L0 100 Z" fill="#F7C948" />
        </svg>
        """);

        File.WriteAllText(markdownPath, """
        ---
        theme: gaia
        backgroundColor: "#F7F3E8"
        ---

        # Quoted Color

        ---

        <!-- backgroundImage: url(accent-wave.svg) -->
        # Svg Background
        """);

        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);

        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions { SourceDirectory = tempRoot });

        using var document = PresentationDocument.Open(outputPath, false);
        var validationErrors = new OpenXmlValidator().Validate(document).ToList();
        Assert.Empty(validationErrors);

        using var archive = ZipFile.OpenRead(outputPath);
        using var contentTypesReader = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open());
        var contentTypes = contentTypesReader.ReadToEnd();
        Assert.Contains("image/svg+xml", contentTypes);
    }
}
