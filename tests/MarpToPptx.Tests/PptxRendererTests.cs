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
}
