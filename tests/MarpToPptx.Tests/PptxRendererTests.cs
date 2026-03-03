using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;
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

        var slideParts = presentationPart.SlideParts.ToArray();
        Assert.Contains("Title Slide", slideParts[0].Slide.Descendants<A.Text>().Select(text => text.Text));
        Assert.Contains("Intro paragraph.", slideParts[0].Slide.Descendants<A.Text>().Select(text => text.Text));
        Assert.Single(slideParts[1].Slide.Descendants<P.Picture>());
    }
}
