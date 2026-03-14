using MarpToPptx.Core;
using MarpToPptx.Pptx.Extraction;
using MarpToPptx.Pptx.Rendering;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Tests;

public class PptxExtractionTests
{
    [Fact]
    public void Extractor_RecoversBasicMarkdownFromRenderedDeck()
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

            <!-- Speaker note -->
            """);

        workspace.WriteFile(
            "pixel.png",
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII="));

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath, new PptxMarkdownExportOptions
        {
            AssetsDirectory = workspace.GetPath("assets"),
            AssetPathPrefix = "assets",
        });

        Assert.Contains("# Title Slide", extracted);
        Assert.Contains("Intro paragraph.", extracted);
        Assert.Contains("# Second Slide", extracted);
        Assert.Contains("- Alpha", extracted);
        Assert.Contains("- Beta", extracted);
        Assert.Contains("![Pixel](assets/pixel.png)", extracted);
        Assert.Contains("<!--", extracted);
        Assert.Contains("Speaker note", extracted);
        Assert.True(File.Exists(workspace.GetPath(Path.Combine("assets", "pixel.png"))));
    }

    [Fact]
    public void Extractor_RecoversNativeTablesAsMarkdownTables()
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

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("| Name | Score | Rank |", extracted);
        Assert.Contains("| --- | ---: | :---: |", extracted);
        Assert.Contains("| Alice | 95 | 1 |", extracted);
        Assert.Contains("| Bob | 87 | 2 |", extracted);
    }

    [Fact]
    public void Extractor_OmitsNotes_WhenIncludeNotesIsFalse()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Intro paragraph.

            <!-- Speaker note -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath, new PptxMarkdownExportOptions
        {
            IncludeNotes = false,
        });

        Assert.DoesNotContain("Speaker note", extracted);
        Assert.DoesNotContain("<!--", extracted);
    }

    [Fact]
    public void Extractor_SuppressesKnownFooterNoiseInNotes()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Title Slide

            Intro paragraph.

            <!-- Useful note -->
            <!-- © Microsoft Corporation. All rights reserved. -->
            <!-- 10/15/2025 1:24 AM -->
            <!-- 7 -->
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("Useful note", extracted);
        Assert.DoesNotContain("© Microsoft Corporation", extracted);
        Assert.DoesNotContain("10/15/2025 1:24 AM", extracted);
        Assert.DoesNotContain(Environment.NewLine + "7" + Environment.NewLine, extracted);
    }

    [Fact]
    public void Extractor_FiltersRepeatedSmallDecorativeImages()
    {
        using var workspace = TestWorkspace.Create();

        workspace.WriteFile(
            "badge.png",
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII="));

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Slide One

            ![Badge](badge.png)

            ---

            # Slide Two

            ![Badge](badge.png)
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        ShrinkPicturesToDecorativeBadges(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath, new PptxMarkdownExportOptions
        {
            AssetsDirectory = workspace.GetPath("assets"),
            AssetPathPrefix = "assets",
        });

        Assert.DoesNotContain("![Badge]", extracted);
    }

    [Fact]
    public void Extractor_InfersBulletListFromShortParagraphStack_WhenBulletMetadataIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # AI and Vector Data Extensions

            - Cloud
            - Web
            - Desktop
            - Mobile
            - AI Model Provider SDKs
            - UI Components
            - AI Libraries
            - Vector Store Provider SDKs
            - Apps
            - Agent Frameworks
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        StripBulletMetadata(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("- Cloud", extracted);
        Assert.Contains("- Web", extracted);
        Assert.Contains("- Agent Frameworks", extracted);
    }

    [Fact]
    public void Extractor_InfersBulletListFromSeparateStackedShapes_WhenBulletMetadataIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # AI and Vector Data Extensions

            - Cloud
            - Web
            - Desktop
            - Mobile
            - AI Model Provider SDKs
            - UI Components
            - AI Libraries
            - Vector Store Provider SDKs
            - Apps
            - Agent Frameworks
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        StripBulletMetadata(outputPath);
        SplitBodyParagraphsIntoSeparateShapes(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("- Cloud", extracted);
        Assert.Contains("- Web", extracted);
        Assert.Contains("- Desktop", extracted);
        Assert.Contains("- Agent Frameworks", extracted);
    }

    [Fact]
    public void Extractor_InfersBulletListFromGridOfShortLabels_WhenBulletMetadataIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # AI and Vector Data Extensions

            - Cloud
            - Web
            - Desktop
            - Mobile
            - AI Model Provider SDKs
            - UI Components
            - AI Libraries
            - Vector Store Provider SDKs
            - Apps
            - Agent Frameworks
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        StripBulletMetadata(outputPath);
        SplitBodyParagraphsIntoGridShapes(outputPath, columns: 4);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("- Cloud", extracted);
        Assert.Contains("- Web", extracted);
        Assert.Contains("- Desktop", extracted);
        Assert.Contains("- Mobile", extracted);
        Assert.Contains("- Agent Frameworks", extracted);
    }

    [Fact]
    public void Extractor_InfersBulletListFromThreeShortParagraphs_WhenBulletMetadataIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Status

            - Monitoring
            - Evaluations
            - Deployment
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        StripBulletMetadata(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("- Monitoring", extracted);
        Assert.Contains("- Evaluations", extracted);
        Assert.Contains("- Deployment", extracted);
    }

    [Fact]
    public void Extractor_InfersBulletListFromSingleRowLabelStrip_WhenBulletMetadataIsMissing()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Platforms

            - Cloud
            - Web
            - Desktop
            - Mobile
            - Games
            - IoT
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        StripBulletMetadata(outputPath);
        SplitBodyParagraphsIntoGridShapes(outputPath, columns: 6);
        CollapseAllGridRowsToSingleRow(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("- Cloud", extracted);
        Assert.Contains("- Web", extracted);
        Assert.Contains("- Games", extracted);
        Assert.Contains("- IoT", extracted);
    }

    [Fact]
    public void Extractor_InfersBulletListWithDescriptions_WhenLabelsAlternateWithBodyText()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # What are agents?

            Placeholder
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        ReplaceFirstBodyShapeWithParagraphs(
            outputPath,
            "Retrieval",
            "Retrieve information from grounding data, reason, summarize, and answer user questions.");

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("- Retrieval", extracted);
        Assert.Contains("Retrieve information from grounding data", extracted);
    }

    [Fact]
    public void Extractor_InfersCodeBlockFromCodeLikeText_WhenShapeNameIsGeneric()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Creating an Agent

            ```csharp
            using System;

            public sealed class Writer
            {
            public string Name { get; set; } = "Writer";
            }
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        RenameCodeShapesToText(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("```csharp", extracted);
        Assert.Contains("using System;", extracted);
        Assert.Contains("public string Name { get; set; } = \"Writer\";", extracted);
        Assert.Contains("```", extracted);
    }

    [Fact]
    public void Extractor_InfersCodeBlockFromConsolasFont_WhenShapeNameIsGeneric()
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            """
            # Creating an Agent

            ```text
            using System;

            public sealed class Writer
            {
            public string Name { get; set; } = "Writer";
            }
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        RenameCodeShapesToText(outputPath);
        ChangeTextFontsToConsolas(outputPath);

        var exporter = new PptxMarkdownExporter();
        var extracted = exporter.ExportToMarkdown(outputPath);

        Assert.Contains("```csharp", extracted);
        Assert.Contains("using System;", extracted);
        Assert.Contains("public string Name { get; set; } = \"Writer\";", extracted);
    }

    [Fact]
    public void Extractor_InfersJavaScriptCodeBlock_WhenShapeNameIsGeneric()
    {
        var extracted = ExtractGenericCodeBlock(
            """
            const agent = createAgent();
            console.log(agent);
            export default agent;
            """);

        Assert.Contains("```javascript", extracted);
        Assert.Contains("console.log(agent);", extracted);
    }

    [Fact]
    public void Extractor_InfersTypeScriptCodeBlock_WhenShapeNameIsGeneric()
    {
        var extracted = ExtractGenericCodeBlock(
            """
            interface AgentConfig {
                name: string;
            }
            const config: AgentConfig = { name: "Writer" };
            """);

        Assert.Contains("```typescript", extracted);
        Assert.Contains("interface AgentConfig", extracted);
    }

    [Fact]
    public void Extractor_InfersPythonCodeBlock_WhenShapeNameIsGeneric()
    {
        var extracted = ExtractGenericCodeBlock(
            """
            def greet(name):
                print(name)
            greet("agent")
            """);

        Assert.Contains("```python", extracted);
        Assert.Contains("def greet(name):", extracted);
    }

    [Fact]
    public void Extractor_InfersJavaCodeBlock_WhenShapeNameIsGeneric()
    {
        var extracted = ExtractGenericCodeBlock(
            """
            public class Demo {
                public static void main(String[] args) {
                    System.out.println("agent");
                }
            }
            """);

        Assert.Contains("```java", extracted);
        Assert.Contains("System.out.println", extracted);
    }

    private static void RenderDeck(string markdownPath, string outputPath, string sourceDirectory)
    {
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(File.ReadAllText(markdownPath), markdownPath);
        var renderer = new OpenXmlPptxRenderer();
        renderer.Render(deck, outputPath, new PptxRenderOptions
        {
            SourceDirectory = sourceDirectory,
        });
    }

    private static string ExtractGenericCodeBlock(string code)
    {
        using var workspace = TestWorkspace.Create();

        var markdownPath = workspace.WriteMarkdown(
            "deck.md",
            $$"""
            # Creating an Agent

            ```text
            {{code}}
            ```
            """);

        var outputPath = workspace.GetPath("deck.pptx");
        RenderDeck(markdownPath, outputPath, workspace.RootPath);
        RenameCodeShapesToText(outputPath);
        ChangeTextFontsToConsolas(outputPath);

        var exporter = new PptxMarkdownExporter();
        return exporter.ExportToMarkdown(outputPath);
    }

    private static void ShrinkPicturesToDecorativeBadges(string outputPath)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        foreach (var slidePart in document.PresentationPart!.SlideParts)
        {
            foreach (var picture in slidePart.Slide!.Descendants<P.Picture>())
            {
                var transform = picture.ShapeProperties?.Transform2D;
                if (transform?.Offset is null || transform.Extents is null)
                {
                    continue;
                }

                transform.Offset.X = 100000L;
                transform.Offset.Y = 100000L;
                transform.Extents.Cx = 900000L;
                transform.Extents.Cy = 900000L;
            }

            slidePart.Slide!.Save();
        }
    }

    private static void StripBulletMetadata(string outputPath)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        foreach (var shape in slidePart.Slide!.Descendants<P.Shape>())
        {
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            if (name.Contains("Title", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (shape.TextBody is null)
            {
                continue;
            }

            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
            {
                var properties = paragraph.ParagraphProperties;
                if (properties is null)
                {
                    continue;
                }

                properties.RemoveAllChildren<A.CharacterBullet>();
                properties.RemoveAllChildren<A.AutoNumberedBullet>();
                properties.RemoveAllChildren<A.NoBullet>();
                properties.Level = null;
            }
        }

        slidePart.Slide!.Save();
    }

    private static void RenameCodeShapesToText(string outputPath)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        foreach (var shape in slidePart.Slide!.Descendants<P.Shape>())
        {
            var drawingProperties = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            if (drawingProperties?.Name?.Value?.StartsWith("Code", StringComparison.OrdinalIgnoreCase) == true)
            {
                drawingProperties.Name = "Text";
            }
        }

        slidePart.Slide!.Save();
    }

    private static void ChangeTextFontsToConsolas(string outputPath)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        foreach (var shape in slidePart.Slide!.Descendants<P.Shape>())
        {
            var drawingProperties = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            if (drawingProperties?.Name?.Value?.Contains("Title", StringComparison.OrdinalIgnoreCase) == true)
            {
                continue;
            }

            foreach (var runProperties in shape.Descendants<A.RunProperties>())
            {
                var latinFont = runProperties.GetFirstChild<A.LatinFont>();
                if (latinFont is null)
                {
                    runProperties.Append(new A.LatinFont { Typeface = "Consolas" });
                }
                else
                {
                    latinFont.Typeface = "Consolas";
                }
            }
        }

        slidePart.Slide!.Save();
    }

    private static void SplitBodyParagraphsIntoSeparateShapes(string outputPath)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
        var bodyShape = shapeTree.Elements<P.Shape>()
            .First(shape => !((shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty)
                .Contains("Title", StringComparison.OrdinalIgnoreCase)));

        var transform = bodyShape.ShapeProperties?.Transform2D;
        if (transform?.Offset is null || transform.Extents is null || bodyShape.TextBody is null)
        {
            return;
        }

        var paragraphs = bodyShape.TextBody.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length <= 1)
        {
            return;
        }

        var baseX = transform.Offset.X?.Value ?? 0L;
        var baseY = transform.Offset.Y?.Value ?? 0L;
        var baseCx = transform.Extents.Cx?.Value ?? 0L;
        var baseCy = transform.Extents.Cy?.Value ?? 0L;
        var paragraphHeight = Math.Max(300000L, baseCy / paragraphs.Length);
        var nextId = shapeTree.Descendants<P.NonVisualDrawingProperties>().Max(props => props.Id?.Value ?? 0U) + 1U;

        bodyShape.TextBody.RemoveAllChildren<A.Paragraph>();
        bodyShape.TextBody.Append(paragraphs[0].CloneNode(true));
        transform.Extents.Cy = paragraphHeight;

        for (var index = 1; index < paragraphs.Length; index++)
        {
            var clone = (P.Shape)bodyShape.CloneNode(true);
            clone.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = nextId++;
            clone.NonVisualShapeProperties.NonVisualDrawingProperties.Name = $"Text {index}";
            clone.TextBody!.RemoveAllChildren<A.Paragraph>();
            clone.TextBody.Append(paragraphs[index].CloneNode(true));
            clone.ShapeProperties!.Transform2D!.Offset!.X = baseX;
            clone.ShapeProperties.Transform2D.Offset.Y = baseY + (paragraphHeight * index);
            clone.ShapeProperties.Transform2D.Extents!.Cx = baseCx;
            clone.ShapeProperties.Transform2D.Extents.Cy = paragraphHeight;
            shapeTree.Append(clone);
        }

        slidePart.Slide!.Save();
    }

    private static void SplitBodyParagraphsIntoGridShapes(string outputPath, int columns)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
        var bodyShape = shapeTree.Elements<P.Shape>()
            .First(shape => !((shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty)
                .Contains("Title", StringComparison.OrdinalIgnoreCase)));

        var transform = bodyShape.ShapeProperties?.Transform2D;
        if (transform?.Offset is null || transform.Extents is null || bodyShape.TextBody is null)
        {
            return;
        }

        var paragraphs = bodyShape.TextBody.Elements<A.Paragraph>().ToArray();
        if (paragraphs.Length <= 1)
        {
            return;
        }

        var baseX = transform.Offset.X?.Value ?? 0L;
        var baseY = transform.Offset.Y?.Value ?? 0L;
        var baseCx = transform.Extents.Cx?.Value ?? 0L;
        var baseCy = transform.Extents.Cy?.Value ?? 0L;
        var rows = (int)Math.Ceiling(paragraphs.Length / (double)columns);
        var cellWidth = Math.Max(600000L, baseCx / columns);
        var cellHeight = Math.Max(300000L, baseCy / rows);
        var nextId = shapeTree.Descendants<P.NonVisualDrawingProperties>().Max(props => props.Id?.Value ?? 0U) + 1U;

        bodyShape.TextBody.RemoveAllChildren<A.Paragraph>();
        bodyShape.TextBody.Append(paragraphs[0].CloneNode(true));
        transform.Extents.Cx = cellWidth;
        transform.Extents.Cy = cellHeight;

        for (var index = 1; index < paragraphs.Length; index++)
        {
            var clone = (P.Shape)bodyShape.CloneNode(true);
            clone.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = nextId++;
            clone.NonVisualShapeProperties.NonVisualDrawingProperties.Name = $"Grid Text {index}";
            clone.TextBody!.RemoveAllChildren<A.Paragraph>();
            clone.TextBody.Append(paragraphs[index].CloneNode(true));

            var column = index % columns;
            var row = index / columns;

            clone.ShapeProperties!.Transform2D!.Offset!.X = baseX + (cellWidth * column);
            clone.ShapeProperties.Transform2D.Offset.Y = baseY + (cellHeight * row);
            clone.ShapeProperties.Transform2D.Extents!.Cx = cellWidth;
            clone.ShapeProperties.Transform2D.Extents.Cy = cellHeight;
            shapeTree.Append(clone);
        }

        slidePart.Slide!.Save();
    }

    private static void CollapseAllGridRowsToSingleRow(string outputPath)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var shapes = slidePart.Slide!.Descendants<P.Shape>()
            .Where(shape => (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty).StartsWith("Grid Text", StringComparison.OrdinalIgnoreCase) ||
                (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty).StartsWith("Text", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (shapes.Length == 0)
        {
            return;
        }

        var minY = shapes.Min(shape => shape.ShapeProperties?.Transform2D?.Offset?.Y?.Value ?? 0L);
        foreach (var shape in shapes)
        {
            if (shape.ShapeProperties?.Transform2D?.Offset is null)
            {
                continue;
            }

            shape.ShapeProperties.Transform2D.Offset.Y = minY;
        }

        slidePart.Slide!.Save();
    }

    private static void ReplaceFirstBodyShapeWithParagraphs(string outputPath, params string[] paragraphs)
    {
        using var document = PresentationDocument.Open(outputPath, true);
        var slidePart = document.PresentationPart!.SlideParts.First();
        var bodyShape = slidePart.Slide!.Descendants<P.Shape>()
            .First(shape => !string.Equals(GetShapeText(shape), "What are agents?", StringComparison.OrdinalIgnoreCase));
        if (bodyShape.TextBody is null)
        {
            return;
        }

        bodyShape.TextBody.RemoveAllChildren<A.Paragraph>();
        foreach (var paragraphText in paragraphs)
        {
            bodyShape.TextBody.Append(new A.Paragraph(new A.Run(new A.Text(paragraphText))));
        }

        slidePart.Slide!.Save();
    }

    private static string GetShapeText(P.Shape shape)
        => string.Join(" ", shape.TextBody?.Elements<A.Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text)).Trim()) ?? []);

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