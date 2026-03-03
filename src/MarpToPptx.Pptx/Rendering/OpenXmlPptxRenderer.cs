using DocumentFormat.OpenXml.Packaging;
using MarpToPptx.Core.Layout;
using MarpToPptx.Core.Models;
using MarpToPptx.Core.Themes;
using A = DocumentFormat.OpenXml.Drawing;
using IOPath = System.IO.Path;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Rendering;

public sealed class OpenXmlPptxRenderer
{
    private const long SlideWidthEmu = 12192000L;
    private const long SlideHeightEmu = 6858000L;
    private const int LayoutScale = 12700;

    private readonly LayoutEngine _layoutEngine = new();

    public void Render(SlideDeck deck, string outputPath, PptxRenderOptions? options = null)
    {
        options ??= new PptxRenderOptions();
        var outputDirectory = IOPath.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var document = OpenPresentation(outputPath, options.TemplatePath);
        var presentationPart = document.PresentationPart ?? document.AddPresentationPart();
        var slideLayoutPart = EnsurePresentationScaffold(presentationPart);

        ClearSlides(presentationPart);

        foreach (var slideModel in deck.Slides)
        {
            AddSlide(presentationPart, slideLayoutPart, slideModel, deck.Theme, options.SourceDirectory ?? GetSourceDirectory(deck.SourcePath));
        }

        presentationPart.Presentation!.Save();
    }

    private static string? GetSourceDirectory(string? sourcePath)
        => string.IsNullOrWhiteSpace(sourcePath) ? null : IOPath.GetDirectoryName(sourcePath);

    private static PresentationDocument OpenPresentation(string outputPath, string? templatePath)
    {
        if (!string.IsNullOrWhiteSpace(templatePath))
        {
            File.Copy(templatePath, outputPath, overwrite: true);
            return PresentationDocument.Open(outputPath, true);
        }

        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        return PresentationDocument.Create(outputPath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
    }

    private static SlideLayoutPart EnsurePresentationScaffold(PresentationPart presentationPart)
    {
        if (presentationPart.Presentation is null)
        {
            presentationPart.Presentation = new P.Presentation();
        }

        var existingLayout = presentationPart.SlideMasterParts.FirstOrDefault()?.SlideLayoutParts.FirstOrDefault();
        if (existingLayout is not null)
        {
            presentationPart.Presentation.SlideIdList ??= new P.SlideIdList();
            presentationPart.Presentation.SlideSize ??= new P.SlideSize { Cx = (int)SlideWidthEmu, Cy = (int)SlideHeightEmu, Type = P.SlideSizeValues.Screen16x9 };
            presentationPart.Presentation.NotesSize ??= new P.NotesSize { Cx = 6858000, Cy = 9144000 };
            return existingLayout;
        }

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        slideLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                CreateRootGroupShapeProperties(),
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

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(new P.ShapeTree(
                CreateRootGroupShapeProperties(),
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
            new P.SlideLayoutIdList(),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));

        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId5");
        themePart.Theme = CreateTheme();

        var slideMasterRelId = presentationPart.GetIdOfPart(slideMasterPart);
        var slideLayoutRelId = slideMasterPart.GetIdOfPart(slideLayoutPart);
        slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = slideLayoutRelId });
        slideMasterPart.SlideMaster.Save();
        slideLayoutPart.SlideLayout.Save();

        presentationPart.Presentation.SlideMasterIdList = new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = slideMasterRelId });
        presentationPart.Presentation.SlideIdList = new P.SlideIdList();
        presentationPart.Presentation.SlideSize = new P.SlideSize { Cx = (int)SlideWidthEmu, Cy = (int)SlideHeightEmu, Type = P.SlideSizeValues.Screen16x9 };
        presentationPart.Presentation.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 };
        presentationPart.Presentation.DefaultTextStyle = new P.DefaultTextStyle();
        presentationPart.Presentation.Save();
        return slideLayoutPart;
    }

    private static void ClearSlides(PresentationPart presentationPart)
    {
        var slideIdList = presentationPart.Presentation?.SlideIdList;
        if (slideIdList is null)
        {
            presentationPart.Presentation!.SlideIdList = new P.SlideIdList();
            return;
        }

        foreach (var slideId in slideIdList.Elements<P.SlideId>().ToList())
        {
            if (!string.IsNullOrWhiteSpace(slideId.RelationshipId))
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                presentationPart.DeletePart(slidePart);
            }

            slideId.Remove();
        }
    }

    private void AddSlide(PresentationPart presentationPart, SlideLayoutPart slideLayoutPart, MarpToPptx.Core.Models.Slide slideModel, ThemeDefinition theme, string? sourceDirectory)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(slideLayoutPart);

        var shapeTree = new P.ShapeTree(
            CreateRootGroupShapeProperties(),
            new P.GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = 0L, Cy = 0L })));

        var slide = new P.Slide(
            new P.CommonSlideData(shapeTree),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        slidePart.Slide = slide;
        var context = new SlideRenderContext(slidePart, shapeTree, sourceDirectory, theme);
        AddBackground(slideModel.Style, context);

        var plan = _layoutEngine.LayoutSlide(slideModel, theme);
        foreach (var placed in plan.Elements)
        {
            switch (placed.Element)
            {
                case HeadingElement heading:
                    AddTextShape(context, placed.Frame, heading.Text, ResolveHeadingStyle(theme, heading.Level), isTitle: heading.Level == 1 && slideModel.Elements.IndexOf(heading) == 0);
                    break;
                case ParagraphElement paragraph:
                    AddTextShape(context, placed.Frame, paragraph.Text, theme.Body);
                    break;
                case BulletListElement list:
                    AddBulletList(context, placed.Frame, list, theme.Body);
                    break;
                case ImageElement image:
                    AddImage(context, placed.Frame, image.Source, image.AltText);
                    break;
                case CodeBlockElement code:
                    AddCodeBlock(context, placed.Frame, code, theme.Code);
                    break;
                case TableElement table:
                    AddTable(context, placed.Frame, table, theme.Body);
                    break;
            }
        }

        slidePart.Slide.Save();
        AppendSlideId(presentationPart, slidePart);
    }

    private static TextStyle ResolveHeadingStyle(ThemeDefinition theme, int level)
        => theme.Headings.TryGetValue(level, out var style) ? style : theme.Headings[1];

    private static void AddBackground(SlideStyle style, SlideRenderContext context)
    {
        var backgroundColor = style.BackgroundColor ?? context.Theme.BackgroundColor;
        if (!string.IsNullOrWhiteSpace(backgroundColor))
        {
            context.ShapeTree.Append(CreateRectangleShape(
                context.NextShapeId(),
                "Background",
                new Rect(0, 0, SlideWidthEmu / LayoutScale, SlideHeightEmu / LayoutScale),
                NormalizeColor(backgroundColor),
                string.Empty,
                1,
                false,
                noFill: false,
                noOutline: true));
        }

        if (!string.IsNullOrWhiteSpace(style.BackgroundImage))
        {
            AddImage(context, new Rect(0, 0, SlideWidthEmu / LayoutScale, SlideHeightEmu / LayoutScale), style.BackgroundImage, string.Empty, useFullBleed: true);
        }
    }

    private static void AddTextShape(SlideRenderContext context, Rect frame, string text, TextStyle style, bool isTitle = false)
    {
        var paragraphs = text
            .Split('\n', StringSplitOptions.None)
            .Select(line => CreateParagraph(line, style, null, false, 1))
            .ToArray();

        context.ShapeTree.Append(CreateTextShape(
            context.NextShapeId(),
            isTitle ? "Title" : "Text",
            frame,
            paragraphs,
            noFill: true,
            fillColor: null,
            lineColor: null));
    }

    private static void AddBulletList(SlideRenderContext context, Rect frame, BulletListElement list, TextStyle style)
    {
        var paragraphs = list.Items
            .Select((item, index) => CreateParagraph(item.Text, style, item.Depth, list.Ordered, index + 1))
            .ToArray();

        context.ShapeTree.Append(CreateTextShape(
            context.NextShapeId(),
            list.Ordered ? "Ordered List" : "Bullet List",
            frame,
            paragraphs,
            noFill: true,
            fillColor: null,
            lineColor: null));
    }

    private static void AddCodeBlock(SlideRenderContext context, Rect frame, CodeBlockElement code, TextStyle style)
    {
        var paragraphs = code.Code
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n', StringSplitOptions.None)
            .Select(line => CreateParagraph(line, style, null, false, 1))
            .ToArray();

        context.ShapeTree.Append(CreateTextShape(
            context.NextShapeId(),
            string.IsNullOrWhiteSpace(code.Language) ? "Code" : $"Code ({code.Language})",
            frame,
            paragraphs,
            noFill: false,
            fillColor: NormalizeColor(style.BackgroundColor ?? "#0F172A"),
            lineColor: NormalizeColor(context.Theme.AccentColor)));
    }

    private static void AddTable(SlideRenderContext context, Rect frame, TableElement table, TextStyle style)
    {
        if (table.Rows.Count == 0)
        {
            return;
        }

        var columnCount = table.Rows.Max(row => row.Cells.Count);
        if (columnCount == 0)
        {
            return;
        }

        var tableText = string.Join("\n", table.Rows.Select(row => string.Join(" | ", row.Cells)));
        AddTextShape(context, frame, tableText, style);
    }

    private static void AddImage(SlideRenderContext context, Rect frame, string source, string altText, bool useFullBleed = false)
    {
        var resolved = ResolvePath(context.SourceDirectory, source);
        if (resolved is null || !File.Exists(resolved))
        {
            AddTextShape(context, frame, string.IsNullOrWhiteSpace(source) ? "Missing image" : $"Missing image: {source}", context.Theme.Body);
            return;
        }

        var contentType = GetImageContentType(resolved);
        var imagePart = context.SlidePart.AddImagePart(contentType);
        using (var imageStream = File.OpenRead(resolved))
        {
            imagePart.FeedData(imageStream);
        }

        var (x, y, width, height) = CalculateImagePlacement(frame, resolved, useFullBleed);
        var relationshipId = context.SlidePart.GetIdOfPart(imagePart);

        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = context.NextShapeId(), Name = IOPath.GetFileName(resolved), Description = altText },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmu(x), Y = ToEmu(y) },
                    new A.Extents { Cx = ToEmu(width), Cy = ToEmu(height) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        context.ShapeTree.Append(picture);
    }

    private static (double X, double Y, double Width, double Height) CalculateImagePlacement(Rect frame, string imagePath, bool useFullBleed)
    {
        if (!ImageMetadataReader.TryReadSize(imagePath, out var pixelWidth, out var pixelHeight) || pixelWidth <= 0 || pixelHeight <= 0)
        {
            return (frame.X, frame.Y, frame.Width, frame.Height);
        }

        var imageAspect = (double)pixelWidth / pixelHeight;
        var frameAspect = frame.Width / frame.Height;

        if (useFullBleed)
        {
            if (imageAspect > frameAspect)
            {
                var scaledWidth = frame.Height * imageAspect;
                return (frame.X - ((scaledWidth - frame.Width) / 2), frame.Y, scaledWidth, frame.Height);
            }

            var scaledHeight = frame.Width / imageAspect;
            return (frame.X, frame.Y - ((scaledHeight - frame.Height) / 2), frame.Width, scaledHeight);
        }

        if (imageAspect > frameAspect)
        {
            var fittedHeight = frame.Width / imageAspect;
            return (frame.X, frame.Y + ((frame.Height - fittedHeight) / 2), frame.Width, fittedHeight);
        }

        var fittedWidth = frame.Height * imageAspect;
        return (frame.X + ((frame.Width - fittedWidth) / 2), frame.Y, fittedWidth, frame.Height);
    }

    private static string? ResolvePath(string? sourceDirectory, string source)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            return null;
        }

        if (Uri.TryCreate(source, UriKind.Absolute, out var uri) && !uri.IsFile)
        {
            return null;
        }

        if (IOPath.IsPathFullyQualified(source))
        {
            return source;
        }

        return string.IsNullOrWhiteSpace(sourceDirectory)
            ? IOPath.GetFullPath(source)
            : IOPath.GetFullPath(IOPath.Combine(sourceDirectory, source));
    }

    private static P.Shape CreateRectangleShape(uint shapeId, string name, Rect frame, string fillColor, string text, double fontSize, bool bold, bool noFill, bool noOutline)
        => CreateTextShape(
            shapeId,
            name,
            frame,
            [CreateParagraph(text, new TextStyle(fontSize, "#000000", "Aptos", bold), null, false, 1)],
            noFill,
            noFill ? null : fillColor,
            noOutline ? null : fillColor);

    private static P.Shape CreateTextShape(uint shapeId, string name, Rect frame, IEnumerable<A.Paragraph> paragraphs, bool noFill, string? fillColor, string? lineColor)
    {
        var shapeProperties = new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = ToEmu(frame.X), Y = ToEmu(frame.Y) },
                new A.Extents { Cx = ToEmu(frame.Width), Cy = ToEmu(frame.Height) }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

        if (noFill)
        {
            shapeProperties.Append(new A.NoFill());
        }
        else if (!string.IsNullOrWhiteSpace(fillColor))
        {
            shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = fillColor }));
        }

        shapeProperties.Append(string.IsNullOrWhiteSpace(lineColor)
            ? new A.Outline(new A.NoFill())
            : new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = lineColor })));

        var textBody = new P.TextBody(
            new A.BodyProperties { Wrap = A.TextWrappingValues.Square, VerticalOverflow = A.TextVerticalOverflowValues.Overflow },
            new A.ListStyle());

        foreach (var paragraph in paragraphs)
        {
            textBody.Append(paragraph.CloneNode(true));
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            textBody);
    }

    private static A.Paragraph CreateParagraph(string text, TextStyle style, int? level, bool ordered, int orderNumber)
    {
        var paragraph = new A.Paragraph();
        var paragraphProperties = new A.ParagraphProperties();

        if (level is not null)
        {
            paragraphProperties.Level = level.Value;
            if (ordered)
            {
                paragraphProperties.Append(new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod, StartAt = orderNumber });
            }
            else
            {
                paragraphProperties.Append(new A.CharacterBullet { Char = "•" });
            }
        }

        paragraph.Append(paragraphProperties);

        var runProperties = new A.RunProperties
        {
            Language = "en-US",
            FontSize = (int)Math.Round(style.FontSize * 100),
            Bold = style.Bold,
        };
        runProperties.Append(new A.LatinFont { Typeface = style.FontFamily });
        runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = NormalizeColor(style.Color) }));

        paragraph.Append(new A.Run(runProperties, new A.Text(text)));
        paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US", FontSize = (int)Math.Round(style.FontSize * 100) });
        return paragraph;
    }

    private static P.NonVisualGroupShapeProperties CreateRootGroupShapeProperties()
        => new(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties());

    private static A.Theme CreateTheme()
        => new()
        {
            Name = "MarpToPptx Theme",
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
                { Name = "MarpToPptx" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = "Aptos Display" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                    new A.MinorFont(new A.LatinFont { Typeface = "Aptos" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "MarpToPptx" },
                new A.FormatScheme(
                    new A.FillStyleList(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }), new A.NoFill(), new A.GradientFill()),
                    new A.LineStyleList(new A.Outline(), new A.Outline(), new A.Outline()),
                    new A.EffectStyleList(new A.EffectStyle(new A.EffectList()), new A.EffectStyle(new A.EffectList()), new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 }), new A.NoFill(), new A.GradientFill()))
                { Name = "MarpToPptx" }),
            ObjectDefaults = new A.ObjectDefaults(),
            ExtraColorSchemeList = new A.ExtraColorSchemeList(),
        };

    private static void AppendSlideId(PresentationPart presentationPart, SlidePart slidePart)
    {
        var slideIdList = presentationPart.Presentation!.SlideIdList ??= new P.SlideIdList();
        uint nextId = slideIdList.Elements<P.SlideId>().Select(id => id.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1;
        slideIdList.Append(new P.SlideId { Id = nextId, RelationshipId = presentationPart.GetIdOfPart(slidePart) });
    }

    private static long ToEmu(double value) => (long)Math.Round(value * LayoutScale);

    private static string NormalizeColor(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "000000";
        }

        var trimmed = value.Trim();
        if (trimmed.StartsWith('#'))
        {
            trimmed = trimmed[1..];
        }

        if (trimmed.Length == 3)
        {
            trimmed = string.Concat(trimmed.Select(ch => new string(ch, 2)));
        }

        return trimmed.Length >= 6 ? trimmed[..6].ToUpperInvariant() : "000000";
    }

    private static string GetImageContentType(string path)
        => IOPath.GetExtension(path).ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            _ => "image/png",
        };

    private sealed class SlideRenderContext(SlidePart slidePart, P.ShapeTree shapeTree, string? sourceDirectory, ThemeDefinition theme)
    {
        private uint _shapeId = 1;

        public SlidePart SlidePart { get; } = slidePart;

        public P.ShapeTree ShapeTree { get; } = shapeTree;

        public string? SourceDirectory { get; } = sourceDirectory;

        public ThemeDefinition Theme { get; } = theme;

        public uint NextShapeId() => ++_shapeId;
    }
}
