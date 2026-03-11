using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DiagramForge;
using DiagramForge.Abstractions;
using MarpToPptx.Core.Layout;
using MarpToPptx.Core.Models;
using MarpToPptx.Core.Themes;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using IOPath = System.IO.Path;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarpToPptx.Pptx.Rendering;

public sealed class OpenXmlPptxRenderer
{
    private const long SlideWidthEmu = 12192000L;
    private const long SlideHeightEmu = 6858000L;
    private const int LayoutScale = 12700;
    private const string DefaultTableStyleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
    private const string SvgBlipExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    private const double MermaidErrorLabelVerticalGap = 4;
    private const double MermaidErrorLabelHeight = 20;
    private static readonly byte[] MediaPlaceholderImage = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnV9a4AAAAASUVORK5CYII=");
    private static readonly DiagramRenderer _diagramRenderer = new();

    private readonly LayoutEngine _layoutEngine = new();

    public void Render(SlideDeck deck, string outputPath, PptxRenderOptions? options = null)
    {
        options ??= new PptxRenderOptions();

        var outputDirectory = IOPath.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var remoteAssets = options.AllowRemoteAssets
            ? new RemoteAssetResolver(options.RemoteAssetHandler)
            : null;

        using (var document = OpenPresentation(outputPath, options.TemplatePath))
        {
            var presentationPart = document.PresentationPart ?? document.AddPresentationPart();
            if (string.IsNullOrEmpty(options.TemplatePath))
            {
                EnsureRelationshipId(document, presentationPart, "rId1");
            }
            var allLayouts = EnsurePresentationScaffold(presentationPart);
            var templateSlides = GetSlidesInPresentationOrder(presentationPart);

            // Pre-clone all template slide parts before ClearSlides runs.
            // In some SDK environments (e.g., DocumentFormat.OpenXml 3.x on .NET 10),
            // removing a slide's <p:sldId> XML reference automatically destroys the
            // SlidePart and ALL its sub-parts (images, layouts, etc.). By pre-cloning
            // each template slide into a new SlidePart first, we add an extra OPC
            // relationship from the pre-clone to every sub-part, keeping those sub-parts
            // alive even after the original SlidePart is destroyed.
            var preClonedTemplateSlides = templateSlides.Count > 0
                ? templateSlides.Select(sp => PreCloneTemplateSlidePart(presentationPart, sp)).ToArray()
                : [];

            var templateSelector = new SlideTemplateSelector(allLayouts, preClonedTemplateSlides);

            ClearSlides(presentationPart);

            var language = deck.Language ?? "en-US";
            var slideNumber = 1;
            foreach (var slideModel in deck.Slides)
            {
                var slideKind = SlideTemplateSelector.Classify(slideModel);
                var selectedLayout = templateSelector.SelectLayout(slideModel, slideKind, deck.DefaultContentLayout);
                AddSlide(presentationPart, selectedLayout.LayoutPart, slideModel, deck.Theme, options.SourceDirectory ?? GetSourceDirectory(deck.SourcePath), remoteAssets, selectedLayout.UseTemplateStyle, slideNumber, language, selectedLayout.TemplateSlide);
                slideNumber++;
            }

            DeleteSlideParts(presentationPart, templateSlides);
            // Remove pre-cloned orphaned slides. Shared sub-parts (images, layouts) are
            // preserved because they are also referenced by the rendered slide parts.
            DeleteSlideParts(presentationPart, preClonedTemplateSlides);

            EnsureDocumentProperties(document, deck, options.TemplatePath);
            presentationPart.Presentation!.Save();
        }

        NormalizePackage(outputPath);
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

    private static IReadOnlyList<SlideLayoutPart> EnsurePresentationScaffold(PresentationPart presentationPart)
    {
        if (presentationPart.Presentation is null)
        {
            presentationPart.Presentation = new P.Presentation();
        }

        var existingLayouts = presentationPart.SlideMasterParts
            .SelectMany(master => master.SlideLayoutParts)
            .ToList();
        if (existingLayouts.Count > 0)
        {
            EnsurePresentationMetadataParts(presentationPart);
            presentationPart.Presentation.SlideIdList ??= new P.SlideIdList();
            presentationPart.Presentation.SlideSize ??= new P.SlideSize { Cx = (int)SlideWidthEmu, Cy = (int)SlideHeightEmu, Type = P.SlideSizeValues.Screen16x9 };
            presentationPart.Presentation.NotesSize ??= new P.NotesSize { Cx = 6858000, Cy = 9144000 };
            return existingLayouts;
        }

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
        var contentLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        contentLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(new P.ShapeTree(
                CreateRootGroupShapeProperties(),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })),
                CreatePlaceholderShape(2U, "Title 1", P.PlaceholderValues.Title),
                CreatePlaceholderShape(
                    3U,
                    "Text Placeholder 2",
                    P.PlaceholderValues.Body,
                    1U,
                    new Rect(66, 144, 828, 343)))),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Text,
        };

        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId2");
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
                    new A.ChildExtents { Cx = 0L, Cy = 0L })),
                CreatePlaceholderShape(2U, "Title Placeholder 1", P.PlaceholderValues.Title))),
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

        var themePart = presentationPart.AddNewPart<ThemePart>("rId6");
        themePart.Theme = CreateTheme();
        slideMasterPart.AddPart(themePart, "rId3");
        contentLayoutPart.AddPart(slideMasterPart, "rId1");
        slideLayoutPart.AddPart(slideMasterPart, "rId1");

        var slideMasterRelId = presentationPart.GetIdOfPart(slideMasterPart);
        var contentLayoutRelId = slideMasterPart.GetIdOfPart(contentLayoutPart);
        var blankLayoutRelId = slideMasterPart.GetIdOfPart(slideLayoutPart);
        slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = contentLayoutRelId });
        slideMasterPart.SlideMaster.SlideLayoutIdList!.Append(new P.SlideLayoutId { Id = 2147483650U, RelationshipId = blankLayoutRelId });
        slideMasterPart.SlideMaster.Save();
        contentLayoutPart.SlideLayout.Save();
        slideLayoutPart.SlideLayout.Save();

        EnsurePresentationMetadataParts(presentationPart);

        presentationPart.Presentation.SlideMasterIdList = new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = slideMasterRelId });
        presentationPart.Presentation.SlideIdList = new P.SlideIdList();
        presentationPart.Presentation.SlideSize = new P.SlideSize { Cx = (int)SlideWidthEmu, Cy = (int)SlideHeightEmu, Type = P.SlideSizeValues.Screen16x9 };
        presentationPart.Presentation.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 };
        presentationPart.Presentation.DefaultTextStyle = new P.DefaultTextStyle();
        presentationPart.Presentation.Save();
        return [contentLayoutPart, slideLayoutPart];
    }

    private static void EnsurePresentationMetadataParts(PresentationPart presentationPart)
    {
        ThemePart themePart;
        if (presentationPart.ThemePart is not null)
        {
            themePart = presentationPart.ThemePart;
            themePart.Theme ??= CreateTheme();
        }
        else
        {
            themePart = presentationPart.SlideMasterParts.FirstOrDefault()?.ThemePart is { } existingTheme
                ? presentationPart.AddPart(existingTheme, "rId6")
                : presentationPart.AddNewPart<ThemePart>("rId6");

            themePart.Theme ??= CreateTheme();
        }

        foreach (var slideMasterPart in presentationPart.SlideMasterParts)
        {
            if (slideMasterPart.ThemePart is null)
            {
                slideMasterPart.AddPart(themePart, "rId3");
            }

            slideMasterPart.SlideMaster?.Save();
        }

        var presentationPropertiesPart = presentationPart.PresentationPropertiesPart ?? presentationPart.AddNewPart<PresentationPropertiesPart>("rId4");
        if (presentationPropertiesPart.PresentationProperties is null)
        {
            presentationPropertiesPart.PresentationProperties = new P.PresentationProperties();
        }

        presentationPropertiesPart.PresentationProperties.Save();

        var viewPropertiesPart = presentationPart.ViewPropertiesPart ?? presentationPart.AddNewPart<ViewPropertiesPart>("rId5");
        WriteXmlPart(viewPropertiesPart, CreateViewPropertiesDocument());

        var tableStylesPart = presentationPart.TableStylesPart ?? presentationPart.AddNewPart<TableStylesPart>("rId7");
        if (tableStylesPart.TableStyleList is null)
        {
            tableStylesPart.TableStyleList = new A.TableStyleList { Default = DefaultTableStyleId };
        }

        tableStylesPart.TableStyleList.Save();
    }

    private static void EnsureDocumentProperties(PresentationDocument document, SlideDeck deck, string? templatePath)
    {
        var now = DateTime.UtcNow;
        var corePropertiesPart = document.CoreFilePropertiesPart ?? document.AddCoreFilePropertiesPart();
        if (string.IsNullOrEmpty(templatePath))
        {
            EnsureRelationshipId(document, corePropertiesPart, "rId2");
        }
        WriteXmlPart(corePropertiesPart, CreateCorePropertiesDocument(deck, now));

        var appPropertiesPart = document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart();
        if (string.IsNullOrEmpty(templatePath))
        {
            EnsureRelationshipId(document, appPropertiesPart, "rId3");
        }
        WriteXmlPart(appPropertiesPart, CreateExtendedPropertiesDocument(deck));
    }

    private static void ClearSlides(PresentationPart presentationPart)
    {
        var slideIdList = presentationPart.Presentation?.SlideIdList;
        if (slideIdList is null)
        {
            presentationPart.Presentation!.SlideIdList = new P.SlideIdList();

            foreach (var slidePart in presentationPart.SlideParts.ToList())
            {
                presentationPart.DeletePart(slidePart);
            }

            return;
        }

        // For each slide ID, delete the underlying SlidePart (and its dependent parts)
        // before removing the slide ID entry to avoid leaving orphaned parts in the package.
        foreach (var slideId in slideIdList.Elements<P.SlideId>().ToList())
        {
            var relId = slideId.RelationshipId?.Value;
            if (!string.IsNullOrEmpty(relId))
            {
                if (presentationPart.TryGetPartById(relId, out var part) && part is SlidePart slidePart)
                {
                    presentationPart.DeletePart(slidePart);
                }
            }

            slideId.Remove();
        }
    }

    private void AddSlide(PresentationPart presentationPart, SlideLayoutPart slideLayoutPart, MarpToPptx.Core.Models.Slide slideModel, ThemeDefinition theme, string? sourceDirectory, RemoteAssetResolver? remoteAssets, bool useTemplateStyle, int slideNumber, string language, TemplateSlideReference? templateSlide = null)
    {
        SlidePart slidePart;
        P.ShapeTree shapeTree;
        if (templateSlide is null)
        {
            slidePart = presentationPart.AddNewPart<SlidePart>(GetNextRelationshipId(presentationPart));
            slidePart.AddPart(slideLayoutPart, "rId1");

            shapeTree = new P.ShapeTree(
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
        }
        else
        {
            slidePart = CloneTemplateSlidePart(presentationPart, templateSlide.SlidePart);
            slidePart.Slide!.CommonSlideData ??= new P.CommonSlideData();
            shapeTree = slidePart.Slide.CommonSlideData.ShapeTree ??= new P.ShapeTree(
                CreateRootGroupShapeProperties(),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })));
        }
        // Resolve class variant: when the slide has a className, look up overrides.
        ClassVariant? classVariant = null;
        if (!useTemplateStyle && !string.IsNullOrWhiteSpace(slideModel.Style.ClassName))
        {
            theme.ClassVariants.TryGetValue(slideModel.Style.ClassName!, out classVariant);
        }

        var effectiveTheme = useTemplateStyle
            ? ThemeDefinition.Default
            : theme.ApplyClassVariant(classVariant);
        var context = new SlideRenderContext(slidePart, shapeTree, sourceDirectory, effectiveTheme, remoteAssets, useTemplateStyle, language);

        AddBackground(slideModel.Style, context);

        if (templateSlide is not null &&
            TryRenderIntoTemplateSlideTextShapes(context, slideModel, effectiveTheme))
        {
            if (!useTemplateStyle)
            {
                AddHeaderFooterAndPageNumber(context, slideModel.Style, effectiveTheme.Body, slideNumber);
            }
            slidePart.Slide.Save();
            AppendSlideId(presentationPart, slidePart);
            if (!string.IsNullOrWhiteSpace(slideModel.Notes))
            {
                AddNotesSlide(presentationPart, slidePart, slideModel.NoteSpans, slideModel.Notes!, effectiveTheme, language);
            }
            return;
        }

        // Placeholder-based rendering: when a named layout matched, write text content
        // into slide-level placeholder shapes that inherit geometry and text styling
        // from the template layout. Non-text elements fall back to standalone shapes.
        // If the layout lacks a title or body placeholder, that portion degrades to
        // the standard standalone-shape path. See doc/template-authoring-guidelines.md.
        if (useTemplateStyle &&
            TryRenderIntoTemplatePlaceholders(context, slideLayoutPart, slideModel, effectiveTheme))
        {
            AddHeaderFooterAndPageNumber(context, slideModel.Style, effectiveTheme.Body, slideNumber);
            slidePart.Slide.Save();
            AppendSlideId(presentationPart, slidePart);
            if (!string.IsNullOrWhiteSpace(slideModel.Notes))
            {
                AddNotesSlide(presentationPart, slidePart, slideModel.NoteSpans, slideModel.Notes!, effectiveTheme, language);
            }
            return;
        }

        // Read placeholder bounds from the selected layout. When the title placeholder
        // carries an explicit transform, use it for the first top-level heading so that
        // the slide respects the template's intended title position. When there is
        // exactly one non-heading element, and a usable body placeholder rect is
        // available (see canUseBodyRect), place that single body element inside the
        // body placeholder; otherwise, fall back to LayoutEngine positioning.
        var titleRect = SlideTemplateSelector.GetTitlePlaceholderRect(slideLayoutPart);
        var bodyRect = SlideTemplateSelector.GetBodyPlaceholderRect(slideLayoutPart);
        var nonHeadingCount = slideModel.Elements.Count(e => e is not HeadingElement);
        var hasHeading = slideModel.Elements.OfType<HeadingElement>().Any();

        // Only use the body rect when the title rect is also known, so that the two
        // placeholder regions are coordinated and won't overlap each other. If there
        // is no heading on the slide at all, the body rect is safe to apply on its own.
        var canUseBodyRect = bodyRect is not null &&
            (titleRect is not null || !hasHeading);

        var firstElement = slideModel.Elements.FirstOrDefault();
        var plan = _layoutEngine.LayoutSlide(slideModel, effectiveTheme);
        var bodyStyle = effectiveTheme.Body;
        foreach (var placed in plan.Elements)
        {
            // Resolve the frame: prefer template placeholder rect when available.
            var frame = placed.Element is HeadingElement ph &&
                ph.Level == 1 && ReferenceEquals(ph, firstElement) &&
                titleRect is not null
                ? titleRect
                : canUseBodyRect && placed.Element is not HeadingElement && nonHeadingCount == 1
                    ? bodyRect!
                    : placed.Frame;

            switch (placed.Element)
            {
                case HeadingElement heading:
                    AddTextShape(context, frame, heading.Spans, ResolveHeadingStyle(effectiveTheme, heading.Level), effectiveTheme.InlineCode, isTitle: heading.Level == 1 && ReferenceEquals(heading, firstElement));
                    break;
                case ParagraphElement paragraph:
                    AddTextShape(context, frame, paragraph.Spans, bodyStyle, effectiveTheme.InlineCode);
                    break;
                case BulletListElement list:
                    AddBulletList(context, frame, list, bodyStyle);
                    break;
                case ImageElement image:
                    AddImage(context, frame, image.Source, image.AltText);
                    break;
                case VideoElement video:
                    AddVideo(context, frame, video.Source, video.AltText);
                    break;
                case AudioElement audio:
                    AddAudio(context, frame, audio.Source, audio.AltText);
                    break;
                case CodeBlockElement code:
                    AddCodeBlock(context, frame, code, effectiveTheme.Code);
                    break;
                case MermaidDiagramElement mermaid:
                    AddMermaidDiagram(context, frame, mermaid, effectiveTheme.Code);
                    break;
                case TableElement table:
                    AddTable(context, frame, table, bodyStyle);
                    break;
            }
        }

        AddHeaderFooterAndPageNumber(context, slideModel.Style, bodyStyle, slideNumber);

        slidePart.Slide.Save();
        AppendSlideId(presentationPart, slidePart);

        if (!string.IsNullOrWhiteSpace(slideModel.Notes))
        {
            AddNotesSlide(presentationPart, slidePart, slideModel.NoteSpans, slideModel.Notes!, effectiveTheme, language);
        }
    }

    private static TextStyle ResolveHeadingStyle(ThemeDefinition theme, int level)
        => theme.GetHeadingStyle(level);

    /// <summary>
    /// Attempts to render slide content into template placeholder shapes. Returns
    /// <c>true</c> when the placeholder path was taken (text content written into one
    /// or both of the layout's title/body placeholders). Returns <c>false</c> when the
    /// layout exposes neither a title-like nor a body-like placeholder, in which case
    /// the caller should fall back to the standalone-shape path unchanged.
    ///
    /// Content mapping:
    ///   - first heading (any level: #, ##, etc.) -> title placeholder (title | ctrTitle)
    ///   - remaining headings, paragraphs, bullet/numbered lists -> body placeholder
    ///     (body | subTitle), collapsed into a single shape with multiple paragraphs
    ///   - images, video, audio, code blocks, tables -> standalone positioned shapes
    ///
    /// Paragraphs emitted here deliberately omit font size, color, and font family so
    /// the template's layout+master text styles cascade; only inline run formatting
    /// (bold, italic, strike, hyperlink) and bullet level / nobullet are set.
    /// </summary>
    private bool TryRenderIntoTemplatePlaceholders(SlideRenderContext context, SlideLayoutPart slideLayoutPart, MarpToPptx.Core.Models.Slide slideModel, ThemeDefinition effectiveTheme)
    {
        var titlePlaceholder = SlideTemplateSelector.GetTitlePlaceholder(slideLayoutPart);
        var bodyPlaceholder = SlideTemplateSelector.GetBodyPlaceholder(slideLayoutPart);
        if (titlePlaceholder is null && bodyPlaceholder is null)
        {
            return false;
        }

        // Split elements into: optional title heading, body text, and non-text remainder.
        // The title placeholder receives the very first element when it is a heading of
        // any level. Level is intentionally ignored because a template-bound "Title and
        // Content" slide typically uses an H2 for its per-slide heading and still wants
        // it routed to the title placeholder.
        HeadingElement? titleHeading = null;
        var bodyTextElements = new List<ISlideElement>();
        var nonTextElements = new List<ISlideElement>();
        foreach (var element in slideModel.Elements)
        {
            if (titleHeading is null && titlePlaceholder is not null &&
                element is HeadingElement h &&
                bodyTextElements.Count == 0 && nonTextElements.Count == 0)
            {
                titleHeading = h;
                continue;
            }

            switch (element)
            {
                case HeadingElement or ParagraphElement or BulletListElement:
                    bodyTextElements.Add(element);
                    break;
                default:
                    nonTextElements.Add(element);
                    break;
            }
        }

        // Title placeholder shape.
        if (titleHeading is not null)
        {
            var titleParagraphs = SplitSpansIntoParagraphs(titleHeading.Spans)
                .Select(group => CreateTemplateParagraphFromSpans(group, context.SlidePart, level: null, ordered: false, forceBold: false, context.Language))
                .ToArray();
            context.ShapeTree.Append(CreateSlidePlaceholderShape(
                context.NextShapeId(),
                "Title",
                titlePlaceholder!,
                titleParagraphs));
        }

        // Body placeholder shape: collapse all body-text elements into one shape.
        if (bodyPlaceholder is not null && bodyTextElements.Count > 0)
        {
            var bodyParagraphs = new List<A.Paragraph>();
            foreach (var element in bodyTextElements)
            {
                switch (element)
                {
                    case HeadingElement heading:
                        foreach (var group in SplitSpansIntoParagraphs(heading.Spans))
                        {
                            bodyParagraphs.Add(CreateTemplateParagraphFromSpans(group, context.SlidePart, level: null, ordered: false, forceBold: true, context.Language));
                        }
                        break;
                    case ParagraphElement paragraph:
                        foreach (var group in SplitSpansIntoParagraphs(paragraph.Spans))
                        {
                            bodyParagraphs.Add(CreateTemplateParagraphFromSpans(group, context.SlidePart, level: null, ordered: false, forceBold: false, context.Language));
                        }
                        break;
                    case BulletListElement list:
                        var orderNumber = 1;
                        foreach (var item in list.Items)
                        {
                            bodyParagraphs.Add(CreateTemplateParagraphFromSpans(item.Spans, context.SlidePart, level: item.Depth, list.Ordered, forceBold: false, context.Language, orderNumber));
                            orderNumber++;
                        }
                        break;
                }
            }

            context.ShapeTree.Append(CreateSlidePlaceholderShape(
                context.NextShapeId(),
                "Content Placeholder",
                bodyPlaceholder,
                bodyParagraphs));
        }
        else if (bodyPlaceholder is null && bodyTextElements.Count > 0)
        {
            // Layout has no body placeholder: route body text back through the
            // standalone path so content is not silently dropped.
            nonTextElements.InsertRange(0, bodyTextElements);
        }

        // Non-text elements: render as standalone shapes using the layout engine for
        // positioning, just as the non-placeholder path does.
        if (nonTextElements.Count > 0)
        {
            var residualSlide = new MarpToPptx.Core.Models.Slide { Style = slideModel.Style };
            residualSlide.Elements.AddRange(nonTextElements);
            var plan = _layoutEngine.LayoutSlide(residualSlide, effectiveTheme);
            var bodyStyle = effectiveTheme.Body;
            foreach (var placed in plan.Elements)
            {
                switch (placed.Element)
                {
                    case HeadingElement heading:
                        AddTextShape(context, placed.Frame, heading.Spans, ResolveHeadingStyle(effectiveTheme, heading.Level), effectiveTheme.InlineCode);
                        break;
                    case ParagraphElement paragraph:
                        AddTextShape(context, placed.Frame, paragraph.Spans, bodyStyle, effectiveTheme.InlineCode);
                        break;
                    case BulletListElement list:
                        AddBulletList(context, placed.Frame, list, bodyStyle);
                        break;
                    case ImageElement image:
                        AddImage(context, placed.Frame, image.Source, image.AltText);
                        break;
                    case VideoElement video:
                        AddVideo(context, placed.Frame, video.Source, video.AltText);
                        break;
                    case AudioElement audio:
                        AddAudio(context, placed.Frame, audio.Source, audio.AltText);
                        break;
                    case CodeBlockElement code:
                        AddCodeBlock(context, placed.Frame, code, effectiveTheme.Code);
                        break;
                    case MermaidDiagramElement mermaid:
                        AddMermaidDiagram(context, placed.Frame, mermaid, effectiveTheme.Code);
                        break;
                    case TableElement table:
                        AddTable(context, placed.Frame, table, bodyStyle);
                        break;
                }
            }
        }

        return true;
    }

    /// <summary>
    /// Reuses a real template slide as the slide base, preserving authored artwork and
    /// replacing the existing text boxes with slide content.
    /// </summary>
    private bool TryRenderIntoTemplateSlideTextShapes(SlideRenderContext context, MarpToPptx.Core.Models.Slide slideModel, ThemeDefinition effectiveTheme)
    {
        var textShapes = GetTemplateSlideTextShapes(context.ShapeTree);
        if (textShapes.Count == 0)
        {
            return false;
        }

        var titleShape = SelectTemplateSlideTitleShape(textShapes);
        var bodyShapes = textShapes
            .Where(candidate => !ReferenceEquals(candidate.Shape, titleShape?.Shape))
            .OrderBy(candidate => candidate.Y)
            .ThenBy(candidate => candidate.X)
            .ToArray();

        HeadingElement? titleHeading = null;
        var bodyTextElements = new List<ISlideElement>();
        var nonTextElements = new List<ISlideElement>();
        foreach (var element in slideModel.Elements)
        {
            if (titleHeading is null && titleShape is not null &&
                element is HeadingElement heading &&
                bodyTextElements.Count == 0 && nonTextElements.Count == 0)
            {
                titleHeading = heading;
                continue;
            }

            switch (element)
            {
                case HeadingElement or ParagraphElement or BulletListElement:
                    bodyTextElements.Add(element);
                    break;
                default:
                    nonTextElements.Add(element);
                    break;
            }
        }

        if (titleShape is not null)
        {
            var titleParagraphs = titleHeading is null
                ? []
                : SplitSpansIntoParagraphs(titleHeading.Spans)
                    .Select(group => new TemplateTextParagraph(group, ForceBold: false))
                    .ToArray();
            ReplaceTemplateSlideTextShape(titleShape.Shape, titleParagraphs, context.SlidePart, context.Language);
        }
        else if (titleHeading is not null)
        {
            bodyTextElements.Insert(0, titleHeading);
        }

        if (bodyShapes.Length > 0)
        {
            var bodyAssignments = AssignElementsToTemplateSlideTextShapes(bodyTextElements, bodyShapes.Length);
            for (var index = 0; index < bodyShapes.Length; index++)
            {
                ReplaceTemplateSlideTextShape(
                    bodyShapes[index].Shape,
                    BuildTemplateSlideParagraphs(bodyAssignments[index]),
                    context.SlidePart,
                    context.Language);
            }
        }
        else if (bodyTextElements.Count > 0)
        {
            nonTextElements.InsertRange(0, bodyTextElements);
        }

        if (nonTextElements.Count > 0)
        {
            var residualSlide = new MarpToPptx.Core.Models.Slide { Style = slideModel.Style };
            residualSlide.Elements.AddRange(nonTextElements);
            var plan = _layoutEngine.LayoutSlide(residualSlide, effectiveTheme);
            var bodyStyle = effectiveTheme.Body;
            foreach (var placed in plan.Elements)
            {
                switch (placed.Element)
                {
                    case HeadingElement heading:
                        AddTextShape(context, placed.Frame, heading.Spans, ResolveHeadingStyle(effectiveTheme, heading.Level), effectiveTheme.InlineCode);
                        break;
                    case ParagraphElement paragraph:
                        AddTextShape(context, placed.Frame, paragraph.Spans, bodyStyle, effectiveTheme.InlineCode);
                        break;
                    case BulletListElement list:
                        AddBulletList(context, placed.Frame, list, bodyStyle);
                        break;
                    case ImageElement image:
                        AddImage(context, placed.Frame, image.Source, image.AltText);
                        break;
                    case VideoElement video:
                        AddVideo(context, placed.Frame, video.Source, video.AltText);
                        break;
                    case AudioElement audio:
                        AddAudio(context, placed.Frame, audio.Source, audio.AltText);
                        break;
                    case CodeBlockElement code:
                        AddCodeBlock(context, placed.Frame, code, effectiveTheme.Code);
                        break;
                    case MermaidDiagramElement mermaid:
                        AddMermaidDiagram(context, placed.Frame, mermaid, effectiveTheme.Code);
                        break;
                    case TableElement table:
                        AddTable(context, placed.Frame, table, bodyStyle);
                        break;
                }
            }
        }

        return true;
    }

    /// <summary>
    /// Creates a slide-level placeholder shape that inherits geometry and text styling
    /// from the matching layout/master placeholder. The shape carries an empty
    /// <c>&lt;p:spPr/&gt;</c> (no transform) and a text body with the supplied paragraphs.
    /// </summary>
    private static P.Shape CreateSlidePlaceholderShape(uint shapeId, string name, TemplatePlaceholder placeholder, IEnumerable<A.Paragraph> paragraphs)
    {
        // Echo the layout placeholder's identity exactly. For typeless content
        // placeholders (<p:ph idx="..."/> on obj layouts such as "Title and Content"),
        // the slide-level ph must ALSO omit the type attribute or PowerPoint may
        // not resolve the inheritance chain.
        var ph = new P.PlaceholderShape();
        if (placeholder.Type is { } phType)
        {
            ph.Type = phType;
        }
        if (placeholder.Index is { } idx)
        {
            ph.Index = idx;
        }

        var textBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in paragraphs)
        {
            textBody.Append(paragraph.CloneNode(true));
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(ph)),
            new P.ShapeProperties(),
            textBody);
    }

    /// <summary>
    /// Builds a paragraph for use inside a template placeholder shape. Unlike
    /// <see cref="CreateParagraphFromSpans"/>, runs here omit font size, colour fill,
    /// and font family so the layout/master text styles cascade. Plain (non-list)
    /// paragraphs emit <c>&lt;a:buNone/&gt;</c> so they do not pick up the body
    /// placeholder's default bullet.
    /// </summary>
    private static A.Paragraph CreateTemplateParagraphFromSpans(
        IReadOnlyList<InlineSpan> spans,
        SlidePart slidePart,
        int? level,
        bool ordered,
        bool forceBold,
        string language,
        int orderNumber = 1)
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
        }
        else
        {
            paragraphProperties.Append(new A.NoBullet());
        }
        paragraph.Append(paragraphProperties);

        foreach (var span in spans.Where(s => s.Text.Length > 0))
        {
            if (span.Text == "\n")
            {
                paragraph.Append(new A.Break());
                continue;
            }

            var runProperties = new A.RunProperties { Language = language };
            if (span.Bold || forceBold)
            {
                runProperties.Bold = true;
            }
            if (span.Italic)
            {
                runProperties.Italic = true;
            }
            if (span.Strikethrough)
            {
                runProperties.Strike = A.TextStrikeValues.SingleStrike;
            }
            if (span.HyperlinkUrl is not null &&
                Uri.TryCreate(span.HyperlinkUrl, UriKind.Absolute, out var hlinkUri))
            {
                var hlinkRel = slidePart.AddHyperlinkRelationship(hlinkUri, true);
                runProperties.Append(new A.HyperlinkOnClick { Id = hlinkRel.Id });
            }

            paragraph.Append(new A.Run(runProperties, new A.Text(span.Text)));
        }

        paragraph.Append(new A.EndParagraphRunProperties { Language = language });
        return paragraph;
    }

    private static IReadOnlyList<List<ISlideElement>> AssignElementsToTemplateSlideTextShapes(IReadOnlyList<ISlideElement> elements, int shapeCount)
    {
        var assignments = Enumerable.Range(0, shapeCount)
            .Select(_ => new List<ISlideElement>())
            .ToArray();

        if (shapeCount == 0)
        {
            return assignments;
        }

        for (var index = 0; index < elements.Count; index++)
        {
            var targetIndex = shapeCount == 1 || index < shapeCount - 1
                ? Math.Min(index, shapeCount - 1)
                : shapeCount - 1;
            assignments[targetIndex].Add(elements[index]);
        }

        return assignments;
    }

    private static IReadOnlyList<TemplateTextParagraph> BuildTemplateSlideParagraphs(IEnumerable<ISlideElement> elements)
    {
        var paragraphs = new List<TemplateTextParagraph>();
        foreach (var element in elements)
        {
            switch (element)
            {
                case HeadingElement heading:
                    paragraphs.AddRange(SplitSpansIntoParagraphs(heading.Spans)
                        .Select(group => new TemplateTextParagraph(group, ForceBold: true)));
                    break;
                case ParagraphElement paragraph:
                    paragraphs.AddRange(SplitSpansIntoParagraphs(paragraph.Spans)
                        .Select(group => new TemplateTextParagraph(group, ForceBold: false)));
                    break;
                case BulletListElement list:
                    var orderNumber = 1;
                    foreach (var item in list.Items)
                    {
                        var prefix = list.Ordered ? $"{orderNumber}. " : "• ";
                        paragraphs.Add(new TemplateTextParagraph([new InlineSpan(prefix), .. item.Spans], ForceBold: false));
                        orderNumber++;
                    }
                    break;
            }
        }

        return paragraphs;
    }

    private static void ReplaceTemplateSlideTextShape(P.Shape shape, IReadOnlyList<TemplateTextParagraph> paragraphs, SlidePart slidePart, string language)
    {
        var existingTextBody = shape.TextBody ?? new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties()));
        var paragraphTemplates = existingTextBody.Elements<A.Paragraph>().ToArray();
        if (paragraphTemplates.Length == 0)
        {
            paragraphTemplates = [new A.Paragraph(new A.EndParagraphRunProperties())];
        }

        var replacementTextBody = new P.TextBody(
            existingTextBody.BodyProperties is null ? new A.BodyProperties() : (A.BodyProperties)existingTextBody.BodyProperties.CloneNode(true),
            existingTextBody.ListStyle is null ? new A.ListStyle() : (A.ListStyle)existingTextBody.ListStyle.CloneNode(true));

        if (paragraphs.Count == 0)
        {
            replacementTextBody.Append(CreateTemplateSlideParagraphFromTemplate(paragraphTemplates[0], new TemplateTextParagraph([], ForceBold: false), slidePart, language));
        }
        else
        {
            for (var index = 0; index < paragraphs.Count; index++)
            {
                var template = paragraphTemplates[Math.Min(index, paragraphTemplates.Length - 1)];
                replacementTextBody.Append(CreateTemplateSlideParagraphFromTemplate(template, paragraphs[index], slidePart, language));
            }
        }

        var existingTextBodyElement = shape.TextBody;
        if (existingTextBodyElement is not null)
        {
            shape.ReplaceChild(replacementTextBody, existingTextBodyElement);
        }
        else
        {
            // Insert the TextBody at the correct position: after spPr/style and before extLst.
            DocumentFormat.OpenXml.OpenXmlElement? insertAfter = null;
            var shapeProperties = shape.GetFirstChild<P.ShapeProperties>();
            var shapeStyle = shape.GetFirstChild<P.ShapeStyle>();

            if (shapeStyle is not null)
            {
                insertAfter = shapeStyle;
            }
            else if (shapeProperties is not null)
            {
                insertAfter = shapeProperties;
            }

            if (insertAfter is not null)
            {
                shape.InsertAfter(replacementTextBody, insertAfter);
            }
            else
            {
                var extLst = shape.GetFirstChild<P.ExtensionList>();
                if (extLst is not null)
                {
                    shape.InsertBefore(replacementTextBody, extLst);
                }
                else
                {
                    shape.Append(replacementTextBody);
                }
            }
        }
    }

    private static A.Paragraph CreateTemplateSlideParagraphFromTemplate(A.Paragraph template, TemplateTextParagraph content, SlidePart slidePart, string language)
    {
        var paragraph = new A.Paragraph();
        if (template.ParagraphProperties is not null)
        {
            paragraph.Append((A.ParagraphProperties)template.ParagraphProperties.CloneNode(true));
        }

        var runTemplate = template.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        foreach (var span in content.Spans.Where(span => span.Text.Length > 0))
        {
            if (span.Text == "\n")
            {
                paragraph.Append(new A.Break());
                continue;
            }

            var runProperties = runTemplate is null
                ? new A.RunProperties()
                : (A.RunProperties)runTemplate.CloneNode(true);
            runProperties.Language = language;

            if (content.ForceBold || span.Bold)
            {
                runProperties.Bold = true;
            }
            if (span.Italic)
            {
                runProperties.Italic = true;
            }
            if (span.Strikethrough)
            {
                runProperties.Strike = A.TextStrikeValues.SingleStrike;
            }

            runProperties.RemoveAllChildren<A.HyperlinkOnClick>();
            if (span.HyperlinkUrl is not null && Uri.TryCreate(span.HyperlinkUrl, UriKind.Absolute, out var hyperlinkUri))
            {
                var hyperlinkRelationship = slidePart.AddHyperlinkRelationship(hyperlinkUri, true);
                runProperties.Append(new A.HyperlinkOnClick { Id = hyperlinkRelationship.Id });
            }

            paragraph.Append(new A.Run(runProperties, new A.Text(span.Text)));
        }

        var templateEndParagraphRunProperties = template.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
        var endParagraphRunProperties = templateEndParagraphRunProperties is null
            ? new A.EndParagraphRunProperties()
            : (A.EndParagraphRunProperties)templateEndParagraphRunProperties.CloneNode(true);
        endParagraphRunProperties.Language = language;
        paragraph.Append(endParagraphRunProperties);
        return paragraph;
    }

    private static IReadOnlyList<TemplateTextShapeCandidate> GetTemplateSlideTextShapes(P.ShapeTree shapeTree)
    {
        var candidates = new List<TemplateTextShapeCandidate>();
        foreach (var shape in shapeTree.Elements<P.Shape>())
        {
            // Skip shapes without text or without valid bounds.
            if (shape.TextBody is null ||
                !TryGetShapeBounds(shape, out var x, out var y, out var cx, out var cy))
            {
                continue;
            }

            // Allow most placeholder-based text boxes (e.g., title/content), but skip
            // known non-content placeholders such as footer, date, and slide number.
            var placeholder = shape.NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<P.PlaceholderShape>();

            if (placeholder is not null)
            {
                var placeholderType = placeholder.Type?.Value;
                if (placeholderType == P.PlaceholderValues.Footer ||
                    placeholderType == P.PlaceholderValues.DateAndTime ||
                    placeholderType == P.PlaceholderValues.SlideNumber)
                {
                    continue;
                }
            }

            candidates.Add(new TemplateTextShapeCandidate(shape, x, y, cx, cy));
        }

        return candidates;
    }

    private static TemplateTextShapeCandidate? SelectTemplateSlideTitleShape(IReadOnlyList<TemplateTextShapeCandidate> textShapes)
    {
        return textShapes
            .Where(shape => shape.Y + (shape.Cy / 2) <= SlideHeightEmu / 2)
            .OrderByDescending(shape => shape.Cx * shape.Cy)
            .ThenBy(shape => shape.Y)
            .FirstOrDefault()
            ?? textShapes.OrderByDescending(shape => shape.Cx * shape.Cy)
                .ThenBy(shape => shape.Y)
                .FirstOrDefault();
    }

    private static bool TryGetShapeBounds(P.Shape shape, out long x, out long y, out long cx, out long cy)
    {
        x = 0L;
        y = 0L;
        cx = 0L;
        cy = 0L;

        var transform = shape.ShapeProperties?.Transform2D;
        if (transform?.Offset is null || transform.Extents is null)
        {
            return false;
        }

        x = transform.Offset.X?.Value ?? 0L;
        y = transform.Offset.Y?.Value ?? 0L;
        cx = transform.Extents.Cx?.Value ?? 0L;
        cy = transform.Extents.Cy?.Value ?? 0L;
        return cx > 0L && cy > 0L;
    }

    private static IReadOnlyList<SlidePart> GetSlidesInPresentationOrder(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>()
            .Where(slideId => !string.IsNullOrWhiteSpace(slideId.RelationshipId))
            .ToArray();
        if (slideIds is null || slideIds.Length == 0)
        {
            return [];
        }

        return slideIds
            .Select(slideId => (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!))
            .ToArray();
    }

    private static void DeleteSlideParts(PresentationPart presentationPart, IReadOnlyList<SlidePart> slideParts)
    {
        foreach (var slidePart in slideParts)
        {
            if (presentationPart.Parts.Any(part => ReferenceEquals(part.OpenXmlPart, slidePart)))
            {
                presentationPart.DeletePart(slidePart);
            }
        }
    }

    private static SlidePart CloneTemplateSlidePart(PresentationPart presentationPart, SlidePart templateSlidePart)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>(GetNextRelationshipId(presentationPart));
        // Clone the slide XML so that each rendered slide gets an independent DOM tree.
        // This is necessary when the same Template[n] is used for multiple slides: the
        // text-replacement pass modifies the DOM in-place, so each slide must start from
        // a fresh copy.
        slidePart.Slide = (P.Slide)templateSlidePart.Slide!.CloneNode(true);
        CopySlidePartRelationships(templateSlidePart, slidePart);
        return slidePart;
    }

    /// <summary>
    /// Creates a pre-clone of <paramref name="source"/> by adding a new <see cref="SlidePart"/>
    /// to <paramref name="presentationPart"/> that shares all sub-parts of <paramref name="source"/>
    /// (images, layouts, etc.). The pre-clone must be created <em>before</em> <c>ClearSlides</c>
    /// runs, because in some SDK environments (e.g., DocumentFormat.OpenXml 3.x on .NET 10)
    /// removing a slide's <c>&lt;p:sldId&gt;</c> XML reference also destroys the originating
    /// <see cref="SlidePart"/> and all its sub-parts. By adding an extra OPC relationship to
    /// each sub-part via the pre-clone, those sub-parts survive the original slide's destruction.
    /// </summary>
    private static SlidePart PreCloneTemplateSlidePart(PresentationPart presentationPart, SlidePart source)
    {
        var clone = presentationPart.AddNewPart<SlidePart>(GetNextRelationshipId(presentationPart));
        clone.Slide = (P.Slide)source.Slide!.CloneNode(true);
        CopySlidePartRelationships(source, clone);
        return clone;
    }

    /// <summary>
    /// Copies all non-notes sub-part relationships, external relationships, and hyperlink
    /// relationships from <paramref name="source"/> to <paramref name="destination"/>.
    /// </summary>
    private static void CopySlidePartRelationships(SlidePart source, SlidePart destination)
    {
        foreach (var relationship in source.Parts)
        {
            if (relationship.OpenXmlPart is NotesSlidePart)
            {
                continue;
            }

            destination.AddPart(relationship.OpenXmlPart, relationship.RelationshipId);
        }

        foreach (var externalRelationship in source.ExternalRelationships)
        {
            destination.AddExternalRelationship(externalRelationship.RelationshipType, externalRelationship.Uri, externalRelationship.Id);
        }

        foreach (var hyperlinkRelationship in source.HyperlinkRelationships)
        {
            destination.AddHyperlinkRelationship(hyperlinkRelationship.Uri, hyperlinkRelationship.IsExternal, hyperlinkRelationship.Id);
        }
    }


    private static void AddNotesSlide(PresentationPart presentationPart, SlidePart slidePart, IReadOnlyList<InlineSpan> noteSpans, string notes, ThemeDefinition theme, string language)
    {
        var notesMasterPart = EnsureNotesMasterPart(presentationPart);
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>(GetNextRelationshipId(slidePart));
        notesSlidePart.AddPart(notesMasterPart, "rId1");
        notesSlidePart.AddPart(slidePart, GetNextRelationshipId(notesSlidePart));

        var effectiveNoteSpans = noteSpans.Count > 0
            ? noteSpans
            : CreateLiteralNoteSpans(notes);

        var noteTextStyle = CreateNotesTextStyle(theme);
        var noteCodeStyle = noteTextStyle with
        {
            Color = theme.InlineCode.Color,
            FontFamily = theme.InlineCode.FontFamily,
        };
        var paragraphs = SplitSpansIntoParagraphs(effectiveNoteSpans)
            .Select(group => CreateParagraphFromSpans(group, noteTextStyle, noteCodeStyle, null, null, false, 1, language))
            .ToArray();

        var notesTextBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in paragraphs)
        {
            notesTextBody.Append(paragraph.CloneNode(true));
        }

        notesSlidePart.NotesSlide = new P.NotesSlide(
            new P.CommonSlideData(new P.ShapeTree(
                CreateRootGroupShapeProperties(),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 2U, Name = "Slide Image 1" },
                        new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true, NoRotation = true, NoChangeAspect = true }),
                        new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.SlideImage })),
                    new P.ShapeProperties()),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 3U, Name = "Notes Placeholder 2" },
                        new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                        new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Body, Index = 1U })),
                    new P.ShapeProperties(),
                    notesTextBody))),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        notesSlidePart.NotesSlide.Save();
    }

    private static NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart)
    {
        if (presentationPart.NotesMasterPart is not null)
        {
            var existingNotesMasterPart = presentationPart.NotesMasterPart;
            EnsureNotesMasterThemePart(presentationPart, existingNotesMasterPart);
            var existingRelId = presentationPart.GetIdOfPart(existingNotesMasterPart);

            var notesMasterIdList = presentationPart.Presentation!.NotesMasterIdList ??= new P.NotesMasterIdList();

            var hasEntry = false;
            foreach (var notesMasterId in notesMasterIdList.Elements<P.NotesMasterId>())
            {
                if (notesMasterId.Id == existingRelId)
                {
                    hasEntry = true;
                    break;
                }
            }

            if (!hasEntry)
            {
                notesMasterIdList.Append(new P.NotesMasterId { Id = existingRelId });
            }

            return existingNotesMasterPart;
        }

        var notesMasterPart = presentationPart.AddNewPart<NotesMasterPart>(GetNextRelationshipId(presentationPart));
        notesMasterPart.NotesMaster = new P.NotesMaster(
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
            });
        EnsureNotesMasterThemePart(presentationPart, notesMasterPart);
        notesMasterPart.NotesMaster.Save();

        var relId = presentationPart.GetIdOfPart(notesMasterPart);
        presentationPart.Presentation!.NotesMasterIdList ??= new P.NotesMasterIdList();
        presentationPart.Presentation.NotesMasterIdList.Append(new P.NotesMasterId { Id = relId });

        return notesMasterPart;
    }

    private static void EnsureNotesMasterThemePart(PresentationPart presentationPart, NotesMasterPart notesMasterPart)
    {
        if (notesMasterPart.ThemePart is not null)
        {
            notesMasterPart.ThemePart.Theme ??= ClonePresentationTheme(presentationPart);
            notesMasterPart.ThemePart.Theme.Save();
            return;
        }

        var themePart = notesMasterPart.AddNewPart<ThemePart>(GetNextRelationshipId(notesMasterPart));
        themePart.Theme = ClonePresentationTheme(presentationPart);
        themePart.Theme.Save();
    }

    private static A.Theme ClonePresentationTheme(PresentationPart presentationPart)
    {
        var sourceTheme = presentationPart.SlideMasterParts.FirstOrDefault()?.ThemePart?.Theme
            ?? presentationPart.ThemePart?.Theme;

        return sourceTheme is null
            ? CreateTheme()
            : (A.Theme)sourceTheme.CloneNode(true);
    }

    private static IReadOnlyList<InlineSpan> CreateLiteralNoteSpans(string notes)
    {
        var spans = new List<InlineSpan>();
        var lines = notes.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n', StringSplitOptions.None);
        for (var index = 0; index < lines.Length; index++)
        {
            if (index > 0)
            {
                spans.Add(new InlineSpan("\n"));
            }

            if (lines[index].Length > 0)
            {
                spans.Add(new InlineSpan(lines[index]));
            }
        }

        return spans;
    }

    private static TextStyle CreateNotesTextStyle(ThemeDefinition theme)
        => new(
            FontSize: 12,
            Color: theme.Body.Color,
            FontFamily: theme.Body.FontFamily,
            Bold: false,
            BackgroundColor: null,
            LineHeight: theme.Body.LineHeight,
            LetterSpacing: theme.Body.LetterSpacing,
            TextTransform: null);

    private static void AddBackground(SlideStyle style, SlideRenderContext context)
    {
        if (context.UseTemplateStyle)
        {
            return;
        }

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

        var backgroundImage = style.BackgroundImage ?? context.Theme.BackgroundImage;
        if (!string.IsNullOrWhiteSpace(backgroundImage))
        {
            var backgroundSize = style.BackgroundSize ?? context.Theme.BackgroundSize;
            var useFullBleed = !string.Equals(backgroundSize, "contain", StringComparison.OrdinalIgnoreCase);
            AddImage(context, new Rect(0, 0, SlideWidthEmu / LayoutScale, SlideHeightEmu / LayoutScale), backgroundImage, string.Empty, useFullBleed: useFullBleed);
        }
    }

    private static void AddTextShape(SlideRenderContext context, Rect frame, string text, TextStyle style, bool isTitle = false)
    {
        var paragraphs = text
            .Split('\n', StringSplitOptions.None)
            .Select(line => CreateParagraph(line, style, null, false, 1, context.Language))
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

    private static void AddTextShape(SlideRenderContext context, Rect frame, IReadOnlyList<InlineSpan> spans, TextStyle style, TextStyle? codeStyle, bool isTitle = false)
    {
        var paragraphGroups = SplitSpansIntoParagraphs(spans);
        var paragraphs = paragraphGroups
            .Select(group => CreateParagraphFromSpans(group, style, codeStyle, context.SlidePart, null, false, 1, context.Language))
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
            .Select((item, index) => CreateParagraphFromSpans(item.Spans, style, context.Theme.InlineCode, context.SlidePart, item.Depth, list.Ordered, index + 1, context.Language))
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
        A.Paragraph[] paragraphs;

        if (SyntaxHighlighter.IsSupported(code.Language))
        {
            var tokenizedLines = SyntaxHighlighter.Tokenize(code.Language, code.Code);
            paragraphs = tokenizedLines
                .Select(runs => CreateHighlightedParagraph(runs, style, context.Language))
                .ToArray();
        }
        else
        {
            paragraphs = code.Code
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Split('\n', StringSplitOptions.None)
                .Select(line => CreateParagraph(line, style, null, false, 1, context.Language))
                .ToArray();
        }

        context.ShapeTree.Append(CreateTextShape(
            context.NextShapeId(),
            string.IsNullOrWhiteSpace(code.Language) ? "Code" : $"Code ({code.Language})",
            frame,
            paragraphs,
            noFill: false,
            fillColor: NormalizeColor(style.BackgroundColor ?? "#0F172A"),
            lineColor: NormalizeColor(context.Theme.AccentColor)));
    }

    private static void AddMermaidDiagram(SlideRenderContext context, Rect frame, MermaidDiagramElement diagram, TextStyle fallbackStyle)
    {
        string svg;
        try
        {
            svg = _diagramRenderer.Render(diagram.Source);
        }
        catch (DiagramParseException ex)
        {
            // Keep the fallback code block and error label within the original frame.
            // If the frame is too small for both (edge case), the code block gets zero height
            // and only the error label is shown.
            var reservedForLabel = MermaidErrorLabelVerticalGap + MermaidErrorLabelHeight;
            var availableCodeHeight = Math.Max(0, frame.Height - reservedForLabel);
            var codeFrame = new Rect(frame.X, frame.Y, frame.Width, availableCodeHeight);
            var fallbackCode = new CodeBlockElement("mermaid", diagram.Source);
            AddCodeBlock(context, codeFrame, fallbackCode, fallbackStyle);

            var labelY = frame.Y + frame.Height - MermaidErrorLabelHeight;
            AddTextShape(
                context,
                new Rect(frame.X, labelY, frame.Width, MermaidErrorLabelHeight),
                $"Mermaid parse error: {ex.Message}",
                fallbackStyle);
            return;
        }

        const string svgContentType = "image/svg+xml";
        var imagePart = context.SlidePart.AddImagePart(svgContentType);
        var svgBytes = System.Text.Encoding.UTF8.GetBytes(svg);
        using (var stream = new MemoryStream(svgBytes, writable: false))
        {
            imagePart.FeedData(stream);
        }

        // Preserve the SVG's intrinsic aspect ratio using the same contain-fit
        // logic as regular images rather than stretching to fill the frame.
        double x, y, width, height;
        if (ImageMetadataReader.TryReadSvgBytesSize(svgBytes, out var svgW, out var svgH) && svgW > 0 && svgH > 0)
        {
            var imageAspect = (double)svgW / svgH;
            var frameAspect = frame.Width / frame.Height;
            if (imageAspect > frameAspect)
            {
                var fittedHeight = frame.Width / imageAspect;
                (x, y, width, height) = (frame.X, frame.Y + ((frame.Height - fittedHeight) / 2), frame.Width, fittedHeight);
            }
            else
            {
                var fittedWidth = frame.Height * imageAspect;
                (x, y, width, height) = (frame.X + ((frame.Width - fittedWidth) / 2), frame.Y, fittedWidth, frame.Height);
            }
        }
        else
        {
            (x, y, width, height) = (frame.X, frame.Y, frame.Width, frame.Height);
        }

        var relationshipId = context.SlidePart.GetIdOfPart(imagePart);
        var blip = CreateImageBlip(svgContentType, relationshipId);

        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = context.NextShapeId(), Name = "Mermaid Diagram", Description = "Mermaid diagram" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                blip,
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmu(x), Y = ToEmu(y) },
                    new A.Extents { Cx = ToEmu(width), Cy = ToEmu(height) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        context.ShapeTree.Append(picture);
    }

    private static A.Paragraph CreateHighlightedParagraph(IReadOnlyList<TokenizedRun> runs, TextStyle style, string language)
    {
        var paragraph = new A.Paragraph();

        if (style.LineHeight.HasValue)
        {
            var paragraphProperties = new A.ParagraphProperties();
            var lineSpacingValue = (int)Math.Round(style.LineHeight.Value * 100000);
            paragraphProperties.Append(new A.LineSpacing(new A.SpacingPercent { Val = lineSpacingValue }));
            paragraph.Append(paragraphProperties);
        }

        foreach (var run in runs)
        {
            if (run.Text.Length == 0)
            {
                continue;
            }

            var runColor = run.Color ?? NormalizeColor(style.Color);
            var runProperties = new A.RunProperties
            {
                Language = language,
                FontSize = (int)Math.Round(style.FontSize * 100),
                Bold = style.Bold,
            };
            if (style.LetterSpacing.HasValue)
            {
                runProperties.Spacing = (int)Math.Round(style.LetterSpacing.Value * 100);
            }

            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = runColor }));
            runProperties.Append(new A.LatinFont { Typeface = style.FontFamily });
            paragraph.Append(new A.Run(runProperties, new A.Text(run.Text)));
        }

        paragraph.Append(new A.EndParagraphRunProperties { Language = language, FontSize = (int)Math.Round(style.FontSize * 100) });
        return paragraph;
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

        var hasHeader = table.Rows.Any(row => row.IsHeader);
        var colWidth = ToEmu(frame.Width) / columnCount;
        var tableStyle = CreateTableTextStyle(style);
        var headerFillColor = NormalizeColor(context.Theme.AccentColor);
        var headerTextColor = GetContrastingTextColor(headerFillColor);
        const string bodyFillColor = "FFFFFF";
        const string bandFillColor = "F8FAFC";
        const string bodyTextColor = "1F2937";

        var tableProperties = new A.TableProperties { FirstRow = hasHeader, BandRow = true };
        tableProperties.Append(new A.TableStyleId(DefaultTableStyleId));

        var tableGrid = new A.TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            tableGrid.Append(new A.GridColumn { Width = colWidth });
        }

        var aTable = new A.Table();
        aTable.Append(tableProperties);
        aTable.Append(tableGrid);

        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            var rowHeight = ToEmu(CalculateTableRowHeight(row, tableStyle, frame.Width, columnCount));
            var aRow = new A.TableRow { Height = rowHeight };
            for (var col = 0; col < columnCount; col++)
            {
                var cellSpans = col < row.Cells.Count ? row.Cells[col] : Array.Empty<InlineSpan>();
                var alignment = col < table.ColumnAlignments.Count ? table.ColumnAlignments[col] : null;
                var fillColor = row.IsHeader
                    ? headerFillColor
                    : rowIndex % 2 == 0 ? bodyFillColor : bandFillColor;
                var textColor = row.IsHeader ? headerTextColor : bodyTextColor;
                var cellCodeStyle = context.Theme.InlineCode is null ? null : context.Theme.InlineCode with { Color = textColor };
                aRow.Append(CreateTableCell(cellSpans, tableStyle with { Color = textColor }, row.IsHeader, alignment, fillColor, context.SlidePart, cellCodeStyle, context.Language));
            }

            aTable.Append(aRow);
        }

        var graphicData = new A.GraphicData(aTable) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
        var graphicFrame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = context.NextShapeId(), Name = "Table" },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = ToEmu(frame.X), Y = ToEmu(frame.Y) },
                new A.Extents { Cx = ToEmu(frame.Width), Cy = ToEmu(frame.Height) }),
            new A.Graphic(graphicData));

        context.ShapeTree.Append(graphicFrame);
    }

    private static A.TableCell CreateTableCell(IReadOnlyList<InlineSpan> spans, TextStyle style, bool isHeader, TableColumnAlignment? alignment, string fillColor, SlidePart slidePart, TextStyle? codeStyle, string language)
    {
        var paragraph = new A.Paragraph();
        var paragraphProperties = new A.ParagraphProperties();
        if (alignment.HasValue)
        {
            paragraphProperties.Alignment = alignment.Value switch
            {
                TableColumnAlignment.Center => A.TextAlignmentTypeValues.Center,
                TableColumnAlignment.Right => A.TextAlignmentTypeValues.Right,
                _ => A.TextAlignmentTypeValues.Left,
            };
        }

        if (style.LineHeight.HasValue)
        {
            var lineSpacingValue = (int)Math.Round(style.LineHeight.Value * 100000);
            paragraphProperties.Append(new A.LineSpacing(new A.SpacingPercent { Val = lineSpacingValue }));
        }

        if (alignment.HasValue || style.LineHeight.HasValue)
        {
            paragraph.Append(paragraphProperties);
        }

        foreach (var span in spans.Where(s => s.Text.Length > 0))
        {
            if (span.Text == "\n")
            {
                paragraph.Append(new A.Break());
                continue;
            }

            var fontFamily = span.Code && codeStyle is not null ? codeStyle.FontFamily : style.FontFamily;
            var color = span.Code && codeStyle is not null ? NormalizeColor(codeStyle.Color) : NormalizeColor(style.Color);
            var runProperties = new A.RunProperties
            {
                Language = language,
                FontSize = (int)Math.Round(style.FontSize * 100),
                Bold = isHeader || span.Bold || style.Bold,
            };
            if (span.Italic)
            {
                runProperties.Italic = true;
            }

            if (span.Strikethrough)
            {
                runProperties.Strike = A.TextStrikeValues.SingleStrike;
            }

            if (style.LetterSpacing.HasValue)
            {
                runProperties.Spacing = (int)Math.Round(style.LetterSpacing.Value * 100);
            }

            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            runProperties.Append(new A.LatinFont { Typeface = fontFamily });

            if (span.HyperlinkUrl is not null && Uri.TryCreate(span.HyperlinkUrl, UriKind.Absolute, out var tableCellHlinkUri))
            {
                var hlinkRel = slidePart.AddHyperlinkRelationship(tableCellHlinkUri, true);
                runProperties.Append(new A.HyperlinkOnClick { Id = hlinkRel.Id });
            }

            var displayText = ApplyTextTransform(span.Text, style.TextTransform);
            paragraph.Append(new A.Run(runProperties, new A.Text(displayText)));
        }

        paragraph.Append(new A.EndParagraphRunProperties { Language = language, FontSize = (int)Math.Round(style.FontSize * 100) });

        var textBody = new A.TextBody(new A.BodyProperties(), new A.ListStyle());
        textBody.Append(paragraph);

        var cell = new A.TableCell();
        cell.Append(textBody);
        cell.Append(new A.TableCellProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = fillColor })));
        return cell;
    }

    private static TextStyle CreateTableTextStyle(TextStyle style)
        => style with
        {
            FontSize = Math.Min(style.FontSize, 18),
            LineHeight = style.LineHeight ?? 1.2,
            BackgroundColor = null,
        };

    private static double CalculateTableRowHeight(TableRowModel row, TextStyle style, double tableWidth, int columnCount)
    {
        var cellWidth = Math.Max(48, tableWidth / Math.Max(1, columnCount) - 12);
        var maxHeight = 0d;
        foreach (var cell in row.Cells)
        {
            var text = string.Concat(cell.Select(span => span.Text));
            maxHeight = Math.Max(maxHeight, EstimateTextBoxHeight(text, style.FontSize, cellWidth, style.LineHeight ?? 1.2) + 10);
        }

        return Math.Max(style.FontSize * 1.6, maxHeight);
    }

    private static double EstimateTextBoxHeight(string text, double fontSize, double width, double lineSpacing)
    {
        // Delegate to the shared layout engine heuristic to keep layout and rendering in sync.
        return LayoutEngine.EstimateTextHeight(text, fontSize, width, lineSpacing);
    }

    private static string GetContrastingTextColor(string backgroundColor)
    {
        var normalized = NormalizeColor(backgroundColor);
        var red = int.Parse(normalized[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var green = int.Parse(normalized[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var blue = int.Parse(normalized[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var luminance = (0.299 * red) + (0.587 * green) + (0.114 * blue);
        return luminance >= 160 ? "1F2937" : "FFFFFF";
    }

    private static void AddImage(SlideRenderContext context, Rect frame, string source, string altText, bool useFullBleed = false)
    {
        if (!TryResolveMediaSource(context, frame, source, "image", out var resolved))
        {
            return;
        }

        var contentType = GetImageContentType(resolved);
        if (contentType is null)
        {
            AddTextShape(context, frame, $"Unsupported image format: {source}", context.Theme.Body);
            return;
        }

        var imagePart = context.SlidePart.AddImagePart(contentType);
        using (var imageStream = File.OpenRead(resolved))
        {
            imagePart.FeedData(imageStream);
        }

        var (x, y, width, height) = CalculateImagePlacement(frame, resolved, useFullBleed);
        var relationshipId = context.SlidePart.GetIdOfPart(imagePart);
        var blip = CreateImageBlip(contentType, relationshipId);

        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = context.NextShapeId(), Name = IOPath.GetFileName(resolved), Description = altText },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                blip,
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmu(x), Y = ToEmu(y) },
                    new A.Extents { Cx = ToEmu(width), Cy = ToEmu(height) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        context.ShapeTree.Append(picture);
    }

    private static A.Blip CreateImageBlip(string contentType, string relationshipId)
    {
        if (!string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase))
        {
            return new A.Blip { Embed = relationshipId };
        }

        var svgBlip = new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip
        {
            Embed = relationshipId,
        };

        return new A.Blip(
            new A.BlipExtensionList(
                new A.BlipExtension(svgBlip)
                {
                    Uri = SvgBlipExtensionUri,
                }));
    }

    private static void AddVideo(SlideRenderContext context, Rect frame, string source, string altText)
    {
        if (!TryResolveMediaSource(context, frame, source, "video", out var resolved))
        {
            return;
        }

        var ext = IOPath.GetExtension(resolved).ToLowerInvariant();
        if (ext != ".mp4")
        {
            AddTextShape(context, frame, $"Unsupported video format: {source}", context.Theme.Body);
            return;
        }

        var mediaDataPart = context.SlidePart.OpenXmlPackage.CreateMediaDataPart("video/mp4", ".mp4");
        using (var videoStream = File.OpenRead(resolved))
        {
            mediaDataPart.FeedData(videoStream);
        }

        var videoRel = context.SlidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRel = context.SlidePart.AddMediaReferenceRelationship(mediaDataPart);
        var placeholderImageRelId = AddMediaPlaceholderImage(context.SlidePart);

        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                CreateMediaDrawingProperties(context.NextShapeId(), IOPath.GetFileName(resolved), altText),
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                CreateMediaApplicationProperties(new A.VideoFromFile { Link = videoRel.Id }, mediaRel.Id)),
            new P.BlipFill(
                new A.Blip { Embed = placeholderImageRelId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmu(frame.X), Y = ToEmu(frame.Y) },
                    new A.Extents { Cx = ToEmu(frame.Width), Cy = ToEmu(frame.Height) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        context.ShapeTree.Append(picture);
    }

    private static void AddAudio(SlideRenderContext context, Rect frame, string source, string altText)
    {
        if (!TryResolveMediaSource(context, frame, source, "audio", out var resolved))
        {
            return;
        }

        var ext = IOPath.GetExtension(resolved).ToLowerInvariant();
        (string ContentType, string Extension)? mediaPartInfo = ext switch
        {
            ".mp3" => (ContentType: "audio/mp3", Extension: ".mp3"),
            ".wav" => (ContentType: "audio/wav", Extension: ".wav"),
            ".m4a" => (ContentType: "audio/mp4", Extension: ".m4a"),
            _ => null,
        };

        if (mediaPartInfo is null)
        {
            AddTextShape(context, frame, $"Unsupported audio format: {source}", context.Theme.Body);
            return;
        }

        var mediaDataPart = context.SlidePart.OpenXmlPackage.CreateMediaDataPart(mediaPartInfo.Value.ContentType, mediaPartInfo.Value.Extension);
        using (var audioStream = File.OpenRead(resolved))
        {
            mediaDataPart.FeedData(audioStream);
        }

        var audioRel = context.SlidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRel = context.SlidePart.AddMediaReferenceRelationship(mediaDataPart);
        var placeholderImageRelId = AddMediaPlaceholderImage(context.SlidePart);

        // Position a small audio icon shape in the center of the frame.
        var iconSize = 60.0;
        var iconX = frame.X + ((frame.Width - iconSize) / 2);
        var iconY = frame.Y + ((frame.Height - iconSize) / 2);

        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                CreateMediaDrawingProperties(context.NextShapeId(), IOPath.GetFileName(resolved), altText),
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                CreateMediaApplicationProperties(new A.AudioFromFile { Link = audioRel.Id }, mediaRel.Id)),
            new P.BlipFill(
                new A.Blip { Embed = placeholderImageRelId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmu(iconX), Y = ToEmu(iconY) },
                    new A.Extents { Cx = ToEmu(iconSize), Cy = ToEmu(iconSize) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        context.ShapeTree.Append(picture);
    }

    private static P.NonVisualDrawingProperties CreateMediaDrawingProperties(uint shapeId, string name, string altText)
    {
        var drawingProperties = new P.NonVisualDrawingProperties
        {
            Id = shapeId,
            Name = name,
            Description = altText,
        };

        // PowerPoint uses this action marker to open the embedded media on click.
        drawingProperties.Append(new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" });
        return drawingProperties;
    }

    private static P.ApplicationNonVisualDrawingProperties CreateMediaApplicationProperties(DocumentFormat.OpenXml.OpenXmlElement mediaFileElement, string mediaEmbedRelationshipId)
    {
        var extension = new P.ApplicationNonVisualDrawingPropertiesExtension
        {
            Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}",
        };

        var media = new Media { Embed = mediaEmbedRelationshipId };
        media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
        extension.Append(media);

        return new P.ApplicationNonVisualDrawingProperties(
            mediaFileElement,
            new P.ApplicationNonVisualDrawingPropertiesExtensionList(extension));
    }

    private static string AddMediaPlaceholderImage(SlidePart slidePart)
    {
        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using var imageStream = new MemoryStream(MediaPlaceholderImage, writable: false);
        imagePart.FeedData(imageStream);
        return slidePart.GetIdOfPart(imagePart);
    }

    private static bool TryResolveMediaSource(SlideRenderContext context, Rect frame, string source, string mediaKind, out string resolved)
    {
        resolved = string.Empty;

        var candidate = ResolvePath(context.SourceDirectory, source, context.RemoteAssets, out var resolveError);
        if (candidate is null || !File.Exists(candidate))
        {
            var errorText = resolveError is not null
                ? $"Missing {mediaKind}: {source} ({resolveError})"
                : string.IsNullOrWhiteSpace(source) ? $"Missing {mediaKind}" : $"Missing {mediaKind}: {source}";
            AddTextShape(context, frame, errorText, context.Theme.Body);
            return false;
        }

        resolved = candidate;
        return true;
    }

    private static void AddHeaderFooterAndPageNumber(SlideRenderContext context, SlideStyle style, TextStyle bodyStyle, int slideNumber)
    {
        // Slide dimensions in points: width ≈ 960pt, height ≈ 540pt.
        const double slideWidth = SlideWidthEmu / (double)LayoutScale;
        const double slideHeight = SlideHeightEmu / (double)LayoutScale;
        const double marginX = 30.0;
        const double footerY = slideHeight - 20.0;
        const double footerHeight = 18.0;
        const double headerY = 4.0;
        const double headerHeight = 18.0;
        const double pageNumWidth = 60.0;

        var footerStyle = new TextStyle(10, bodyStyle.Color, bodyStyle.FontFamily, false);

        if (!context.UseTemplateStyle && !string.IsNullOrWhiteSpace(style.Header))
        {
            var headerWidth = slideWidth - (marginX * 2);
            AddTextShape(context, new Rect(marginX, headerY, headerWidth, headerHeight), style.Header!, footerStyle);
        }

        if (!context.UseTemplateStyle && !string.IsNullOrWhiteSpace(style.Footer))
        {
            var footerWidth = style.Paginate == true
                ? slideWidth - (marginX * 2) - pageNumWidth - 8.0
                : slideWidth - (marginX * 2);
            AddTextShape(context, new Rect(marginX, footerY, footerWidth, footerHeight), style.Footer!, footerStyle);
        }

        // For template-bound slides, rely on the template's own sldNum placeholder for slide
        // numbering and styling; emitting an explicit standalone number shape would duplicate
        // the counter and force the default MarpToPptx theme colors onto the slide number.
        if (!context.UseTemplateStyle && style.Paginate == true)
        {
            var pageNumX = slideWidth - marginX - pageNumWidth;
            var fieldId = Guid.NewGuid().ToString("B").ToUpperInvariant();
            var fieldRunProperties = new A.RunProperties { Language = context.Language, FontSize = (int)Math.Round(footerStyle.FontSize * 100) };
            fieldRunProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = NormalizeColor(footerStyle.Color) }));
            fieldRunProperties.Append(new A.LatinFont { Typeface = footerStyle.FontFamily });

            var field = new A.Field(
                fieldRunProperties,
                new A.Text(slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)))
            {
                Id = fieldId,
                Type = "slidenum",
            };

            var endRunProperties = new A.EndParagraphRunProperties { Language = context.Language, FontSize = (int)Math.Round(footerStyle.FontSize * 100) };
            endRunProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = NormalizeColor(footerStyle.Color) }));
            endRunProperties.Append(new A.LatinFont { Typeface = footerStyle.FontFamily });

            var paragraphProperties = new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Right };
            var paragraph = new A.Paragraph(paragraphProperties, field, endRunProperties);

            context.ShapeTree.Append(CreateTextShape(
                context.NextShapeId(),
                "Slide Number",
                new Rect(pageNumX, footerY, pageNumWidth, footerHeight),
                [paragraph],
                noFill: true,
                fillColor: null,
                lineColor: null));
        }
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

    private static string? ResolvePath(string? sourceDirectory, string source, RemoteAssetResolver? remoteAssets, out string? errorMessage)
    {
        errorMessage = null;

        if (string.IsNullOrWhiteSpace(source))
        {
            return null;
        }

        if (Uri.TryCreate(source, UriKind.Absolute, out var uri) && !uri.IsFile)
        {
            if (uri.Scheme is "http" or "https" && remoteAssets is not null)
            {
                return remoteAssets.Resolve(source, out errorMessage);
            }

            if (uri.Scheme is "http" or "https")
            {
                errorMessage = "remote assets are disabled";
            }

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

    private static A.Paragraph CreateParagraph(string text, TextStyle style, int? level, bool ordered, int orderNumber, string language = "en-US")
    {
        var paragraph = new A.Paragraph();

        if (level is not null)
        {
            var paragraphProperties = new A.ParagraphProperties();
            paragraphProperties.Level = level.Value;
            if (ordered)
            {
                paragraphProperties.Append(new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod, StartAt = orderNumber });
            }
            else
            {
                paragraphProperties.Append(new A.CharacterBullet { Char = "•" });
            }

            paragraph.Append(paragraphProperties);
        }
        else if (style.LineHeight.HasValue)
        {
            var paragraphProperties = new A.ParagraphProperties();
            var lineSpacingValue = (int)Math.Round(style.LineHeight.Value * 100000);
            paragraphProperties.Append(new A.LineSpacing(new A.SpacingPercent { Val = lineSpacingValue }));
            paragraph.Append(paragraphProperties);
        }

        if (!string.IsNullOrEmpty(text))
        {
            var runProperties = new A.RunProperties
            {
                Language = language,
                FontSize = (int)Math.Round(style.FontSize * 100),
                Bold = style.Bold,
            };
            if (style.LetterSpacing.HasValue)
            {
                runProperties.Spacing = (int)Math.Round(style.LetterSpacing.Value * 100);
            }

            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = NormalizeColor(style.Color) }));
            runProperties.Append(new A.LatinFont { Typeface = style.FontFamily });

            var displayText = ApplyTextTransform(text, style.TextTransform);
            paragraph.Append(new A.Run(runProperties, new A.Text(displayText)));
        }

        paragraph.Append(new A.EndParagraphRunProperties { Language = language, FontSize = (int)Math.Round(style.FontSize * 100) });
        return paragraph;
    }

    private static A.Paragraph CreateParagraphFromSpans(
        IReadOnlyList<InlineSpan> spans,
        TextStyle style,
        TextStyle? codeStyle,
        SlidePart? slidePart,
        int? level,
        bool ordered,
        int orderNumber,
        string language = "en-US")
    {
        var paragraph = new A.Paragraph();

        if (level is not null)
        {
            var paragraphProperties = new A.ParagraphProperties();
            paragraphProperties.Level = level.Value;
            if (ordered)
            {
                paragraphProperties.Append(new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod, StartAt = orderNumber });
            }
            else
            {
                paragraphProperties.Append(new A.CharacterBullet { Char = "•" });
            }

            paragraph.Append(paragraphProperties);
        }
        else if (style.LineHeight.HasValue)
        {
            var paragraphProperties = new A.ParagraphProperties();
            var lineSpacingValue = (int)Math.Round(style.LineHeight.Value * 100000);
            paragraphProperties.Append(new A.LineSpacing(new A.SpacingPercent { Val = lineSpacingValue }));
            paragraph.Append(paragraphProperties);
        }

        foreach (var span in spans.Where(s => s.Text.Length > 0))
        {
            if (span.Text == "\n")
            {
                paragraph.Append(new A.Break());
                continue;
            }

            var isCode = span.Code && codeStyle is not null;
            var fontFamily = isCode ? codeStyle!.FontFamily : style.FontFamily;
            var color = isCode ? NormalizeColor(codeStyle!.Color) : NormalizeColor(style.Color);

            var runProperties = new A.RunProperties
            {
                Language = language,
                FontSize = (int)Math.Round(style.FontSize * 100),
                Bold = span.Bold || style.Bold,
            };
            if (span.Italic)
            {
                runProperties.Italic = true;
            }

            if (span.Strikethrough)
            {
                runProperties.Strike = A.TextStrikeValues.SingleStrike;
            }

            if (style.LetterSpacing.HasValue)
            {
                runProperties.Spacing = (int)Math.Round(style.LetterSpacing.Value * 100);
            }

            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            runProperties.Append(new A.LatinFont { Typeface = fontFamily });

            if (span.HyperlinkUrl is not null && slidePart is not null &&
                Uri.TryCreate(span.HyperlinkUrl, UriKind.Absolute, out var hlinkUri))
            {
                var hlinkRel = slidePart.AddHyperlinkRelationship(hlinkUri, true);
                runProperties.Append(new A.HyperlinkOnClick { Id = hlinkRel.Id });
            }

            var displayText = ApplyTextTransform(span.Text, style.TextTransform);
            paragraph.Append(new A.Run(runProperties, new A.Text(displayText)));
        }

        paragraph.Append(new A.EndParagraphRunProperties { Language = language, FontSize = (int)Math.Round(style.FontSize * 100) });
        return paragraph;
    }

    /// <summary>
    /// Splits a flat span list into paragraph groups, using spans whose <c>Text</c> equals
    /// <c>"\n"</c> as paragraph-break markers.
    /// </summary>
    private static IReadOnlyList<IReadOnlyList<InlineSpan>> SplitSpansIntoParagraphs(IReadOnlyList<InlineSpan> spans)
    {
        var result = new List<List<InlineSpan>>();
        var current = new List<InlineSpan>();

        foreach (var span in spans)
        {
            if (span.Text == "\n")
            {
                result.Add(current);
                current = new List<InlineSpan>();
            }
            else
            {
                current.Add(span);
            }
        }

        result.Add(current);
        return result;
    }

    private static string ApplyTextTransform(string text, string? textTransform)
        => textTransform switch
        {
            "uppercase" => text.ToUpperInvariant(),
            "lowercase" => text.ToLowerInvariant(),
            // ToTitleCase capitalises the first letter of each word, which matches CSS capitalize intent.
            // Unlike CSS, it may downcase other letters for some casing-sensitive locales.
            "capitalize" => System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(text.ToLowerInvariant()),
            _ => text,
        };

    private static P.NonVisualGroupShapeProperties CreateRootGroupShapeProperties()
        => new(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties());

    private static P.Shape CreatePlaceholderShape(uint shapeId, string name, P.PlaceholderValues? placeholderType, uint? index = null, Rect? frame = null)
    {
        var placeholder = placeholderType is null
            ? new P.PlaceholderShape()
            : new P.PlaceholderShape { Type = placeholderType };

        if (index is not null)
        {
            placeholder.Index = index;
        }

        var appProperties = new P.ApplicationNonVisualDrawingProperties(placeholder);

        var shapeProperties = frame is Rect frameValue
            ? new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmu(frameValue.X), Y = ToEmu(frameValue.Y) },
                    new A.Extents { Cx = ToEmu(frameValue.Width), Cy = ToEmu(frameValue.Height) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
            : new P.ShapeProperties();

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                appProperties),
            shapeProperties,
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties())));
    }

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
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 50000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 0 },
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 37000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 35000 },
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 15000 },
                                        new A.SaturationModulation { Val = 350000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 100000 }),
                            new A.LinearGradientFill { Angle = 16200000, Scaled = true }),
                        new A.NoFill(),
                        new A.PatternFill(),
                        new A.GroupFill()),
                    new A.LineStyleList(
                        CreateThemeOutline(),
                        CreateThemeOutline(),
                        CreateThemeOutline()),
                    new A.EffectStyleList(
                        CreateThemeEffectStyle(),
                        CreateThemeEffectStyle(),
                        CreateThemeEffectStyle()),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 50000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 0 },
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 50000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 50000 },
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 50000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 100000 }),
                            new A.LinearGradientFill { Angle = 16200000, Scaled = true }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 50000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 0 },
                                new A.GradientStop(
                                    new A.SchemeColor(
                                        new A.Tint { Val = 50000 },
                                        new A.SaturationModulation { Val = 300000 })
                                    { Val = A.SchemeColorValues.PhColor })
                                { Position = 100000 }),
                            new A.LinearGradientFill { Angle = 16200000, Scaled = true }),
                        new A.NoFill(),
                        new A.PatternFill(),
                        new A.GroupFill()))
                { Name = "MarpToPptx" }),
            ObjectDefaults = new A.ObjectDefaults(),
            ExtraColorSchemeList = new A.ExtraColorSchemeList(),
        };

    private static string GetPresentationTitle(SlideDeck deck)
        => deck.Slides
            .Select(GetSlideTitle)
            .FirstOrDefault(static title => !string.IsNullOrWhiteSpace(title))
            ?? (string.IsNullOrWhiteSpace(deck.SourcePath) ? "PowerPoint Presentation" : IOPath.GetFileNameWithoutExtension(deck.SourcePath));

    private static string GetSlideTitle(MarpToPptx.Core.Models.Slide slide)
        => slide.Elements.OfType<HeadingElement>().FirstOrDefault()?.Text?.Trim()
            ?? "PowerPoint Presentation";

    private static XDocument CreateExtendedPropertiesDocument(SlideDeck deck)
    {
        XNamespace ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        var slideTitles = deck.Slides.Select(GetSlideTitle).ToList();
        var titlesOfParts = new List<string> { "Aptos", "Aptos Display", "MarpToPptx Theme" };
        titlesOfParts.AddRange(slideTitles);

        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(
                ep + "Properties",
                new XAttribute(XNamespace.Xmlns + "vt", vt),
                new XElement(ep + "TotalTime", 0),
                new XElement(ep + "Words", 0),
                new XElement(ep + "Application", "Microsoft Office PowerPoint"),
                new XElement(ep + "PresentationFormat", "On-screen Show (16:9)"),
                new XElement(ep + "Paragraphs", 0),
                new XElement(ep + "Slides", slideTitles.Count),
                new XElement(ep + "Notes", 0),
                new XElement(ep + "HiddenSlides", 0),
                new XElement(ep + "MMClips", 0),
                new XElement(ep + "ScaleCrop", "false"),
                new XElement(
                    ep + "HeadingPairs",
                    new XElement(
                        vt + "vector",
                        new XAttribute("size", 6),
                        new XAttribute("baseType", "variant"),
                        CreateVariantString(vt, "Fonts Used"),
                        CreateVariantInt(vt, 2),
                        CreateVariantString(vt, "Theme"),
                        CreateVariantInt(vt, 1),
                        CreateVariantString(vt, "Slide Titles"),
                        CreateVariantInt(vt, slideTitles.Count))),
                new XElement(
                    ep + "TitlesOfParts",
                    new XElement(
                        vt + "vector",
                        new XAttribute("size", titlesOfParts.Count),
                        new XAttribute("baseType", "lpstr"),
                        titlesOfParts.Select(title => new XElement(vt + "lpstr", title)))),
                new XElement(ep + "Company", "Created by MarpToPptx"),
                new XElement(ep + "LinksUpToDate", "false"),
                new XElement(ep + "SharedDoc", "false"),
                new XElement(ep + "HyperlinksChanged", "false"),
                new XElement(ep + "AppVersion", "16.0000")));
    }

    private static XDocument CreateCorePropertiesDocument(SlideDeck deck, DateTime now)
    {
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        XNamespace dcterms = "http://purl.org/dc/terms/";
        XNamespace dcmitype = "http://purl.org/dc/dcmitype/";
        XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";

        var coreProperties = new XElement(
            cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", cp),
            new XAttribute(XNamespace.Xmlns + "dc", dc),
            new XAttribute(XNamespace.Xmlns + "dcterms", dcterms),
            new XAttribute(XNamespace.Xmlns + "dcmitype", dcmitype),
            new XAttribute(XNamespace.Xmlns + "xsi", xsi),
            new XElement(dc + "title", GetPresentationTitle(deck)),
            new XElement(dc + "subject", "PowerPoint Presentation"),
            new XElement(dc + "creator", "MarpToPptx"),
            new XElement(cp + "lastModifiedBy", "MarpToPptx"),
            new XElement(cp + "revision", "1"),
            new XElement(dcterms + "created", new XAttribute(xsi + "type", "dcterms:W3CDTF"), now.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)),
            new XElement(dcterms + "modified", new XAttribute(xsi + "type", "dcterms:W3CDTF"), now.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)));

        if (deck.Language is not null)
        {
            coreProperties.Add(new XElement(dc + "language", deck.Language));
        }

        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            coreProperties);
    }

    private static XElement CreateVariantString(XNamespace variantNamespace, string value)
        => new(variantNamespace + "variant", new XElement(variantNamespace + "lpstr", value));

    private static XElement CreateVariantInt(XNamespace variantNamespace, int value)
        => new(variantNamespace + "variant", new XElement(variantNamespace + "i4", value));

    private static XDocument CreateViewPropertiesDocument()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(
                p + "viewPr",
                new XAttribute(XNamespace.Xmlns + "a", a),
                new XAttribute(XNamespace.Xmlns + "r", r),
                new XElement(
                    p + "normalViewPr",
                    new XAttribute("horzBarState", "maximized"),
                    new XElement(p + "restoredLeft", new XAttribute("sz", "15611")),
                    new XElement(p + "restoredTop", new XAttribute("sz", "94610"))),
                new XElement(
                    p + "slideViewPr",
                    new XElement(
                        p + "cSldViewPr",
                        new XAttribute("snapToGrid", "0"),
                        new XAttribute("snapToObjects", "1"),
                        new XElement(
                            p + "cViewPr",
                            new XAttribute("varScale", "1"),
                            new XElement(
                                p + "scale",
                                new XElement(a + "sx", new XAttribute("n", "136"), new XAttribute("d", "100")),
                                new XElement(a + "sy", new XAttribute("n", "136"), new XAttribute("d", "100"))),
                            new XElement(p + "origin", new XAttribute("x", "216"), new XAttribute("y", "312"))),
                        new XElement(p + "guideLst"))),
                new XElement(
                    p + "notesTextViewPr",
                    new XElement(
                        p + "cViewPr",
                        new XElement(
                            p + "scale",
                            new XElement(a + "sx", new XAttribute("n", "1"), new XAttribute("d", "1")),
                            new XElement(a + "sy", new XAttribute("n", "1"), new XAttribute("d", "1"))),
                        new XElement(p + "origin", new XAttribute("x", "0"), new XAttribute("y", "0")))),
                new XElement(p + "gridSpacing", new XAttribute("cx", "76200"), new XAttribute("cy", "76200"))));
    }

    private static void WriteXmlPart(OpenXmlPart part, XDocument document)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        document.Save(stream);
    }

    private static void NormalizePackage(string outputPath)
    {
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        NormalizeContentTypes(archive);
        NormalizeRelationships(archive);
    }

    private static void NormalizeContentTypes(ZipArchive archive)
    {
        var entry = archive.GetEntry("[Content_Types].xml");
        if (entry is null)
        {
            return;
        }

        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        XDocument document;
        using (var stream = entry.Open())
        {
            document = XDocument.Load(stream);
        }

        var root = document.Root;
        if (root is null)
        {
            return;
        }

        foreach (var defaultElement in root.Elements(ct + "Default").Where(element => (string?)element.Attribute("Extension") == "xml").ToList())
        {
            defaultElement.Remove();
        }

        root.AddFirst(new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")));

        if (!root.Elements(ct + "Override").Any(element => (string?)element.Attribute("PartName") == "/ppt/presentation.xml"))
        {
            root.Add(new XElement(
                ct + "Override",
                new XAttribute("PartName", "/ppt/presentation.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")));
        }

        ReplaceArchiveEntry(archive, entry.FullName, document);
    }

    private static void NormalizeRelationships(ZipArchive archive)
    {
        foreach (var entry in archive.Entries.Where(static entry => entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            XDocument document;
            using (var stream = entry.Open())
            {
                document = XDocument.Load(stream);
            }

            var root = document.Root;
            if (root is null)
            {
                continue;
            }

            var sourcePartPath = GetRelationshipSourcePartPath(entry.FullName);
            var changed = false;
            foreach (var relationship in root.Elements().Where(element => element.Name.LocalName == "Relationship"))
            {
                var target = relationship.Attribute("Target")?.Value;
                var targetMode = relationship.Attribute("TargetMode")?.Value;
                if (string.IsNullOrWhiteSpace(target) || string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase) || !target.StartsWith("/", StringComparison.Ordinal))
                {
                    continue;
                }

                relationship.SetAttributeValue("Target", MakeRelativeTarget(sourcePartPath, target));
                changed = true;
            }

            if (changed)
            {
                ReplaceArchiveEntry(archive, entry.FullName, document);
            }
        }
    }

    private static string GetRelationshipSourcePartPath(string relationshipPath)
    {
        if (string.Equals(relationshipPath, "_rels/.rels", StringComparison.OrdinalIgnoreCase))
        {
            return "/";
        }

        var marker = "/_rels/";
        var markerIndex = relationshipPath.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            return "/";
        }

        var directory = relationshipPath[..markerIndex];
        var fileName = relationshipPath[(markerIndex + marker.Length)..];
        return $"/{directory}/{fileName[..^5]}";
    }

    private static string MakeRelativeTarget(string sourcePartPath, string absoluteTarget)
    {
        if (sourcePartPath == "/")
        {
            return absoluteTarget.TrimStart('/');
        }

        var sourceDirectory = sourcePartPath[..(sourcePartPath.LastIndexOf('/') + 1)];
        var baseUri = new Uri($"https://package{sourceDirectory}", UriKind.Absolute);
        var targetUri = new Uri($"https://package{absoluteTarget}", UriKind.Absolute);
        return Uri.UnescapeDataString(baseUri.MakeRelativeUri(targetUri).ToString());
    }

    private static void ReplaceArchiveEntry(ZipArchive archive, string entryName, XDocument document)
    {
        archive.GetEntry(entryName)?.Delete();
        var replacement = archive.CreateEntry(entryName);
        using var stream = replacement.Open();
        document.Save(stream);
    }

    private static string GetNextRelationshipId(OpenXmlPartContainer container)
    {
        var usedIds = container.Parts
            .Select(part => part.RelationshipId)
            .Where(static id => id.StartsWith("rId", StringComparison.Ordinal))
            .Select(static id => int.TryParse(id[3..], out var value) ? value : 0)
            .Where(static value => value > 0)
            .ToHashSet();

        var next = 1;
        while (usedIds.Contains(next))
        {
            next++;
        }

        return $"rId{next}";
    }

    private static void EnsureRelationshipId(OpenXmlPartContainer container, OpenXmlPart part, string relationshipId)
    {
        if (container.GetIdOfPart(part) == relationshipId)
        {
            return;
        }

        container.ChangeIdOfPart(part, relationshipId);
    }

    private static A.Outline CreateThemeOutline()
        => new(
            new A.SolidFill(
                new A.SchemeColor(
                    new A.Shade { Val = 95000 },
                    new A.SaturationModulation { Val = 105000 })
                { Val = A.SchemeColorValues.PhColor }),
            new A.PresetDash { Val = A.PresetLineDashValues.Solid })
        {
            Width = 9525,
            CapType = A.LineCapValues.Flat,
            CompoundLineType = A.CompoundLineValues.Single,
            Alignment = A.PenAlignmentValues.Center,
        };

    private static A.EffectStyle CreateThemeEffectStyle()
        => new(
            new A.EffectList(
                new A.OuterShadow(
                    new A.RgbColorModelHex(
                        new A.Alpha { Val = 38000 })
                    { Val = "000000" })
                {
                    BlurRadius = 40000L,
                    Distance = 20000L,
                    Direction = 5400000,
                    RotateWithShape = false,
                }));

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

        var trimmed = value.Trim().Trim('"', '\'');
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

    private static string? GetImageContentType(string path)
    {
        var ext = IOPath.GetExtension(path).ToLowerInvariant();
        var typeFromExtension = ext switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            ".svg" => "image/svg+xml",
            ".webp" => "image/webp",
            _ => null,
        };

        if (typeFromExtension is not null)
        {
            return typeFromExtension;
        }

        // Fall back to magic-byte detection for files without a recognized extension
        // (e.g. remote images downloaded to a .bin temp file).
        try
        {
            using var stream = File.OpenRead(path);
            return ImageMetadataReader.TryDetectContentType(stream);
        }
        catch
        {
            return null;
        }
    }

    private sealed record TemplateTextParagraph(IReadOnlyList<InlineSpan> Spans, bool ForceBold);

    private sealed record TemplateTextShapeCandidate(P.Shape Shape, long X, long Y, long Cx, long Cy);

    private sealed class SlideRenderContext(SlidePart slidePart, P.ShapeTree shapeTree, string? sourceDirectory, ThemeDefinition theme, RemoteAssetResolver? remoteAssets, bool useTemplateStyle, string language = "en-US")
    {
        private uint _shapeId = shapeTree.Descendants<P.NonVisualDrawingProperties>()
            .Select(properties => properties.Id?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max();

        public SlidePart SlidePart { get; } = slidePart;

        public P.ShapeTree ShapeTree { get; } = shapeTree;

        public string? SourceDirectory { get; } = sourceDirectory;

        public ThemeDefinition Theme { get; } = theme;

        public RemoteAssetResolver? RemoteAssets { get; } = remoteAssets;

        public bool UseTemplateStyle { get; } = useTemplateStyle;

        public string Language { get; } = language;

        public uint NextShapeId() => ++_shapeId;
    }
}
