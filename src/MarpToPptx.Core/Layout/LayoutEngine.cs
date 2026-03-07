using MarpToPptx.Core.Models;
using MarpToPptx.Core.Themes;

namespace MarpToPptx.Core.Layout;

public sealed class LayoutEngine
{
    public LayoutPlan LayoutSlide(Slide slide, ThemeDefinition theme, LayoutOptions? options = null)
    {
        options ??= LayoutOptions.Default;

        var plan = new LayoutPlan(options.SlideWidth, options.SlideHeight);
        var contentX = theme.SlidePadding.Left;
        var contentWidth = options.SlideWidth - theme.SlidePadding.Left - theme.SlidePadding.Right;
        var y = theme.SlidePadding.Top;
        var isTitleSlide = slide.Elements.Count > 0 && slide.Elements[0] is HeadingElement && slide.Elements.Count <= 2;

        foreach (var element in slide.Elements)
        {
            var frame = element switch
            {
                HeadingElement heading => CreateHeadingFrame(heading, theme, contentX, contentWidth, ref y, isTitleSlide),
                ParagraphElement paragraph => CreateParagraphFrame(paragraph.Text, theme.Body.FontSize, contentX, contentWidth, ref y),
                BulletListElement list => CreateParagraphFrame(string.Join("\n", list.Items.Select(item => item.Text)), theme.Body.FontSize, contentX, contentWidth, ref y, 1.2),
                ImageElement => CreateFixedFrame(contentX, contentWidth, ref y, 220),
                VideoElement => CreateFixedFrame(contentX, contentWidth, ref y, 220),
                AudioElement => CreateFixedFrame(contentX, contentWidth, ref y, 80),
                CodeBlockElement code => CreateCodeBlockFrame(code, theme, contentX, contentWidth, ref y),
                TableElement table => CreateFixedFrame(contentX, contentWidth, ref y, Math.Max(120, table.Rows.Count * 26 + 20)),
                _ => CreateFixedFrame(contentX, contentWidth, ref y, 80),
            };

            plan.Elements.Add(new PlacedElement(element, frame));
        }

        return plan;
    }

    private static Rect CreateHeadingFrame(HeadingElement heading, ThemeDefinition theme, double x, double width, ref double y, bool isTitleSlide)
    {
        var style = theme.GetHeadingStyle(heading.Level);
        var height = EstimateTextHeight(heading.Text, style.FontSize, width, isTitleSlide ? 1.15 : 1.1);
        var top = isTitleSlide ? Math.Max(y, 110) : y;
        y = top + height + (isTitleSlide ? 18 : 10);
        return new Rect(x, top, width, height);
    }

    private static Rect CreateParagraphFrame(string text, double fontSize, double x, double width, ref double y, double lineSpacing = 1.35)
    {
        var height = EstimateTextHeight(text, fontSize, width, lineSpacing);
        var frame = new Rect(x, y, width, height);
        y += height + 12;
        return frame;
    }

    private static Rect CreateFixedFrame(double x, double width, ref double y, double height)
    {
        var frame = new Rect(x, y, width, height);
        y += height + 16;
        return frame;
    }

    private static Rect CreateCodeBlockFrame(CodeBlockElement code, ThemeDefinition theme, double x, double width, ref double y)
    {
        var fontSize = theme.Code.FontSize;
        var lineHeight = theme.Code.LineHeight ?? 1.45;
        var estimatedHeight = EstimateTextHeight(code.Code, fontSize, width, lineHeight);
        var height = Math.Clamp(estimatedHeight + 18, 40, 400); // +18 for code block padding/borders
        var frame = new Rect(x, y, width, height);
        y += height + 16;
        return frame;
    }

    private static double EstimateTextHeight(string text, double fontSize, double width, double lineSpacing)
    {
        var safeText = string.IsNullOrWhiteSpace(text) ? " " : text;
        var approxCharsPerLine = Math.Max(8, (int)(width / Math.Max(8, fontSize * 0.55)));
        var lineCount = safeText
            .Split('\n', StringSplitOptions.None)
            .Sum(line => Math.Max(1, (int)Math.Ceiling((double)Math.Max(1, line.Length) / approxCharsPerLine)));

        return lineCount * fontSize * lineSpacing + 6;
    }
}

public sealed record LayoutOptions(double SlideWidth, double SlideHeight)
{
    public static LayoutOptions Default => new(960, 540);
}

public sealed record LayoutPlan(double SlideWidth, double SlideHeight)
{
    public List<PlacedElement> Elements { get; } = [];
}

public sealed record PlacedElement(ISlideElement Element, Rect Frame);

public sealed record Rect(double X, double Y, double Width, double Height);
