using MarpToPptx.Core.Models;
using System.Security.Cryptography;
using System.Text;

namespace MarpToPptx.Core.Authoring;

public sealed record ComputedSlideIdentity(string SlideId, string Title, string Hash, string SourceSlide);

public static class SlideIdentityGenerator
{
    public static IReadOnlyList<ComputedSlideIdentity> BuildSlideIdentities(SlideDeck deck)
    {
        var seenSlideIds = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var identities = new ComputedSlideIdentity[deck.Slides.Count];

        for (var i = 0; i < deck.Slides.Count; i++)
        {
            var slide = deck.Slides[i];
            var title = GetSlideTitle(slide);
            var baseSlideId = ComputeBaseSlideId(slide, title);
            if (!seenSlideIds.TryGetValue(baseSlideId, out var seenCount))
            {
                seenSlideIds[baseSlideId] = 1;
            }
            else
            {
                seenCount++;
                seenSlideIds[baseSlideId] = seenCount;
                baseSlideId = $"{baseSlideId}-{seenCount}";
            }

            identities[i] = new ComputedSlideIdentity(
                baseSlideId,
                title,
                ComputeSlideContentHash(slide),
                $"{Path.GetFileName(deck.SourcePath) ?? string.Empty}#slide-{i + 1}");
        }

        return identities;
    }

    public static string ComputeDeckId(SlideDeck deck)
    {
        if (deck.FrontMatter.TryGetValue("deckid", out var explicitDeckId) ||
            deck.FrontMatter.TryGetValue("deck-id", out explicitDeckId))
        {
            var normalizedExplicitDeckId = NormalizeIdentifier(explicitDeckId);
            if (!string.IsNullOrWhiteSpace(normalizedExplicitDeckId))
            {
                return normalizedExplicitDeckId;
            }
        }

        var sourceKey = string.IsNullOrWhiteSpace(deck.SourcePath)
            ? GetPresentationTitle(deck)
            : deck.SourcePath.Replace('\\', '/').ToLowerInvariant();
        return "deck-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sourceKey)))[..12].ToLowerInvariant();
    }

    public static string GetPresentationTitle(SlideDeck deck)
        => deck.Slides
            .Select(GetSlideTitle)
            .FirstOrDefault(static title => !string.IsNullOrWhiteSpace(title))
            ?? (string.IsNullOrWhiteSpace(deck.SourcePath) ? "PowerPoint Presentation" : Path.GetFileNameWithoutExtension(deck.SourcePath));

    public static string GetSlideTitle(Slide slide)
        => slide.Elements.OfType<HeadingElement>().FirstOrDefault()?.Text?.Trim()
            ?? "PowerPoint Presentation";

    public static string ComputeBaseSlideId(Slide slide, string title)
    {
        var explicitSlideId = NormalizeIdentifier(slide.Style.SlideId);
        if (!string.IsNullOrWhiteSpace(explicitSlideId))
        {
            return explicitSlideId;
        }

        var normalizedTitle = NormalizeIdentifier(title);
        if (!string.IsNullOrWhiteSpace(normalizedTitle))
        {
            return normalizedTitle;
        }

        return "slide-" + ComputeSlideContentHash(slide)[7..19];
    }

    public static string NormalizeIdentifier(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length);
        var previousWasSeparator = false;
        foreach (var ch in value.Trim().ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(ch))
            {
                builder.Append(ch);
                previousWasSeparator = false;
            }
            else if (!previousWasSeparator)
            {
                builder.Append('-');
                previousWasSeparator = true;
            }
        }

        return builder.ToString().Trim('-');
    }

    public static string ComputeSlideContentHash(Slide slideModel)
    {
        var sb = new StringBuilder();

        foreach (var element in slideModel.Elements)
        {
            AppendElementHashContent(sb, element);
            sb.Append('\x00');
        }

        sb.Append("notes\x01").Append(slideModel.Notes ?? string.Empty);
        sb.Append("\x01bg\x01").Append(slideModel.Style.BackgroundColor ?? string.Empty);
        sb.Append("\x01color\x01").Append(slideModel.Style.Color ?? string.Empty);
        sb.Append("\x01bgImg\x01").Append(slideModel.Style.BackgroundImage ?? string.Empty);
        sb.Append("\x01bgSize\x01").Append(slideModel.Style.BackgroundSize ?? string.Empty);
        sb.Append("\x01bgPos\x01").Append(slideModel.Style.BackgroundPosition ?? string.Empty);
        sb.Append("\x01layout\x01").Append(slideModel.Style.Layout ?? string.Empty);
        sb.Append("\x01class\x01").Append(slideModel.Style.ClassName ?? string.Empty);
        sb.Append("\x01theme\x01").Append(slideModel.Style.ThemeName ?? string.Empty);
        sb.Append("\x01header\x01").Append(slideModel.Style.Header ?? string.Empty);
        sb.Append("\x01footer\x01").Append(slideModel.Style.Footer ?? string.Empty);
        sb.Append("\x01paginate\x01").Append(slideModel.Style.Paginate?.ToString() ?? string.Empty);
        if (slideModel.Style.Transition is { } t)
        {
            // \x01 separates this field from others; \x01 again separates the sub-fields
            // (type, direction, durationMs) within the transition value.
            sb.Append("\x01transition\x01").Append(t.Type).Append('\x01').Append(t.Direction ?? string.Empty).Append('\x01').Append(t.DurationMs?.ToString() ?? string.Empty);
        }
        if (slideModel.Style.FontSize is { } fs)
        {
            sb.Append("\x01fontSize\x01").Append(fs);
        }
        if (slideModel.Style.SmartArtHint is { } smartArt)
        {
            sb.Append("\x01smartArt\x01").Append(smartArt);
        }
        if (slideModel.Style.SplitBackgroundLeft is { } splitLeft)
        {
            sb.Append("\x01splitBgLeft\x01").Append(splitLeft);
        }
        if (slideModel.Style.SplitBackgroundRight is { } splitRight)
        {
            sb.Append("\x01splitBgRight\x01").Append(splitRight);
        }

        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString()));
        return "sha256:" + Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static void AppendElementHashContent(StringBuilder sb, ISlideElement element)
    {
        switch (element)
        {
            case HeadingElement heading:
                sb.Append('H').Append(heading.Level).Append(':');
                foreach (var span in heading.Spans) { AppendSpanHashContent(sb, span); }
                break;
            case ParagraphElement paragraph:
                sb.Append("P:");
                foreach (var span in paragraph.Spans) { AppendSpanHashContent(sb, span); }
                break;
            case BulletListElement bullets:
                sb.Append(bullets.Ordered ? "OL:" : "UL:");
                foreach (var item in bullets.Items)
                {
                    sb.Append(item.Depth).Append(':');
                    foreach (var span in item.Spans) { AppendSpanHashContent(sb, span); }
                    sb.Append('\n');
                }
                break;
            case ImageElement image:
                sb.Append("IMG:").Append(image.Source).Append(':').Append(image.AltText);
                break;
            case VideoElement video:
                sb.Append("VID:").Append(video.Source);
                break;
            case AudioElement audio:
                sb.Append("AUD:").Append(audio.Source);
                break;
            case CodeBlockElement code:
                sb.Append("CODE:").Append(code.Language).Append(':').Append(code.Code);
                break;
            case MermaidDiagramElement mermaid:
                sb.Append("MERMAID:").Append(mermaid.Source);
                break;
            case DiagramElement diagram:
                sb.Append("DIAGRAM:").Append(diagram.Source);
                break;
            case TableElement table:
                sb.Append("TABLE:");
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var span in cell) { AppendSpanHashContent(sb, span); }
                        sb.Append('\t');
                    }
                    sb.Append('\n');
                }
                break;
        }
    }

    private static void AppendSpanHashContent(StringBuilder sb, InlineSpan span)
    {
        sb.Append(span.Text);
        if (span.Bold) { sb.Append(":B"); }
        if (span.Italic) { sb.Append(":I"); }
        if (span.Code) { sb.Append(":C"); }
        if (span.Strikethrough) { sb.Append(":S"); }
    }
}