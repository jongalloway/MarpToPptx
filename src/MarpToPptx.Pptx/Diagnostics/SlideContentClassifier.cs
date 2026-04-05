using MarpToPptx.Core.Models;

namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Classifies a slide's content into a <see cref="SlideContentKind"/> using heuristic rules
/// based on the slide's element structure.
/// </summary>
public static class SlideContentClassifier
{
    private const int MaxTaglineLength = 60;
    private const int MaxBoldSpanLength = 10;
    private const int MaxBigNumberParagraphLength = 20;
    private const int MaxNumericValueLength = 10;
    private const int DenseContentThreshold = 6;
    /// <summary>
    /// Classifies the content of <paramref name="slide"/> and returns the most appropriate
    /// <see cref="SlideContentKind"/> along with an optional short explanation.
    /// </summary>
    /// <param name="slide">The slide to classify.</param>
    /// <param name="isFirst"><see langword="true"/> when this is the first slide in the deck.</param>
    /// <param name="isLast"><see langword="true"/> when this is the last slide in the deck.</param>
    /// <returns>A tuple of (<see cref="SlideContentKind"/>, reason hint string or <see langword="null"/>).</returns>
    public static (SlideContentKind Kind, string? Reason) Classify(Slide slide, bool isFirst, bool isLast)
    {
        var elements = slide.Elements;

        var headings = elements.OfType<HeadingElement>().ToList();
        var paragraphs = elements.OfType<ParagraphElement>().ToList();
        var bulletLists = elements.OfType<BulletListElement>().ToList();
        var images = elements.OfType<ImageElement>().ToList();
        var tables = elements.OfType<TableElement>().ToList();
        var blockquotes = elements.OfType<BlockquoteElement>().ToList();

        // Body elements = everything that is not an H1 heading.
        var bodyElements = elements.Where(e => e is not HeadingElement h || h.Level > 1).ToList();

        // ── Title: first slide with H1 and at most one short subtitle paragraph ──────
        if (isFirst && headings.Any(h => h.Level == 1))
        {
            var nonH1Elements = bodyElements
                .Where(e => e is not HeadingElement)
                .ToList();

            if (nonH1Elements.Count <= 1 && blockquotes.Count == 0 && bulletLists.Count == 0 && tables.Count == 0 && images.Count == 0)
            {
                return (SlideContentKind.Title, null);
            }
        }

        // ── Conclusion: last slide, short/empty body ─────────────────────────────────
        if (isLast && headings.Any(h => h.Level == 1))
        {
            var nonH1Elements = bodyElements.Where(e => e is not HeadingElement).ToList();
            if (nonH1Elements.Count <= 1 && blockquotes.Count == 0 && bulletLists.Count == 0 && tables.Count == 0 && images.Count == 0)
            {
                return (SlideContentKind.Conclusion, null);
            }
        }

        // ── Quote: has at least one blockquote ───────────────────────────────────────
        if (blockquotes.Count > 0)
        {
            return (SlideContentKind.Quote, "has blockquote");
        }

        // ── ImageFocused: has image(s) as primary non-heading content ────────────────
        if (images.Count > 0)
        {
            var nonImageBodyElements = bodyElements.Where(e => e is not ImageElement).ToList();
            if (images.Count >= nonImageBodyElements.Count)
            {
                return (SlideContentKind.ImageFocused, "has image");
            }
        }

        // ── Big Number: short body with a single bold/all-numeric span ───────────────
        if (IsBigNumber(paragraphs, bulletLists))
        {
            return (SlideContentKind.BigNumber, "short body, bold number");
        }

        // ── SectionHeader: H1 only (no meaningful body content) ──────────────────────
        var meaningfulBody = bodyElements.Where(e => e is not HeadingElement).ToList();
        if (meaningfulBody.Count == 0 || (meaningfulBody.Count == 1 && paragraphs.Count == 1 && IsShortTagline(paragraphs[0])))
        {
            return (SlideContentKind.SectionHeader, null);
        }

        // ── Table: has a table ───────────────────────────────────────────────────────
        if (tables.Count > 0)
        {
            return (SlideContentKind.Content, "has table");
        }

        // ── Agenda: ordered list with 2–5 items ──────────────────────────────────────
        var orderedList = bulletLists.FirstOrDefault(l => l.Ordered);
        if (orderedList is not null && orderedList.Items.Count is >= 2 and <= 5 && bulletLists.Count == 1 && paragraphs.Count == 0)
        {
            return (SlideContentKind.Agenda, null);
        }

        // ── Dense content: more than DenseContentThreshold body elements ─────────────────
        if (CountBodyUnits(bodyElements) > DenseContentThreshold)
        {
            return (SlideContentKind.WideContent, "dense text");
        }

        // ── Statement: short unordered bullet list (≤ 4 items) ───────────────────────
        var unorderedList = bulletLists.FirstOrDefault(l => !l.Ordered);
        if (unorderedList is not null && unorderedList.Items.Count <= 4 && bulletLists.Count == 1 && paragraphs.Count == 0)
        {
            return (SlideContentKind.Statement, null);
        }

        // ── Default: standard content ─────────────────────────────────────────────────
        return (SlideContentKind.Content, null);
    }

    // ── Private helpers ──────────────────────────────────────────────────────────────

    private static bool IsShortTagline(ParagraphElement para)
        => para.Text.Length <= MaxTaglineLength && !para.Text.Contains('\n');

    private static bool IsBigNumber(IReadOnlyList<ParagraphElement> paragraphs, IReadOnlyList<BulletListElement> bulletLists)
    {
        if (bulletLists.Count > 0)
        {
            return false;
        }

        if (paragraphs.Count is < 1 or > 2)
        {
            return false;
        }

        // The first paragraph must contain a short bold or numeric token.
        var firstPara = paragraphs[0];
        var spans = firstPara.Spans;
        if (spans.Count == 0)
        {
            return false;
        }

        var hasBoldOrNumericSpan = spans.Any(s =>
            (s.Bold && s.Text.Trim().Length <= MaxBoldSpanLength && s.Text.Trim().Any(char.IsDigit)) ||
            IsNumericValue(s.Text.Trim()));

        return hasBoldOrNumericSpan && firstPara.Text.Trim().Length <= MaxBigNumberParagraphLength;
    }

    private static bool IsNumericValue(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        // Accept values like "98%", "$1.2M", "42", "3x", "100K", "#1"
        return text.Length <= MaxNumericValueLength &&
               text.Any(char.IsDigit) &&
               text.All(c => char.IsDigit(c) || c is '%' or '$' or '.' or ',' or 'x' or 'X' or 'K' or 'M' or 'B' or '#' or '+' or '-');
    }

    /// <summary>
    /// Counts the number of distinct body units (lines, bullet items, images, etc.)
    /// for density estimation.
    /// </summary>
    private static int CountBodyUnits(IReadOnlyList<ISlideElement> bodyElements)
    {
        var count = 0;
        foreach (var element in bodyElements)
        {
            count += element switch
            {
                BulletListElement list => list.Items.Count,
                ParagraphElement para => Math.Max(1, para.Text.Split('\n').Length),
                _ => 1,
            };
        }

        return count;
    }
}
