using MarpToPptx.Core.Models;

namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Matches slide content to template layouts and produces a <see cref="LayoutRecommendationReport"/>
/// that suggests the best layout for each slide in a Marp deck.
/// </summary>
public sealed class LayoutRecommender
{
    /// <summary>
    /// Analyses each slide in <paramref name="deck"/> against the layouts catalogued in
    /// <paramref name="templateReport"/> and returns a recommendation report.
    /// </summary>
    public LayoutRecommendationReport Recommend(SlideDeck deck, TemplateDiagnosticReport templateReport)
    {
        ArgumentNullException.ThrowIfNull(deck);
        ArgumentNullException.ThrowIfNull(templateReport);

        var layouts = templateReport.Layouts;
        var photoLayouts = BuildPhotoLayoutList(layouts);
        var photoIndex = 0;

        // Pre-select named fallback layouts for common roles.
        var defaultContentLayout =
            templateReport.RecommendedDefaultContentLayout ??
            layouts.FirstOrDefault(l => l.HasTitlePlaceholder && l.HasBodyPlaceholder)?.Name ??
            layouts.FirstOrDefault()?.Name ??
            "Title and Content";

        var titleLayout =
            templateReport.RecommendedTitleLayout ??
            defaultContentLayout;

        var sectionLayout =
            templateReport.RecommendedSectionLayout ??
            titleLayout;

        var recommendations = new List<SlideRecommendation>();
        var slideCount = deck.Slides.Count;

        for (var i = 0; i < slideCount; i++)
        {
            var slide = deck.Slides[i];
            var isFirst = i == 0;
            var isLast = i == slideCount - 1;

            // Respect an explicit per-slide layout directive when already set.
            if (!string.IsNullOrWhiteSpace(slide.Style.Layout))
            {
                var explicitTitle = GetSlideTitle(slide);
                recommendations.Add(new SlideRecommendation(
                    i + 1,
                    explicitTitle,
                    SlideContentKind.Content,
                    slide.Style.Layout,
                    Reason: null,
                    IsExplicitLayout: true));
                continue;
            }

            var (kind, reason) = SlideContentClassifier.Classify(slide, isFirst, isLast);
            var layout = PickLayout(kind, layouts, defaultContentLayout, titleLayout, sectionLayout, photoLayouts, ref photoIndex);

            recommendations.Add(new SlideRecommendation(i + 1, GetSlideTitle(slide), kind, layout, reason));
        }

        var suggestedFrontMatter = SuggestFrontMatterLayout(recommendations, defaultContentLayout);
        return new LayoutRecommendationReport(
            deck.SourcePath ?? "deck.md",
            templateReport.TemplatePath,
            recommendations,
            suggestedFrontMatter,
            photoLayouts);
    }

    // ── Layout selection ─────────────────────────────────────────────────────────────

    private static string PickLayout(
        SlideContentKind kind,
        IReadOnlyList<LayoutDiagnostic> layouts,
        string defaultContentLayout,
        string titleLayout,
        string sectionLayout,
        IReadOnlyList<string> photoLayouts,
        ref int photoIndex)
    {
        return kind switch
        {
            SlideContentKind.Title      => titleLayout,
            SlideContentKind.Conclusion => FindConclusionLayout(layouts) ?? titleLayout,
            SlideContentKind.SectionHeader => sectionLayout,
            SlideContentKind.Quote      => FindNamedLayout(layouts, "quote") ?? sectionLayout,
            SlideContentKind.BigNumber  => FindNamedLayout(layouts, "big number", "big-number", "bignumber") ?? defaultContentLayout,
            SlideContentKind.Agenda     => FindNamedLayout(layouts, "agenda") ?? defaultContentLayout,
            SlideContentKind.Statement  => FindNamedLayout(layouts, "statement") ?? sectionLayout,
            SlideContentKind.ImageFocused => PickPhotoLayout(photoLayouts, ref photoIndex) ?? defaultContentLayout,
            SlideContentKind.WideContent => FindWideContentLayout(layouts) ?? defaultContentLayout,
            _                            => defaultContentLayout,
        };
    }

    private static string? FindNamedLayout(IReadOnlyList<LayoutDiagnostic> layouts, params string[] keywords)
    {
        foreach (var keyword in keywords)
        {
            var match = layouts.FirstOrDefault(l =>
                l.Name.Contains(keyword, StringComparison.OrdinalIgnoreCase));
            if (match is not null)
            {
                return match.Name;
            }
        }

        return null;
    }

    private static string? FindConclusionLayout(IReadOnlyList<LayoutDiagnostic> layouts)
        => FindNamedLayout(layouts, "conclusion", "closing", "end", "thank");

    private static string? FindWideContentLayout(IReadOnlyList<LayoutDiagnostic> layouts)
    {
        var contentLayouts = layouts
            .Where(l => l.SemanticRole == LayoutSemanticRole.Content && l.HasTitlePlaceholder && l.HasBodyPlaceholder)
            .ToList();

        // 1. Prefer a content layout whose name ends with a digit (variant numbering convention).
        var digitSuffixedCandidate = contentLayouts
            .FirstOrDefault(l => char.IsDigit(l.Name.LastOrDefault()));
        if (digitSuffixedCandidate is not null)
        {
            return digitSuffixedCandidate.Name;
        }

        // 2. Fall back to a content layout explicitly named as wide/full.
        var keywordCandidate = contentLayouts
            .FirstOrDefault(l =>
                l.Name.Contains("wide", StringComparison.OrdinalIgnoreCase) ||
                l.Name.Contains("full", StringComparison.OrdinalIgnoreCase));

        return keywordCandidate?.Name;
    }

    private static string? PickPhotoLayout(IReadOnlyList<string> photoLayouts, ref int photoIndex)
    {
        if (photoLayouts.Count == 0)
        {
            return null;
        }

        var name = photoLayouts[photoIndex % photoLayouts.Count];
        photoIndex++;
        return name;
    }

    // ── Photo layout discovery ───────────────────────────────────────────────────────

    private static IReadOnlyList<string> BuildPhotoLayoutList(IReadOnlyList<LayoutDiagnostic> layouts)
    {
        var photoLayouts = layouts
            .Where(l => l.SemanticRole == LayoutSemanticRole.PictureCaption || l.HasPicturePlaceholder)
            .Select(l => l.Name)
            .ToList();

        return photoLayouts;
    }

    // ── Front-matter suggestion ──────────────────────────────────────────────────────

    private static string? SuggestFrontMatterLayout(
        IReadOnlyList<SlideRecommendation> recommendations,
        string defaultContentLayout)
    {
        // Pick the most frequently recommended layout across content/statement/wide-content slides.
        var contentKinds = new HashSet<SlideContentKind>
        {
            SlideContentKind.Content,
            SlideContentKind.Statement,
            SlideContentKind.WideContent,
        };

        var grouped = recommendations
            .Where(r => !r.IsExplicitLayout && contentKinds.Contains(r.ContentKind))
            .GroupBy(r => r.RecommendedLayout)
            .OrderByDescending(g => g.Count())
            .FirstOrDefault();

        return grouped?.Key ?? defaultContentLayout;
    }

    // ── Slide title extraction ───────────────────────────────────────────────────────

    private static string? GetSlideTitle(Slide slide)
    {
        var h1 = slide.Elements.OfType<HeadingElement>().FirstOrDefault(h => h.Level == 1);
        if (h1 is not null)
        {
            return h1.Text;
        }

        var h2 = slide.Elements.OfType<HeadingElement>().FirstOrDefault(h => h.Level == 2);
        return h2?.Text;
    }
}
