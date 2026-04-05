namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Semantic content classification of a slide, inferred from its element structure.
/// Used by <see cref="SlideContentClassifier"/> to drive layout recommendations.
/// </summary>
public enum SlideContentKind
{
    /// <summary>Title or cover slide (typically the first slide with H1 + subtitle only).</summary>
    Title,

    /// <summary>Section divider: H1 with no body content or only a very short tagline.</summary>
    SectionHeader,

    /// <summary>Quote slide: H1 with at least one blockquote element.</summary>
    Quote,

    /// <summary>Big-number slide: H1 with a single short bold/numeric span and optional supporting text.</summary>
    BigNumber,

    /// <summary>Image-focused slide: H1 with one or more images forming the dominant content.</summary>
    ImageFocused,

    /// <summary>Statement slide: H1 with a short bullet list (≤ 4 items) or a very short paragraph.</summary>
    Statement,

    /// <summary>Agenda slide: H1 with a numbered list of 2–5 structured items.</summary>
    Agenda,

    /// <summary>Dense-content slide: H1 with more than six body elements (lots of text/bullets).</summary>
    WideContent,

    /// <summary>Standard content slide: H1 with a table or general mixed body content.</summary>
    Content,

    /// <summary>Closing or conclusion slide (typically the last slide with a short farewell).</summary>
    Conclusion,
}
