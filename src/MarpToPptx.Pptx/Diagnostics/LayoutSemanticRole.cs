namespace MarpToPptx.Pptx.Diagnostics;

/// <summary>
/// Semantic role of a slide layout inferred from its OpenXML type and placeholder structure.
/// </summary>
public enum LayoutSemanticRole
{
    /// <summary>Standard content layout with a title and a body text or object area.</summary>
    Content,

    /// <summary>Title or cover slide, typically the first slide in a presentation.</summary>
    Title,

    /// <summary>Section divider or header slide used to introduce a new section.</summary>
    SectionHeader,

    /// <summary>Blank layout with no placeholder scaffolding.</summary>
    Blank,

    /// <summary>Side-by-side comparison or multi-column content layout.</summary>
    Comparison,

    /// <summary>Image or picture placeholder with an accompanying caption text area.</summary>
    PictureCaption,

    /// <summary>Layout that shows only a title heading with no body content area.</summary>
    TitleOnly,

    /// <summary>Custom or unrecognised layout type.</summary>
    Other,
}
