namespace MarpToPptx.Pptx.Rendering;

/// <summary>
/// Per-slide metadata written into a slide's <c>p:extLst</c> by MarpToPptx.
/// Used to identify and reconcile MarpToPptx-managed slides during re-entrant update mode.
/// </summary>
internal sealed record SlideMetadata(
    /// <summary>Metadata schema version for future evolution.</summary>
    string SchemaVersion,
    /// <summary>Diagnostic deck identity.</summary>
    string DeckId,
    /// <summary>Stable per-slide identity.</summary>
    string SlideId,
    /// <summary>Human-readable title extracted from Marp slide content.</summary>
    string Title,
    /// <summary>SHA-256 content hash of the Marp slide model at the time of last render.</summary>
    string Hash,
    /// <summary>Human-readable source reference, e.g. <c>deck.md#slide-2</c>.</summary>
    string SourceSlide,
    /// <summary>Version string of the MarpToPptx generator that emitted this metadata.</summary>
    string GeneratorVersion,
    /// <summary>Requested Marp layout reference, if any.</summary>
    string? LayoutRef,
    /// <summary>Template reference used for diagnostics, if any.</summary>
    string? TemplateRef);
