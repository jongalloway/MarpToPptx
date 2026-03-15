namespace MarpToPptx.Pptx.Rendering;

/// <summary>
/// Per-slide metadata written into a slide's <c>p:extLst</c> by MarpToPptx.
/// Used to identify and reconcile MarpToPptx-managed slides during re-entrant update mode.
/// </summary>
internal sealed record SlideMetadata(
    /// <summary>Stable per-slide identity GUID, deterministic from the deck source path and slide ordinal.</summary>
    string Guid,
    /// <summary>SHA-256 content hash of the Marp slide model at the time of last render.</summary>
    string Hash,
    /// <summary>Human-readable source reference, e.g. <c>deck.md#slide-2</c>.</summary>
    string SourceSlide);
