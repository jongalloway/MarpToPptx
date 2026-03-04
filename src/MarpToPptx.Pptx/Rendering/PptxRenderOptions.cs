namespace MarpToPptx.Pptx.Rendering;

public sealed class PptxRenderOptions
{
    public string? TemplatePath { get; init; }

    public string? SourceDirectory { get; init; }

    /// <summary>
    /// When <c>true</c> (the default), HTTP and HTTPS image URLs are fetched at render time.
    /// Set to <c>false</c> to treat all remote references as missing assets.
    /// </summary>
    public bool AllowRemoteAssets { get; init; } = true;

    /// <summary>
    /// Optional <see cref="HttpMessageHandler"/> used when fetching remote assets.
    /// Primarily intended for unit-testing; when <c>null</c> the renderer uses a
    /// shared <see cref="HttpClient"/> with a 30-second timeout.
    /// </summary>
    public HttpMessageHandler? RemoteAssetHandler { get; init; }
}
