namespace MarpToPptx.Pptx.Rendering;

/// <summary>
/// Downloads and caches remote (HTTP/HTTPS) image assets to temporary local files
/// for the duration of a render operation.
/// </summary>
internal sealed class RemoteAssetResolver : IDisposable
{
    private static readonly Lazy<HttpClient> SharedClientLazy = new(CreateSharedClient);

    private readonly HttpClient _client;
    private readonly bool _ownsClient;
    private readonly Dictionary<string, string> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _tempFiles = [];

    public RemoteAssetResolver(HttpMessageHandler? handler = null)
    {
        if (handler is not null)
        {
            _client = new HttpClient(handler);
            _ownsClient = true;
        }
        else
        {
            _client = SharedClientLazy.Value;
            _ownsClient = false;
        }
    }

    /// <summary>
    /// Downloads the image at <paramref name="url"/> to a temporary file and returns the
    /// local path. Returns <c>null</c> if the download fails, and sets
    /// <paramref name="errorMessage"/> to an actionable description of the failure.
    /// </summary>
    public string? Resolve(string url, out string? errorMessage)
    {
        if (_cache.TryGetValue(url, out var cached))
        {
            errorMessage = null;
            return cached;
        }

        try
        {
            using var response = _client.GetAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
            if (!response.IsSuccessStatusCode)
            {
                errorMessage = $"HTTP {(int)response.StatusCode} {response.ReasonPhrase}";
                return null;
            }

            var contentType = response.Content.Headers.ContentType?.MediaType;
            var extension = GetExtensionFromUrl(url)
                ?? GetExtensionFromContentType(contentType)
                ?? ".bin";

            var tempPath = Path.Combine(Path.GetTempPath(), $"marp2pptx_{Guid.NewGuid():N}{extension}");

            using (var fileStream = File.Create(tempPath))
            using (var contentStream = response.Content.ReadAsStream())
            {
                contentStream.CopyTo(fileStream);
            }

            _tempFiles.Add(tempPath);
            _cache[url] = tempPath;
            errorMessage = null;
            return tempPath;
        }
        catch (HttpRequestException ex)
        {
            errorMessage = ex.Message;
            return null;
        }
        catch (TaskCanceledException)
        {
            errorMessage = "Request timed out";
            return null;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            return null;
        }
    }

    public void Dispose()
    {
        foreach (var path in _tempFiles)
        {
            try
            {
                File.Delete(path);
            }
            catch
            {
                // best-effort cleanup
            }
        }

        _tempFiles.Clear();
        _cache.Clear();

        if (_ownsClient)
        {
            _client.Dispose();
        }
    }

    private static HttpClient CreateSharedClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("MarpToPptx/1.0");
        return client;
    }

    private static string? GetExtensionFromUrl(string url)
    {
        try
        {
            var path = new Uri(url).AbsolutePath;
            var ext = Path.GetExtension(path);
            return string.IsNullOrEmpty(ext) ? null : ext.ToLowerInvariant();
        }
        catch
        {
            return null;
        }
    }

    private static string? GetExtensionFromContentType(string? contentType)
        => contentType?.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/webp" => ".webp",
            "image/tiff" => ".tiff",
            "image/svg+xml" => ".svg",
            _ => null,
        };
}
