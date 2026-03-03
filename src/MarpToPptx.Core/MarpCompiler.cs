using MarpToPptx.Core.Models;
using MarpToPptx.Core.Parsing;

namespace MarpToPptx.Core;

public sealed class MarpCompiler
{
    private readonly MarpMarkdownParser _parser = new();

    public SlideDeck Compile(string markdown, string? sourcePath = null, string? themeCss = null)
        => _parser.Parse(markdown, sourcePath, themeCss);
}
