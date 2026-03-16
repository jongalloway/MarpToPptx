using MarpToPptx.Core.Models;
using MarpToPptx.Core.Parsing;
using System.Text;
using System.Text.RegularExpressions;

namespace MarpToPptx.Core.Authoring;

public sealed record SlideIdRewriteResult(string UpdatedMarkdown, int AddedCount);

public static partial class SlideIdDirectiveWriter
{
    public static SlideIdRewriteResult WriteMissingSlideIds(string markdown, SlideDeck deck)
    {
        var newline = markdown.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
        var (prefix, body) = SplitFrontMatter(markdown);
        var headingDivider = TryParseHeadingDivider(deck.FrontMatter);
        var chunks = SplitSlidesPreservingBody(body, headingDivider);
        if (chunks.Count != deck.Slides.Count)
        {
            return new SlideIdRewriteResult(markdown, 0);
        }

        var computedSlideIds = SlideIdentityGenerator.BuildSlideIdentities(deck)
            .Select(static identity => identity.SlideId)
            .ToArray();
        var updatedChunks = new List<string>(chunks.Count);
        var carryForwardStyle = new SlideStyle();
        var addedCount = 0;

        for (var i = 0; i < chunks.Count; i++)
        {
            var chunk = chunks[i];
            var (effectiveStyle, newCarryForward, _, _) = MarpDirectiveParser.Parse(chunk.Content, carryForwardStyle);
            carryForwardStyle = newCarryForward;

            if (!string.IsNullOrWhiteSpace(effectiveStyle.SlideId))
            {
                updatedChunks.Add(chunk.Content);
                continue;
            }

            updatedChunks.Add(InsertSlideIdDirective(chunk.Content, computedSlideIds[i], newline));
            addedCount++;
        }

        if (addedCount == 0)
        {
            return new SlideIdRewriteResult(markdown, 0);
        }

        var rebuiltBody = RebuildBody(updatedChunks, chunks, newline);
        return new SlideIdRewriteResult(prefix + rebuiltBody, addedCount);
    }

    private static int? TryParseHeadingDivider(Dictionary<string, string> frontMatter)
        => frontMatter.TryGetValue("headingDivider", out var hdValue) && int.TryParse(hdValue, out var hdLevel) && hdLevel is >= 1 and <= 6
            ? hdLevel
            : null;

    private static (string Prefix, string Body) SplitFrontMatter(string markdown)
    {
        var normalized = markdown.Replace("\r\n", "\n");
        if (!normalized.StartsWith("---\n", StringComparison.Ordinal))
        {
            return (string.Empty, normalized);
        }

        var lines = normalized.Split('\n');
        var closingIndex = Array.FindIndex(lines, 1, line => line.Trim() == "---");
        if (closingIndex <= 0)
        {
            return (string.Empty, normalized);
        }

        var prefix = string.Join('\n', lines[..(closingIndex + 1)]) + "\n";
        var body = string.Join('\n', lines[(closingIndex + 1)..]);
        return (prefix, body);
    }

    private static List<RawSlideChunk> SplitSlidesPreservingBody(string body, int? headingDivider)
    {
        var chunks = new List<RawSlideChunk>();
        var buffer = new List<string>();
        var inFence = false;
        string? boundaryBeforeNext = null;

        foreach (var rawLine in body.Replace("\r\n", "\n").Split('\n'))
        {
            var line = rawLine.TrimEnd();
            if (line.StartsWith("```", StringComparison.Ordinal) || line.StartsWith("~~~", StringComparison.Ordinal))
            {
                inFence = !inFence;
            }

            if (!inFence && line.Trim() == "---")
            {
                chunks.Add(new RawSlideChunk(string.Join('\n', buffer), boundaryBeforeNext));
                buffer.Clear();
                boundaryBeforeNext = "separator";
                continue;
            }

            if (!inFence && headingDivider is > 0 and <= 6)
            {
                var headingMatch = HeadingLineRegex().Match(line);
                if (headingMatch.Success && headingMatch.Groups[1].Value.Length <= headingDivider)
                {
                    var pending = string.Join('\n', buffer);
                    if (!string.IsNullOrWhiteSpace(pending))
                    {
                        chunks.Add(new RawSlideChunk(pending, boundaryBeforeNext));
                        buffer.Clear();
                        boundaryBeforeNext = "heading";
                    }
                }
            }

            buffer.Add(rawLine);
        }

        if (buffer.Count > 0 || chunks.Count == 0)
        {
            chunks.Add(new RawSlideChunk(string.Join('\n', buffer), boundaryBeforeNext));
        }

        return chunks.Where(chunk => !string.IsNullOrWhiteSpace(chunk.Content)).ToList();
    }

    private static string InsertSlideIdDirective(string content, string slideId, string newline)
    {
        var normalized = content.Replace("\r\n", "\n");
        var lines = normalized.Split('\n').ToList();
        var insertIndex = 0;
        while (insertIndex < lines.Count && string.IsNullOrWhiteSpace(lines[insertIndex]))
        {
            insertIndex++;
        }

        lines.Insert(insertIndex, $"<!-- slideId: {slideId} -->");
        if (insertIndex + 1 < lines.Count && !string.IsNullOrWhiteSpace(lines[insertIndex + 1]))
        {
            lines.Insert(insertIndex + 1, string.Empty);
        }

        return string.Join(newline, lines);
    }

    private static string RebuildBody(IReadOnlyList<string> updatedChunks, IReadOnlyList<RawSlideChunk> originalChunks, string newline)
    {
        var builder = new StringBuilder();
        for (var i = 0; i < updatedChunks.Count; i++)
        {
            if (i > 0)
            {
                builder.Append(originalChunks[i].BoundaryBefore switch
                {
                    "separator" => $"{newline}---{newline}",
                    _ => newline,
                });
            }

            builder.Append(updatedChunks[i].Replace("\n", newline, StringComparison.Ordinal));
        }

        return builder.ToString();
    }

    [GeneratedRegex(@"^\s{0,3}(#{1,6})\s")]
    private static partial Regex HeadingLineRegex();

    private sealed record RawSlideChunk(string Content, string? BoundaryBefore);
}