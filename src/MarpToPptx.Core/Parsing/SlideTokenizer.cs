using System.Text.RegularExpressions;

namespace MarpToPptx.Core.Parsing;

public static partial class SlideTokenizer
{
    /// <summary>
    /// Splits markdown into slide chunks on <c>---</c> separators and, when
    /// <paramref name="headingDivider"/> is set, on headings at or above that level.
    /// </summary>
    public static IReadOnlyList<string> SplitSlides(string markdown, int? headingDivider = null)
    {
        var slides = new List<string>();
        var buffer = new List<string>();
        var inFence = false;

        foreach (var rawLine in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            var line = rawLine.TrimEnd();
            if (line.StartsWith("```", StringComparison.Ordinal) || line.StartsWith("~~~", StringComparison.Ordinal))
            {
                inFence = !inFence;
            }

            if (!inFence && line.Trim() == "---")
            {
                slides.Add(string.Join('\n', buffer).Trim());
                buffer.Clear();
                continue;
            }

            // headingDivider: split before a heading at or above the specified level.
            if (!inFence && headingDivider is > 0 and <= 6)
            {
                var headingMatch = HeadingLineRegex().Match(line);
                if (headingMatch.Success && headingMatch.Groups[1].Value.Length <= headingDivider)
                {
                    // Flush the current buffer as the previous slide.
                    var pending = string.Join('\n', buffer).Trim();
                    if (!string.IsNullOrWhiteSpace(pending))
                    {
                        slides.Add(pending);
                    }

                    buffer.Clear();
                }
            }

            buffer.Add(rawLine);
        }

        if (buffer.Count > 0)
        {
            slides.Add(string.Join('\n', buffer).Trim());
        }

        return slides.Where(slide => !string.IsNullOrWhiteSpace(slide)).ToArray();
    }

    [GeneratedRegex(@"^(#{1,6})\s")]
    private static partial Regex HeadingLineRegex();
}
