namespace MarpToPptx.Core.Parsing;

public static class SlideTokenizer
{
    public static IReadOnlyList<string> SplitSlides(string markdown)
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

            buffer.Add(rawLine);
        }

        if (buffer.Count > 0)
        {
            slides.Add(string.Join('\n', buffer).Trim());
        }

        return slides.Where(slide => !string.IsNullOrWhiteSpace(slide)).ToArray();
    }
}
