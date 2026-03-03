namespace MarpToPptx.Core.Parsing;

public static class FrontMatterParser
{
    public static (Dictionary<string, string> Values, string Body) Parse(string markdown)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!markdown.StartsWith("---\n", StringComparison.Ordinal) && !markdown.StartsWith("---\r\n", StringComparison.Ordinal))
        {
            return (values, markdown);
        }

        var normalized = markdown.Replace("\r\n", "\n");
        var lines = normalized.Split('\n');
        var closingIndex = Array.FindIndex(lines, 1, line => line.Trim() == "---");
        if (closingIndex <= 0)
        {
            return (values, markdown);
        }

        for (var i = 1; i < closingIndex; i++)
        {
            var line = lines[i];
            var separator = line.IndexOf(':');
            if (separator <= 0)
            {
                continue;
            }

            var key = line[..separator].Trim();
            var value = line[(separator + 1)..].Trim();
            if (key.Length > 0)
            {
                values[key] = value;
            }
        }

        var body = string.Join('\n', lines[(closingIndex + 1)..]);
        return (values, body);
    }
}
