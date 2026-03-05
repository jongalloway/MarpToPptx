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

        var i = 1;
        while (i < closingIndex)
        {
            var line = lines[i];
            i++;

            var separator = line.IndexOf(':');
            if (separator <= 0)
            {
                continue;
            }

            var key = line[..separator].Trim();
            var value = line[(separator + 1)..].Trim();

            // Handle YAML literal block scalar ("|"): collect subsequent indented lines.
            if (value == "|" && key.Length > 0)
            {
                var blockLines = new List<string>();
                var blockIndent = -1;
                while (i < closingIndex)
                {
                    var nextLine = lines[i];
                    if (string.IsNullOrWhiteSpace(nextLine))
                    {
                        blockLines.Add(string.Empty);
                        i++;
                        continue;
                    }

                    var indent = nextLine.Length - nextLine.TrimStart().Length;
                    if (blockIndent < 0)
                    {
                        blockIndent = indent;
                    }
                    else if (indent < blockIndent)
                    {
                        break;
                    }

                    blockLines.Add(blockIndent > 0 ? nextLine[blockIndent..] : nextLine);
                    i++;
                }

                // Strip trailing empty lines per YAML spec.
                while (blockLines.Count > 0 && string.IsNullOrWhiteSpace(blockLines[^1]))
                {
                    blockLines.RemoveAt(blockLines.Count - 1);
                }

                value = string.Join('\n', blockLines);
            }

            if (key.Length > 0)
            {
                values[key] = value;
            }
        }

        var body = string.Join('\n', lines[(closingIndex + 1)..]);
        return (values, body);
    }
}
