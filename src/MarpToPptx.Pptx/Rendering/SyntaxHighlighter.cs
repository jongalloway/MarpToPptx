using TextMateSharp.Grammars;
using TextMateSharp.Registry;
using TextMateSharp.Themes;

namespace MarpToPptx.Pptx.Rendering;

/// <summary>
/// A single syntax-highlighted run within a code line.
/// <see cref="Color"/> is a 6-digit hex string (no leading #) when a theme color applies,
/// or <c>null</c> to use the default code text color from the current theme.
/// </summary>
public sealed record TokenizedRun(string Text, string? Color);

/// <summary>
/// Provides language-aware syntax highlighting for code blocks using TextMateSharp.
/// Supported language identifiers: csharp, javascript, typescript, json, html, css,
/// xml, powershell, python, sql (and common aliases such as cs, js, ts, py, ps1).
/// Unsupported languages produce a single plain run per line.
/// </summary>
public static class SyntaxHighlighter
{
    // Maps user-facing language identifiers (from fenced code blocks) to the
    // canonical identifiers recognised by TextMateSharp.Grammars.
    private static readonly Dictionary<string, string> LanguageAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["cs"] = "csharp",
        ["js"] = "javascript",
        ["ts"] = "typescript",
        ["py"] = "python",
        ["ps1"] = "powershell",
        ["sh"] = "shellscript",
        ["bash"] = "shellscript",
        ["yml"] = "yaml",
    };

    // Lazy so the registry is only initialised when syntax highlighting is first needed.
    private static readonly Lazy<HighlightState?> _state = new(CreateState);

    /// <summary>Returns <c>true</c> when <paramref name="language"/> has grammar support.</summary>
    public static bool IsSupported(string? language)
    {
        if (string.IsNullOrWhiteSpace(language))
        {
            return false;
        }

        var canonicalId = CanonicalId(language);
        var state = _state.Value;
        if (state is null)
        {
            return false;
        }

        return state.Options.GetScopeByLanguageId(canonicalId) is not null;
    }

    /// <summary>
    /// Tokenises <paramref name="code"/> into lines of coloured runs.
    /// Returns one inner list per source line. Falls back to plain runs when
    /// the language is unsupported or grammar loading fails.
    /// </summary>
    public static IReadOnlyList<IReadOnlyList<TokenizedRun>> Tokenize(string language, string code)
    {
        var state = _state.Value;
        if (state is not null && !string.IsNullOrWhiteSpace(language))
        {
            var canonicalId = CanonicalId(language);
            var scopeName = state.Options.GetScopeByLanguageId(canonicalId);
            if (scopeName is not null)
            {
                var grammar = state.Registry.LoadGrammar(scopeName);
                if (grammar is not null)
                {
                    return TokenizeWithGrammar(code, grammar, state.Theme);
                }
            }
        }

        return FallbackTokenize(code);
    }

    private static IReadOnlyList<IReadOnlyList<TokenizedRun>> TokenizeWithGrammar(string code, IGrammar grammar, Theme theme)
    {
        var lines = SplitLines(code);
        var result = new List<IReadOnlyList<TokenizedRun>>(lines.Length);
        IStateStack? ruleStack = null;

        foreach (var line in lines)
        {
            ITokenizeLineResult lineResult = ruleStack is null
                ? grammar.TokenizeLine(new LineText(line))
                : grammar.TokenizeLine(new LineText(line), ruleStack, TimeSpan.FromSeconds(5));

            ruleStack = lineResult.RuleStack;
            result.Add(BuildRuns(line, lineResult.Tokens, theme));
        }

        return result;
    }

    private static IReadOnlyList<TokenizedRun> BuildRuns(string line, IToken[] tokens, Theme theme)
    {
        var runs = new List<TokenizedRun>(tokens.Length);
        foreach (var token in tokens)
        {
            var end = Math.Min(token.EndIndex, line.Length);
            if (token.StartIndex >= end)
            {
                continue;
            }

            var text = line[token.StartIndex..end];
            if (text.Length == 0)
            {
                continue;
            }

            var matches = theme.Match(token.Scopes);
            string? color = null;
            if (matches is { Count: > 0 })
            {
                var foregroundId = matches[0].foreground;
                if (foregroundId > 0)
                {
                    color = NormalizeHex(theme.GetColor(foregroundId));
                }
            }

            runs.Add(new TokenizedRun(text, color));
        }

        return runs.Count > 0 ? runs : [new TokenizedRun(line, null)];
    }

    private static IReadOnlyList<IReadOnlyList<TokenizedRun>> FallbackTokenize(string code)
    {
        var lines = SplitLines(code);
        var result = new List<IReadOnlyList<TokenizedRun>>(lines.Length);
        foreach (var line in lines)
        {
            result.Add([new TokenizedRun(line, null)]);
        }

        return result;
    }

    private static string[] SplitLines(string code)
        => code.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');

    private static string CanonicalId(string language)
        => LanguageAliases.TryGetValue(language, out var canonical) ? canonical : language;

    /// <summary>
    /// Normalises a hex color string returned by TextMateSharp to the 6-digit
    /// uppercase format expected by OpenXML (no leading #, no alpha channel).
    /// Returns <c>null</c> when the input cannot be parsed.
    /// </summary>
    private static string? NormalizeHex(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var hex = value.Trim().TrimStart('#');

        // Some themes include an alpha byte as the first two digits (AARRGGBB).
        // Strip to the last 6 characters to get RRGGBB.
        return hex.Length >= 6 ? hex[^6..].ToUpperInvariant() : null;
    }

    private static HighlightState? CreateState()
    {
        try
        {
            var options = new RegistryOptions(ThemeName.DarkPlus);
            var registry = new Registry(options);
            registry.SetTheme(options.GetDefaultTheme());
            var theme = registry.GetTheme();
            return new HighlightState(options, registry, theme);
        }
        catch
        {
            return null;
        }
    }

    private sealed record HighlightState(RegistryOptions Options, Registry Registry, Theme Theme);
}
