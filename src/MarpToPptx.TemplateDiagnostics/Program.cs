using MarpToPptx.Core;
using MarpToPptx.Pptx.Diagnostics;

if (args.Length == 0 || IsHelpFlag(args[0]))
{
    PrintUsage();
    return 0;
}

var subcommand = args[0].ToLowerInvariant();

if (subcommand == "diagnose")
{
    return await RunDiagnoseAsync(args[1..]);
}

if (subcommand == "recommend")
{
    return await RunRecommendAsync(args[1..]);
}

if (subcommand == "doctor")
{
    return await RunDoctorAsync(args[1..]);
}

// Backward-compatible: no subcommand → treat the first argument as a template path and run diagnose.
if (!args[0].StartsWith('-'))
{
    return await RunDiagnoseAsync(args);
}

Console.Error.WriteLine($"Unknown subcommand '{args[0]}'. Run with --help for usage.");
return 1;

// ── Subcommand: diagnose ────────────────────────────────────────────────────

static Task<int> RunDiagnoseAsync(string[] subArgs)
{
    if (subArgs.Length == 0 || IsHelpFlag(subArgs[0]))
    {
        Console.WriteLine("Usage: template-diagnostics diagnose <path-to-template.pptx>");
        Console.WriteLine();
        Console.WriteLine("Analyzes a .pptx template and reports its layout structure,");
        Console.WriteLine("recommended Markdown directives, and potential warnings.");
        return Task.FromResult(0);
    }

    var templatePath = ResolveTemplatePath(subArgs[0]);
    if (templatePath is null)
    {
        return Task.FromResult(1);
    }

    try
    {
        var diagnoser = new TemplateDiagnoser();
        var report = diagnoser.Diagnose(templatePath);
        PrintDiagnoseReport(report);
        return Task.FromResult(0);
    }
    catch (Exception exception)
    {
        Console.Error.WriteLine($"Template diagnostics failed for '{templatePath}': {exception.Message}");
        return Task.FromResult(1);
    }
}

// ── Subcommand: recommend ───────────────────────────────────────────────────

static async Task<int> RunRecommendAsync(string[] subArgs)
{
    if (subArgs.Length == 0 || IsHelpFlag(subArgs[0]))
    {
        Console.WriteLine("Usage: template-diagnostics recommend <deck.md> --template <template.pptx> [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --template <path>  Path to the .pptx template to match layouts against (required).");
        Console.WriteLine("  --json             Emit the report as JSON to stdout.");
        Console.WriteLine("  --verbose          Include a reason for each recommendation.");
        Console.WriteLine("  --patch            Write _layout: directives back into the Markdown source.");
        Console.WriteLine();
        Console.WriteLine("Analyzes a Marp Markdown deck against a PPTX template and suggests the");
        Console.WriteLine("best layout for each slide based on its content structure.");
        return 0;
    }

    string? deckPath = null;
    string? templatePath = null;
    var jsonMode = false;
    var verbose = false;
    var patch = false;

    for (var i = 0; i < subArgs.Length; i++)
    {
        var arg = subArgs[i];

        switch (arg.ToLowerInvariant())
        {
            case "--template":
                if (i + 1 >= subArgs.Length)
                {
                    Console.Error.WriteLine("--template requires a path argument.");
                    return 1;
                }

                templatePath = subArgs[++i];
                break;

            case "--json":
                jsonMode = true;
                break;

            case "--verbose":
                verbose = true;
                break;

            case "--patch":
                patch = true;
                break;

            default:
                if (arg.StartsWith('-'))
                {
                    Console.Error.WriteLine($"Unknown option '{arg}'. Run 'template-diagnostics recommend --help' for usage.");
                    return 1;
                }

                deckPath ??= arg;
                break;
        }
    }

    if (deckPath is null)
    {
        Console.Error.WriteLine("A deck path is required.");
        return 1;
    }

    if (templatePath is null)
    {
        Console.Error.WriteLine("A template path is required (--template <path>).");
        return 1;
    }

    var resolvedDeckPath = Path.GetFullPath(deckPath);
    if (!File.Exists(resolvedDeckPath))
    {
        Console.Error.WriteLine($"Deck file was not found: {resolvedDeckPath}");
        return 1;
    }

    var resolvedTemplatePath = ResolveTemplatePath(templatePath);
    if (resolvedTemplatePath is null)
    {
        return 1;
    }

    try
    {
        var markdown = await File.ReadAllTextAsync(resolvedDeckPath);
        var compiler = new MarpCompiler();
        var deck = compiler.Compile(markdown, resolvedDeckPath);

        var diagnoser = new TemplateDiagnoser();
        var templateReport = diagnoser.Diagnose(resolvedTemplatePath);

        var recommender = new LayoutRecommender();
        var report = recommender.Recommend(deck, templateReport);

        if (jsonMode)
        {
            PrintRecommendReportJson(report);
        }
        else
        {
            PrintRecommendReport(report, verbose);
        }

        if (patch)
        {
            ApplyLayoutPatches(resolvedDeckPath, report);
        }

        return 0;
    }
    catch (Exception exception)
    {
        Console.Error.WriteLine($"Recommend failed: {exception.Message}");
        return 1;
    }
}

// ── Subcommand: doctor ──────────────────────────────────────────────────────

static Task<int> RunDoctorAsync(string[] subArgs)
{
    if (subArgs.Length == 0 || IsHelpFlag(subArgs[0]))
    {
        Console.WriteLine("Usage: template-diagnostics doctor <path-to-template.pptx> [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --dry-run                    Report issues without writing any file (default).");
        Console.WriteLine("  --write-fixed-template <path>  Write a repaired copy of the template to <path>.");
        Console.WriteLine("  --json                       Emit the report as JSON to stdout.");
        Console.WriteLine();
        Console.WriteLine("Analyzes a .pptx template for structural issues that degrade MarpToPptx output");
        Console.WriteLine("and optionally writes a repaired copy with safe fixups applied.");
        return Task.FromResult(0);
    }

    string? templatePath = null;
    string? outputPath = null;
    var jsonMode = false;

    for (var i = 0; i < subArgs.Length; i++)
    {
        var arg = subArgs[i];

        switch (arg.ToLowerInvariant())
        {
            case "--dry-run":
                // outputPath stays null → dry run.
                break;

            case "--write-fixed-template":
                if (i + 1 >= subArgs.Length)
                {
                    Console.Error.WriteLine("--write-fixed-template requires a path argument.");
                    return Task.FromResult(1);
                }

                outputPath = Path.GetFullPath(subArgs[++i]);
                break;

            case "--json":
                jsonMode = true;
                break;

            default:
                if (arg.StartsWith('-'))
                {
                    Console.Error.WriteLine($"Unknown option '{arg}'. Run 'template-diagnostics doctor --help' for usage.");
                    return Task.FromResult(1);
                }

                templatePath ??= arg;
                break;
        }
    }

    if (templatePath is null)
    {
        Console.Error.WriteLine("A template path is required.");
        return Task.FromResult(1);
    }

    var resolvedTemplatePath = ResolveTemplatePath(templatePath);
    if (resolvedTemplatePath is null)
    {
        return Task.FromResult(1);
    }

    try
    {
        var doctor = new TemplateDoctor();
        var report = doctor.Run(resolvedTemplatePath, outputPath);

        if (jsonMode)
        {
            PrintDoctorReportJson(report);
        }
        else
        {
            PrintDoctorReport(report);
        }

        return Task.FromResult(0);
    }
    catch (Exception exception)
    {
        Console.Error.WriteLine($"Template doctor failed for '{resolvedTemplatePath}': {exception.Message}");
        return Task.FromResult(1);
    }
}

// ── Shared helpers ──────────────────────────────────────────────────────────

static bool IsHelpFlag(string arg)
    => arg is "-h" or "--help" or "-?" or "/?";

static void PrintUsage()
{
    Console.WriteLine("Usage: template-diagnostics <subcommand> [options]");
    Console.WriteLine();
    Console.WriteLine("Subcommands:");
    Console.WriteLine("  diagnose  <template.pptx>                Analyze template layout structure and report recommendations.");
    Console.WriteLine("  doctor    <template.pptx>                Inspect and optionally repair structural issues.");
    Console.WriteLine("  recommend <deck.md> --template <t.pptx>  Suggest the best layout for each slide in a deck.");
    Console.WriteLine();
    Console.WriteLine("Run 'template-diagnostics <subcommand> --help' for subcommand-specific options.");
}

static string? ResolveTemplatePath(string path)
{
    var fullPath = Path.GetFullPath(path);
    if (!File.Exists(fullPath))
    {
        Console.Error.WriteLine($"Template file was not found: {fullPath}");
        return null;
    }

    return fullPath;
}

// ── Print: recommend (human-readable) ───────────────────────────────────────

static void PrintRecommendReport(LayoutRecommendationReport report, bool verbose)
{
    var deckName = Path.GetFileName(report.DeckPath);
    var templateName = Path.GetFileName(report.TemplatePath);
    var title = $"Layout Recommendations for {deckName} + {templateName}";
    Console.WriteLine(title);
    Console.WriteLine(new string('═', title.Length));
    Console.WriteLine();

    if (report.Recommendations.Count == 0)
    {
        Console.WriteLine("No slides found in the deck.");
        return;
    }

    const int MaxTitleDisplayWidth = 40;
    var slideNumWidth = report.Recommendations.Count.ToString().Length;
    var titleWidth = Math.Min(MaxTitleDisplayWidth, report.Recommendations.Max(r => r.SlideTitle?.Length ?? 0));
    var layoutWidth = report.Recommendations.Max(r => r.RecommendedLayout.Length);

    foreach (var rec in report.Recommendations)
    {
        var slideNum = $"Slide {rec.SlideNumber.ToString().PadLeft(slideNumWidth)}";
        var slideTitle = (rec.SlideTitle is not null
            ? $"\"{TruncateTitle(rec.SlideTitle, titleWidth)}\""
            : "(no title)").PadRight(titleWidth + 2);
        var layout = rec.RecommendedLayout.PadRight(layoutWidth);

        if (verbose && rec.Reason is not null)
        {
            Console.WriteLine($"{slideNum}  {slideTitle}  → {layout}  ({rec.Reason})");
        }
        else
        {
            Console.WriteLine($"{slideNum}  {slideTitle}  → {layout}");
        }
    }

    Console.WriteLine();

    if (report.SuggestedFrontMatterLayout is { } frontMatterLayout)
    {
        Console.WriteLine($"Suggested front-matter:  layout: {frontMatterLayout}");
    }

    if (report.PhotoLayoutRotation.Count > 1)
    {
        var photoNames = string.Join(", ", report.PhotoLayoutRotation.Select(n => $"\"{n}\""));
        Console.WriteLine($"Photo layouts will rotate: {photoNames}");
    }
}

static string TruncateTitle(string title, int maxWidth)
{
    if (title.Length <= maxWidth)
    {
        return title;
    }

    return title[..(maxWidth - 1)] + "…";
}

// ── Print: recommend (JSON) ──────────────────────────────────────────────────

static void PrintRecommendReportJson(LayoutRecommendationReport report)
{
    Console.WriteLine("{");
    Console.WriteLine($"  \"deckPath\": {JsonString(report.DeckPath)},");
    Console.WriteLine($"  \"templatePath\": {JsonString(report.TemplatePath)},");
    Console.WriteLine($"  \"suggestedFrontMatterLayout\": {(report.SuggestedFrontMatterLayout is not null ? JsonString(report.SuggestedFrontMatterLayout) : "null")},");

    Console.WriteLine($"  \"photoLayoutRotation\": [");
    for (var i = 0; i < report.PhotoLayoutRotation.Count; i++)
    {
        var comma = i < report.PhotoLayoutRotation.Count - 1 ? "," : "";
        Console.WriteLine($"    {JsonString(report.PhotoLayoutRotation[i])}{comma}");
    }

    Console.WriteLine("  ],");
    Console.WriteLine($"  \"recommendations\": [");

    for (var i = 0; i < report.Recommendations.Count; i++)
    {
        var rec = report.Recommendations[i];
        var comma = i < report.Recommendations.Count - 1 ? "," : "";
        Console.WriteLine("    {");
        Console.WriteLine($"      \"slideNumber\": {rec.SlideNumber},");
        Console.WriteLine($"      \"slideTitle\": {(rec.SlideTitle is not null ? JsonString(rec.SlideTitle) : "null")},");
        Console.WriteLine($"      \"contentKind\": {JsonString(rec.ContentKind.ToString())},");
        Console.WriteLine($"      \"recommendedLayout\": {JsonString(rec.RecommendedLayout)},");
        Console.WriteLine($"      \"reason\": {(rec.Reason is not null ? JsonString(rec.Reason) : "null")}");
        Console.WriteLine($"    }}{comma}");
    }

    Console.WriteLine("  ]");
    Console.WriteLine("}");
}

// ── Patch: write _layout: directives into the Markdown source ───────────────

static void ApplyLayoutPatches(string deckPath, LayoutRecommendationReport report)
{
    var lines = File.ReadAllLines(deckPath).ToList();
    var slideBoundaries = FindSlideBoundaries(lines);
    var patchedCount = 0;

    // Walk in reverse so line insertions don't shift earlier indices.
    for (var idx = report.Recommendations.Count - 1; idx >= 0; idx--)
    {
        var rec = report.Recommendations[idx];

        // Skip slides that already have an explicit layout directive.
        if (rec.IsExplicitLayout)
        {
            continue;
        }

        if (idx >= slideBoundaries.Count)
        {
            continue;
        }

        var insertAt = slideBoundaries[idx];
        var slideEnd = idx + 1 < slideBoundaries.Count ? slideBoundaries[idx + 1] : lines.Count;

        // Check if the slide already has a <!-- _layout: ... --> comment.
        if (SlideHasLayoutDirective(lines, insertAt, slideEnd))
        {
            continue;
        }

        lines.Insert(insertAt, $"<!-- _layout: {rec.RecommendedLayout} -->");
        patchedCount++;
    }

    File.WriteAllLines(deckPath, lines);
    Console.WriteLine($"Patched {patchedCount} slide(s) in {deckPath}");
}

static List<int> FindSlideBoundaries(List<string> lines)
{
    // The first slide starts at line 0 (after front-matter if present).
    // Subsequent slides start after "---" separator lines.
    // Matches SlideTokenizer.SplitSlides: uses line.Trim() == "---" and tracks fenced
    // code blocks (``` or ~~~) to avoid splitting on separators inside code fences.
    // Fence detection uses StartsWith on the TrimEnd()-ed line — same as SlideTokenizer.
    var boundaries = new List<int> { 0 };
    var inFence = false;

    // Skip front-matter block if present (the very first "---" / "---" pair).
    // Track fences within front-matter so a code block containing "---" is not
    // mistaken for the closing delimiter.
    var start = 0;
    if (lines.Count > 0 && lines[0].Trim() == "---")
    {
        for (var i = 1; i < lines.Count; i++)
        {
            var fmLine = lines[i].TrimEnd();
            if (fmLine.StartsWith("```", StringComparison.Ordinal) || fmLine.StartsWith("~~~", StringComparison.Ordinal))
            {
                inFence = !inFence;
            }

            if (!inFence && lines[i].Trim() == "---")
            {
                start = i + 1;
                inFence = false; // reset; any fence opened inside front-matter is now closed
                break;
            }
        }

        boundaries[0] = start;
    }

    for (var i = start; i < lines.Count; i++)
    {
        var line = lines[i].TrimEnd();

        // Track fenced code blocks — same toggle logic as SlideTokenizer.SplitSlides.
        if (line.StartsWith("```", StringComparison.Ordinal) || line.StartsWith("~~~", StringComparison.Ordinal))
        {
            inFence = !inFence;
        }

        if (!inFence && i > start && lines[i].Trim() == "---")
        {
            boundaries.Add(i + 1);
        }
    }

    return boundaries;
}

static bool SlideHasLayoutDirective(List<string> lines, int slideStart, int slideEnd)
{
    // slideEnd is already bounded to lines.Count by the caller.
    for (var i = slideStart; i < slideEnd; i++)
    {
        var trimmed = lines[i].Trim();
        if (trimmed.StartsWith("<!--") && trimmed.Contains("_layout:", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    }

    return false;
}

// ── Print: diagnose ─────────────────────────────────────────────────────────

static void PrintDiagnoseReport(TemplateDiagnosticReport report)
{
    var fileName = Path.GetFileName(report.TemplatePath);
    var title = $"Template Diagnostic Report: {fileName}";
    Console.WriteLine(title);
    Console.WriteLine(new string('=', title.Length));
    Console.WriteLine();

    // ── Overview ──────────────────────────────────────────────────────────
    Console.WriteLine("Overview");
    Console.WriteLine("--------");
    Console.WriteLine($"Slide masters:  {report.SlideMasterCount}");
    Console.WriteLine($"Slide layouts:  {report.Layouts.Count}");
    Console.WriteLine();

    // ── Layouts ───────────────────────────────────────────────────────────
    if (report.Layouts.Count == 0)
    {
        Console.WriteLine("No layouts found in template.");
        Console.WriteLine();
    }
    else
    {
        Console.WriteLine($"Layouts ({report.Layouts.Count})");
        Console.WriteLine(new string('─', 80));

        // Determine column widths dynamically.
        var nameWidth = Math.Max(4, report.Layouts.Max(l => l.Name.Length));
        var typeWidth = Math.Max(4, report.Layouts.Max(l => l.TypeCode?.Length ?? 0));
        var roleWidth = Math.Max(4, report.Layouts.Max(l => l.SemanticRole.ToString().Length));
        var indexWidth = report.Layouts.Count.ToString().Length;

        // Header
        var header = $"  {"#".PadLeft(indexWidth)}  {"Name".PadRight(nameWidth)}  {"Type".PadRight(typeWidth)}  {"Role".PadRight(roleWidth)}  Title  Body  Pic  Shapes";
        Console.WriteLine(header);
        Console.WriteLine(new string('─', Math.Max(header.Length, 80)));

        foreach (var (layout, index) in report.Layouts.Select((l, i) => (l, i + 1)))
        {
            var redundantMarker = layout.LikelyVisuallyRedundant ? " ⚠" : "  ";
            var line = string.Format(
                "  {0}  {1}  {2}  {3}  {4}    {5}   {6}    {7}{8}",
                index.ToString().PadLeft(indexWidth),
                layout.Name.PadRight(nameWidth),
                (layout.TypeCode ?? "—").PadRight(typeWidth),
                layout.SemanticRole.ToString().PadRight(roleWidth),
                layout.HasTitlePlaceholder ? "✓" : "✗",
                layout.HasBodyPlaceholder ? "✓" : "✗",
                layout.HasPicturePlaceholder ? "✓" : "✗",
                layout.NonPlaceholderShapeCount.ToString().PadLeft(6),
                redundantMarker);
            Console.WriteLine(line);
        }

        Console.WriteLine();
    }

    // ── Recommendations ───────────────────────────────────────────────────
    Console.WriteLine("Recommendations");
    Console.WriteLine("---------------");

    if (report.RecommendedDefaultContentLayout is { } defaultLayout)
    {
        Console.WriteLine($"Default content layout:   \"{defaultLayout}\"");
        Console.WriteLine($"  → front-matter:  layout: {defaultLayout}");
    }
    else
    {
        Console.WriteLine("Default content layout:   (none found)");
    }

    if (report.RecommendedTitleLayout is { } titleLayout)
    {
        Console.WriteLine($"Title slide layout:       \"{titleLayout}\"");
        Console.WriteLine($"  → per-slide:     _layout: {titleLayout}");
    }
    else
    {
        Console.WriteLine("Title slide layout:       (none found)");
    }

    if (report.RecommendedSectionLayout is { } sectionLayout)
    {
        Console.WriteLine($"Section header layout:    \"{sectionLayout}\"");
        Console.WriteLine($"  → per-slide:     _layout: {sectionLayout}");
    }
    else
    {
        Console.WriteLine("Section header layout:    (none found — consider using the title layout for section slides)");
    }

    if (report.RecommendedPictureCaptionLayout is { } pictureLayout)
    {
        Console.WriteLine($"Picture/caption layout:   \"{pictureLayout}\"");
        Console.WriteLine($"  → per-slide:     _layout: {pictureLayout}");
    }
    else
    {
        Console.WriteLine("Picture/caption layout:   (none found — image slides will use the default content layout)");
    }

    Console.WriteLine();

    // ── Warnings ──────────────────────────────────────────────────────────
    if (report.Warnings.Count > 0)
    {
        Console.WriteLine("Warnings");
        Console.WriteLine("--------");
        foreach (var warning in report.Warnings)
        {
            Console.WriteLine($"⚠  {warning}");
        }
        Console.WriteLine();
    }
}

// ── Print: doctor (human-readable) ──────────────────────────────────────────

static void PrintDoctorReport(TemplateDoctorReport report)
{
    var fileName = Path.GetFileName(report.TemplatePath);
    var title = $"Template Doctor Report: {fileName}";
    Console.WriteLine(title);
    Console.WriteLine(new string('=', title.Length));
    Console.WriteLine();

    if (report.Issues.Count == 0)
    {
        Console.WriteLine("✓  No structural issues found.");
        Console.WriteLine();
    }
    else
    {
        var fixable = report.Issues.Where(i => i.Severity == IssueSeverity.Fixable).ToList();
        var warnings = report.Issues.Where(i => i.Severity == IssueSeverity.Warning).ToList();
        var infos = report.Issues.Where(i => i.Severity == IssueSeverity.Info).ToList();

        Console.WriteLine($"Issues found: {report.Issues.Count} total  " +
            $"({fixable.Count} fixable, {warnings.Count} warning(s), {infos.Count} informational)");
        Console.WriteLine();

        if (fixable.Count > 0)
        {
            Console.WriteLine("Fixable issues");
            Console.WriteLine("--------------");
            foreach (var issue in fixable)
            {
                var prefix = issue.LayoutName is not null ? $"[{issue.LayoutName}]  " : "";
                Console.WriteLine($"  🔧 {prefix}{issue.Description}");
                if (issue.ProposedFix is not null)
                {
                    Console.WriteLine($"     Fix: {issue.ProposedFix}");
                }
            }

            Console.WriteLine();
        }

        if (warnings.Count > 0)
        {
            Console.WriteLine("Warnings");
            Console.WriteLine("--------");
            foreach (var issue in warnings)
            {
                var prefix = issue.LayoutName is not null ? $"[{issue.LayoutName}]  " : "";
                Console.WriteLine($"  ⚠  {prefix}{issue.Description}");
            }

            Console.WriteLine();
        }

        if (infos.Count > 0)
        {
            Console.WriteLine("Informational");
            Console.WriteLine("-------------");
            foreach (var issue in infos)
            {
                var prefix = issue.LayoutName is not null ? $"[{issue.LayoutName}]  " : "";
                Console.WriteLine($"  ℹ  {prefix}{issue.Description}");
            }

            Console.WriteLine();
        }
    }

    if (report.WroteFixedTemplate)
    {
        Console.WriteLine("Repaired template");
        Console.WriteLine("-----------------");
        Console.WriteLine($"Written to: {report.FixedTemplatePath}");

        if (report.AppliedFixes.Count > 0)
        {
            Console.WriteLine($"Fixes applied ({report.AppliedFixes.Count}):");
            foreach (var fix in report.AppliedFixes)
            {
                Console.WriteLine($"  ✓ {fix}");
            }
        }
        else
        {
            Console.WriteLine("No fixable issues were found; the file is a clean copy of the original.");
        }

        Console.WriteLine();
    }
    else if (report.Issues.Any(i => i.Severity == IssueSeverity.Fixable))
    {
        Console.WriteLine("Tip: Run with --write-fixed-template <output.pptx> to apply the fixable repairs above.");
        Console.WriteLine();
    }
}

// ── Print: doctor (JSON) ────────────────────────────────────────────────────

static void PrintDoctorReportJson(TemplateDoctorReport report)
{
    // Minimal hand-rolled JSON to avoid a System.Text.Json dependency.
    Console.WriteLine("{");
    Console.WriteLine($"  \"templatePath\": {JsonString(report.TemplatePath)},");
    Console.WriteLine($"  \"wroteFixedTemplate\": {(report.WroteFixedTemplate ? "true" : "false")},");
    Console.WriteLine($"  \"fixedTemplatePath\": {(report.FixedTemplatePath is not null ? JsonString(report.FixedTemplatePath) : "null")},");
    Console.WriteLine($"  \"appliedFixes\": [");
    var fixLines = report.AppliedFixes.Select((f, i) => $"    {JsonString(f)}{(i < report.AppliedFixes.Count - 1 ? "," : "")}");
    foreach (var line in fixLines)
    {
        Console.WriteLine(line);
    }

    Console.WriteLine("  ],");
    Console.WriteLine($"  \"issues\": [");

    for (var i = 0; i < report.Issues.Count; i++)
    {
        var issue = report.Issues[i];
        var comma = i < report.Issues.Count - 1 ? "," : "";
        Console.WriteLine("    {");
        Console.WriteLine($"      \"layoutName\": {(issue.LayoutName is not null ? JsonString(issue.LayoutName) : "null")},");
        Console.WriteLine($"      \"severity\": {JsonString(issue.Severity.ToString())},");
        Console.WriteLine($"      \"code\": {JsonString(issue.Code)},");
        Console.WriteLine($"      \"description\": {JsonString(issue.Description)},");
        Console.WriteLine($"      \"proposedFix\": {(issue.ProposedFix is not null ? JsonString(issue.ProposedFix) : "null")}");
        Console.WriteLine($"    }}{comma}");
    }

    Console.WriteLine("  ]");
    Console.WriteLine("}");
}

static string JsonString(string value)
    => "\"" + value
        .Replace("\\", "\\\\")
        .Replace("\"", "\\\"")
        .Replace("\n", "\\n")
        .Replace("\r", "\\r")
        .Replace("\t", "\\t") + "\"";
