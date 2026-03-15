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
    Console.WriteLine("  diagnose <template.pptx>    Analyze template layout structure and report recommendations.");
    Console.WriteLine("  doctor   <template.pptx>    Inspect and optionally repair structural issues.");
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
