using MarpToPptx.Pptx.Diagnostics;

if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0]))
{
    Console.Error.WriteLine("Usage: dotnet run --project src/MarpToPptx.TemplateDiagnostics -- <path-to-template.pptx>");
    return 1;
}

var templatePath = Path.GetFullPath(args[0]);
if (!File.Exists(templatePath))
{
    Console.Error.WriteLine($"Template file was not found: {templatePath}");
    return 1;
}

try
{
    var diagnoser = new TemplateDiagnoser();
    var report = diagnoser.Diagnose(templatePath);
    PrintReport(report);
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine($"Template diagnostics failed for '{templatePath}': {exception.Message}");
    return 1;
}

static void PrintReport(TemplateDiagnosticReport report)
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
