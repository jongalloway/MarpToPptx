using MarpToPptx.Pptx.Contrast;

if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0]))
{
    Console.Error.WriteLine("Usage: dotnet run --project src/MarpToPptx.ContrastAuditor -- <path-to-pptx>");
    return 1;
}

var pptxPath = Path.GetFullPath(args[0]);
if (!File.Exists(pptxPath))
{
    Console.Error.WriteLine($"PPTX file was not found: {pptxPath}");
    return 1;
}

try
{
    var auditor = new ContrastAuditor();
    var results = auditor.Audit(pptxPath);

    var failures = results.Where(r => r.IsFailing).ToList();

    if (failures.Count == 0)
    {
        if (results.Count == 0)
        {
            Console.WriteLine(
                $"Contrast audit found no auditable color pairs in '{pptxPath}'. " +
                "The presentation may use theme or inherited colors that cannot be resolved from solid fills alone.");
            return 0;
        }

        Console.WriteLine($"Contrast audit passed for '{pptxPath}'. {results.Count} color pair(s) checked.");
        return 0;
    }

    Console.Error.WriteLine($"Contrast audit found {failures.Count} failure(s) in '{pptxPath}':");
    foreach (var f in failures)
    {
        var textLabel = f.IsLargeText ? "large text" : "normal text";
        Console.Error.WriteLine(
            $"  Slide {f.SlideNumber} – {f.ShapeContext}: " +
            $"#{f.ForegroundColor} on #{f.BackgroundColor} " +
            $"= {f.ContrastRatio:F2}:1 (requires {f.MinimumRequiredRatio:F1}:1 for {textLabel})");
    }

    return 2;
}
catch (Exception exception)
{
    Console.Error.WriteLine($"Contrast audit failed for '{pptxPath}': {exception.Message}");
    return 1;
}
