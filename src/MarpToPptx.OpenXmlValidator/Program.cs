using MarpToPptx.Pptx.Validation;

if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0]))
{
    Console.Error.WriteLine("Usage: dotnet run --project src/MarpToPptx.OpenXmlValidator -- <path-to-pptx>");
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
    var validator = new OpenXmlPackageValidator();
    var validationErrors = validator.Validate(pptxPath);

    if (validationErrors.Count == 0)
    {
        Console.WriteLine($"Open XML validation passed for '{pptxPath}'.");
        return 0;
    }

    Console.Error.WriteLine($"Open XML validation failed for '{pptxPath}' with {validationErrors.Count} error(s):");
    foreach (var validationError in validationErrors)
    {
        Console.Error.WriteLine($"- {validationError.Description}");
        Console.Error.WriteLine($"  Path: {validationError.Path}");

        if (!string.IsNullOrWhiteSpace(validationError.PartUri))
        {
            Console.Error.WriteLine($"  Part: {validationError.PartUri}");
        }
    }

    return 2;
}
catch (Exception exception)
{
    Console.Error.WriteLine($"Open XML validation failed for '{pptxPath}': {exception.Message}");
    return 1;
}
