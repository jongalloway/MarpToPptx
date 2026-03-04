namespace MarpToPptx.Pptx.Validation;

public sealed record OpenXmlPackageValidationError(
    string Description,
    string Path,
    string? PartUri);
