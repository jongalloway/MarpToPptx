using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace MarpToPptx.Pptx.Validation;

public sealed class OpenXmlPackageValidator
{
    public IReadOnlyList<OpenXmlPackageValidationError> Validate(string pptxPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pptxPath);

        using var document = PresentationDocument.Open(pptxPath, false);
        return Validate(document);
    }

    public IReadOnlyList<OpenXmlPackageValidationError> Validate(PresentationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        return new OpenXmlValidator()
            .Validate(document)
            .Select(error => new OpenXmlPackageValidationError(
                error.Description ?? "Open XML validation error.",
                error.Path?.ToString() ?? "<unknown path>",
                error.Part?.Uri?.ToString()))
            .ToArray();
    }
}
