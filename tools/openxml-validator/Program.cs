using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Xml.Linq;

if (args.Length == 0)
{
    Console.Error.WriteLine("usage: openxml-validator <workbook.xlsx>");
    return 2;
}

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var failed = false;
foreach (var path in args)
{
    try
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var validationErrors = validator.Validate(document).ToList();
        var knownC16R3Errors = validationErrors
            .Where(IsKnownC16R3ExtensionValidatorError)
            .ToList();
        var errors = validationErrors.Except(knownC16R3Errors).ToList();
        if (errors.Count == 0)
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                valid = true,
                file = path,
                validator = "DocumentFormat.OpenXml 3.5.1",
                format = "Microsoft365 SpreadsheetML",
                knownSdkFalsePositives = knownC16R3Errors.Count
            }));
            continue;
        }
        failed = true;
        foreach (var error in errors)
        {
            Console.Error.WriteLine(JsonSerializer.Serialize(new
            {
                valid = false,
                code = error.Id ?? error.ErrorType.ToString(),
                error.Description,
                part = error.Part?.Uri.ToString(),
                xpath = error.Path?.XPath,
                suggestion = "Inspect the reported part/XPath and compare the generated SpreadsheetML with the Microsoft 365 Open XML schema."
            }));
        }
    }
    catch (Exception exception) when (exception is OpenXmlPackageException or IOException or UnauthorizedAccessException or InvalidDataException or FileFormatException)
    {
        failed = true;
        Console.Error.WriteLine(JsonSerializer.Serialize(new
        {
            valid = false,
            code = "package-open-failed",
            description = exception.Message,
            file = path,
            suggestion = "Confirm this is a valid, unencrypted XLSX OPC package with a workbook part and valid relationships."
        }));
    }
}

return failed ? 1 : 0;

static bool IsKnownC16R3ExtensionValidatorError(ValidationErrorInfo error)
{
    // Open XML SDK issue dotnet/Open-XML-SDK#1877: ChartExtensionList is not
    // implemented in the validator and valid c:ext/c16r3 content is misreported.
    if (error.Part is null || !error.Part.Uri.OriginalString.StartsWith("/xl/charts/", StringComparison.Ordinal))
    {
        return false;
    }

    var path = error.Path?.XPath ?? string.Empty;
    var knownShapeError = error.Id == "Sch_InvalidElementContentExpectingComplex"
        && path.Contains("/c:ext[", StringComparison.Ordinal)
        && error.Description?.Contains("dataDisplayOptions16", StringComparison.Ordinal) == true;
    var knownUriError = error.Id == "Sch_UndeclaredAttribute"
        && path.EndsWith("/c:ext[1]", StringComparison.Ordinal)
        && error.Description?.Contains("'uri' attribute", StringComparison.Ordinal) == true;
    var knownValueError = error.Id == "Sch_UndeclaredAttribute"
        && path.EndsWith("/c16r3:dispNaAsBlank[1]", StringComparison.Ordinal)
        && error.Description?.Contains("'val' attribute", StringComparison.Ordinal) == true;
    if (!knownShapeError && !knownUriError && !knownValueError)
    {
        return false;
    }

    try
    {
        using var stream = error.Part.GetStream(FileMode.Open, FileAccess.Read);
        var document = XDocument.Load(stream);
        XNamespace chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace c16r3 = "http://schemas.microsoft.com/office/drawing/2017/03/chart";
        return document.Descendants(chart + "ext").Any(extension =>
            (string?)extension.Attribute("uri") == "{56B9EC1D-385E-4148-901F-78D8002777C0}"
            && extension.GetNamespaceOfPrefix("c16r3") == c16r3
            && extension.Elements().Count() == 1
            && extension.Element(c16r3 + "dataDisplayOptions16") is { } options
            && options.Elements().Count() == 1
            && options.Element(c16r3 + "dispNaAsBlank") is { } display
            && ((string?)display.Attribute("val") is "0" or "1"));
    }
    catch (Exception exception) when (exception is IOException or System.Xml.XmlException)
    {
        return false;
    }
}
