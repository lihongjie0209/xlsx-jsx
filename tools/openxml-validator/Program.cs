using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

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
        var errors = validator.Validate(document).ToList();
        if (errors.Count == 0)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { valid = true, file = path, validator = "DocumentFormat.OpenXml 3.5.1", format = "Microsoft365 SpreadsheetML" }));
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
