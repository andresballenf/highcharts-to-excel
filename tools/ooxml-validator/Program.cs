// Validates .xlsx workbooks with the Open XML SDK (DocumentFormat.OpenXml), the same object model
// Microsoft's own tooling uses. It checks schema conformance plus the SDK's semantic constraints
// (relationship targets, attribute value ranges, element order, references between parts) against
// the Office 2016 file format version. Passing it is still not proof that Microsoft Excel opens
// the file without repair; see docs/manual-qa.md.
//
// Usage: dotnet run --project tools/ooxml-validator -- <file|glob>...
//   A glob may use * and ? in its last path segment (e.g. tests/output/export/v13/*.xlsx); shells
//   usually expand it before the program sees it. A glob that matches nothing is reported and skipped.
// Output: one line per error, "file: part: path: description [id]".
// Exit codes: 0 = every file valid, 1 = validation errors or unreadable files, 2 = usage error.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length == 0)
{
    Console.Error.WriteLine("usage: OoxmlValidator <file.xlsx|glob>...");
    return 2;
}

var files = new List<string>();
var missing = 0;
foreach (var arg in args)
{
    if (arg.IndexOfAny(['*', '?']) >= 0)
    {
        var dir = Path.GetDirectoryName(arg);
        var pattern = Path.GetFileName(arg);
        if (string.IsNullOrEmpty(dir)) dir = ".";
        var matches = Directory.Exists(dir)
            ? Directory.GetFiles(dir, pattern).OrderBy(p => p, StringComparer.Ordinal).ToArray()
            : [];
        if (matches.Length == 0) Console.Error.WriteLine($"warning: {arg}: no files match");
        files.AddRange(matches);
    }
    else if (File.Exists(arg))
    {
        files.Add(arg);
    }
    else
    {
        Console.Error.WriteLine($"{arg}: file not found");
        missing++;
    }
}

var validator = new OpenXmlValidator(FileFormatVersions.Office2016) { MaxNumberOfErrors = 0 };
var errorCount = 0;
var invalidFiles = 0;
foreach (var file in files)
{
    var fileErrors = 0;
    try
    {
        using var doc = SpreadsheetDocument.Open(file, false);
        foreach (var e in validator.Validate(doc))
        {
            var part = e.Part?.Uri.ToString() ?? "(package)";
            var path = e.Path?.XPath ?? "(none)";
            var description = e.Description.ReplaceLineEndings(" ");
            Console.WriteLine($"{file}: {part}: {path}: {description} [{e.Id}]");
            fileErrors++;
        }
    }
    catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or IOException or InvalidOperationException or System.Xml.XmlException)
    {
        Console.WriteLine($"{file}: (package): (none): cannot open: {ex.GetType().Name}: {ex.Message.ReplaceLineEndings(" ")}");
        fileErrors++;
    }
    if (fileErrors > 0) invalidFiles++;
    errorCount += fileErrors;
}

Console.WriteLine(
    $"ooxml-validator: {files.Count} file(s) checked against Office 2016, {errorCount} error(s) in {invalidFiles} file(s)"
        + (missing > 0 ? $", {missing} path(s) not found" : ""));
return errorCount > 0 || missing > 0 ? 1 : 0;
