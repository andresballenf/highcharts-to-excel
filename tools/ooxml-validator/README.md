# ooxml-validator

A minimal .NET 8 console tool that validates `.xlsx` workbooks with the Open XML SDK
([`DocumentFormat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml) 3.x), Microsoft's own library for
Office files and the closest automated check to Excel's parser available without Excel.

## What it validates

Each file is opened with `SpreadsheetDocument.Open(path, false)` (read-only) and checked with
`new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc)`, without an error limit. That covers:

- every part the package reaches through relationships: workbook, worksheets, drawings, charts, styles, shared parts;
- schema conformance (element order, unknown or missing elements and attributes, value types);
- the SDK's semantic constraints (attribute value ranges, unique ids, references between parts);
- elements and attributes that Office 2016 does not know.

A package that cannot be opened at all (bad zip, missing content types, broken relationships) is reported as one error.
Passing is **not** proof that Microsoft Excel opens the file without a repair prompt: see `docs/manual-qa.md`.

## Run

Needs the .NET 8 SDK (`sudo apt-get install -y dotnet-sdk-8.0` on Ubuntu 24.04, or https://dot.net). The first run
restores the NuGet package from api.nuget.org.

```bash
pnpm test        # writes tests/output/export/v13/*.xlsx and tests/output/writer/*.xlsx
dotnet run --project tools/ooxml-validator -- tests/output/export/v13/*.xlsx
pnpm validate:openxml   # export fixtures v13 + writer fixtures
```

Arguments are file paths or globs with `*`/`?` in the last path segment (expanded by the tool when the shell did not).
A glob that matches nothing prints a warning; a missing literal path is an error.

Output: one line per error, `file: part: path: description [id]`, then a summary line.
Exit codes: `0` no errors, `1` validation errors, unreadable packages or missing paths, `2` no arguments.

CI runs it in the `openxml-validate` job of `.github/workflows/ci.yml` on the v13 export fixtures, the writer fixtures
and `examples/workbooks/*.xlsx`.
