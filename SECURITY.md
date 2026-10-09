# Security policy

## Scope

`highcharts-editable-excel` runs entirely in the caller's process (browser or Node). It makes no
network requests, loads no remote code and writes only the workbook it returns. The main risks are
therefore in the generated file:

- **Formula injection.** Every cell that comes from chart data is written as an inline string or a
  number, never as a formula. The only formulas in a workbook are the ones the library generates
  itself for range series (`=High-Low` and helpers); the writer rejects any other formula value.
- **Malformed OOXML.** Workbooks are validated with the Open XML SDK in CI, round-tripped through `@office-kit/xlsx` and rendered in LibreOffice; the ECMA-376 XSD validation (`pnpm validate:xsd`) runs locally.
- **Zip bombs and resource use.** The writer bounds its output by the chart's own data; there is no
  decompression of untrusted input.

If you find a way to make the library emit a workbook that executes something, leaks data, or
crashes a spreadsheet application, that is in scope.

## Reporting a vulnerability

Please do not open a public issue. Use GitHub's private reporting:

https://github.com/andresballenf/highcharts-to-excel/security/advisories/new

Include the chart configuration, the library and Highcharts versions, and the spreadsheet
application affected. You will get an acknowledgement within 7 days. Fixes are released as a patch
version with a changeset entry, and the advisory is published after the fix is available.

## Supported versions

Only the latest minor version receives security fixes. Pin the package and keep it updated.
