# Changelog

All notable changes to `highcharts-editable-excel` are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses semantic versioning.

## Unreleased

- Export a live Highcharts chart, or a plain options object, to an `.xlsx` workbook holding a **native Excel chart** whose series reference cells on a data sheet (`exportHighchartsToXlsx`, `exportHighchartsOptionsToXlsx`, `downloadHighchartsAsXlsx`).
- Context-menu integration: `installHighchartsExcelExport(Highcharts)` adds a "Download editable Excel chart" item to every chart, with per-chart `exporting.editableExcel` options and `addEditableExcelMenuItem` for already-rendered charts.
- Chart types: line, spline, area, column, bar (stacked and percent), scatter, bubble, pie, doughnut, and column/line combos with secondary axes; datetime, logarithmic, reversed and opposite axes.
- Styling carried over where Excel can express it: series and point colors, gradients, dash styles, markers, fonts, data labels, legend position, axis number formats; `themeOverrides` and `hooks.transformModel` customise the result.
- Diagnostics for everything that is approximated, unsupported or blocking (`analyzeChartCompatibility`, `onWarning`, `strictMode`, `ExportError` codes); missing data stays empty and is reported, never fabricated.
- Multi-chart workbooks (`exportChartsToWorkbook`) and an optional PNG reference image next to the native chart (`includeReferenceImage`).
- Data modes: `rendered` (what the chart shows, including Stock grouping and cropping) and `raw` (the source `options.data`).
- Own narrowly scoped OOXML writer on `fflate` behind the `ExcelWriter` interface; a custom writer can be passed via `ExportOptions.writer`. All cell text is written as inline strings (formula-injection safe).
- ESM only, Highcharts as a peer dependency (tested against 11.4.9, 12.6.2 and 13.1.1), no runtime dependencies.
- Validated structurally, by an `@office-kit/xlsx` round-trip and by LibreOffice renders; not yet opened in Microsoft Excel (see `docs/manual-qa.md`).
