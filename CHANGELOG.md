# Changelog

## 0.1.1

### Patch Changes

- 078971d: Releases are now published from GitHub Actions through npm trusted publishing (OIDC): no npm token is stored anywhere, and every version from this one on carries a provenance attestation that links the tarball to the commit and workflow run that built it (`npm audit signatures` verifies it). No code changes.

## 0.1.0

### Minor Changes

- b23e785: Responsive large exports: chunked writing with cooperative yields, worker-based zip in browsers (`zip: auto|sync|async` writer option), `signal` (AbortSignal → `ABORTED`), `onProgress`, and `timings.zipMs`. Output bytes are unchanged.
- b23e785: Chart.js 4 adapter on the `highcharts-editable-excel/chartjs` subpath: `exportChartJsToXlsx`, `analyzeChartJsCompatibility`, `downloadChartJsAsXlsx`, `extractChartJsModel` for live charts or plain configs (line, bar, stacked, pie, doughnut, scatter, bubble, mixed, time scales, secondary axes). `chart.js` is an optional peer dependency.
- b23e785: Polar line/spline/area charts export as native Excel **radar** charts (marker, standard or filled; polar columns stay blocking). `errorbar` series become custom Excel **error bars** on their linked parent series (`+err` / `-err` data columns). `columnrange` / `arearange` export as stacked columns/areas with a hidden base series; the data sheet holds Low, High and a Range column of library-generated formulas (`=High-Low`, or MAX/MIN Base/Up/Down helpers when a column range crosses zero) so editing Low/High updates the chart. Doughnut rings whose slices differ get their own category columns (no more `APPROXIMATED_LAYOUT` union). `lang.thousandsSep` / `lang.decimalPoint` are now read and raise `APPROXIMATED_NUMBER_FORMAT` once per chart when a format groups digits or shows decimals with a different separator. Per-chart menu settings are re-applied after `chart.update({ exporting })`. New `ThemeOverrides.cssVariables` resolves styled-mode colors headless. IR: `SeriesKind` gains `errorbar`, `columnrange`, `arearange`; optional `SeriesModel.linkedTo`, `PointModel.low/high`. Writer: `CellValue` `{ type: "formula" }` (library formulas only; the writer rejects anything else), `PlotGroupSpec` kind `radar`, optional `ExcelSeriesSpec.errorBars`.
- b23e785: Customize the export button and menu: `button` (symbol, custom `svgPath`, fill/stroke, theme, className, title), `menuIcon` (`excel`, custom svg/html), `menuStyle`/`menuItemStyle`/`menuItemHoverStyle`, and i18n via `langKey` (`lang.downloadEditableXLSX`). Per-chart `exporting.editableExcel.menuIcon`. Uninstall restores everything it changed.
  
  Behavior change: without `menuText`, the global menu item definition has no `text`; its text comes from `lang.downloadEditableXLSX` (registered with the default text at install when absent). Text order: per-chart `menuText` → install `menuText` → `lang[langKey]` → `DEFAULT_MENU_TEXT`.
- b23e785: First release of the editable-Excel export, with a stable public surface.
  
  - Export a live Highcharts chart, or a plain options object, to an `.xlsx` workbook holding a **native Excel chart** whose series reference cells on a data sheet (`exportHighchartsToXlsx`, `exportHighchartsOptionsToXlsx`, `downloadHighchartsAsXlsx`, `triggerDownload`).
  - Context-menu integration: `installHighchartsExcelExport(Highcharts)` adds a "Download editable Excel chart" item to every chart, with per-chart `exporting.editableExcel` options and `addEditableExcelMenuItem` for already-rendered charts.
  - Chart types: line, spline, area, column, bar (stacked and percent), scatter, bubble, pie, doughnut, and column/line combos with secondary axes; datetime, logarithmic, reversed and opposite axes.
  - Styling carried over where Excel can express it: series and point colors, gradients, dash styles, markers, fonts, data labels, legend position, axis number formats; `themeOverrides` and `hooks.transformModel` customise the result.
  - Diagnostics for everything that is approximated, unsupported or blocking (`analyzeChartCompatibility`, `onWarning`, `strictMode`, `ExportError` codes); missing data stays empty and is reported, never fabricated.
  - Multi-chart workbooks (`exportChartsToWorkbook`) and an optional PNG reference image next to the native chart (`includeReferenceImage`).
  - Data modes: `rendered` (what the chart shows, including Stock grouping and cropping) and `raw` (the source `options.data`).
  - Own narrowly scoped OOXML writer on `fflate` behind the `ExcelWriter` interface; a custom writer can be passed via `ExportOptions.writer`. All cell text is written as inline strings (formula-injection safe).
  - ESM and CommonJS builds, Highcharts as a peer dependency (tested against 11.4.9, 12.6.2 and 13.1.1), no runtime dependencies.
  - **Public surface split.** The main entry holds the stable API (export functions, install helpers, `ExportError`, constants, option / result / diagnostic types, the `ExcelWriter` interface). The experimental pipeline moved to the new subpath `highcharts-editable-excel/internals`: `extractChartModel`, `extractChartModelFromOptions`, `applyThemeOverrides`, `translateChartModel`, `resolveChartType`, `createDefaultExcelWriter`, `DiagnosticCollector`, `createDiagnostic`, `buildCompatibilityReport` and the `WorkbookSpec` family of types. `ChartModel` and its member types stay importable from the main entry (for `hooks.transformModel`) but are experimental. Update imports of those names to `highcharts-editable-excel/internals`.
  - New `highcharts-editable-excel/augment`: `import 'highcharts-editable-excel/augment'` adds `exporting.editableExcel` to `Highcharts.Options` for TypeScript; the runtime module is empty.
  - Validated structurally, by an `@office-kit/xlsx` round-trip and by LibreOffice renders; not yet opened in Microsoft Excel (see `docs/manual-qa.md`).
