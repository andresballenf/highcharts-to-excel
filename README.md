# highcharts-editable-excel

[![npm version](https://img.shields.io/npm/v/highcharts-editable-excel.svg)](https://www.npmjs.com/package/highcharts-editable-excel)
[![CI](https://github.com/andresballenf/highcharts-to-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/andresballenf/highcharts-to-excel/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Sponsor](https://img.shields.io/badge/sponsor-%E2%9D%A4-ea4aaa?logo=githubsponsors)](https://github.com/sponsors/andresballenf)

Export live Highcharts charts to `.xlsx` workbooks that contain a **native, editable Excel chart** linked to worksheet cells.

## Overview

`highcharts-editable-excel` reads a rendered Highcharts chart (or a plain Highcharts options object), writes its data to a worksheet and builds a real Excel chart whose series reference those cells. It is not a picture of your chart: you can change a value in Excel and the chart redraws. You can also change the chart type, colors, titles and axes with Excel's own tools.

How this differs from the other ways to get a chart out of Highcharts:

| Export | What you get | Editable chart in Excel? |
| --- | --- | --- |
| `export-data` module (CSV / XLS) | The data table only | No chart at all |
| Exporting module (PNG / JPEG / SVG / PDF) | An image of the chart | No, it is a picture |
| **This library** | A workbook with a `Chart` sheet holding a native chart (`xl/charts/chartN.xml`, a DrawingML `c:chart` part) whose series formulas point at the `Data` sheet, e.g. `Data!$B$2:$B$5` | Yes. Excel recalculates the chart from the cells, and copying or pasting the chart into other workbooks or PowerPoint keeps it as editable as Excel itself allows |

Workbooks are generated locally, in the browser or in Node. There is no server component, no network request and no telemetry. The bundle has no runtime dependencies: the zip library (`fflate`) is bundled, and Highcharts is a peer dependency you provide.

The same pipeline also exports **Chart.js 4** charts through the `highcharts-editable-excel/chartjs` subpath (`chart.js` is an optional peer dependency); see [docs/chartjs.md](docs/chartjs.md).

## Install

```bash
pnpm add highcharts-editable-excel
# or
npm install highcharts-editable-excel
# or
yarn add highcharts-editable-excel
```

- Peer dependency: `highcharts >= 11`. The tested versions are 11.4.9, 12.6.2 and 13.1.1 (see [Compatibility matrix](#compatibility-matrix)). On Highcharts 11 the CommonJS module files are factories: call them with the namespace (`import Exporting from 'highcharts/modules/exporting'; Exporting(Highcharts);`) instead of relying on a side-effect import.
- The context-menu integration needs the Highcharts **exporting** module (`highcharts/modules/exporting`). The programmatic API does not need it, except for the optional reference image.
- The `export-data` module is **not** required.
- The package ships ESM (`dist/index.js`, for `import` and bundlers) and CommonJS (`dist/index.cjs`, for `require`) builds with matching type declarations, and has no side effects on import. It bundles [fflate](https://github.com/101arrowz/fflate) (MIT); see `THIRD_PARTY_LICENSES.md`.
- Entry points: `highcharts-editable-excel` (the stable API), `highcharts-editable-excel/internals` (the experimental lower-level pipeline) and `highcharts-editable-excel/augment` (a TypeScript augmentation for `exporting.editableExcel`, an empty module at runtime). See [Stability and versioning](#stability-and-versioning). In TypeScript the subpaths need `moduleResolution` set to `bundler`, `node16` or `nodenext`; with the legacy `node` (`node10`) setting only the main entry resolves.

## Quick start

```js
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { installHighchartsExcelExport } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts);
```

Every chart created after this call gets a **"Download editable Excel chart"** item in its context menu (the ☰ button). The library places the item after `downloadXLS`, `downloadCSV` or `downloadSVG`, whichever of those comes last in the menu. Clicking it downloads `<chart title>.xlsx`, or `chart.xlsx` when the chart has no title, with two sheets:

- **Chart**: the chart title in cell A1 and the native Excel chart below it, sized like the rendered chart.
- **Data**: a bold, frozen header row followed by the values. Category charts get one category column and one column per series. Pie charts get a `Category` column. Scatter and bubble charts get `<name> X`, `<name> Y` and `<name> Size` columns.

### Customizing the button and menu

The install can brand the ☰ button, style the dropdown, put an icon before the item and translate its text. Everything is optional and global (it applies to every chart), and `installation.uninstall()` restores the previous Highcharts defaults:

```js
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { installHighchartsExcelExport } from 'highcharts-editable-excel';

// Translate the item like any Highcharts text (before or after the install).
Highcharts.setOptions({ lang: { downloadEditableXLSX: 'Als Excel-Diagramm herunterladen' } });

installHighchartsExcelExport(Highcharts, {
  menuIcon: 'excel', // built-in spreadsheet glyph; or { svg: '<svg …>' }, { html: '…' }, null
  button: {
    // A download arrow in a 0..1 box, registered as the renderer symbol 'editableExcelButton'.
    svgPath: 'M0.5 0.05 V0.66 M0.2 0.38 L0.5 0.68 L0.8 0.38 M0.08 0.94 H0.92',
    symbolFill: '#1d6f42',
    symbolStroke: '#1d6f42',
    symbolStrokeWidth: 2,
    title: 'Export', // tooltip (lang.contextButtonTitle)
  },
  menuStyle: { border: '1px solid #1d6f42', borderRadius: '6px', boxShadow: '0 4px 12px rgba(0,0,0,.15)' },
  menuItemStyle: { fontFamily: 'system-ui', fontSize: '13px', padding: '6px 14px' },
  menuItemHoverStyle: { background: '#e8f3ec', color: '#1d6f42' },
});
```

The item text is resolved in this order:

1. the chart's `exporting.editableExcel.menuText`;
2. the install `menuText`;
3. `lang[langKey]` (`langKey` defaults to `'downloadEditableXLSX'`), read from the chart's merged `lang`, so `Highcharts.setOptions({ lang })` or a chart's own `lang` translates it;
4. `DEFAULT_MENU_TEXT` (`"Download editable Excel chart"`). The install registers it under `lang[langKey]` only when that key is absent, and `uninstall()` removes it again.

Without `menuText` and `menuIcon`, the item is defined with `textKey: langKey` and Highcharts looks the text up itself. A `lang` change after the install reaches charts created afterwards, not charts already rendered (Highcharts copies `lang` into each chart). With an icon, the item renders as `<span class="highcharts-editable-excel-item"><svg class="hc-excel-menu-icon" …/><span class="highcharts-editable-excel-label">TEXT</span></span>`, so the item's `textContent` is still the plain label.

Icons and item text go through Highcharts' HTML sanitizer (`Highcharts.AST`): tags and attributes outside `AST.allowedTags` / `AST.allowedAttributes` are dropped. That list has `svg`, `path`, `rect`, `circle`, `img` and `span` but not `viewBox` or `rx`, so size a custom SVG with `width`/`height` and draw in pixel coordinates (see `MENU_ICON_EXCEL_SVG`). `button.svgPath` takes the commands M, L, H, V, C, Q, A and Z (absolute or relative; not S or T). A path whose coordinates all lie within 0..1 fills the symbol box; any other path is scaled from its bounding box, made square and centered. The menu styles are inline CSS that Highcharts ignores in [styled mode](https://www.highcharts.com/docs/chart-design-and-style/style-by-css); style `.highcharts-menu` and `.highcharts-menu-item` there instead.

For a single chart, use plain Highcharts options: `exporting.buttons.contextButton` (symbol, colors, `theme`) and `navigation.menuStyle` / `menuItemStyle` / `menuItemHoverStyle` in that chart's options, and `exporting.editableExcel.menuIcon` / `menuText` for the item. The demo's "Branded export button" card (`pnpm demo`) shows the whole set, including a language toggle.

## Programmatic usage

```ts
import {
  exportHighchartsToXlsx,
  downloadHighchartsAsXlsx,
  analyzeChartCompatibility,
  exportHighchartsOptionsToXlsx,
  exportChartsToWorkbook,
  triggerDownload,
} from 'highcharts-editable-excel';

// 1. Bytes only: nothing is downloaded.
const result = await exportHighchartsToXlsx(chart, { filename: 'sales.xlsx' });
result.bytes;      // Uint8Array (the .xlsx package)
result.filename;   // "sales.xlsx"
result.mimeType;   // "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
result.warnings;   // Diagnostic[]: everything that was approximated or dropped
result.report;     // CompatibilityReport (editable, excelChartType, supported/approximated/...)
result.timings;    // { extractMs, translateMs, writeMs, totalMs }
triggerDownload(result.bytes, result.filename); // or upload the bytes somewhere

// 2. Export and download in one step (browser only).
await downloadHighchartsAsXlsx(chart, { fidelity: 'best-effort' });

// 3. Dry run: is this chart exportable, and what will be approximated? (synchronous)
const report = analyzeChartCompatibility(chart);
if (!report.editable) console.warn('Blocked by', report.blocking);

// 4. No chart instance (Node, a worker, a server): export from a plain options object.
const { bytes } = await exportHighchartsOptionsToXlsx({
  chart: { type: 'bar', width: 800, height: 450 },
  title: { text: 'Headcount' },
  xAxis: { categories: ['Sales', 'R&D', 'Support'] },
  series: [{ type: 'bar', name: '2026', data: [42, 77, 18] }],
});

// 5. Several charts in one workbook (one chart sheet + one data sheet per chart).
const multi = await exportChartsToWorkbook([{ chart: salesChart }, { chart: trafficChart }], { filename: 'dashboard' });
```

`exportHighchartsOptionsToXlsx` never renders anything. It always uses the **raw** `series[i].data` (it forces `dataMode: 'raw'`). Sizes come from `chart.width` and `chart.height`, defaulting to 600×400. Styles are the Highcharts defaults, overridden by whatever the options set explicitly. Because nothing is rendered, styled-mode CSS cannot be read and text measurements are estimated.

## Framework integrations

The wrapper packages below (`highcharts-react-official`, `highcharts-angular`, `highcharts-vue`) are your choice. They are **not** dependencies of this library. All you need is the Highcharts `Chart` instance. Complete files are in [`examples/`](examples/README.md).

**JavaScript (script tags + ESM)**: [`examples/javascript/index.html`](examples/javascript/index.html)

```html
<script src="https://code.highcharts.com/13.1.1/highcharts.js"></script>
<script src="https://code.highcharts.com/13.1.1/modules/exporting.js"></script>
<script type="module">
  import { installHighchartsExcelExport, downloadHighchartsAsXlsx } from 'highcharts-editable-excel';
  installHighchartsExcelExport(window.Highcharts);
  const chart = window.Highcharts.chart('container', {
    title: { text: 'Fruit consumption' },
    xAxis: { categories: ['Apples', 'Bananas', 'Oranges'] },
    series: [{ type: 'column', name: 'Jane', data: [1, 0, 4] }],
  });
  document.getElementById('download').onclick = () => downloadHighchartsAsXlsx(chart);
</script>
```

The bare specifier needs a bundler or an import map. The example file imports `../../dist/index.js` instead.

**TypeScript**: [`examples/typescript/main.ts`](examples/typescript/main.ts). This file is typechecked in this repository (`pnpm examples:typecheck`) and contains every plain-TypeScript snippet in this README. The React and Angular snippets below are shortened from their example files, which are typechecked too. The `.vue` file is not.

```ts
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { exportHighchartsToXlsx, installHighchartsExcelExport, type ExportResult } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts);
const chart = Highcharts.chart('sales', { series: [{ type: 'column', data: [1, 2, 3] }] });
const result: ExportResult = await exportHighchartsToXlsx(chart);
```

**React** (`highcharts-react-official`): [`examples/react/ChartWithExcelExport.tsx`](examples/react/ChartWithExcelExport.tsx)

```tsx
import { useRef } from 'react';
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import HighchartsReact from 'highcharts-react-official';
import { downloadHighchartsAsXlsx, installHighchartsExcelExport } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts);
const options: Highcharts.Options = { series: [{ type: 'column', name: 'Revenue', data: [12, 19, 15, 22] }] };

export function ChartWithExcelExport() {
  const chartComponentRef = useRef<HighchartsReact.RefObject>(null);
  const onDownload = async () => {
    const chart = chartComponentRef.current?.chart;
    if (chart) await downloadHighchartsAsXlsx(chart, { filename: 'monthly-revenue' });
  };
  return (
    <div>
      <HighchartsReact highcharts={Highcharts} options={options} ref={chartComponentRef} />
      <button onClick={onDownload}>Download editable Excel chart</button>
    </div>
  );
}
```

**Angular** (plain `ViewChild` + `Highcharts.chart`): [`examples/angular/chart-excel.component.ts`](examples/angular/chart-excel.component.ts)

```ts
@Component({
  selector: 'app-chart-excel',
  standalone: true,
  template: `<div #chartHost></div><button (click)="download()">Download editable Excel chart</button>`,
})
export class ChartExcelComponent implements AfterViewInit {
  @ViewChild('chartHost', { static: true }) chartHost!: ElementRef<HTMLDivElement>;
  private chart?: Highcharts.Chart;
  ngAfterViewInit(): void {
    this.chart = Highcharts.chart(this.chartHost.nativeElement, { series: [{ type: 'bar', data: [320, 210, 180] }] });
  }
  async download(): Promise<void> {
    if (this.chart) await downloadHighchartsAsXlsx(this.chart, { filename: 'orders' });
  }
}
```

With `highcharts-angular`, capture the chart instance in the component's `callbackFunction` input and pass it to the same functions.

**Vue 3** (`highcharts-vue`): [`examples/vue/ChartWithExcelExport.vue`](examples/vue/ChartWithExcelExport.vue)

```vue
<script setup lang="ts">
import { ref } from 'vue';
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { Chart } from 'highcharts-vue';
import { downloadHighchartsAsXlsx, installHighchartsExcelExport } from 'highcharts-editable-excel';

installHighchartsExcelExport(Highcharts);
const chartRef = ref<{ chart: Highcharts.Chart } | null>(null);
const options: Highcharts.Options = { series: [{ type: 'line', name: 'Signups', data: [31, 42, 38] }] };
const download = () => chartRef.value && downloadHighchartsAsXlsx(chartRef.value.chart, { filename: 'signups' });
</script>

<template>
  <Chart ref="chartRef" :highcharts="Highcharts" :options="options" />
  <button type="button" @click="download">Download editable Excel chart</button>
</template>
```

## API reference

Every function is a named export of `highcharts-editable-excel`. Charts are typed as `unknown` so that the package does not depend on a particular Highcharts type version. Pass a rendered `Highcharts.Chart`.

### Functions

| Function | Returns | Throws (`ExportError` code) |
| --- | --- | --- |
| `installHighchartsExcelExport(Highcharts: unknown, options?: InstallOptions)` | `Installation` | `INVALID_OPTIONS` if the argument is not the Highcharts namespace, or for an invalid `menuIcon`, `button.svgPath` or menu style (`details.property` names it; nothing is changed); `EXPORTING_MODULE_MISSING` if the exporting module is not loaded |
| `addEditableExcelMenuItem(chart: unknown, options?: InstallOptions)` | `void` | `INVALID_CHART` if the argument is not a rendered chart with `update()` |
| `exportHighchartsToXlsx(chart: unknown, options?: ExportOptions)` | `Promise<ExportResult>` | `INVALID_CHART`, `CHART_NOT_EDITABLE`, `INVALID_OPTIONS` (invalid option value, bad `transformModel` result), `WRITER_FAILURE` |
| `downloadHighchartsAsXlsx(chart: unknown, options?: ExportOptions)` | `Promise<ExportResult>` | Same as above, plus `BROWSER_REQUIRED` |
| `triggerDownload(bytes: Uint8Array, filename: string, mimeType = XLSX_MIME_TYPE)` | `void` | `BROWSER_REQUIRED` if `document`, `Blob` or `URL.createObjectURL` is missing |
| `analyzeChartCompatibility(chart: unknown, options?: ExportOptions)` | `CompatibilityReport` (synchronous) | `INVALID_CHART` and `INVALID_OPTIONS`; a blocked chart returns `editable: false` instead of throwing |
| `exportHighchartsOptionsToXlsx(highchartsOptions: object, options?: ExportOptions)` | `Promise<ExportResult>` | `INVALID_OPTIONS` (not an object), `CHART_NOT_EDITABLE`, `WRITER_FAILURE` |
| `exportChartsToWorkbook(entries: MultiChartExportEntry[], options?: { filename?: string; properties?: { title?: string; creator?: string }; strictMode?: boolean; writer?: ExcelWriter; signal?: AbortSignal; onProgress?: (p: ExportProgress) => void })` | `Promise<MultiChartExportResult>` | `INVALID_OPTIONS` (empty or non-array input, non-object entry), `INVALID_CHART` and `CHART_NOT_EDITABLE` (the message names the entry index), `WRITER_FAILURE` |

Notes:

- `installHighchartsExcelExport` is **idempotent per Highcharts namespace**. A second call returns the first `Installation`, and the first call's options win. `installation.uninstall()` removes the item definition and restores the global menu, the context-button options, `navigation` menu styles, `lang.contextButtonTitle`, the custom button symbol and `lang[langKey]` (deleted only when the install added it and it still holds the default text). Charts created after that are unaffected, and a new install is then allowed. `installation.langKey` and `installation.menuItemKey` give the keys in use.
- `addEditableExcelMenuItem` patches an already-rendered chart through `chart.update()`. Use it for charts created before the install. It is a no-op when the chart already lists the item or has `exporting.editableExcel.enabled === false`.
- `analyzeChartCompatibility` runs the same extraction and translation as an export but writes nothing and never renders the reference image.
- Errors thrown by **your own callbacks** (`onWarning`, `onProgress`, `hooks.transformModel`) are not wrapped: they propagate from every function above exactly as thrown.
- In a browser, the async exports yield to the event loop between the extract, translate and write phases and, in the built-in writer, between chunks of worksheet cells and chart cache points; the package is zipped in Web Workers. Extraction and translation are each still one synchronous task. See [Large charts, cancellation and progress](#large-charts-cancellation-and-progress).
- Every async export rejects with `ABORTED` once `options.signal` is aborted (for `exportChartsToWorkbook`, `signal` and `onProgress` go in its second argument).
- `exportChartsToWorkbook` names the sheets `Chart 1`/`Data 1`, `Chart 2`/`Data 2` and so on unless an entry sets `chartSheetName`/`dataSheetName`. Names that collide are de-duplicated with ` (2)`, ` (3)` and reported as `SHEET_NAME_ADJUSTED`. The default filename is `charts.xlsx`. The workbook title comes from the first chart.

Lower-level building blocks are **experimental** and live in the `highcharts-editable-excel/internals` subpath: `extractChartModel(chart, options)`, `extractChartModelFromOptions(options, extractOptions)`, `applyThemeOverrides(model, overrides)`, `resolveChartType(model)`, `translateChartModel(model, options)`, `createDefaultExcelWriter()`, `DiagnosticCollector`, `createDiagnostic`, `buildCompatibilityReport`, and the types `WorkbookSpec`, `SheetSpec`, `ExcelChartSpec` and related, `TranslateOptions`, `TranslationResult` and `ExtractOptions`. They may change in minor versions (see [Stability and versioning](#stability-and-versioning)) and are covered in [Custom writer and pipeline](#custom-writer-and-pipeline). The main entry also exports `CHART_TYPE_MATRIX`, the `ExcelWriter` interface (stable, for `ExportOptions.writer`) and, as types only, `ChartModel` and its member types (experimental; `hooks.transformModel` receives and returns a `ChartModel`).

Constants: `XLSX_MIME_TYPE`, `DEFAULT_MENU_TEXT` (`"Download editable Excel chart"`), `DEFAULT_MENU_ITEM_KEY` (`"downloadEditableXLSX"`), `DEFAULT_LANG_KEY` (`"downloadEditableXLSX"`), `DEFAULT_BUTTON_SYMBOL` (`"editableExcelButton"`), `MENU_ICON_EXCEL_SVG` (the markup of `menuIcon: 'excel'`).

### `ExportOptions`

| Option | Type | Default | Semantics |
| --- | --- | --- | --- |
| `filename` | `string` | Chart title, else `"chart"` | Sanitized: `\ / : * ? " < > \|` and control characters are removed and the name is limited to 120 characters. `.xlsx` is appended if missing. |
| `chartSheetName` | `string` | `"Chart"` | Worksheet that holds the chart. Invalid characters are removed, the name is cut to 31 characters and made unique (`SHEET_NAME_ADJUSTED`). |
| `dataSheetName` | `string` | `"Data"` | Worksheet that holds the values the chart references. |
| `includeSourceData` | `boolean` | `true` | `false` **hides** the data sheet. It is still written, because the chart must reference cells to stay editable. |
| `fidelity` | `'best-effort' \| 'minimal'` | `'best-effort'` | `'minimal'` keeps chart type, data, titles, axis types and scales, number formats, data-label content and legend position. It leaves colors, fonts, lines, markers, backgrounds and plot-area layout to Excel's defaults and suppresses style diagnostics. |
| `dataMode` | `'rendered' \| 'raw'` | `'rendered'` | See [Data semantics](#data-semantics). |
| `seriesVisibility` | `'visible' \| 'all'` | `'visible'` | `'visible'` skips hidden series (`HIDDEN_SERIES_EXCLUDED`). `'all'` exports them, and they are visible in Excel (`HIDDEN_SERIES_INCLUDED`). |
| `onWarning` | `(d: Diagnostic) => void` | none | Called once per distinct diagnostic (same code, property and series) as it is raised. |
| `themeOverrides` | `ThemeOverrides` | none | Explicit styling applied after extraction. See [Theme overrides](#theme-overrides). |
| `chartWidth`, `chartHeight` | `number` (CSS px) | Rendered size (`chart.chartWidth`/`chartHeight`; options path: `chart.width`/`height`, else 600×400) | Size of the Excel chart object: a finite number from 50 to 20000, otherwise `INVALID_OPTIONS`. |
| `strictMode` | `boolean` | `false` | Also throw `CHART_NOT_EDITABLE` when any diagnostic has outcome `unsupported` at `warning` or `error` severity (for example annotations, plot bands, formatter callbacks). Blocking diagnostics always throw, with or without strict mode. Reference-image diagnostics are ignored. |
| `includeReferenceImage` | `boolean` | `false` | Opt-in. Embeds a PNG of the Highcharts rendering to the right of the native chart for side-by-side comparison. Needs a real browser and `chart.getSVG` (exporting module). Otherwise the export continues without it and reports `WRITER_LIMITATION`. |
| `includeModel` | `boolean` | `false` | Adds the normalized `ChartModel` to `result.model` (debugging). |
| `hooks` | `{ transformModel?: (model: ChartModel) => ChartModel }` | none | Called after extraction, before theme overrides and translation. Must return a model with a `series` array (otherwise `INVALID_OPTIONS`). Never mutate the Highcharts chart in it. |
| `properties` | `{ title?: string; creator?: string }` | Title is the chart title. Creator is `"highcharts-editable-excel"`. | Workbook document properties. |
| `writer` | `ExcelWriter` | Built-in OOXML writer | Custom writer that receives the `WorkbookSpec` and returns the bytes. For advanced use; see [Custom writer and pipeline](#custom-writer-and-pipeline). |
| `signal` | `AbortSignal` | none | Cancels the export: checked between phases and between write chunks, and it terminates the worker zip. The export then rejects with `ExportError` `ABORTED` (`error.cause` is `signal.reason`); `downloadHighchartsAsXlsx` never downloads after an abort. See [Large charts, cancellation and progress](#large-charts-cancellation-and-progress). |
| `onProgress` | `(p: ExportProgress) => void` | none | Progress events `{ phase: 'extract' \| 'translate' \| 'write' \| 'zip' \| 'done'; fraction: number; detail?: string }`: at least one per phase and one per write chunk, `fraction` (0 to 1, overall) never decreases, `'done'` (1) comes once, last. Errors it throws propagate unwrapped. |

### `ThemeOverrides`

| Field | Type | Effect |
| --- | --- | --- |
| `colors` | `string[]` | Palette by series position. Pie and doughnut use it per slice. |
| `chartBackground`, `plotBackground` | `string` | Solid fills. |
| `fontFamily` | `string` | Family for every text element. |
| `title`, `subtitle`, `axisTitle`, `axisLabels`, `legend`, `dataLabels` | `FontOverride` (`{ family?, size? (CSS px), bold?, italic?, color? }`) | Per-element font overrides. |
| `gridLineColor`, `gridLineWidth` | `string`, `number` | Applied to existing grid lines. A width > 0 also adds y-axis grid lines where there were none. |
| `series` | `Record<number, { color?, lineWidth?, fillOpacity? }>` | Per-series overrides keyed by series index. |
| `cssVariables` | `Record<string, string>` | CSS custom property values for styled-mode charts, e.g. `{ '--highcharts-color-0': '#8e44ad' }`. Consulted first (headless and in a browser) when series colors, slice colors and `--highcharts-background-color` are resolved. `STYLED_MODE_FALLBACK` is not raised for series whose color variables all come from here. |

Colors that cannot be parsed are ignored and reported as `UNRESOLVED_COLOR`.

### `InstallOptions`

| Option | Type | Default | Semantics |
| --- | --- | --- | --- |
| `menuText` | `string` | none: `lang[langKey]`, else `"Download editable Excel chart"` | Menu item text for every chart. Wins over `lang`; a chart's `exporting.editableExcel.menuText` wins over it (see [Customizing the button and menu](#customizing-the-button-and-menu)). |
| `menuItemKey` | `string` | `"downloadEditableXLSX"` | Key in `exporting.menuItemDefinitions` and `menuItems`. |
| `langKey` | `string` | `"downloadEditableXLSX"` | `lang` key holding the item text (the definition's `textKey`). Translate with `Highcharts.setOptions({ lang: { [langKey]: '…' } })`. |
| `menuIcon` | `'excel' \| { svg: string } \| { html: string } \| null` | `null` | Icon before the item text. `'excel'` is the built-in 14×14 spreadsheet glyph in `currentColor` (class `hc-excel-menu-icon`). Sanitized by Highcharts' AST. |
| `button` | `ContextButtonOptions` | none | Global context-button branding, merged into `exporting.buttons.contextButton` (table below). |
| `menuStyle` | `Record<string, string \| number>` | none | CSS merged into `navigation.menuStyle` (the dropdown box). Not applied in styled mode. |
| `menuItemStyle` | `Record<string, string \| number>` | none | CSS merged into `navigation.menuItemStyle` (every item). Not applied in styled mode. |
| `menuItemHoverStyle` | `Record<string, string \| number>` | none | CSS merged into `navigation.menuItemHoverStyle` (item under the pointer). Not applied in styled mode. |
| `exportOptions` | `ExportOptions` | `{}` | Defaults for every menu export. |
| `onExport` | `(result: ExportResult, chart: unknown) => void` | none | Called after a successful menu download. |
| `onError` | `(error: unknown, chart: unknown) => void` | `console.error(...)` | Called when a menu export fails. |
| `insertAfter` | `string` | Last of `downloadXLS`/`downloadCSV`/`downloadSVG` present, else the end | Menu key to insert the item after. |

`ContextButtonOptions` (all optional; unset fields keep Highcharts' values):

| Field | Type | Semantics |
| --- | --- | --- |
| `symbol` | `string` | A renderer symbol name (`'menu'`, `'menuball'`, `'circle'`…), or the name to register `svgPath` under. |
| `svgPath` | `string \| Array<string \| number>` | An SVG path `d` (or a flat `['M', 0, 0, 'L', 1, 1]` array) registered as a renderer symbol named `symbol` (default `'editableExcelButton'`) and used as the button symbol. |
| `symbolFill`, `symbolStroke` | `string` | Symbol colors. |
| `symbolStrokeWidth`, `symbolSize` | `number` | Symbol stroke width and box size in px (Highcharts defaults 3 and 14). |
| `theme` | `Record<string, unknown>` | SVG attributes of the button box (`fill`, `stroke`, `r`, `states.hover.fill`…). |
| `text` | `string` | Text drawn next to the symbol. |
| `className` | `string` | Extra class, added after `highcharts-contextbutton`. |
| `title` | `string` | Button tooltip, written to `lang.contextButtonTitle`. |

### Per-chart `exporting.editableExcel`

Set in a chart's options. It is read when the chart is created and after every `chart.update({ exporting: … })` (menu items), and again when the menu item is clicked (export options):

```js
Highcharts.chart('container', {
  exporting: {
    editableExcel: {
      enabled: true,                       // false: no menu item on this chart
      menuText: 'Excel (editable chart)',  // this chart only
      menuIcon: 'excel',                   // this chart only; null removes the install icon
      filename: 'sales-2026',              // ...plus any ExportOptions; they override install exportOptions
      dataSheetName: 'Sales data',
    },
  },
});
```

If the chart sets its own `exporting.buttons.contextButton.menuItems`, the item is appended to that list. The caller's options object is never modified (frozen or shared option objects work); the change is made on the chart's own copy. `enabled`, `menuText`, `menuIcon` and `menuItems` are applied when the chart is created and again after each `chart.update()` that touches `exporting` (the context button is redrawn): `chart.update({ exporting: { editableExcel: { enabled: false } } })` removes the item, and a new `buttons.contextButton.menuItems` list gets it re-inserted. TypeScript users import the shipped augmentation once, next to their Highcharts import, so that `Highcharts.Options` accepts `exporting.editableExcel` (typed as `PerChartExportConfig`); see [`examples/typescript/main.ts`](examples/typescript/main.ts):

```ts
import Highcharts from 'highcharts';
import 'highcharts-editable-excel/augment';
```

### Result shapes

```ts
interface ExportResult {
  bytes: Uint8Array;
  filename: string;
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  warnings: Diagnostic[];
  report: CompatibilityReport;
  timings: { extractMs: number; translateMs: number; writeMs: number; zipMs: number; imageMs?: number; totalMs: number }; // zipMs: compression, part of writeMs; imageMs: reference image, 0 when none
  model?: ChartModel;                 // with includeModel: true
}

interface MultiChartExportResult {
  bytes: Uint8Array;
  filename: string;
  mimeType: ExportResult['mimeType'];
  charts: Array<{ chartSheetName: string; dataSheetName: string; warnings: Diagnostic[]; report: CompatibilityReport }>;
}

interface CompatibilityReport {
  editable: boolean;                  // a native editable chart can be / was produced
  sourceChartType: string;            // e.g. "column"
  excelChartType: string | null;      // e.g. "line", "stackedColumn", "percentStackedBar", "combo:column+line"; null when blocked
  supported: string[];                // option paths translated without loss
  approximated: string[];
  unsupported: string[];
  blocking: string[];
  dataConcerns: Diagnostic[];         // grouping, cropping, hidden series, nulls, limits, datetime
  warnings: Diagnostic[];             // every diagnostic, in emission order
}

interface Diagnostic {
  code: DiagnosticCode;               // stable identifier, e.g. "APPROXIMATED_LEGEND_POSITION"
  severity: 'info' | 'warning' | 'error';
  outcome: 'translated' | 'approximated' | 'unsupported' | 'blocking';
  property: string;                   // option path, e.g. "series[1].marker.symbol"
  message: string;
  seriesIndex?: number;
  details?: Record<string, unknown>;  // JSON-serializable
}
```

### `ExportError`

`class ExportError extends Error` with `name: 'HighchartsExcelExportError'`, `code: ExportErrorCode`, `details: { chartId?, property?, diagnostics? }` and the standard `cause` (the underlying error, when there is one).

| Code | Thrown when |
| --- | --- |
| `INVALID_CHART` | The argument is not a rendered chart (it needs `series`, `xAxis` and `yAxis` arrays, `options`, and `renderTo` or `container`). |
| `INVALID_OPTIONS` | An `ExportOptions` value is invalid (unknown `fidelity`/`dataMode`/`seriesVisibility`, `chartWidth`/`chartHeight` not a finite number from 50 to 20000, a non-string sheet name or filename, a non-function `onWarning` or `hooks.transformModel`, a `writer` without `write`); `details.property` names the option. Also: `installHighchartsExcelExport` did not get a Highcharts namespace; `exportHighchartsOptionsToXlsx` did not get an object; `hooks.transformModel` returned something that is not a model; `exportChartsToWorkbook` got an empty or invalid list. |
| `CHART_NOT_EDITABLE` | Any **blocking** diagnostic: polar columns/bars (or another polar type without a radar form), no series or no data, only unsupported series types, pie or bubble combined with other types, or worksheet row or column limit exceeded. In strict mode, also any `unsupported` warning. `details.diagnostics` lists the causes. |
| `EXPORTING_MODULE_MISSING` | `installHighchartsExcelExport` was called before `highcharts/modules/exporting` was loaded. |
| `BROWSER_REQUIRED` | `triggerDownload` or `downloadHighchartsAsXlsx` ran outside a browser. |
| `WRITER_FAILURE` | The writer (built-in or `options.writer`) threw. For example, the built-in writer refuses a plot group with more than 255 series. `error.cause` holds the original error. |
| `ABORTED` | `options.signal` was aborted before the export finished. `error.cause` is `signal.reason`. |

Unsupported chart types and worksheet row or column overflows are not error codes of their own. They arrive as diagnostic codes (`UNSUPPORTED_CHART_TYPE`, `ROW_LIMIT_EXCEEDED`, `COLUMN_LIMIT_EXCEEDED`) in `details.diagnostics` of a `CHART_NOT_EDITABLE` error.

## Stability and versioning

The package follows [semantic versioning](https://semver.org/). Until 1.0.0, a breaking change of the stable surface bumps the minor version and is called out in [CHANGELOG.md](CHANGELOG.md).

| Surface | Status | What that means |
| --- | --- | --- |
| Main entry `highcharts-editable-excel`: every runtime export (functions, `ExportError`, constants, `CHART_TYPE_MATRIX`) and the option, result and diagnostic types (`ExportOptions`, `InstallOptions`, `PerChartExportConfig`, `ExportResult`, `MultiChartExportResult`, `ExportErrorCode`, `Diagnostic`, `DiagnosticCode`, `CompatibilityReport`, …) and the `ExcelWriter` interface | **Stable** | Breaking changes only in a major version |
| `highcharts-editable-excel/augment` | **Stable** | Follows `PerChartExportConfig` |
| `highcharts-editable-excel/internals` (pipeline functions, `DiagnosticCollector`, `createDiagnostic`, `buildCompatibilityReport`, `createDefaultExcelWriter`, `WorkbookSpec` and the other spec types) | **Experimental** | May change in any minor version. Pin an exact version (`"highcharts-editable-excel": "x.y.z"`, not `^x.y.z`) when you use it |
| IR types `ChartModel`, `SeriesModel`, `AxisModel`, … (exported as types from both entries, for `hooks.transformModel`) | **Experimental** | May change in any minor version; a `transformModel` hook that reads or rewrites the model should pin an exact version |
| Diagnostic codes (`DiagnosticCode`) | **Append-only** | New codes can appear in a minor version; existing codes are not renamed or removed outside a major version. Treat unknown codes as informational |
| Excel output structure: sheet names, cell layout of the data sheet, chart XML | **Documented, not covered by semver** | Described in this README and `docs/compatibility.md`, but it can change in a minor or patch version (for example to fix fidelity). Do not parse the generated workbook by cell address in production code |

Release notes are generated with [Changesets](https://github.com/changesets/changesets): each pull request that changes the package adds a changeset (`pnpm changeset`).

## Compatibility matrix

### Chart types

Generated from `CHART_TYPE_MATRIX` (`src/core/chart-type-registry.ts`). Series types that are not supported are dropped from the chart (`UNSUPPORTED_SERIES_TYPE`). The export is blocked (`UNSUPPORTED_CHART_TYPE`) only when no supported series is left.

| Highcharts type | Excel chart | Support | Notes |
| --- | --- | --- | --- |
| line | line | native | Line chart; stacking maps to stacked/percent-stacked line. |
| spline | line (smoothed) | native | Excel smoothing differs slightly from Highcharts splines. |
| area | area | native | Stacking and percent stacking supported; fill opacity kept. |
| areaspline | area | approximated | Excel area charts cannot be smoothed; drawn with straight segments. |
| column | column (clustered/stacked/100%) | native | groupPadding/pointPadding map to gap width/overlap. |
| bar | bar (clustered/stacked/100%) | native | Horizontal bars. |
| pie | pie | native | One series per pie; slice colors and sliced points (explosion) kept. |
| pie + innerSize | doughnut | native | Hole size from innerSize (10-90%); multiple rings share one hole size. |
| scatter | scatter | native | X/Y columns per series; marker-only unless lineWidth > 0. |
| bubble | bubble | native | X/Y/Size columns per series; cannot be combined with other types. |
| columnrange | stacked column (hidden base) | approximated | Floating bars: a hidden base series stacked under a Range column (=High-Low formula cells, so editing Low/High updates the chart). Ranges crossing zero use Base/Up/Down helper columns. One range series per axis. |
| arearange | stacked area (hidden base) | approximated | A hidden Low area with a Range area (=High-Low formula cells) stacked on top. One range series per axis. |
| boxplot | - | unsupported | Excel box & whisker is a chartex type that cannot reference this layout. |
| heatmap | - | unsupported | Excel has no heatmap chart (only conditional formatting). |
| treemap | - | unsupported | Excel treemap is a chartex type not produced by this library. |
| waterfall | - | unsupported | Excel waterfall is a chartex type not produced by this library. |
| funnel | - | unsupported | Excel funnel is a chartex type not produced by this library. |
| gauge | - | unsupported | Excel has no gauge chart. |
| polar (line/spline/area) | radar (marker/standard/filled) | approximated | Categories are spaced evenly around the circle and lines are straight. Polar columns/bars and other types stay blocking (Excel has no polar columns). |
| variablepie | - | unsupported | Excel pies cannot vary slice radius. |
| sankey | - | unsupported | Excel has no flow diagrams. |
| networkgraph | - | unsupported | Excel has no network/graph layout chart. |
| timeline | - | unsupported | Excel has no timeline chart type. |
| histogram | - | unsupported | Excel histogram is a chartex type computed from raw data, not from this series. |
| bellcurve | - | unsupported | Derived series (computed in the browser) with no Excel equivalent. |
| errorbar | error bars (custom) | approximated | Drawn as custom Excel error bars on the linked parent series (bar, line, area, scatter, bubble); +err/-err columns on the data sheet. Unlinked error bars are not exported. |
| lollipop | - | unsupported | No Excel equivalent without helper series and error bars. |
| dumbbell | - | unsupported | No Excel equivalent without helper series and high-low lines. |

Combination rules (`resolveChartType`):

- Column, bar, line, area and scatter series combine into one Excel **combo chart**, for example `combo:column+line`, reported as `MIXED_SERIES_TYPES` (info).
- Pie, or bubble, combined with any other type **blocks** the export.
- Several plain pie series: only the first is drawn and the others stay on the data sheet. A pie together with a doughnut ring is drawn as doughnut rings.
- Excel has at most two value axes. The lowest y-axis index in use is primary, and every other one is collapsed onto the secondary axis (`UNSUPPORTED_AXIS_FEATURE`).
- Series of one type on one axis form one stack. Separate `stack` groups are merged (`APPROXIMATED_CHART_TYPE`).
- A range series (`columnrange`, `arearange`) takes its own stacked group: one per chart type and axis. Another range series, or a plain column/area series, on the same axis drops it from the chart (`UNSUPPORTED_SERIES_TYPE`).
- `errorbar` series are not drawn as series: they become error bars of the series they are linked to.
- Polar charts draw one radar group (the first series decides the radar style) on a single value axis.
- Doughnut rings whose slices differ (names or order) each get their own category column (`Category 1`, `Category 2`…), so every ring keeps its own slices in its own order.

### Features

| Feature | Outcome | Diagnostic / notes |
| --- | --- | --- |
| Stacking `normal` / `percent` | native | Stacked and percent-stacked column, bar, line and area charts. Multiple stack groups: approximated (`APPROXIMATED_CHART_TYPE`). |
| Markers (line, spline, scatter) | native | Circle, square, diamond, triangle, with size, fill and outline. `triangle-down` becomes triangle, custom or `url()` symbols become circle (`APPROXIMATED_MARKER`). Area-series markers are not exported. |
| Dash styles | native | All 11 Highcharts dash styles map to OOXML presets (for example `shortdash` → `sysDash`). |
| Per-point colors | native | Column, bar and bubble fills, pie slices, line and scatter markers, and `colorByPoint`. |
| Combo charts | native | See the combination rules above. Scatter inside a category combo gets its own hidden axes (`APPROXIMATED_AXIS_SCALE`). Horizontal bars plus lines share rotated axes (`APPROXIMATED_CHART_TYPE`, info). |
| Secondary axis | native | `SECONDARY_AXIS` (info). More than 2 y axes: approximated. Multiple x axes: the first is used (`MULTIPLE_X_AXES`). |
| Datetime x axis | native / approximated | Whole-day timestamps produce an Excel date axis with base unit days, months or years. Intraday timestamps produce an evenly spaced category axis (`APPROXIMATED_DATETIME`). |
| Logarithmic axis | native / approximated | Native on value axes and on the scatter x axis. A log x axis on a line or column chart becomes evenly spaced categories (`APPROXIMATED_AXIS_SCALE`). |
| Reversed / opposite axes | native | Excel axis orientation and crossing (including reversed horizontal bars). |
| Axis min/max, tickInterval, minorTickInterval, label rotation | native | On value axes. On category axes min/max are ignored (`APPROXIMATED_AXIS_SCALE`, info). |
| Axis and data-label `format` strings | native / approximated | Translated to Excel number formats where representable (`UNSUPPORTED_NUMBER_FORMAT`, `APPROXIMATED_NUMBER_FORMAT`). |
| `formatter` callbacks | unsupported | Never executed (`UNSUPPORTED_FORMATTER`). |
| Data labels | native / approximated | Value, category name, series name and percentage, with font, background and border. Positions Excel does not allow for the chart type fall back to Excel's default (`APPROXIMATED_DATA_LABELS`). |
| Legend | native / approximated | Positions top, bottom, left, right and top-right are native. Top-left, bottom-left and bottom-right become top or bottom (`APPROXIMATED_LEGEND_POSITION`). `floating` becomes an overlay. `reversed` is unsupported. |
| Gradients | native / approximated | Linear gradients are native (angle from the gradient vector). Radial gradients become top-to-bottom linear (`UNSUPPORTED_GRADIENT`). Patterns are unsupported (`UNRESOLVED_COLOR`). |
| Styled mode | approximated | `themeOverrides.cssVariables` (`--highcharts-color-N`, `--highcharts-background-color`) are used first. Otherwise, in a real browser the rendered SVG colors are read; headless, the palette by color index is used (`STYLED_MODE_FALLBACK`). |
| Data grouping (Stock) | approximated | Rendered mode exports the grouped points (`DATA_GROUPED`, info). Raw mode exports source data. |
| Zoom / cropping | approximated | The zoomed range becomes fixed axis bounds (`APPROXIMATED_AXIS_SCALE`, info). Cropped points: `DATA_CROPPED` (info). |
| `negativeColor`, `zones` | approximated | The series color is used (`UNSUPPORTED_STYLE`). |
| Rounded bars (`borderRadius`) | approximated | Square corners (`UNSUPPORTED_STYLE`, info, only when set explicitly). |
| Inverted line/area/scatter, semi-circle pie | approximated | `APPROXIMATED_CHART_TYPE`. |
| 3D (`options3d`) | ignored | Exported flat (`UNSUPPORTED_3D`). |
| Polar / spider | approximated / blocking | Line, spline, area and areaspline become an Excel radar chart (marker, standard or filled style; `APPROXIMATED_CHART_TYPE`, info): categories are spaced evenly and lines are straight. Polar columns, bars and other types stay blocking (`UNSUPPORTED_POLAR`). |
| Error bars (`errorbar`) | approximated | Custom Excel error bars on the linked parent series (bar, line, area, scatter, bubble), from `<parent> +err` / `<parent> -err` data columns (plus = high − y, minus = y − low). Unlinked error bars are not exported (`UNSUPPORTED_SERIES_TYPE`). |
| Range series (`columnrange`, `arearange`) | approximated | A stacked column/area with a hidden base series (`APPROXIMATED_CHART_TYPE`, info). The data sheet has `<name> Low`, `<name> High` and a `<name> Range` column of formulas `=High−Low`, so editing Low/High updates the chart. Column ranges crossing zero use `Base`/`Up`/`Down` formula columns (`MAX`/`MIN`) instead. One range series per chart type and axis; range data labels are not exported. |
| Annotations | unsupported | `UNSUPPORTED_ANNOTATION`. |
| Plot bands / plot lines | unsupported | `UNSUPPORTED_PLOT_BAND`. |
| Tooltips | unsupported | Never exported. Custom formatters are reported (`UNSUPPORTED_TOOLTIP`, info). `tooltip.valueDecimals`/`valuePrefix`/`valueSuffix` become the data cells' number format. |

Per-option detail (option path, outcome, diagnostic) is in [docs/compatibility.md](docs/compatibility.md).

### Tested versions and environments

- **Highcharts 11.4.9, 12.6.2 and 13.1.1.** The extraction, menu-install, export-fixture and compatibility-matrix suites run once per version (`*.v11.test.ts` / `*.v12.test.ts` / `*.v13.test.ts`); tested against 11.4.9. The other suites that use Highcharts run on 13.1.1. Other versions within the `>= 11` peer range are untested.
- **Node 22 and 24 + jsdom** (vitest; CI runs both). The `.xlsx` packages are inspected structurally, and the writer output round-trips through `@office-kit/xlsx`'s strict chart parser.
- **Chromium** via Playwright (demo end-to-end tests).
- **LibreOffice** headless render to PDF as a smoke check, when `soffice` is on the PATH (CI installs it and sets `REQUIRE_RENDER=1`, so a missing toolchain fails there instead of skipping).

**The generated workbooks were not opened in Microsoft Excel in the development environment.** Structural validation and LibreOffice rendering do not replace a check in Excel. Follow [docs/manual-qa.md](docs/manual-qa.md) before relying on a chart type in production.

## Fidelity limitations

Excel's chart engine is not Highcharts. Expect these differences:

- **Fonts.** The first usable family of the CSS stack is used. The Highcharts default stack maps to `Segoe UI`. Generic families map to `Arial`, `Times New Roman` and so on (`APPROXIMATED_FONT`). Excel substitutes fonts that are not installed. Sizes convert at 1 px = 0.75 pt.
- **One title.** Excel charts have no subtitle, so the subtitle is merged into the title as a second paragraph in its own font (size, weight and color from the subtitle style). Excel honours per-paragraph fonts. LibreOffice draws the whole rich title in the first paragraph's style.
- **No rounded bars, no zones, no negative colors.**
- **Data-label positions are restricted per chart type:**
  - clustered bars: center, inside end, inside base, outside end
  - stacked bars: no outside end
  - line, scatter and bubble: center, left, right, above, below
  - pie: center, inside end, outside end, best fit
  - area, doughnut and radar: no positioning at all
- **Legend** positions are limited to top, bottom, left, right and top-right.
- **Plot area.** The plot area is pinned to the source proportions (`APPROXIMATED_LAYOUT`, info). Excel still positions titles, labels and the legend around it itself, so spacing differs slightly.
- **No interactivity:** tooltips, drilldown, events, animation and the navigator are dropped.
- **Spline smoothing** uses Excel's algorithm. Areasplines are straight.
- **Area opacity.** Highcharts' `fillOpacity` (default 0.75) becomes fill transparency. As in Highcharts, it is not applied to an explicit `fillColor`.
- **Styled mode.** Colors live in CSS, and only a real browser can resolve them. Headless or server-side exports fall back to the palette plus `themeOverrides`. To reproduce the app's theme on the server, pass the CSS variables themselves: `themeOverrides: { cssVariables: { '--highcharts-color-0': '#8e44ad', '--highcharts-color-1': '#16a085' } }`. They are consulted before anything else, and `STYLED_MODE_FALLBACK` is not raised when every color variable a series needs is given.

## Data semantics

- **`dataMode: 'rendered'` (default)** exports the points the chart currently shows: after `setData`/`update`, after data grouping and after cropping. Grouped series report `DATA_GROUPED` and cropped series `DATA_CROPPED`. Both are info diagnostics with source and rendered point counts in `details`.
- **`dataMode: 'raw'`** parses `series.options.data` the way Highcharts does: numbers, `null`, `[x, y]`, `[x, y, z]`, `[name, y]`, arrays mapped through `series.keys`, point objects and typed arrays (y values). Missing x values follow `pointStart`/`pointInterval`/`pointIntervalUnit` (calendar units are stepped on the dates the chart displays), and `relativeXValue` x values count intervals from `pointStart`. On datetime axes, date strings (`'2024-01-01'`, also as `pointStart`) are parsed; strings that are not dates are skipped (`NON_NUMERIC_VALUE`). If the source data is not available (for example data loaded through a data table), the rendered points are used (`DATA_MODE_FALLBACK`). The options-object API always uses raw data.
- **Boost module.** A boosted series keeps only pixel positions in `series.points`, so its values are read from the series data instead (`DATA_MODE_FALLBACK`, info).
- **Hidden series:** see `seriesVisibility`. A hidden pie slice is exported and shown (`HIDDEN_POINT`).
- **Nulls vs missing points.** A `null` y becomes an empty cell, drawn as a gap (`NULL_VALUES`, info). A series without a value for a category another series has also gets an empty cell (`UNALIGNED_X_VALUES`). Values that are not numeric are skipped (`NON_NUMERIC_VALUE`).
- **Zoom.** A zoomed axis (`userMin`/`userMax`) is exported as fixed axis bounds for the visible window.
- **Dates** are written as Excel serial numbers of the **wall-clock time the chart displays**. Excel has no time zones, so with `time.timezone` (or the legacy `useUTC: false` / `timezoneOffset`) each datetime x value is moved by the zone's offset at that instant (`APPROXIMATED_DATETIME`, info, property `time.timezone`; the offset of the first point is in `model.meta.datetimeOffsetMinutes`). UTC charts are unchanged. Dates are formatted from the axis label format when it can be translated.
- **Category charts** share one category column. Numeric and datetime x values are the sorted union across series. Categories follow the axis order, or first appearance when points are matched by name. Pie series share a `Category` column. Scatter and bubble series get their own X/Y(/Size) blocks.
- **Nothing is fabricated.** No values are interpolated, no unknown point shapes are guessed, and no rows are invented. Missing data stays empty, and every gap is reported.

## Advanced usage

All snippets below are copied from [`examples/typescript/main.ts`](examples/typescript/main.ts), which is typechecked.

### Theme overrides

```ts
const brandOptions: ExportOptions = {
  themeOverrides: {
    colors: ['#0b5394', '#e69138', '#6aa84f'],
    fontFamily: 'Calibri',
    title: { size: 18, bold: true, color: '#222222' },
    gridLineColor: '#d9d9d9',
    series: { 1: { lineWidth: 3 } },
  },
  onWarning: (d: Diagnostic) => console.info(`[excel] ${d.severity} ${d.code} at ${d.property}`),
  includeSourceData: false, // data sheet is hidden; the chart still references it
  hooks: {
    transformModel: (model) => ({
      ...model,
      series: model.series.map((s) => ({ ...s, name: s.name.toUpperCase() })),
    }),
  },
  properties: { title: 'Sales report', creator: 'ACME BI' },
};
```

### Reading warnings

There are three ways to read them:

- Live, through `onWarning`.
- After the export, through `result.warnings`.
- Grouped, through `result.report` (`approximated`, `unsupported`, `blocking` and `dataConcerns`).

Switch on `diagnostic.code`; the codes are stable. They are grouped as follows:

- chart type: `UNSUPPORTED_CHART_TYPE`, `APPROXIMATED_CHART_TYPE`, `MIXED_SERIES_TYPES`, `UNSUPPORTED_SERIES_TYPE`, `UNSUPPORTED_POLAR`, `UNSUPPORTED_3D`, `EMPTY_CHART`, `EMPTY_SERIES`
- data: `DATA_GROUPED`, `DATA_CROPPED`, `DATA_MODE_FALLBACK`, `HIDDEN_SERIES_EXCLUDED`, `HIDDEN_SERIES_INCLUDED`, `HIDDEN_POINT`, `UNALIGNED_X_VALUES`, `ROW_LIMIT_EXCEEDED`, `COLUMN_LIMIT_EXCEEDED`, `NULL_VALUES`, `NON_NUMERIC_VALUE`
- axes: `APPROXIMATED_AXIS_SCALE`, `UNSUPPORTED_AXIS_FEATURE`, `MULTIPLE_X_AXES`, `SECONDARY_AXIS`, `APPROXIMATED_DATETIME`
- styling: `UNRESOLVED_COLOR`, `UNSUPPORTED_GRADIENT`, `APPROXIMATED_FONT`, `APPROXIMATED_LEGEND_POSITION`, `APPROXIMATED_MARKER`, `APPROXIMATED_LAYOUT`, `UNSUPPORTED_STYLE`, `STYLED_MODE_FALLBACK`, `UNSUPPORTED_NUMBER_FORMAT`, `APPROXIMATED_NUMBER_FORMAT`, `UNSUPPORTED_FORMATTER`, `APPROXIMATED_DATA_LABELS`
- not representable: `UNSUPPORTED_TOOLTIP`, `UNSUPPORTED_ANNOTATION`, `UNSUPPORTED_PLOT_BAND`
- integration: `FORMULA_LIKE_TEXT_ESCAPED`, `SHEET_NAME_ADJUSTED`, `WRITER_LIMITATION`

The `DiagnosticCode` type also declares `APPROXIMATED_COLOR`, `APPROXIMATED_FONT_SIZE` and `APPROXIMATED_DASH_STYLE`. These are **reserved**: no code path raises them today. A missing exporting module is reported as an `ExportError` (`EXPORTING_MODULE_MISSING`), not as a diagnostic.

### Strict mode

```ts
try {
  await downloadHighchartsAsXlsx(sales, { ...brandOptions, strictMode: true });
} catch (error) {
  if (error instanceof ExportError) {
    console.error(error.code, error.message, error.details.diagnostics?.map((d) => d.code));
  } else {
    throw error;
  }
}
```

### Custom filenames and sheet names

`filename` (`.xlsx` is appended), `chartSheetName` and `dataSheetName` can be set per call, in `InstallOptions.exportOptions` or per chart in `exporting.editableExcel`. Sheet names follow Excel's rules:

- no `[ ] : * ? / \`
- no leading or trailing apostrophe
- 1-31 characters
- not `History`
- unique, ignoring case

Adjusted names are reported as `SHEET_NAME_ADJUSTED`.

### Runtime-updated charts

The export reads the chart at call time, so `chart.update()`, `series.setData()`, `addSeries()`, toggling series visibility and zooming are all reflected. Nothing is cached.

### Multi-chart workbook

```ts
const result = await exportChartsToWorkbook(
  [
    { chart: sales, options: { chartSheetName: 'Sales', dataSheetName: 'Sales data' } },
    { chart: traffic }, // sheets "Chart 2" / "Data 2"
  ],
  { filename: 'dashboard' },
);
triggerDownload(result.bytes, result.filename, result.mimeType);
```

### Reference image

`includeReferenceImage: true` renders `chart.getSVG()` to a PNG at 2x through a `<canvas>`. The PNG goes on the chart sheet to the right of the native chart. It is meant for comparing the two renderings and is off by default.

### Large charts, cancellation and progress

```ts
import { ExportError, exportHighchartsToXlsx, triggerDownload } from 'highcharts-editable-excel';

const controller = new AbortController();
cancelButton.onclick = () => controller.abort();
try {
  const result = await exportHighchartsToXlsx(chart, {
    signal: controller.signal,
    onProgress: ({ phase, fraction }) => {
      progressBar.value = fraction; // 0..1, never decreases
      progressLabel.textContent = phase; // 'extract' | 'translate' | 'write' | 'zip' | 'done'
    },
  });
  triggerDownload(result.bytes, result.filename, result.mimeType);
} catch (error) {
  if (error instanceof ExportError && error.code === 'ABORTED') return; // error.cause === controller.signal.reason
  throw error;
}
```

- **Chunked writing.** The built-in writer serializes worksheet cells and chart cache points in chunks of 10,000. After each chunk it reports progress, checks the signal and, in a browser (`typeof window !== 'undefined'`), yields a macrotask (`scheduler.yield()`, else a `MessageChannel` message) so a spinner or progress bar can paint. Large parts are UTF-8 encoded piece by piece, never as one giant string. Node and workers do not pause.
- **Zip in Web Workers.** In browsers the package is compressed by fflate's async `zip`, which deflates every part of 160 KB or more in a Web Worker started from a `blob:` URL. The writer first checks once per page that such a worker starts. A Content-Security-Policy without `worker-src blob:` (or `child-src`/`script-src` allowing it) makes that check fail, and the writer then falls back to the synchronous `zipSync`. Node and jsdom have no `Worker` and always use `zipSync`. Every path writes the same bytes. To force a path, pass `writer: createDefaultExcelWriter({ zip: 'sync' | 'async' | 'auto' })` (from `highcharts-editable-excel/internals`; default `'auto'`; `'async'` rejects where `Worker` is undefined). The worker path runs in a real browser only. It was checked once in headless Chromium with a scratch script; the bundled Playwright downloads use small charts, which `'auto'` zips synchronously.
- **Progress.** At least one event per phase and one per write chunk. `fraction` is overall: extract and translate take the first 20 %, writing 20 to 60 %, zipping 60 to 100 %. `'done'` (1) is emitted once, after a successful export. `result.timings.zipMs` is the zip share of `writeMs`.
- **Cancellation.** The signal is checked before each phase, between write chunks and around the zip, and an abort terminates the zip workers. Every export function rejects with `ExportError` `ABORTED` (`error.cause` is `signal.reason`). `downloadHighchartsAsXlsx` never starts a download after an abort. `exportChartsToWorkbook` takes `signal` and `onProgress` in its second argument; an entry's own are ignored.
- **Custom writers** receive `write(workbook, { signal, onProgress, yieldEvery })`. The second argument is optional: a writer that ignores it still gets cancelled before and after `write`, and the phase boundaries are still reported.
- **What still blocks.** Extraction and translation each run as one synchronous task. In headless Chromium on a 5 × 100,000-point export they took about 0.15 s and 0.35 s. Before handing the parts to its workers, fflate computes CRC-32 checksums and copies the data on the main thread, about 0.2 s per 60 MB. Garbage collection of the large intermediate data adds pauses. In that Chromium run the longest main-thread stall fell from about 3.5 s (`zip: 'sync'`) to 1.3 to 1.4 s (default), and painted frames during the export rose from 8 or 9 to about 120. For 1 × 100,000 points the stall fell from 1.1 to 1.6 s to 0.3 to 0.4 s. `pnpm bench` measures the same stall in Node as "max tick gap", for both zip paths.
- **Roadmap.** Streaming (writing rows straight into a streaming zip without holding the whole package in memory) and running the entire export in a Worker are not implemented. Memory still peaks at a few times the XLSX's uncompressed size.

### Custom writer and pipeline

The `ExcelWriter` interface is `{ name: string; write(workbook: WorkbookSpec, context?: WriteContext): Promise<Uint8Array> }`, where the optional `context` is `{ signal?, onProgress?, yieldEvery? }` (see [Large charts, cancellation and progress](#large-charts-cancellation-and-progress)). Pass an implementation as `options.writer` to `exportHighchartsToXlsx`, `downloadHighchartsAsXlsx`, `exportHighchartsOptionsToXlsx` or `exportChartsToWorkbook`. `createDefaultExcelWriter()` returns the built-in OOXML writer, so a custom writer can wrap it. `ExcelWriter` is a stable type of the main entry; `createDefaultExcelWriter`, `DiagnosticCollector`, `extractChartModel`, `translateChartModel` and the `WorkbookSpec` type are experimental and come from the subpath:

```ts
import { exportHighchartsToXlsx, type ExcelWriter } from 'highcharts-editable-excel';
import {
  DiagnosticCollector,
  createDefaultExcelWriter,
  extractChartModel,
  translateChartModel,
} from 'highcharts-editable-excel/internals';
```


```ts
const loggingWriter: ExcelWriter = {
  name: 'logging-default',
  async write(workbook) {
    console.log('writing sheets', workbook.sheets.map((s) => s.name));
    return createDefaultExcelWriter().write(workbook);
  },
};
const { bytes } = await exportHighchartsToXlsx(sales, { writer: loggingWriter });
```

For full control, run the three pipeline steps yourself with the exported `DiagnosticCollector`:

1. `extractChartModel` produces the `ChartModel`.
2. `translateChartModel` produces the sheets.
3. Any `ExcelWriter` writes them.

Unlike the high-level functions, this path does not throw on blocking diagnostics, so check `translation.blocking` yourself:

```ts
const diagnostics = new DiagnosticCollector((d) => console.info(d.code, d.property));
// 1. Extract the neutral ChartModel.
const model = extractChartModel(chart, { dataMode: 'rendered', seriesVisibility: 'visible', diagnostics });
// 2. Translate it into worksheet specs with a native chart.
const translation = translateChartModel(model, {
  chartSheetName: 'Chart',
  dataSheetName: 'Data',
  includeSourceData: true,
  fidelity: 'best-effort',
  diagnostics,
});
if (translation.blocking) throw new Error(`Not editable: ${diagnostics.items.map((d) => d.code).join(', ')}`);
// 3. Write the WorkbookSpec with any ExcelWriter.
return loggingWriter.write({ properties: { title: 'Manual' }, sheets: translation.sheets });
```

`buildCompatibilityReport` and `createDiagnostic` are in `highcharts-editable-excel/internals` too, so you can build a report or raise your own diagnostics in `hooks.transformModel`.

## Security and privacy

- **Local only.** The workbook is generated in memory, and downloads use a temporary `blob:` URL that is revoked afterwards. No data leaves the page or process.
- **No formula injection.** Every text cell is written as an inline string (`t="inlineStr"`), never as a formula. A category such as `=SUM(A1)` or `@cmd` stays literal text, and the first such value is reported as `FORMULA_LIKE_TEXT_ESCAPED` (info). Worksheet cells contain no formulas at all. Only the chart part has series references (`c:f`) to the data sheet.
- **No macros, no external links.** The package contains only workbook, worksheet, styles, theme, drawing, chart, PNG (reference image only) and document-property parts.
- **XML safety.** All text and attributes are XML-escaped. Characters not allowed in XML 1.0 are stripped. Cell text is cut at Excel's 32,767-character limit.
- **Excel limits.**
  - More than 1,048,576 rows or 16,384 columns blocks the export (`ROW_LIMIT_EXCEEDED` / `COLUMN_LIMIT_EXCEEDED` inside `CHART_NOT_EDITABLE`).
  - Excel 2007 capped chart series at 32,000 points. Excel 2010 and later are limited by memory only. A longer series is still exported, with a non-blocking `ROW_LIMIT_EXCEEDED` warning (outcome `approximated`), because older Excel versions may truncate it or render it slowly. To stay under the guidance, export grouped data (Highcharts Stock, `dataMode: 'rendered'`) or reduce the points first.

## Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| `ExportError: EXPORTING_MODULE_MISSING` | Load `highcharts/modules/exporting` (or `modules/exporting.js` from the CDN) **before** calling `installHighchartsExcelExport`. |
| The menu item does not show | Charts created *before* the install are not patched. Call `addEditableExcelMenuItem(chart)`. Also check `exporting.editableExcel.enabled !== false` and that `exporting.enabled` is not `false`. |
| `chart.update({ exporting: … })` did not change the item | Per-chart `exporting.editableExcel` (`enabled`, `menuText`, `menuIcon`) and `menuItems` are re-applied on Chart `afterUpdate` while the install is active. A chart updated after `uninstall()`, or one that was never patched, needs `addEditableExcelMenuItem(chart)`. Export options in `editableExcel` (filename, sheet names…) are read on each click and need nothing. |
| The menu icon does not show, or a custom SVG icon renders empty | Highcharts' AST sanitizer drops tags and attributes it does not allow (Highcharts logs error #33 for a tag). `viewBox` and `rx` are not allowed: give the `<svg>` `width`/`height` and pixel coordinates, as `MENU_ICON_EXCEL_SVG` does, or use `{ html: '<img src="…">' }`. |
| The menu text is not translated | The text comes from `lang[langKey]` only when neither the chart's `exporting.editableExcel.menuText` nor the install `menuText` is set. A `lang` change reaches charts created after it (Highcharts copies `lang` into each chart): re-create the charts already rendered, as the demo's "Deutsch" button does. |
| `menuStyle` / `menuItemStyle` / button colors have no effect | Styled mode (`chart.styledMode`) ignores inline styles: style `.highcharts-menu`, `.highcharts-menu-item` and `.highcharts-contextbutton` in CSS. A chart's own `navigation` or `exporting.buttons.contextButton` options override the install's global ones. |
| `INVALID_OPTIONS` naming `button.svgPath` | Only M, L, H, V, C, Q, A and Z are supported (absolute or relative), and every command needs complete coordinates. Convert S/T curves to C/Q. |
| `CHART_NOT_EDITABLE` | Inspect `error.details.diagnostics` or run `analyzeChartCompatibility(chart)`. The usual causes are polar column/bar charts, only unsupported series types, pie or bubble mixed with other types, an empty chart, or `strictMode` with an unsupported feature. |
| Nothing downloads / `BROWSER_REQUIRED` | `downloadHighchartsAsXlsx` needs `document` and `URL.createObjectURL`. In Node or a worker, use `exportHighchartsToXlsx` or `exportHighchartsOptionsToXlsx` and handle `bytes` yourself. Browsers can also block downloads that are not triggered by a user gesture, or in sandboxed iframes without `allow-downloads`. Trigger the export from a click. |
| Wrong colors in styled mode on the server | CSS is not readable without a real browser (`STYLED_MODE_FALLBACK`). Pass the app's CSS variables in `themeOverrides.cssVariables` (or `themeOverrides.colors`, backgrounds and fonts). |
| Dates shifted by some hours | Cells hold the wall-clock time the chart displays in its `time.timezone` (`APPROXIMATED_DATETIME` reports the shift). On the options path the zone comes from `options.time` only: a zone set globally with `Highcharts.setOptions({ time })` is not visible there, so pass `time` in the options. An unknown zone name is exported as UTC (`APPROXIMATED_DATETIME`, warning). |
| Large charts are slow | All points are written to the sheet and to the chart cache. Use `dataMode: 'rendered'` with data grouping, or fewer points. `result.timings` shows where the time goes (`zipMs` is the compression share of `writeMs`). In browsers the writer yields between chunks and zips in Web Workers. Show `onProgress` in a progress bar and offer cancellation with `signal` (see [Large charts, cancellation and progress](#large-charts-cancellation-and-progress)). Extraction and translation still run as single tasks, and a CSP that forbids `blob:` workers makes the zip synchronous. For measurements, run `pnpm bench`: it prints a table (timings, max tick gap for the main-thread and worker zip paths) and writes `tests/output/bench.json`. Streaming export is on the roadmap. |
| Vite / CJS errors importing Highcharts modules | Highcharts ships UMD/CJS files without an exports map. Import modules as side effects (`import 'highcharts/modules/exporting'`, not as factories). With Vite, list them in `optimizeDeps.include`, as `demo/vite.config.ts` does. |

## Support the project

`highcharts-editable-excel` is free and MIT-licensed, and will stay that way. It is maintained in the open, and the work that remains (new chart types, Excel fidelity, testing against every Excel build, keeping up with Highcharts and Chart.js releases) is funded by sponsors and donations rather than by a paid tier.

If the library saves you from writing an OOXML chart writer, or if your product ships it, please consider sponsoring:

- **[GitHub Sponsors](https://github.com/sponsors/andresballenf)**: monthly or one-time; the preferred route.
- `npm fund highcharts-editable-excel` prints the same link from the command line.

Sponsors at a company tier are listed in this README with a logo and link, get priority on bug reports and compatibility requests, and can ask for a chart type or option to move up the fidelity backlog. Open an issue or a discussion to talk about anything else, including support agreements.

Other ways to help: open workbooks from `docs/manual-qa.md` in your copy of Microsoft Excel and report what you see, star the repository, and tell people who export charts for a living that this exists. See [CONTRIBUTING.md](CONTRIBUTING.md).

## Development

```bash
pnpm install
pnpm typecheck     # tsc on src, tests, e2e, demo (incl. demo/vite.config.ts), scripts
pnpm lint          # biome check . (lint rules + formatting; biome.json)
pnpm lint:fix      # biome check --write . (safe fixes + formatting)
pnpm format        # biome format --write .
pnpm build         # tsup → dist/{index,internals,augment}.js (ESM) + .cjs (CJS) + .d.ts/.d.cts
pnpm test          # vitest (jsdom): unit + integration, writes tests/output/**
pnpm coverage      # same suite with v8 coverage; fails below lines 90 / functions 92 / branches 78 / statements 87
pnpm test:e2e      # Playwright against the Vite demo (set PW_CHROMIUM_EXECUTABLE to use a preinstalled Chromium)
pnpm demo          # Vite demo on http://127.0.0.1:4173
pnpm demo:build    # production build of the demo into dist-demo/
pnpm bench         # export timing benchmark (sets BENCH=1; prints a table, writes tests/output/bench.json)
pnpm pack-check    # build (skip with --no-build when dist/ exists), pack, run publint + attw on the tarball,
                   # install it (with highcharts) in a temp project, check every src/index.ts and src/internals.ts
                   # export from ESM and CJS, that the main entry does not export experimental names, that
                   # /augment typechecks exporting.editableExcel, and that Highcharts is not bundled
pnpm examples:typecheck  # tsc -p examples/tsconfig.json (examples against src/)
pnpm validate:xsd  # XSD-validate tests/output/{export/v13,writer}/*.xlsx (run pnpm test first; needs OOXML_SCHEMA_DIR + lxml)
pnpm validate:openxml  # Open XML SDK validation of the same workbooks (run pnpm test first; needs the .NET 8 SDK)
pnpm test:visual   # LibreOffice render visual regression vs tests/baselines/render (UPDATE_BASELINES=1 rewrites baselines)
pnpm check         # typecheck + build + test
pnpm changeset     # add a changeset (release note + semver bump) for your pull request
pnpm run version   # changeset version: apply pending changesets to package.json and CHANGELOG.md
pnpm release       # pnpm check, then changeset publish (publishes to npm and tags the release)
```

- **Render checks.** `tests/integration/render-libreoffice.test.ts` and `writer-render.test.ts` render workbooks with LibreOffice (`soffice`) and poppler (`pdftoppm`, `pdftotext`). They skip with a printed reason when the tools are missing, unless `REQUIRE_RENDER=1` is set, in which case they fail. CI installs `libreoffice-calc` and `poppler-utils` and sets it.
- **Visual regression.** `tests/integration/visual-regression.test.ts` (`pnpm test:visual`, also part of `pnpm test`) exports eight fixtures with `includeSourceData: false`, renders them with LibreOffice at 60 dpi, crops to the chart and compares with `tests/baselines/render/*.png` (pixelmatch threshold 0.1, at most 1.5% differing pixels). Mismatches write `tests/output/visual/<name>.actual.png` and `<name>.diff.png`. A missing baseline fails unless `UPDATE_BASELINES=1`, which writes missing baselines and rewrites mismatching ones; review them before committing. Baselines come from Ubuntu 24.04 with apt LibreOffice and the `fonts-inter`/`fonts-crosextra-carlito` fonts (as in CI): font substitution on another machine shows up as a mismatch. Same skip/`REQUIRE_RENDER=1` rules as the render checks.
- **Open XML SDK validation.** `tools/ooxml-validator` is a .NET 8 console tool on `DocumentFormat.OpenXml` 3.x that runs `OpenXmlValidator` (Office 2016) on each workbook and prints `file: part: path: description`; it exits 1 on any error. `pnpm validate:openxml` runs it on `tests/output/{export/v13,writer}/*.xlsx`; the CI job `openxml-validate` also covers `examples/workbooks/*.xlsx`. See `tools/ooxml-validator/README.md`. Valid for the SDK is still not proof that Excel opens the file.
- **OOXML schema validation (optional).** `scripts/validate-ooxml.py` validates every XML part of the given workbooks against the ECMA-376 schemas with Python `lxml`. The schemas are not shipped: download ECMA-376 Part 1/4 schemas from ecma-international.org and set `OOXML_SCHEMA_DIR` to a folder with `ISO-IEC29500-4_2016/*.xsd` and `ecma/fouth-edition/opc-*.xsd`. Exit codes: 0 valid, 1 schema errors (printed with file, part and line; only the known `xml:space` on `<t>` false positive is ignored), 2 missing schemas/lxml. `writer-render.test.ts` runs it on the writer fixtures when both are available and skips with a printed reason otherwise (also under `REQUIRE_RENDER=1`). Schema-valid is still not proof that Excel opens the file.
- **Benchmark.** `tests/integration/bench.test.ts` only runs when `BENCH=1` (`pnpm bench` sets it).
- **Publishing.** Releases use [Changesets](https://github.com/changesets/changesets) (`.changeset/config.json`: public access, base branch `main`). Every pull request that changes the package adds a changeset with `pnpm changeset`; `pnpm run version` turns the pending changesets into the version bump and the `CHANGELOG.md` section, and `pnpm release` runs `pnpm check` and `changeset publish`. `prepublishOnly` runs `pnpm check && pnpm pack-check`. Pushing a `v*` tag runs `.github/workflows/release.yml`: install, `pnpm check`, `pnpm pack-check`, then `pnpm publish --provenance --access public` with the `NPM_TOKEN` secret. `publishConfig` sets public access and provenance. See [CHANGELOG.md](CHANGELOG.md).
- **CI** (`.github/workflows/ci.yml`) runs on Node 22 and 24: typecheck, build, examples typecheck, demo build, coverage with thresholds and `REQUIRE_RENDER=1`, the visual regression (uploading `tests/output/visual/**` on failure), e2e and pack-check; a separate `openxml-validate` job generates the workbooks and runs the Open XML SDK validator, and uploads the generated workbooks, renders and e2e output.
- `.npmrc` sets `engine-strict=true`, so installs fail on Node < 22.

Repository layout:

| Folder | Contents |
| --- | --- |
| `src/api` | Public export functions (single, options-only, analyze, multi-chart). |
| `src/browser` | Download trigger and reference-image rendering. |
| `src/highcharts` | Chart/options → `ChartModel` extraction, guards, CSS resolution, menu install. |
| `src/core` | Chart-type resolution, data-sheet layout, theme overrides, model → `WorkbookSpec` translation. |
| `src/translators` | Pure translators: colors, fonts, dash/marker, number formats. |
| `src/excel` | `WorkbookSpec`/`ExcelWriter` interface and the OOXML writer (fflate zip). |
| `src/types` | The IR (`chart-model.ts`), diagnostics and public API types. |
| `src/utils` | Colors, dates, filenames/sheet names, text, units. |
| `tests` | `unit/` (per Highcharts version where relevant), `integration/`, `fixtures/`, `helpers/`, `setup/`. |
| `demo`, `e2e` | Vite demo and the Playwright tests that drive it. |
| `examples`, `docs`, `scripts` | Framework examples, manual QA and compatibility docs, pack-check and bench scripts. |

Architecture:

```
Highcharts chart / options
  → extractor (src/highcharts)            reads series, points, axes, styles; raises diagnostics
  → ChartModel IR (src/types/chart-model) plain JSON, library-neutral
  → hooks.transformModel, themeOverrides
  → translators (src/core, src/translators)  chart-type plan, data layout, cell references, styles
  → WorkbookSpec (src/excel/writer-interface)
  → OOXML writer (src/excel)              SpreadsheetML + DrawingML chart parts, zipped with fflate
  → Uint8Array → triggerDownload (src/browser)
```

**Adding a chart type:**

1. Add the `SeriesKind` to `src/types/chart-model.ts` and the mapping in `src/highcharts/extract-series.ts` (`DIRECT_KINDS`).
2. Map it to a plot group in `src/core/chart-type-registry.ts` (`KIND_TO_GROUP`, combination rules) and add a `CHART_TYPE_MATRIX` row.
3. If Excel needs a new plot-group kind, extend `PlotGroupSpec` in `src/excel/writer-interface.ts` and emit it in `src/excel/chart-xml.ts`, following the schema element order.
4. Handle it in `translateChartModel` and, if it has its own data shape, in `src/core/data-layout.ts`.
5. Add tests for every tested Highcharts version and update the matrix here and in `docs/compatibility.md`.

**Adding a translator:** write a pure function in `src/translators/` that returns its result plus an optional `Diagnostic` built with `createDiagnostic`. Call it from the extractor or `src/core/style-mapping.ts`, and add a new code to `DiagnosticCode` in `src/types/diagnostics.ts` if needed. Data-related codes also go into the `dataConcerns` set there.

**Writer decision.** `@office-kit/xlsx` 0.24.1 was evaluated first. It writes native charts, but it had no combo charts and no secondary axes, a broken date axis, nulls cached as the text `"null"`, and out-of-order `logBase`. The library therefore ships its own narrowly scoped OOXML chart writer on `fflate`, behind the `ExcelWriter` interface. `@office-kit/xlsx` remains a dev-only dependency used to round-trip-validate the output in tests.

## License

[MIT](LICENSE). The bundled zip library (`fflate`) is also MIT; see [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md).

Highcharts is a trademark of Highsoft AS and Chart.js is maintained by the Chart.js contributors. This project is an independent, community-maintained library and is not affiliated with or endorsed by either. Using Highcharts itself requires a license from Highsoft for commercial use; this library does not change that.
