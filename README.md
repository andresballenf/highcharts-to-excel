# highcharts-editable-excel

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

## Install

```bash
pnpm add highcharts-editable-excel
# or
npm install highcharts-editable-excel
# or
yarn add highcharts-editable-excel
```

- Peer dependency: `highcharts >= 11`. The tested versions are 12.6.2 and 13.1.1 (see [Compatibility matrix](#compatibility-matrix)).
- The context-menu integration needs the Highcharts **exporting** module (`highcharts/modules/exporting`). The programmatic API does not need it, except for the optional reference image.
- The `export-data` module is **not** required.
- The package is ESM only (`"type": "module"`) and has no side effects on import. `dist/index.js` is about 253 KB minified, about 69 KB gzipped.

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

**TypeScript**: [`examples/typescript/main.ts`](examples/typescript/main.ts). This file is typechecked in this repository (`npx tsc -p examples/tsconfig.json`) and contains every plain-TypeScript snippet in this README. The React and Angular snippets below are shortened from their example files, which are typechecked too. The `.vue` file is not.

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
| `installHighchartsExcelExport(Highcharts: unknown, options?: InstallOptions)` | `Installation` | `INVALID_OPTIONS` if the argument is not the Highcharts namespace; `EXPORTING_MODULE_MISSING` if the exporting module is not loaded |
| `addEditableExcelMenuItem(chart: unknown, options?: InstallOptions)` | `void` | `INVALID_CHART` if the argument is not a rendered chart with `update()` |
| `exportHighchartsToXlsx(chart: unknown, options?: ExportOptions)` | `Promise<ExportResult>` | `INVALID_CHART`, `CHART_NOT_EDITABLE`, `INVALID_OPTIONS` (bad `transformModel` result), `WRITER_FAILURE` |
| `downloadHighchartsAsXlsx(chart: unknown, options?: ExportOptions)` | `Promise<ExportResult>` | Same as above, plus `BROWSER_REQUIRED` |
| `triggerDownload(bytes: Uint8Array, filename: string, mimeType = XLSX_MIME_TYPE)` | `void` | `BROWSER_REQUIRED` if `document`, `Blob` or `URL.createObjectURL` is missing |
| `analyzeChartCompatibility(chart: unknown, options?: ExportOptions)` | `CompatibilityReport` (synchronous) | `INVALID_CHART` only; a blocked chart returns `editable: false` |
| `exportHighchartsOptionsToXlsx(highchartsOptions: object, options?: ExportOptions)` | `Promise<ExportResult>` | `INVALID_OPTIONS` (not an object), `CHART_NOT_EDITABLE`, `WRITER_FAILURE` |
| `exportChartsToWorkbook(entries: MultiChartExportEntry[], options?: { filename?: string; properties?: { title?: string; creator?: string }; strictMode?: boolean })` | `Promise<MultiChartExportResult>` | `INVALID_OPTIONS` (empty or non-array input, non-object entry), `INVALID_CHART` and `CHART_NOT_EDITABLE` (the message names the entry index), `WRITER_FAILURE` |

Notes:

- `installHighchartsExcelExport` is **idempotent per Highcharts namespace**. A second call returns the first `Installation`, and the first call's options win. `installation.uninstall()` removes the item definition and restores the global menu. Charts created after that are unaffected, and a new install is then allowed.
- `addEditableExcelMenuItem` patches an already-rendered chart through `chart.update()`. Use it for charts created before the install. It is a no-op when the chart already lists the item or has `exporting.editableExcel.enabled === false`.
- `analyzeChartCompatibility` runs the same extraction and translation as an export but writes nothing and never renders the reference image.
- `exportChartsToWorkbook` names the sheets `Chart 1`/`Data 1`, `Chart 2`/`Data 2` and so on unless an entry sets `chartSheetName`/`dataSheetName`. Names that collide are de-duplicated with ` (2)`, ` (3)` and reported as `SHEET_NAME_ADJUSTED`. The default filename is `charts.xlsx`. The workbook title comes from the first chart.

Lower-level building blocks are also exported: `CHART_TYPE_MATRIX`, `resolveChartType(model)`, `translateChartModel(model, options)`, `extractChartModel(chart, options)`, `extractChartModelFromOptions(options, extractOptions)`, `createDefaultExcelWriter()`, and the types `ChartModel`, `WorkbookSpec`, `ExcelWriter` and related. See [Custom writer](#custom-writer-excelwriter) for the current limits.

Constants: `XLSX_MIME_TYPE`, `DEFAULT_MENU_TEXT` (`"Download editable Excel chart"`), `DEFAULT_MENU_ITEM_KEY` (`"downloadEditableXLSX"`).

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
| `chartWidth`, `chartHeight` | `number` (CSS px) | Rendered size (`chart.chartWidth`/`chartHeight`; options path: `chart.width`/`height`, else 600×400) | Size of the Excel chart object (minimum 50). |
| `strictMode` | `boolean` | `false` | Also throw `CHART_NOT_EDITABLE` when any diagnostic has outcome `unsupported` at `warning` or `error` severity (for example annotations, plot bands, formatter callbacks). Blocking diagnostics always throw, with or without strict mode. Reference-image diagnostics are ignored. |
| `includeReferenceImage` | `boolean` | `false` | Opt-in. Embeds a PNG of the Highcharts rendering to the right of the native chart for side-by-side comparison. Needs a real browser and `chart.getSVG` (exporting module). Otherwise the export continues without it and reports `WRITER_LIMITATION`. |
| `includeModel` | `boolean` | `false` | Adds the normalized `ChartModel` to `result.model` (debugging). |
| `hooks` | `{ transformModel?: (model: ChartModel) => ChartModel }` | none | Called after extraction, before theme overrides and translation. Must return a model with a `series` array (otherwise `INVALID_OPTIONS`). Never mutate the Highcharts chart in it. |
| `properties` | `{ title?: string; creator?: string }` | Title is the chart title. Creator is `"highcharts-editable-excel"`. | Workbook document properties. |

### `ThemeOverrides`

| Field | Type | Effect |
| --- | --- | --- |
| `colors` | `string[]` | Palette by series position. Pie and doughnut use it per slice. |
| `chartBackground`, `plotBackground` | `string` | Solid fills. |
| `fontFamily` | `string` | Family for every text element. |
| `title`, `subtitle`, `axisTitle`, `axisLabels`, `legend`, `dataLabels` | `FontOverride` (`{ family?, size? (CSS px), bold?, italic?, color? }`) | Per-element font overrides. |
| `gridLineColor`, `gridLineWidth` | `string`, `number` | Applied to existing grid lines. A width > 0 also adds y-axis grid lines where there were none. |
| `series` | `Record<number, { color?, lineWidth?, fillOpacity? }>` | Per-series overrides keyed by series index. |

Colors that cannot be parsed are ignored and reported as `UNRESOLVED_COLOR`.

### `InstallOptions`

| Option | Type | Default | Semantics |
| --- | --- | --- | --- |
| `menuText` | `string` | `"Download editable Excel chart"` | Menu item text. |
| `menuItemKey` | `string` | `"downloadEditableXLSX"` | Key in `exporting.menuItemDefinitions` and `menuItems`. |
| `exportOptions` | `ExportOptions` | `{}` | Defaults for every menu export. |
| `onExport` | `(result: ExportResult, chart: unknown) => void` | none | Called after a successful menu download. |
| `onError` | `(error: unknown, chart: unknown) => void` | `console.error(...)` | Called when a menu export fails. |
| `insertAfter` | `string` | Last of `downloadXLS`/`downloadCSV`/`downloadSVG` present, else the end | Menu key to insert the item after. |

### Per-chart `exporting.editableExcel`

Set in a chart's options. It is read when the chart is created (menu items) and again when the menu item is clicked (export options):

```js
Highcharts.chart('container', {
  exporting: {
    editableExcel: {
      enabled: true,                       // false: no menu item on this chart
      menuText: 'Excel (editable chart)',  // this chart only
      filename: 'sales-2026',              // ...plus any ExportOptions; they override install exportOptions
      dataSheetName: 'Sales data',
    },
  },
});
```

If the chart sets its own `exporting.buttons.contextButton.menuItems`, the item is appended to that list (the caller's array is not mutated). TypeScript users can declare the option with a module augmentation. The library does not ship one; see [`examples/typescript/main.ts`](examples/typescript/main.ts):

```ts
declare module 'highcharts' {
  interface ExportingOptions {
    editableExcel?: PerChartExportConfig;
  }
}
```

### Result shapes

```ts
interface ExportResult {
  bytes: Uint8Array;
  filename: string;
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  warnings: Diagnostic[];
  report: CompatibilityReport;
  timings: { extractMs: number; translateMs: number; writeMs: number; totalMs: number };
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

`class ExportError extends Error` with `name: 'HighchartsExcelExportError'`, `code: ExportErrorCode` and `details: { chartId?, property?, diagnostics?, cause? }`.

| Code | Thrown when |
| --- | --- |
| `INVALID_CHART` | The argument is not a rendered chart (it needs `series`, `xAxis` and `yAxis` arrays, `options`, and `renderTo` or `container`). |
| `INVALID_OPTIONS` | `installHighchartsExcelExport` did not get a Highcharts namespace; `exportHighchartsOptionsToXlsx` did not get an object; `hooks.transformModel` returned something that is not a model; `exportChartsToWorkbook` got an empty or invalid list. |
| `CHART_NOT_EDITABLE` | Any **blocking** diagnostic: polar chart, no series or no data, only unsupported series types, pie or bubble combined with other types, or worksheet row or column limit exceeded. In strict mode, also any `unsupported` warning. `details.diagnostics` lists the causes. |
| `EXPORTING_MODULE_MISSING` | `installHighchartsExcelExport` was called before `highcharts/modules/exporting` was loaded. |
| `BROWSER_REQUIRED` | `triggerDownload` or `downloadHighchartsAsXlsx` ran outside a browser. |
| `WRITER_FAILURE` | The XLSX writer refused the workbook, for example a chart series with more than 32,000 points or a plot group with more than 255 series. `details.cause` holds the original error. |
| `UNSUPPORTED_CHART_TYPE`, `ROW_LIMIT_EXCEEDED`, `COLUMN_LIMIT_EXCEEDED` | Reserved in the `ExportErrorCode` type but not thrown today. These conditions surface as **diagnostic** codes inside a `CHART_NOT_EDITABLE` error. |

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
| columnrange | - | unsupported | Excel has no floating range columns without helper series. |
| arearange | - | unsupported | Excel has no band/range area type. |
| boxplot | - | unsupported | Excel box & whisker is a chartex type that cannot reference this layout. |
| heatmap | - | unsupported | Excel has no heatmap chart (only conditional formatting). |
| treemap | - | unsupported | Excel treemap is a chartex type not produced by this library. |
| waterfall | - | unsupported | Excel waterfall is a chartex type not produced by this library. |
| funnel | - | unsupported | Excel funnel is a chartex type not produced by this library. |
| gauge | - | unsupported | Excel has no gauge chart. |
| polar (any type) | - | unsupported | Excel radar charts do not match polar/spider geometry. |
| variablepie | - | unsupported | Excel pies cannot vary slice radius. |
| sankey | - | unsupported | Excel has no flow diagrams. |
| networkgraph | - | unsupported | Excel has no network/graph layout chart. |
| timeline | - | unsupported | Excel has no timeline chart type. |
| histogram | - | unsupported | Excel histogram is a chartex type computed from raw data, not from this series. |
| bellcurve | - | unsupported | Derived series (computed in the browser) with no Excel equivalent. |
| errorbar | - | unsupported | Excel error bars are attached to another series, not standalone series. |
| lollipop | - | unsupported | No Excel equivalent without helper series and error bars. |
| dumbbell | - | unsupported | No Excel equivalent without helper series and high-low lines. |

Combination rules (`resolveChartType`):

- Column, bar, line, area and scatter series combine into one Excel **combo chart**, for example `combo:column+line`, reported as `MIXED_SERIES_TYPES` (info).
- Pie, or bubble, combined with any other type **blocks** the export.
- Several plain pie series: only the first is drawn and the others stay on the data sheet. A pie together with a doughnut ring is drawn as doughnut rings.
- Excel has at most two value axes. The lowest y-axis index in use is primary, and every other one is collapsed onto the secondary axis (`UNSUPPORTED_AXIS_FEATURE`).
- Series of one type on one axis form one stack. Separate `stack` groups are merged (`APPROXIMATED_CHART_TYPE`).

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
| Styled mode | approximated | In a real browser the rendered SVG colors are read. Headless, the palette by color index is used (`STYLED_MODE_FALLBACK`). |
| Data grouping (Stock) | approximated | Rendered mode exports the grouped points (`DATA_GROUPED`, info). Raw mode exports source data. |
| Zoom / cropping | approximated | The zoomed range becomes fixed axis bounds (`APPROXIMATED_AXIS_SCALE`, info). Cropped points: `DATA_CROPPED` (info). |
| `negativeColor`, `zones` | approximated | The series color is used (`UNSUPPORTED_STYLE`). |
| Rounded bars (`borderRadius`) | approximated | Square corners (`UNSUPPORTED_STYLE`, info, only when set explicitly). |
| Inverted line/area/scatter, semi-circle pie | approximated | `APPROXIMATED_CHART_TYPE`. |
| 3D (`options3d`) | ignored | Exported flat (`UNSUPPORTED_3D`). |
| Polar / spider | unsupported (blocking) | `UNSUPPORTED_POLAR`. |
| Annotations | unsupported | `UNSUPPORTED_ANNOTATION`. |
| Plot bands / plot lines | unsupported | `UNSUPPORTED_PLOT_BAND`. |
| Tooltips | unsupported | Never exported. Custom formatters are reported (`UNSUPPORTED_TOOLTIP`, info). `tooltip.valueDecimals`/`valuePrefix`/`valueSuffix` become the data cells' number format. |

Per-option detail (option path, outcome, diagnostic) is in [docs/compatibility.md](docs/compatibility.md).

### Tested versions and environments

- **Highcharts 12.6.2 and 13.1.1.** The extraction, menu-install, export-fixture and compatibility-matrix suites run once per version (`*.v12.test.ts` / `*.v13.test.ts`). The other suites that use Highcharts run on 13.1.1. Other versions within the `>= 11` peer range are untested.
- **Node 22 + jsdom** (vitest). The `.xlsx` packages are inspected structurally, and the writer output round-trips through `@office-kit/xlsx`'s strict chart parser.
- **Chromium** via Playwright (demo end-to-end tests).
- **LibreOffice** headless render to PDF as a smoke check, when `soffice` is on the PATH.

**The generated workbooks were not opened in Microsoft Excel in the development environment.** Structural validation and LibreOffice rendering do not replace a check in Excel. Follow [docs/manual-qa.md](docs/manual-qa.md) before relying on a chart type in production.

## Fidelity limitations

Excel's chart engine is not Highcharts. Expect these differences:

- **Fonts.** The first usable family of the CSS stack is used. The Highcharts default stack maps to `Segoe UI`. Generic families map to `Arial`, `Times New Roman` and so on (`APPROXIMATED_FONT`). Excel substitutes fonts that are not installed. Sizes convert at 1 px = 0.75 pt.
- **One title.** Excel charts have no subtitle, so the subtitle is merged into the title as a second line, in the title font.
- **No rounded bars, no zones, no negative colors.**
- **Data-label positions are restricted per chart type:**
  - clustered bars: center, inside end, inside base, outside end
  - stacked bars: no outside end
  - line, scatter and bubble: center, left, right, above, below
  - pie: center, inside end, outside end, best fit
  - area and doughnut: no positioning at all
- **Legend** positions are limited to top, bottom, left, right and top-right.
- **Plot area.** The plot area is pinned to the source proportions (`APPROXIMATED_LAYOUT`, info). Excel still positions titles, labels and the legend around it itself, so spacing differs slightly.
- **No interactivity:** tooltips, drilldown, events, animation and the navigator are dropped.
- **Spline smoothing** uses Excel's algorithm. Areasplines are straight.
- **Area opacity.** Highcharts' `fillOpacity` (default 0.75) becomes fill transparency. As in Highcharts, it is not applied to an explicit `fillColor`.
- **Styled mode.** Colors live in CSS, and only a real browser can resolve them. Headless or server-side exports fall back to the palette plus `themeOverrides`.

## Data semantics

- **`dataMode: 'rendered'` (default)** exports the points the chart currently shows: after `setData`/`update`, after data grouping and after cropping. Grouped series report `DATA_GROUPED` and cropped series `DATA_CROPPED`. Both are info diagnostics with source and rendered point counts in `details`.
- **`dataMode: 'raw'`** parses `series.options.data` the way Highcharts does: numbers, `null`, `[x, y]`, `[x, y, z]`, `[name, y]` and point objects. Missing x values follow `pointStart`/`pointInterval`/`pointIntervalUnit`, with calendar units computed in UTC. If the source data is not available (for example data loaded through a data table), the rendered points are used (`DATA_MODE_FALLBACK`). The options-object API always uses raw data.
- **Hidden series:** see `seriesVisibility`. A hidden pie slice is exported and shown (`HIDDEN_POINT`).
- **Nulls vs missing points.** A `null` y becomes an empty cell, drawn as a gap (`NULL_VALUES`, info). A series without a value for a category another series has also gets an empty cell (`UNALIGNED_X_VALUES`). Values that are not numeric are skipped (`NON_NUMERIC_VALUE`).
- **Zoom.** A zoomed axis (`userMin`/`userMax`) is exported as fixed axis bounds for the visible window.
- **Dates** are written as Excel serial numbers computed from the UTC timestamp, with no time-zone shift, and formatted from the axis label format when it can be translated.
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

The `DiagnosticCode` type also declares `UNSUPPORTED_AXIS_TYPE`, `APPROXIMATED_COLOR`, `APPROXIMATED_FONT_SIZE`, `APPROXIMATED_DASH_STYLE`, `UNSUPPORTED_INTERACTIVITY`, `EXPORTING_MODULE_MISSING` and `HEADLESS_STYLE_FALLBACK`. They are reserved: no code path raises them as a diagnostic today. A missing exporting module is reported as an `ExportError`.

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

### Custom writer (`ExcelWriter`)

The `ExcelWriter` interface (`{ name: string; write(workbook: WorkbookSpec): Promise<Uint8Array> }`) and the `WorkbookSpec` types are exported for advanced users. `createDefaultExcelWriter()` returns the built-in OOXML writer, and you can call it with a `WorkbookSpec` you build yourself (see `writeCustomWorkbook` in the TypeScript example).

The high-level export functions always use the built-in writer: there is currently **no option to inject another writer**. `extractChartModel` and `translateChartModel` are exported, but their options require a `DiagnosticCollector` instance that is not exported. Treat them as internal for now.

## Security and privacy

- **Local only.** The workbook is generated in memory, and downloads use a temporary `blob:` URL that is revoked afterwards. No data leaves the page or process.
- **No formula injection.** Every text cell is written as an inline string (`t="inlineStr"`), never as a formula. A category such as `=SUM(A1)` or `@cmd` stays literal text, and the first such value is reported as `FORMULA_LIKE_TEXT_ESCAPED` (info). Worksheet cells contain no formulas at all. Only the chart part has series references (`c:f`) to the data sheet.
- **No macros, no external links.** The package contains only workbook, worksheet, styles, theme, drawing, chart, PNG (reference image only) and document-property parts.
- **XML safety.** All text and attributes are XML-escaped. Characters not allowed in XML 1.0 are stripped. Cell text is cut at Excel's 32,767-character limit.
- **Excel limits.**
  - More than 1,048,576 rows or 16,384 columns blocks the export (`ROW_LIMIT_EXCEEDED` / `COLUMN_LIMIT_EXCEEDED` inside `CHART_NOT_EDITABLE`).
  - Excel allows at most 32,000 points per chart series. A longer series is reported as `ROW_LIMIT_EXCEEDED` (approximated), but the writer then refuses it, so the export fails with `WRITER_FAILURE`. To stay under the limit, export grouped data (Highcharts Stock, `dataMode: 'rendered'`) or reduce the points first.

## Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| `ExportError: EXPORTING_MODULE_MISSING` | Load `highcharts/modules/exporting` (or `modules/exporting.js` from the CDN) **before** calling `installHighchartsExcelExport`. |
| The menu item does not show | Charts created *before* the install, or charts whose `menuItems` were replaced later with `chart.update()`, are not patched. Call `addEditableExcelMenuItem(chart)`. Also check `exporting.editableExcel.enabled !== false` and that `exporting.enabled` is not `false`. |
| `CHART_NOT_EDITABLE` | Inspect `error.details.diagnostics` or run `analyzeChartCompatibility(chart)`. The usual causes are polar charts, only unsupported series types, pie or bubble mixed with other types, an empty chart, or `strictMode` with an unsupported feature. |
| Nothing downloads / `BROWSER_REQUIRED` | `downloadHighchartsAsXlsx` needs `document` and `URL.createObjectURL`. In Node or a worker, use `exportHighchartsToXlsx` or `exportHighchartsOptionsToXlsx` and handle `bytes` yourself. Browsers can also block downloads that are not triggered by a user gesture, or in sandboxed iframes without `allow-downloads`. Trigger the export from a click. |
| Wrong colors in styled mode on the server | CSS is not readable without a real browser (`STYLED_MODE_FALLBACK`). Pass `themeOverrides.colors` (and backgrounds or fonts). |
| Dates shifted by some hours | Timestamps are written as UTC serials. `chart.time` time-zone settings are not applied to the cells. Export timestamps at UTC midnight for whole days, or adjust the data before exporting. |
| Large charts are slow | All points are written to the sheet and to the chart cache. Use `dataMode: 'rendered'` with data grouping, or fewer points. `result.timings` shows where the time goes. For measurements, see `pnpm bench`. |
| Vite / CJS errors importing Highcharts modules | Highcharts ships UMD/CJS files without an exports map. Import modules as side effects (`import 'highcharts/modules/exporting'`, not as factories). With Vite, list them in `optimizeDeps.include`, as `demo/vite.config.ts` does. |

## Development

```bash
pnpm install
pnpm typecheck     # tsc on src, tests, e2e, demo, scripts
pnpm build         # tsup → dist/index.js (ESM) + dist/index.d.ts
pnpm test          # vitest (jsdom): unit + integration, writes tests/output/**
pnpm test:e2e      # Playwright against the Vite demo (set PW_CHROMIUM_EXECUTABLE to use a preinstalled Chromium)
pnpm demo          # Vite demo on http://127.0.0.1:4173
pnpm bench         # export timing benchmark
pnpm pack-check    # pack, install the tarball in a temp project, verify exports and that Highcharts is not bundled
```

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
5. Add tests for both Highcharts versions and update the matrix here and in `docs/compatibility.md`.

**Adding a translator:** write a pure function in `src/translators/` that returns its result plus an optional `Diagnostic` built with `createDiagnostic`. Call it from the extractor or `src/core/style-mapping.ts`, and add a new code to `DiagnosticCode` in `src/types/diagnostics.ts` if needed. Data-related codes also go into the `dataConcerns` set there.

**Writer decision.** `@office-kit/xlsx` 0.24.1 was evaluated first. It writes native charts, but it had no combo charts and no secondary axes, a broken date axis, nulls cached as the text `"null"`, and out-of-order `logBase`. The library therefore ships its own narrowly scoped OOXML chart writer on `fflate`, behind the `ExcelWriter` interface. `@office-kit/xlsx` remains a dev-only dependency used to round-trip-validate the output in tests.

## License

MIT, see [LICENSE](LICENSE).
