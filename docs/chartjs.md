# Chart.js adapter

`highcharts-editable-excel/chartjs` exports [Chart.js](https://www.chartjs.org/) 4 charts to an
`.xlsx` workbook holding a **native Excel chart** whose series reference cells on a data sheet.
It is the same pipeline as the Highcharts entry points: the adapter turns a Chart.js chart into the
library's source-independent `ChartModel`, and the shared translator and OOXML writer do the rest.

Chart.js is an optional peer dependency. The adapter never imports `chart.js`: it reads a live
chart (or a configuration object) through its public shape, so it works with `chart.js/auto` and
with tree-shaken registrations alike.

## Usage

```ts
import { Chart } from 'chart.js/auto';
import {
  exportChartJsToXlsx,
  downloadChartJsAsXlsx,
  analyzeChartJsCompatibility,
  extractChartJsModel,
} from 'highcharts-editable-excel/chartjs';

const chart = new Chart(canvas, config);

// Browser: export and download in one call.
await downloadChartJsAsXlsx(chart, { filename: 'sales' });

// Anywhere: bytes only (a plain configuration needs no canvas, so this also runs in Node).
const { bytes, report, warnings } = await exportChartJsToXlsx(config);

// Dry run: is it editable, and what is approximated?
const compatibility = analyzeChartJsCompatibility(chart);

// The intermediate model, for debugging or custom pipelines.
const model = extractChartJsModel(chart, { seriesVisibility: 'all' });
```

All `ExportOptions` of the Highcharts API apply (`fidelity`, `dataMode`, `seriesVisibility`,
`strictMode`, `themeOverrides`, `hooks.transformModel`, `chartWidth`/`chartHeight`,
`includeSourceData`, `writer`, `signal`, `onProgress`, `onWarning`, `includeModel`). Errors are
`ExportError`s with the same codes (`INVALID_CHART`, `CHART_NOT_EDITABLE`, `INVALID_OPTIONS`,
`WRITER_FAILURE`, `BROWSER_REQUIRED`, `ABORTED`). The input is never mutated.

### Live chart or configuration?

| Input | What is read |
| --- | --- |
| `Chart` instance | `chart.config` (Chart.js's merged scale config), `chart.scales`, `chart.getDatasetMeta(i)` (parsed values, resolved element options), `chart.isDatasetVisible(i)` / `getDataVisibility(i)` (legend toggles), `chart.width/height`, `chart.chartArea`, `Chart.defaults`, `Chart.version`. Scriptable options are already evaluated by Chart.js and are exported as rendered. |
| `{ type, data, options }` | The configuration, with Chart.js 4 defaults filled in (Helvetica Neue 12px `#666`, element colors `rgba(0,0,0,0.1)`, the `colors` plugin palette, controller and scale defaults). Scriptable (function) options cannot be evaluated and are reported. Size: `chartWidth` or 600px wide, height from `options.aspectRatio` (2, or 1 for pie/doughnut). |

## Mapping

| Chart.js | Excel | Notes |
| --- | --- | --- |
| `line` | line | `tension > 0` or `cubicInterpolationMode: 'monotone'` → smoothed line. |
| `line` + `fill` | area | `true`/`'origin'`/`'start'` (and `'-1'`/`'stack'` on a stacked scale) are exact; other targets are drawn to the axis (`APPROXIMATED_CHART_TYPE`). Smoothed areas are drawn straight (Excel limitation). |
| `bar` | column | `indexAxis: 'y'` → horizontal bar. `categoryPercentage`/`barPercentage` → gap width/overlap. |
| stacked scales (`stacked: true`) | stacked column/line/area | `stack` names become stack groups (Excel merges groups, reported). |
| `pie` | pie | Per-slice colors, `offset` → explosion, `rotation`/`circumference` → angles. Hidden slices (legend) are reported. |
| `doughnut` | doughnut | `cutout` `'NN%'` → hole size; a pixel cutout is converted against the chart size (`APPROXIMATED_LAYOUT`). Rings keep Chart.js order (first dataset outside). |
| `scatter` | scatter | `{x, y}` data; `showLine` → connecting line. |
| `bubble` | bubble | `{x, y, r}` / `[x, y, r]`; `r` (pixels in Chart.js) becomes the bubble size value. |
| mixed (`dataset.type`) | combo | Bar + line on one or two value axes. |
| `radar`, `polarArea` | — | Not editable: `UNSUPPORTED_SERIES_TYPE` + `UNSUPPORTED_POLAR` (blocking). |
| `labels` | category column | Multi-line labels are joined with spaces. |
| `parsing: { xAxisKey, yAxisKey, key }` | — | Honoured, including dotted paths. |
| `xAxisID` / `yAxisID` | axis binding | `position: 'right'` / `'top'` → opposite axis; more than two value axes collapse onto the secondary axis. |
| `type: 'time'` / `'timeseries'` | date axis | ISO strings, `Date`s and ms numbers. Zone-less date strings are kept as written; instants are moved to the browser's local wall-clock time (what the date adapters show). `time.unit` + `displayFormats` (date-fns tokens) → Excel date format. |
| `type: 'logarithmic'` | log axis (base 10) | |
| `min`, `max`, `suggestedMin/Max`, `beginAtZero`, `reverse`, `ticks.stepSize` | axis bounds, orientation, major unit | `suggested*` and `beginAtZero` apply only when they extend the data range, as in Chart.js. |
| `ticks.format` (Intl options) | number format | `percent` → `0%`, `currency` → `[$€]#,##0.00`, fraction digits, grouping, scientific; `compact`/`unit` approximated. |
| `title`, `grid`, `border`, `ticks.font/color` | axis title, gridlines, axis line, label font | |
| `backgroundColor`, `borderColor` | fill, line/border | Scalars, per-point arrays (point colors), rgba alpha kept. |
| `borderWidth`, `borderDash` | line width, nearest preset dash | |
| `pointStyle`, `pointRadius`, `pointBackgroundColor`… | markers | circle, rect → square, rectRot → diamond, triangle; others → circle (`APPROXIMATED_MARKER`). |
| `hidden` | — | Skipped (`HIDDEN_SERIES_EXCLUDED`) unless `seriesVisibility: 'all'`. |
| `plugins.title` / `plugins.subtitle` | chart title (two lines) | Only when `display: true`. |
| `plugins.legend` | legend | `display`, `position`, `reverse`, `labels.font/color`. |
| background | chart area fill | Chart.js has none: `plugins.customCanvasBackgroundColor.color` (the documented sample plugin), else the canvas's computed background in a real browser, else white. |
| decimation plugin | — | `dataMode: 'rendered'` exports the decimated points (`DATA_GROUPED`); `'raw'` the original data. |

Diagnostic paths point into the Chart.js configuration, e.g. `data.datasets[1].borderDash`,
`options.scales.y.ticks.callback`, `options.plugins.legend.position`.

## Limitations

- **Callbacks are never executed**: `ticks.callback` (`UNSUPPORTED_FORMATTER`), tooltip callbacks
  (`UNSUPPORTED_TOOLTIP`), legend `labels.filter`/`generateLabels`. In a configuration, scriptable
  colors/widths fall back to the next defined value (`UNSUPPORTED_STYLE`); a live chart's rendered
  values are used instead.
- Canvas gradients and patterns are not exported (`UNSUPPORTED_GRADIENT`).
- Floating bars (`[start, end]`) leave their cells empty (`NON_NUMERIC_VALUE`).
- `stepped` lines, `spanGaps`, `grace`, fixed `barThickness`, centered axes, `chartArea` legends and
  non-centered legend/title alignment are approximated and reported.
- `chartjs-plugin-datalabels` and `chartjs-plugin-annotation` options are reported, not exported.
- Time parsing without the chart's date adapter: custom `time.parser` formats are not run, and the
  adapter `zone` option is ignored (reported). Live charts use the adapter-parsed values.
- The reference image (`includeReferenceImage`) uses the chart's own canvas, so it needs a live
  chart in a real browser.
- Validation: the adapter is covered by configuration tests, live `Chart` instances rendered in
  jsdom with a stub 2D context (no time scales: no date adapter is installed in the test
  environment), structural XLSX inspection and LibreOffice renders. **None of this is Microsoft
  Excel**; see `docs/manual-qa.md`.
