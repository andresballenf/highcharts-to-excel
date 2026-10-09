# highcharts-editable-excel demo

An interactive gallery of Highcharts charts. Every chart can be downloaded as an `.xlsx`
workbook containing a **native, editable Excel chart** whose series point at cells on a data sheet.

The demo imports `highcharts-editable-excel`, which `demo/vite.config.ts` aliases to
`../src/index.ts`: it always runs the library **source**, so edits in `src/` hot-reload here.
Only `highcharts`, `highcharts/modules/exporting` and `highcharts/highcharts-more` are loaded:
the export-data module is not needed.

## Run it

```sh
pnpm install
pnpm demo          # http://127.0.0.1:4173
pnpm demo:build    # static build into dist-demo/
```

## What is on the page

`installHighchartsExcelExport(Highcharts, { onExport, onError })` is called once at startup, so
every chart's ☰ menu has a **Download editable Excel chart** item next to Highcharts' own items.
Each card also has:

- **Export to Excel** — `downloadHighchartsAsXlsx(chart, { filename: '<name>.xlsx', onWarning })`.
- **Analyze** — `analyzeChartCompatibility(chart)`: a dry run, nothing is downloaded.
- **Warnings** panel — the diagnostics of the last export/analysis as `code — property — message`
  lines, plus the full JSON compatibility report (editable?, Excel chart type, which properties
  were supported, approximated, unsupported or blocking) and the timings.

| Card | Shows |
| --- | --- |
| `line`, `multi-line`, `column`, `area`, `stacked`, `combo`, `secondary` | Category charts: line, clustered/stacked columns, area, column+line+spline combo, a secondary (opposite) value axis. |
| `bar` | Per-chart menu text: `exporting: { editableExcel: { menuText: 'Save as Excel chart' } }`. |
| `pie` | Its own `exporting.buttons.contextButton.menuItems` with an app item `myCustomItem`; the Excel item is added and the custom item survives. |
| `doughnut`, `scatter`, `datetime` | Pie with `innerSize`, x/y scatter, a datetime x axis (exported as Excel dates). |
| `custom` | Georgia font, custom palette, dashed orange gridlines, data labels, vertical legend on the right, 800×450 px, plot background. |
| `dynamic` | **Add point** (`series.addPoint`) and **Randomize** (`setData` with a deterministic sequence): exports contain what the chart currently shows. |
| `styled` | `chart.styledMode: true`: colors come from CSS. |

**Export all charts to one workbook** uses `exportChartsToWorkbook` + `triggerDownload`: one
chart sheet and one data sheet per chart.

`window.__demo = { charts, lastResults }` exposes the chart instances and the last result per card
(used by the Playwright tests in `e2e/`, handy in the devtools console too).

## Inspecting warnings

Open a card's **Warnings** panel after exporting or analyzing. Every export also logs each
diagnostic with `console.debug` (enable "Verbose" in the devtools console). Codes are stable
(`APPROXIMATED_LAYOUT`, `MIXED_SERIES_TYPES`, `STYLED_MODE_FALLBACK`, …); see
`src/types/diagnostics.ts`.

## Styled mode and CSS variables

In styled mode Highcharts sets no colors itself; they come from `highcharts/css/highcharts.css`
and your own CSS. That stylesheet is global (`.highcharts-*` rules would also restyle the
non-styled charts on the page), so `main.ts` nests it under the styled card's
`.hc-styled-scope` class. The `<style id="styled-mode-theme">` block in `index.html` then defines
`--highcharts-color-0/1/2` and `.highcharts-background` etc. for `#chart-styled`.

The exporter reads the **browser-computed** styles, so the Excel series use exactly those CSS
colors. Without a browser (e.g. `exportHighchartsOptionsToXlsx` on a server) it falls back to
Highcharts' default palette and reports `STYLED_MODE_FALLBACK`.

All data is deterministic and animation is off, so exports and screenshots are reproducible.
