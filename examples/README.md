# Examples

Small, complete integrations of `highcharts-editable-excel`. Each file imports from `highcharts-editable-excel`, except the plain-JavaScript page, which imports the built bundle `../../dist/index.js`. Highcharts and its exporting module are loaded the usual way for each setup.

| File | Setup | Needs in the host app |
| --- | --- | --- |
| `javascript/index.html` | Script tags (Highcharts CDN) + ES module | `pnpm build` first, then serve the **repository root** over HTTP (for example `npx vite` or `python3 -m http.server`) and open `/examples/javascript/`. ES modules do not load from `file://`. |
| `typescript/main.ts` | Plain TypeScript | `highcharts`. Contains every TypeScript snippet from the main README. |
| `react/ChartWithExcelExport.tsx` | React + `highcharts-react-official` | `react`, `highcharts`, `highcharts-react-official` |
| `angular/chart-excel.component.ts` | Angular standalone component, `ViewChild` + `Highcharts.chart` | `@angular/core`, `highcharts` |
| `vue/ChartWithExcelExport.vue` | Vue 3 `<script setup>` + `highcharts-vue` | `vue`, `highcharts`, `highcharts-vue` |

The framework wrappers (`highcharts-react-official`, `highcharts-angular`, `highcharts-vue`) are your choice. They are **not** dependencies of this library, which only needs the Highcharts `Chart` instance. To use an example, copy the file into an app that has those packages installed.

## Typechecking in this repository

```bash
npx tsc -p examples/tsconfig.json
```

`examples/tsconfig.json` maps `highcharts-editable-excel` to `../src/index.ts`, so the examples are checked against the current source. React, Angular and the wrapper packages are not installed here, so `examples/shims.d.ts` declares the few symbols the examples use. Delete it in a real app. The `.vue` file and the HTML page are not typechecked.
