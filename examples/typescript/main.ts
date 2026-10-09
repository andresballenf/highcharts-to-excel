/**
 * TypeScript usage of highcharts-editable-excel. Every TypeScript snippet in the README is copied
 * from this file, so `pnpm examples:typecheck` (tsc -p examples/tsconfig.json) checks them.
 * Runs in a browser page that has <div id="sales"></div> and <div id="traffic"></div>.
 */
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import {
  DiagnosticCollector,
  ExportError,
  extractChartModel,
  translateChartModel,
  analyzeChartCompatibility,
  createDefaultExcelWriter,
  downloadHighchartsAsXlsx,
  exportChartsToWorkbook,
  exportHighchartsOptionsToXlsx,
  exportHighchartsToXlsx,
  installHighchartsExcelExport,
  addEditableExcelMenuItem,
  triggerDownload,
  type Diagnostic,
  type ExcelWriter,
  type ExportOptions,
  type ExportResult,
  type PerChartExportConfig,
  type WorkbookSpec,
} from 'highcharts-editable-excel';

// Optional: let TypeScript accept `exporting.editableExcel` in Highcharts options.
// (The library does not ship this augmentation; add it once in your app.)
declare module 'highcharts' {
  interface ExportingOptions {
    editableExcel?: PerChartExportConfig;
  }
}

// --- Quick start: menu integration --------------------------------------------------------------
installHighchartsExcelExport(Highcharts);

// --- Per-chart configuration ------------------------------------------------------------------------
const sales = Highcharts.chart('sales', {
  title: { text: 'Quarterly sales' },
  xAxis: { categories: ['Q1', 'Q2', 'Q3', 'Q4'] },
  series: [
    { type: 'column', name: 'Revenue', data: [120, 150, 170, 210] },
    { type: 'line', name: 'Target', data: [130, 140, 180, 200] },
  ],
  exporting: {
    editableExcel: { menuText: 'Excel (editable chart)', filename: 'sales-2026', dataSheetName: 'Sales data' },
  },
});

const traffic = Highcharts.chart('traffic', {
  title: { text: 'Traffic' },
  xAxis: { type: 'datetime' },
  series: [{ type: 'area', name: 'Visits', pointStart: Date.UTC(2026, 0, 1), pointInterval: 86_400_000, data: [5, 7, 3, 9, 4] }],
  exporting: { editableExcel: { enabled: false } }, // no menu item on this chart
});

// --- Programmatic export ----------------------------------------------------------------------------
async function exportSales(): Promise<void> {
  const result = await exportHighchartsToXlsx(sales, { filename: 'sales.xlsx' });
  console.log(result.filename, result.bytes.byteLength, result.mimeType);
  console.log(result.report.excelChartType); // e.g. "combo:column+line"
  for (const w of result.warnings) console.log(w.code, w.property, w.message);
  triggerDownload(result.bytes, result.filename); // or upload result.bytes somewhere
}

async function downloadSales(): Promise<void> {
  await downloadHighchartsAsXlsx(sales, { fidelity: 'best-effort' });
}

// --- Dry run --------------------------------------------------------------------------------------------
function checkTraffic(): boolean {
  const report = analyzeChartCompatibility(traffic);
  if (!report.editable) console.warn('Blocked by', report.blocking);
  console.log('approximated:', report.approximated, 'unsupported:', report.unsupported);
  return report.editable;
}

// --- No chart instance (Node or a worker) ---------------------------------------------------------------
async function exportFromOptions(): Promise<Uint8Array> {
  const { bytes } = await exportHighchartsOptionsToXlsx({
    chart: { type: 'bar', width: 800, height: 450 },
    title: { text: 'Headcount' },
    xAxis: { categories: ['Sales', 'R&D', 'Support'] },
    series: [{ type: 'bar', name: '2026', data: [42, 77, 18] }],
  });
  return bytes; // e.g. writeFileSync('headcount.xlsx', bytes) in Node
}

// --- Several charts in one workbook ---------------------------------------------------------------------
async function exportDashboard(): Promise<void> {
  const result = await exportChartsToWorkbook(
    [
      { chart: sales, options: { chartSheetName: 'Sales', dataSheetName: 'Sales data' } },
      { chart: traffic }, // sheets "Chart 2" / "Data 2"
    ],
    { filename: 'dashboard' },
  );
  triggerDownload(result.bytes, result.filename, result.mimeType);
  for (const c of result.charts) console.log(c.chartSheetName, c.report.excelChartType, c.warnings.length);
}

// --- Theme overrides, warnings, hidden data sheet, hooks ------------------------------------------------
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

// --- Strict mode and error handling ---------------------------------------------------------------------
async function exportStrict(): Promise<void> {
  try {
    await downloadHighchartsAsXlsx(sales, { ...brandOptions, strictMode: true });
  } catch (error) {
    if (error instanceof ExportError) {
      console.error(error.code, error.message, error.details.diagnostics?.map((d) => d.code));
    } else {
      throw error;
    }
  }
}

// --- Install with defaults and callbacks; patch a chart rendered before install -------------------------
// installHighchartsExcelExport is idempotent per Highcharts namespace: because the quick start above already
// installed, this call returns that installation and its options are ignored. Use one call in a real app.
function installWithCallbacks(): void {
  const installation = installHighchartsExcelExport(Highcharts, {
    menuText: 'Download editable Excel chart',
    exportOptions: { dataMode: 'rendered', seriesVisibility: 'visible' },
    onExport: (result) => console.log('exported', result.filename, result.warnings.length),
    onError: (error) => console.error('export failed', error),
  });
  addEditableExcelMenuItem(traffic); // no-op here: traffic opts out with enabled: false
  installation.uninstall();
}

// --- Writing a WorkbookSpec of your own with the built-in writer -----------------------------------------
async function writeCustomWorkbook(): Promise<Uint8Array> {
  const spec: WorkbookSpec = {
    properties: { title: 'Notes' },
    sheets: [
      {
        name: 'Notes',
        hidden: false,
        columns: [{ col0: 0, widthChars: 30 }],
        rows: [{ row0: 0, cells: [{ col0: 0, row0: 0, value: { type: 'string', value: '=SUM(A1) stays text' } }] }],
        freezeHeaderRow: false,
        drawings: [],
      },
    ],
  };
  return createDefaultExcelWriter().write(spec);
}

// --- README "Programmatic usage" block, verbatim ----------------------------------------------------------
async function programmaticUsage(chart: Highcharts.Chart, salesChart: Highcharts.Chart, trafficChart: Highcharts.Chart): Promise<void> {
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
  console.log(bytes.byteLength, multi.charts.length);
}

// --- README "TypeScript" framework snippet ---------------------------------------------------------------
async function typescriptSnippet(): Promise<ExportResult> {
  const chart = Highcharts.chart('sales', { series: [{ type: 'column', data: [1, 2, 3] }] });
  const result: ExportResult = await exportHighchartsToXlsx(chart);
  return result;
}

// --- Custom writer and the three-step pipeline (README "Custom writer") ----------------------------------
const loggingWriter: ExcelWriter = {
  name: 'logging-default',
  async write(workbook) {
    console.log('writing sheets', workbook.sheets.map((s) => s.name));
    return createDefaultExcelWriter().write(workbook);
  },
};

async function exportWithCustomWriter(): Promise<Uint8Array> {
  const { bytes } = await exportHighchartsToXlsx(sales, { writer: loggingWriter });
  return bytes;
}

async function manualPipeline(chart: Highcharts.Chart): Promise<Uint8Array> {
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
}

export { exportWithCustomWriter, manualPipeline, programmaticUsage, typescriptSnippet, exportSales, downloadSales, checkTraffic, exportFromOptions, exportDashboard, exportStrict, installWithCallbacks, writeCustomWorkbook };
