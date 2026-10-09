/**
 * highcharts-editable-excel demo: a gallery of charts, each exportable to an editable XLSX chart.
 *
 * `highcharts-editable-excel` is aliased to ../src/index.ts (see demo/vite.config.ts), so this
 * page always runs the library source. Note: no export-data module is loaded — not needed.
 */

import Highcharts, { type Chart, type Options } from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import highchartsStyledModeCss from 'highcharts/css/highcharts.css?raw';
import {
  analyzeChartCompatibility,
  downloadHighchartsAsXlsx,
  exportChartsToWorkbook,
  exportHighchartsToXlsx,
  installHighchartsExcelExport,
  triggerDownload,
  type CompatibilityReport,
  type Diagnostic,
  type ExportTimings,
} from 'highcharts-editable-excel';
import {
  BRAND_COLOR,
  DOWNLOAD_ARROW_PATH,
  GERMAN_MENU_TEXT,
  demoCharts,
  deterministicValue,
  type DemoChart,
} from './charts';

/** What the page remembers about the last export/analysis of each chart (read by the e2e tests). */
export interface DemoResult {
  kind: 'export' | 'analysis';
  filename?: string;
  warnings: Diagnostic[];
  report: CompatibilityReport;
  timings?: ExportTimings;
}

/** Test hooks read by the e2e tests (e2e/export.spec.ts). */
export interface DemoHandle {
  charts: Record<string, Chart>;
  lastResults: Record<string, DemoResult | { error: string }>;
  /** True once every gallery card has rendered. */
  ready: boolean;
  /** Exports one chart with `includeReferenceImage: true` (no download) and returns the bytes. */
  exportWithImage: (name: string) => Promise<{ bytes: number[]; warnings: Diagnostic[] }>;
  /** The page's Highcharts namespace (lets the e2e tests render an unbranded reference button). */
  highcharts: typeof Highcharts;
}

declare global {
  interface Window {
    __demo: DemoHandle;
  }
}

const charts: Record<string, Chart> = {};
const lastResults: DemoHandle['lastResults'] = {};
window.__demo = {
  charts,
  lastResults,
  ready: false,
  exportWithImage: async (name) => {
    const chart = charts[name];
    if (!chart) throw new Error(`[demo] no chart named "${name}"`);
    const result = await exportHighchartsToXlsx(chart, { includeReferenceImage: true });
    return { bytes: Array.from(result.bytes), warnings: result.warnings };
  },
  highcharts: Highcharts,
};

// Deterministic rendering: no animation; the accessibility module is not loaded in this demo.
Highcharts.setOptions({
  accessibility: { enabled: false },
  chart: { animation: false },
  plotOptions: { series: { animation: false } },
});

// Styled mode needs Highcharts' stylesheet. It is global by design, so nest it under the
// styled card's class: otherwise its `.highcharts-*` rules would recolor every other chart.
const scopedCss = document.createElement('style');
scopedCss.id = 'highcharts-styled-mode-css';
scopedCss.textContent = `.hc-styled-scope {\n${highchartsStyledModeCss.replaceAll(':root', '&')}\n}`;
document.head.prepend(scopedCss);

// One install for the whole page: every chart created afterwards gets the menu item, with the
// branded context button (a download arrow), an Excel icon and green menu styles. The item text
// comes from lang.downloadEditableXLSX (the default langKey): see the "Deutsch" button.
installHighchartsExcelExport(Highcharts, {
  onExport: (result, chart) => showResult(nameOf(chart), { kind: 'export', ...result }),
  onError: (error, chart) => showError(nameOf(chart), error),
  menuIcon: 'excel',
  button: { svgPath: DOWNLOAD_ARROW_PATH, symbolFill: BRAND_COLOR, symbolStroke: BRAND_COLOR, symbolStrokeWidth: 2 },
  menuStyle: { border: `1px solid ${BRAND_COLOR}`, borderRadius: '6px', boxShadow: '0 4px 12px rgba(0,0,0,.15)' },
  menuItemStyle: { fontFamily: 'system-ui', fontSize: '13px', padding: '6px 14px' },
  menuItemHoverStyle: { background: '#e8f3ec', color: BRAND_COLOR },
});

// ---------------------------------------------------------------------------------------------
// Rendering helpers

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Partial<HTMLElementTagNameMap[K]> = {},
  ...children: Array<Node | string>
): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
}

function button(text: string, testId: string, onClick: () => void | Promise<void>): HTMLButtonElement {
  const b = el('button', { type: 'button', textContent: text });
  b.dataset.testid = testId;
  b.addEventListener('click', () => void onClick());
  return b;
}

function nameOf(chart: unknown): string {
  const found = Object.entries(charts).find(([, c]) => c === chart);
  return found?.[0] ?? 'unknown';
}

function panel(name: string): {
  details: HTMLDetailsElement;
  summary: HTMLElement;
  list: HTMLUListElement;
  pre: HTMLPreElement;
  status: HTMLElement;
} {
  const details = document.querySelector<HTMLDetailsElement>(`[data-testid="warnings-${name}"]`)!;
  return {
    details,
    summary: details.querySelector('summary')!,
    list: details.querySelector('ul')!,
    pre: details.querySelector('pre')!,
    status: document.querySelector<HTMLElement>(`[data-testid="status-${name}"]`)!,
  };
}

/** Renders the warnings (`code — property — message`) and the JSON report of the last run. */
function showResult(name: string, result: DemoResult): void {
  lastResults[name] = result;
  const p = panel(name);
  if (!p.details) return;
  const { warnings, report } = result;
  p.list.replaceChildren(
    ...(warnings.length === 0
      ? [el('li', { className: 'empty', textContent: 'No warnings: everything was translated.' })]
      : warnings.map((w) => el('li', {}, el('code', { textContent: w.code }), ` — ${w.property} — ${w.message}`))),
  );
  p.pre.textContent = JSON.stringify(
    {
      kind: result.kind,
      filename: result.filename,
      editable: report.editable,
      excelChartType: report.excelChartType,
      report,
      timings: result.timings,
    },
    null,
    2,
  );
  p.summary.textContent = `${result.kind === 'export' ? 'Export' : 'Analysis'}: ${warnings.length} warning${warnings.length === 1 ? '' : 's'}`;
  p.status.classList.remove('error');
  p.status.textContent =
    result.kind === 'export'
      ? `Downloaded ${result.filename} (${report.editable ? 'editable' : 'not editable'} ${report.excelChartType ?? ''} chart)`
      : `Analysis: ${report.editable ? `editable as ${report.excelChartType}` : 'not editable'}`;
  p.details.open = true;
}

function showError(name: string, error: unknown): void {
  const message = error instanceof Error ? `${error.name}: ${error.message}` : String(error);
  lastResults[name] = { error: message };
  console.error(`[demo] ${name}:`, error);
  const status = document.querySelector<HTMLElement>(`[data-testid="status-${name}"]`);
  if (status) {
    status.textContent = message;
    status.classList.add('error');
  }
}

/** One gallery card: title, note, chart, buttons and the warnings panel. Returns the chart. */
function createCard(
  name: string,
  title: string,
  options: Options,
  note = '',
  extraButtons: (chart: Chart) => HTMLButtonElement[] = () => [],
): Chart {
  const chartDiv = el('div', { id: `chart-${name}`, className: 'chart' });
  const actions = el('div', { className: 'actions' });
  const status = el('span', { className: 'status' });
  status.dataset.testid = `status-${name}`;
  status.setAttribute('aria-live', 'polite');
  const details = el(
    'details',
    { className: 'warnings' },
    el('summary', { textContent: 'Warnings (none yet: export or analyze)' }),
    el('ul', { className: 'warning-list' }),
    el('pre', { className: 'report' }),
  );
  details.dataset.testid = `warnings-${name}`;

  const card = el(
    'article',
    { className: 'card' },
    el('h2', { textContent: title }),
    el('p', { className: 'note', textContent: note }),
    el('div', { className: 'chart-frame' }, chartDiv),
    actions,
    status,
    details,
  );
  card.dataset.chart = name;
  if (typeof options.chart?.width === 'number' && options.chart.width > 600) card.classList.add('wide');
  if (options.chart?.styledMode) card.classList.add('hc-styled-scope');
  document.getElementById('gallery')!.append(card);

  // Render after the card is in the DOM so Highcharts can measure the container.
  const chart = Highcharts.chart(chartDiv, options);
  charts[name] = chart;

  // Handlers read `charts[name]`: a card may re-create its chart (the branded one does).
  actions.append(
    button('Export to Excel', `export-${name}`, async () => {
      try {
        const result = await downloadHighchartsAsXlsx(charts[name] ?? chart, {
          filename: `${name}.xlsx`,
          onWarning: (d) => console.debug(`[demo] ${name}: ${d.code} — ${d.property} — ${d.message}`),
        });
        showResult(name, { kind: 'export', ...result });
      } catch (error) {
        showError(name, error);
      }
    }),
    button('Analyze', `analyze-${name}`, () => {
      try {
        const report = analyzeChartCompatibility(charts[name] ?? chart);
        showResult(name, { kind: 'analysis', warnings: report.warnings, report });
      } catch (error) {
        showError(name, error);
      }
    }),
    ...extraButtons(chart),
  );
  return chart;
}

// ---------------------------------------------------------------------------------------------
// Gallery

let dynamicStep = 0;
const dynamicButtons = (chart: Chart): HTMLButtonElement[] => [
  button('Add point', 'add-point-chart-dynamic', () => {
    chart.series[0]!.addPoint(deterministicValue(++dynamicStep));
  }),
  button('Randomize', 'randomize-chart-dynamic', () => {
    const n = chart.series[0]!.points.length;
    const start = ++dynamicStep;
    chart.series[0]!.setData(Array.from({ length: n }, (_, i) => deterministicValue(start * 7 + i)));
  }),
];

/** Re-creates a card's chart in place (new charts pick up the current global options, e.g. lang). */
function recreateChart(name: string, options: Options): Chart {
  const container = document.getElementById(`chart-${name}`)!;
  charts[name]?.destroy();
  const chart = Highcharts.chart(container, options);
  charts[name] = chart;
  return chart;
}

/** Language toggle: the menu text is `lang.downloadEditableXLSX`, translated like any Highcharts text. */
const brandedButtons = (options: Options) => (): HTMLButtonElement[] => [
  button('Deutsch', 'lang-de', () => {
    Highcharts.setOptions({ lang: { downloadEditableXLSX: GERMAN_MENU_TEXT } });
    recreateChart('branded', options);
  }),
  button('English', 'lang-en', () => {
    Highcharts.setOptions({ lang: { downloadEditableXLSX: 'Download editable Excel chart' } });
    recreateChart('branded', options);
  }),
];

const extraButtonsFor: Record<string, (def: DemoChart) => ((chart: Chart) => HTMLButtonElement[]) | undefined> = {
  dynamic: () => dynamicButtons,
  branded: (def) => brandedButtons(def.options),
};

for (const def of demoCharts as DemoChart[]) {
  createCard(def.name, def.title, def.options, def.note, extraButtonsFor[def.name]?.(def));
}

// ---------------------------------------------------------------------------------------------
// Export all charts into one workbook (one chart sheet + one data sheet per chart).

document.querySelector<HTMLButtonElement>('[data-testid="export-all"]')!.addEventListener('click', async () => {
  const status = document.querySelector<HTMLElement>('[data-testid="export-all-status"]')!;
  try {
    const names = Object.keys(charts);
    const result = await exportChartsToWorkbook(
      names.map((name) => ({
        chart: charts[name],
        options: { chartSheetName: `${name} chart`, dataSheetName: `${name} data` },
      })),
      { filename: 'all-charts.xlsx' },
    );
    triggerDownload(result.bytes, result.filename, result.mimeType);
    const warnings = result.charts.reduce((n, c) => n + c.warnings.length, 0);
    status.classList.remove('error');
    status.textContent = `Downloaded ${result.filename}: ${result.charts.length} charts, ${warnings} warnings`;
  } catch (error) {
    showError('all', error);
    status.textContent = error instanceof Error ? error.message : String(error);
    status.classList.add('error');
  }
});

// Every card is rendered and every button is wired: the e2e tests wait for this flag.
window.__demo.ready = true;
