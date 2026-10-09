/**
 * Chart.js → XLSX end to end: real workbooks written by the built-in writer, inspected without
 * mocks, saved under tests/output/chartjs/, and (when LibreOffice is installed) rendered to PNG.
 *
 * The LibreOffice render is a smoke check that a third-party spreadsheet application draws the
 * native charts. It is NOT a Microsoft Excel fidelity claim (see docs/manual-qa.md).
 */

import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { Chart } from 'chart.js/auto';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { analyzeChartJsCompatibility, downloadChartJsAsXlsx, exportChartJsToXlsx } from '../../src/chartjs';
import { ExportError, type ExportProgress } from '../../src/types/public-api';
import * as F from '../fixtures/chartjs-configs';
import { inspectXlsx, type XlsxInspection } from '../helpers/inspect-xlsx';

const OUT = resolve(__dirname, '../output/chartjs');

function which(bin: string): string | null {
  const r = spawnSync('which', [bin], { encoding: 'utf8' });
  return r.status === 0 && r.stdout.trim() !== '' ? r.stdout.trim() : null;
}

/** Cell values of a sheet-qualified range formula such as `Data!$B$2:$B$7`. */
function rangeValues(x: XlsxInspection, formula: string): Array<string | number | null> {
  const m = /^'?(.+?)'?!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/.exec(formula);
  if (!m) throw new Error(`Unexpected formula ${formula}`);
  const [, sheet, col, from, col2, to] = m;
  if (col2 !== undefined && col2 !== col) throw new Error(`Multi-column range ${formula}`);
  const path = x.sheetPath(sheet!);
  const out: Array<string | number | null> = [];
  for (let r = Number(from); r <= Number(to ?? from); r++) out.push(x.cellValue(path, `${col}${r}`));
  return out;
}

async function exportAndSave(name: string, config: unknown, options = {}) {
  const result = await exportChartJsToXlsx(config, { includeModel: true, ...options });
  mkdirSync(OUT, { recursive: true });
  writeFileSync(join(OUT, `${name}.xlsx`), result.bytes);
  const x = await inspectXlsx(result.bytes);
  x.assertWellFormed();
  return { result, x };
}

describe('exportChartJsToXlsx writes native charts whose series reference the data', () => {
  it('line chart: categories, values and names come from the dataset', async () => {
    const config = F.lineChart();
    const { result, x } = await exportAndSave('line', config);
    expect(result.report.editable).toBe(true);
    expect(result.report.excelChartType).toBe('line');
    expect(result.filename).toBe('Monthly revenue.xlsx');
    expect(x.sheetNames().map((s) => s.name)).toEqual(['Chart', 'Data']);
    const chart = x.chartPaths()[0]!;
    expect(x.plotGroupKinds(chart)).toEqual(['lineChart']);
    const formulas = x.seriesFormulas(chart);
    expect(formulas).toHaveLength(2);
    for (const [i, f] of formulas.entries()) {
      const ds = config.data.datasets[i]!;
      expect(rangeValues(x, f.name!)).toEqual([ds.label]);
      expect(rangeValues(x, f.cat!)).toEqual(F.MONTHS);
      expect(rangeValues(x, f.val!)).toEqual(ds.data);
    }
    expect(x.chartXml(0)).toContain('Monthly revenue');
  });

  it('stacked bars: one stacked column group, values per dataset', async () => {
    const config = F.stackedBarChart();
    const { result, x } = await exportAndSave('stacked-bar', config);
    expect(result.report.excelChartType).toBe('stackedColumn');
    const chart = x.chartPaths()[0]!;
    expect(x.plotGroupKinds(chart)).toEqual(['barChart']);
    expect(x.chartXml(0)).toContain('<c:grouping val="stacked"/>');
    const formulas = x.seriesFormulas(chart);
    expect(formulas.map((f) => rangeValues(x, f.val!))).toEqual(config.data.datasets.map((d) => d.data));
    expect(rangeValues(x, formulas[0]!.cat!)).toEqual(['Q1', 'Q2', 'Q3']);
  });

  it('pie chart: slice labels and values, slice colors in the chart', async () => {
    const config = F.pieChart();
    const { result, x } = await exportAndSave('pie', config);
    expect(result.report.excelChartType).toBe('pie');
    const f = x.seriesFormulas(x.chartPaths()[0]!)[0]!;
    expect(rangeValues(x, f.cat!)).toEqual(['Chrome', 'Firefox', 'Safari']);
    expect(rangeValues(x, f.val!)).toEqual([62, 20, 18]);
    const xml = x.chartXml(0);
    for (const color of ['36A2EB', 'FF6384', 'FFCD56']) expect(xml.toUpperCase()).toContain(color);
  });

  it('mixed bar + line becomes a combo chart', async () => {
    const { result, x } = await exportAndSave('combo', F.mixedBarLineChart());
    expect(x.plotGroupKinds(x.chartPaths()[0]!)).toEqual(['barChart', 'lineChart']);
    expect(result.report.excelChartType).toMatch(/combo/);
  });

  it('secondary y axis: two value axes', async () => {
    const { result, x } = await exportAndSave('secondary-axis', F.secondaryAxisChart());
    expect(result.report.editable).toBe(true);
    expect((x.chartXml(0).match(/<c:valAx>/g) ?? []).length).toBe(2);
    expect(result.warnings.some((d) => d.code === 'SECONDARY_AXIS')).toBe(true);
  });

  it('time scale: x values become Excel dates', async () => {
    const { x } = await exportAndSave('time-scale', F.timeScaleChart());
    const f = x.seriesFormulas(x.chartPaths()[0]!)[0]!;
    // 2024-01-01 is Excel serial 45292.
    expect(rangeValues(x, f.cat!)).toEqual([45292, 45293, 45296]);
    expect(rangeValues(x, f.val!)).toEqual([5, 8, 3]);
    expect(x.chartXml(0)).toContain('<c:dateAx>');
  });

  it('scatter: X and Y columns', async () => {
    const { x } = await exportAndSave('scatter', F.scatterChart());
    const f = x.seriesFormulas(x.chartPaths()[0]!)[0]!;
    expect(rangeValues(x, f.xVal!)).toEqual([-10, 0, 10, 0.5]);
    expect(rangeValues(x, f.yVal!)).toEqual([0, 10, 5, 5.5]);
  });

  it('nulls stay empty cells', async () => {
    const { result, x } = await exportAndSave('nulls', F.nullsChart());
    const f = x.seriesFormulas(x.chartPaths()[0]!)[0]!;
    expect(rangeValues(x, f.val!)).toEqual([1, null, 3, null]);
    expect(result.report.dataConcerns.some((d) => d.code === 'NULL_VALUES' || d.code === 'NON_NUMERIC_VALUE')).toBe(
      true,
    );
  });

  it('a live Chart instance exports the same workbook data', async () => {
    const restore = F.installCanvasStub();
    try {
      const config = F.stackedBarChart();
      const chart = new Chart(F.makeCanvas(), F.staticOptions(config));
      const { x, result } = await exportAndSave('live-stacked-bar', chart);
      expect(result.model!.meta.sourceVersion).toBe(Chart.version);
      const formulas = x.seriesFormulas(x.chartPaths()[0]!);
      expect(formulas.map((f) => rangeValues(x, f.val!))).toEqual([
        [10, 20, 30],
        [5, 15, 25],
        [8, 6, 4],
      ]);
      chart.destroy();
    } finally {
      restore();
    }
  });
});

describe('editability, options and download', () => {
  it('radar and polarArea are not editable: export rejects, analyze reports', async () => {
    await expect(exportChartJsToXlsx(F.radarChart())).rejects.toMatchObject({ code: 'CHART_NOT_EDITABLE' });
    const report = analyzeChartJsCompatibility(F.polarAreaChart());
    expect(report.editable).toBe(false);
    expect(report.unsupported).toContain('data.datasets[0].type');
  });

  it('strictMode rejects unsupported tick callbacks', async () => {
    await expect(exportChartJsToXlsx(F.callbackChart())).resolves.toBeDefined();
    const error = await exportChartJsToXlsx(F.callbackChart(), { strictMode: true }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ExportError);
    expect((error as ExportError).code).toBe('CHART_NOT_EDITABLE');
    expect((error as ExportError).message).toContain('options.scales.y.ticks.callback');
  });

  it('rejects non-charts with INVALID_CHART and bad options with INVALID_OPTIONS', async () => {
    await expect(exportChartJsToXlsx({ nope: true })).rejects.toMatchObject({ code: 'INVALID_CHART' });
    await expect(exportChartJsToXlsx(F.lineChart(), { fidelity: 'max' as never })).rejects.toMatchObject({
      code: 'INVALID_OPTIONS',
    });
  });

  it('reports progress, honours an aborted signal and fires onWarning', async () => {
    const phases: ExportProgress['phase'][] = [];
    const warnings: string[] = [];
    await exportChartJsToXlsx(F.lineChart(), {
      onProgress: (p) => phases.push(p.phase),
      onWarning: (d) => warnings.push(d.code),
    });
    expect(phases[0]).toBe('extract');
    expect(phases.at(-1)).toBe('done');
    expect(warnings).toContain('UNSUPPORTED_STYLE');
    const controller = new AbortController();
    controller.abort();
    await expect(exportChartJsToXlsx(F.lineChart(), { signal: controller.signal })).rejects.toMatchObject({
      code: 'ABORTED',
    });
  });

  it('applies theme overrides and the transformModel hook', async () => {
    const { result } = await exportAndSave('overrides', F.lineChart(), {
      themeOverrides: { colors: ['#00aa00'] },
      hooks: { transformModel: (m: { title: unknown }) => ({ ...m, title: null }) },
    });
    expect(result.model!.title).toBeNull();
    expect(result.model!.series[0]!.color).toMatchObject({ r: 0, g: 0xaa, b: 0 });
  });

  it('reports the reference image as unavailable for a configuration', async () => {
    const result = await exportChartJsToXlsx(F.lineChart(), { includeReferenceImage: true });
    expect(result.warnings.some((d) => d.property === 'includeReferenceImage')).toBe(true);
  });

  it('downloadChartJsAsXlsx triggers a browser download', async () => {
    const create = vi.fn(() => 'blob:chartjs');
    const revoke = vi.fn();
    const urlApi = URL as unknown as { createObjectURL?: unknown; revokeObjectURL?: unknown };
    const saved = { create: urlApi.createObjectURL, revoke: urlApi.revokeObjectURL };
    urlApi.createObjectURL = create;
    urlApi.revokeObjectURL = revoke;
    const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
    try {
      const result = await downloadChartJsAsXlsx(F.pieChart(), { filename: 'shares' });
      expect(result.filename).toBe('shares.xlsx');
      expect(create).toHaveBeenCalledOnce();
      expect(click).toHaveBeenCalledOnce();
    } finally {
      urlApi.createObjectURL = saved.create;
      urlApi.revokeObjectURL = saved.revoke;
    }
  });
});

// ---------------------------------------------------------------------------
// LibreOffice render smoke check (skipped when soffice / poppler are missing)
// ---------------------------------------------------------------------------

const SOFFICE = which('soffice') ?? which('libreoffice');
const PDFTOPPM = which('pdftoppm');
const PDFTOTEXT = which('pdftotext');
const RENDER_SKIP = !SOFFICE || !PDFTOPPM || !PDFTOTEXT ? 'soffice or poppler-utils not on PATH' : null;
const RENDER_OUT = join(OUT, 'render');
const RENDER_CASES = [
  { name: 'line', make: F.lineChart, text: ['Monthly revenue', 'Revenue', 'Costs'] },
  { name: 'stacked-bar', make: F.stackedBarChart, text: ['Stacked sales', 'North', 'South'] },
  { name: 'pie', make: F.pieChart, text: ['Browser share', 'Chrome', 'Firefox'] },
];
const pdfText = new Map<string, string>();

describe.skipIf(RENDER_SKIP !== null)('LibreOffice render of Chart.js exports (not an Excel fidelity claim)', () => {
  beforeAll(async () => {
    mkdirSync(RENDER_OUT, { recursive: true });
    const files: string[] = [];
    for (const c of RENDER_CASES) {
      const result = await exportChartJsToXlsx(c.make(), { includeSourceData: false });
      const p = join(RENDER_OUT, `${c.name}.xlsx`);
      writeFileSync(p, result.bytes);
      rmSync(join(RENDER_OUT, `${c.name}.pdf`), { force: true });
      files.push(p);
    }
    const profile = mkdtempSync(join(tmpdir(), 'cj2xl-lo-'));
    try {
      execFileSync(
        SOFFICE!,
        [
          `-env:UserInstallation=${pathToFileURL(profile).href}`,
          '--headless',
          '--convert-to',
          'pdf',
          '--outdir',
          RENDER_OUT,
          ...files,
        ],
        { stdio: 'pipe', timeout: 180_000 },
      );
    } finally {
      rmSync(profile, { recursive: true, force: true });
    }
    for (const c of RENDER_CASES) {
      const pdf = join(RENDER_OUT, `${c.name}.pdf`);
      if (!existsSync(pdf)) continue;
      execFileSync(PDFTOPPM!, ['-r', '60', '-png', '-f', '1', '-singlefile', pdf, join(RENDER_OUT, c.name)], {
        stdio: 'pipe',
        timeout: 60_000,
      });
      pdfText.set(c.name, execFileSync(PDFTOTEXT!, [pdf, '-'], { encoding: 'utf8', timeout: 60_000 }));
    }
  }, 240_000);

  afterAll(() => {
    console.info(`[chartjs-export] LibreOffice renders in ${RENDER_OUT}`);
  });

  it.each(RENDER_CASES.map((c) => [c.name, c] as const))('%s renders to a PNG with the chart text', (_n, c) => {
    const png = join(RENDER_OUT, `${c.name}.png`);
    expect(existsSync(png), `${png} not produced`).toBe(true);
    expect(statSync(png).size).toBeGreaterThan(3 * 1024);
    const text = pdfText.get(c.name) ?? '';
    for (const s of c.text) expect(text, `"${s}" in ${c.name}.pdf`).toContain(s);
  });
});

if (RENDER_SKIP !== null) {
  if (process.env.REQUIRE_RENDER === '1') {
    describe('LibreOffice render of Chart.js exports', () => {
      it('requires the render toolchain (REQUIRE_RENDER=1)', () => {
        throw new Error(`[chartjs-export] REQUIRE_RENDER=1 but ${RENDER_SKIP}.`);
      });
    });
  } else {
    console.warn(`[chartjs-export] render skipped: ${RENDER_SKIP}`);
  }
}
