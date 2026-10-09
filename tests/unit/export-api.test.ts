import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  ExportError,
  XLSX_MIME_TYPE,
  analyzeChartCompatibility,
  exportChartsToWorkbook,
  exportHighchartsOptionsToXlsx,
  exportHighchartsToXlsx,
} from '../../src/index';
import type { Diagnostic } from '../../src/types/diagnostics';
import * as F from '../fixtures/highcharts-options';
import { inspectXlsx } from '../helpers/inspect-xlsx';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

const H = Highcharts as unknown as HighchartsLike;
const render = (o: Parameters<typeof renderChart>[1]) => renderChart(H, o);

async function expectExportError(p: Promise<unknown>, code: string): Promise<ExportError> {
  const error = await p.then(
    () => null,
    (e: unknown) => e,
  );
  expect(error).toBeInstanceOf(ExportError);
  expect((error as ExportError).code).toBe(code);
  return error as ExportError;
}

describe('exportHighchartsToXlsx (Highcharts 13)', () => {
  afterEach(() => destroyAll());

  it('produces an XLSX package with filename from the title, report, warnings and timings', async () => {
    const chart = render({ ...F.simpleLine, title: { text: 'Sales 2024' } });
    const result = await exportHighchartsToXlsx(chart);
    expect(result.bytes).toBeInstanceOf(Uint8Array);
    expect(result.bytes[0]).toBe(0x50); // P
    expect(result.bytes[1]).toBe(0x4b); // K
    expect(result.filename).toBe('Sales 2024.xlsx');
    expect(result.mimeType).toBe(XLSX_MIME_TYPE);
    expect(result.report.editable).toBe(true);
    expect(result.report.excelChartType).toBe('line');
    expect(Array.isArray(result.warnings)).toBe(true);
    for (const v of Object.values(result.timings)) expect(v).toBeGreaterThanOrEqual(0);
    expect(result.timings.totalMs).toBeGreaterThanOrEqual(result.timings.writeMs);
    expect(result.model).toBeUndefined();
    const x = await inspectXlsx(result.bytes);
    x.assertWellFormed();
    expect(x.sheetNames().map((s) => s.name)).toEqual(['Chart', 'Data']);
    expect(x.chartPaths()).toHaveLength(1);
  });

  it('appends .xlsx to an explicit filename and returns the model on request', async () => {
    const chart = render(F.columnChart);
    const result = await exportHighchartsToXlsx(chart, { filename: 'my report', includeModel: true });
    expect(result.filename).toBe('my report.xlsx');
    expect(result.model?.series.length).toBeGreaterThan(0);
    expect(result.model?.meta.sourceChartType).toBe('column');
  });

  it('falls back to chart.xlsx without a title', async () => {
    const chart = render({ ...F.simpleLine, title: { text: '' } });
    expect((await exportHighchartsToXlsx(chart)).filename).toBe('chart.xlsx');
  });

  it('does not mutate the chart options', async () => {
    const chart = render(F.multiLine);
    const before = JSON.stringify(chart.userOptions);
    await exportHighchartsToXlsx(chart, { themeOverrides: { colors: ['#ff0000'] }, hooks: { transformModel: (m) => ({ ...m }) } });
    expect(JSON.stringify(chart.userOptions)).toBe(before);
  });

  it('throws INVALID_CHART for non-charts', async () => {
    await expectExportError(exportHighchartsToXlsx({ foo: 1 }), 'INVALID_CHART');
    await expectExportError(exportHighchartsToXlsx(null), 'INVALID_CHART');
  });

  it('throws CHART_NOT_EDITABLE for a polar chart even without strictMode', async () => {
    const chart = render(F.polarChart);
    const error = await expectExportError(exportHighchartsToXlsx(chart), 'CHART_NOT_EDITABLE');
    expect(error.details.diagnostics?.some((d) => d.code === 'UNSUPPORTED_POLAR')).toBe(true);
    expect(error.message).toContain('UNSUPPORTED_POLAR');
  });

  it('strictMode throws CHART_NOT_EDITABLE for polar and unknown-only charts, not for a plain line chart', async () => {
    await expectExportError(exportHighchartsToXlsx(render(F.polarChart), { strictMode: true }), 'CHART_NOT_EDITABLE');
    await expectExportError(exportHighchartsToXlsx(render(F.unsupportedType), { strictMode: true }), 'CHART_NOT_EDITABLE');
    const ok = await exportHighchartsToXlsx(render(F.simpleLine), { strictMode: true });
    expect(ok.report.editable).toBe(true);
  });

  it('invokes onWarning once per reported warning', async () => {
    const seen: Diagnostic[] = [];
    const chart = render(F.hiddenSeries);
    const result = await exportHighchartsToXlsx(chart, { onWarning: (d) => seen.push(d) });
    expect(result.warnings.length).toBeGreaterThan(0);
    expect(seen).toHaveLength(result.warnings.length);
    expect(new Set(seen)).toEqual(new Set(result.warnings));
  });

  it('reports theme-override warnings through onWarning', async () => {
    const onWarning = vi.fn();
    const result = await exportHighchartsToXlsx(render(F.simpleLine), { onWarning, themeOverrides: { chartBackground: 'not-a-color' } });
    const unresolved = result.warnings.filter((d) => d.code === 'UNRESOLVED_COLOR');
    expect(unresolved).toHaveLength(1);
    expect(onWarning).toHaveBeenCalledWith(unresolved[0]);
  });

  it('includeReferenceImage outside a browser reports WRITER_LIMITATION and still exports', async () => {
    const result = await exportHighchartsToXlsx(render(F.simpleLine), { includeReferenceImage: true, strictMode: true });
    const d = result.warnings.find((w) => w.code === 'WRITER_LIMITATION');
    expect(d?.message).toMatch(/requires a browser/);
    expect((await inspectXlsx(result.bytes)).hasImages()).toBe(false);
  });

  it('hooks.transformModel can rename a series (data sheet header changes)', async () => {
    const chart = render(F.simpleLine);
    const result = await exportHighchartsToXlsx(chart, {
      hooks: { transformModel: (m) => ({ ...m, series: m.series.map((s, i) => (i === 0 ? { ...s, name: 'Renamed' } : s)) }) },
    });
    const x = await inspectXlsx(result.bytes);
    const data = x.sheetPath('Data');
    expect(x.cellValue(data, 'B1')).toBe('Renamed');
    expect(chart.series[0]?.name).toBe('Sales');
  });

  it('wraps unknown writer failures in WRITER_FAILURE', async () => {
    const writer = await import('../../src/excel/ooxml-writer');
    const spy = vi.spyOn(writer.OoxmlExcelWriter.prototype, 'write').mockRejectedValueOnce(new Error('disk full'));
    const error = await expectExportError(exportHighchartsToXlsx(render(F.simpleLine)), 'WRITER_FAILURE');
    expect((error.details.cause as Error).message).toBe('disk full');
    spy.mockRestore();
  });
});

describe('analyzeChartCompatibility (Highcharts 13)', () => {
  afterEach(() => destroyAll());

  it('reports spline → line as editable', () => {
    const report = analyzeChartCompatibility(render(F.splineChart));
    expect(report.sourceChartType).toBe('spline');
    expect(report.excelChartType).toBe('line');
    expect(report.editable).toBe(true);
    expect(report.supported.length).toBeGreaterThan(0);
  });

  it('reports polar as not editable without throwing', () => {
    const report = analyzeChartCompatibility(render(F.polarChart));
    expect(report.editable).toBe(false);
    expect(report.excelChartType).toBeNull();
    expect(report.blocking.length).toBeGreaterThan(0);
    expect(report.warnings.some((d) => d.code === 'UNSUPPORTED_POLAR' && d.outcome === 'blocking')).toBe(true);
  });

  it('throws INVALID_CHART for non-charts', () => {
    expect(() => analyzeChartCompatibility('nope')).toThrow(ExportError);
  });
});

describe('exportHighchartsOptionsToXlsx', () => {
  it('exports from a plain options object without a chart instance', async () => {
    const result = await exportHighchartsOptionsToXlsx(F.columnChart);
    expect(result.bytes[0]).toBe(0x50);
    expect(result.filename).toBe('Column.xlsx');
    expect(result.report.editable).toBe(true);
    const x = await inspectXlsx(result.bytes);
    expect(x.chartPaths()).toHaveLength(1);
  });

  it('skips the reference image with an info diagnostic', async () => {
    const result = await exportHighchartsOptionsToXlsx(F.simpleLine, { includeReferenceImage: true, strictMode: true });
    const d = result.warnings.find((w) => w.code === 'WRITER_LIMITATION');
    expect(d?.severity).toBe('info');
  });
});

describe('exportChartsToWorkbook (Highcharts 13)', () => {
  afterEach(() => destroyAll());

  it('writes one chart/data sheet pair per chart', async () => {
    const result = await exportChartsToWorkbook(
      [{ chart: render(F.simpleLine) }, { chart: render(F.columnChart) }, { chart: render(F.pieChart) }],
      { filename: 'dashboard' },
    );
    expect(result.filename).toBe('dashboard.xlsx');
    expect(result.charts.map((c) => [c.chartSheetName, c.dataSheetName])).toEqual([
      ['Chart 1', 'Data 1'],
      ['Chart 2', 'Data 2'],
      ['Chart 3', 'Data 3'],
    ]);
    const x = await inspectXlsx(result.bytes);
    x.assertWellFormed();
    expect(x.sheetNames()).toHaveLength(6);
    expect(x.chartPaths()).toHaveLength(3);
    // Each chart references its own data sheet.
    const formulas = x.chartPaths().map((p) => x.seriesFormulas(p)[0]?.val ?? '');
    expect(formulas[0]).toContain("'Data 1'");
    expect(formulas[1]).toContain("'Data 2'");
    expect(formulas[2]).toContain("'Data 3'");
  });

  it('honours entry sheet names and de-duplicates collisions', async () => {
    const result = await exportChartsToWorkbook([
      { chart: render(F.simpleLine), options: { chartSheetName: 'Sales', dataSheetName: 'Sales data' } },
      { chart: render(F.columnChart), options: { chartSheetName: 'Sales', dataSheetName: 'Sales data' } },
    ]);
    const names = result.charts.flatMap((c) => [c.chartSheetName, c.dataSheetName]);
    expect(new Set(names.map((n) => n.toLowerCase())).size).toBe(4);
    expect(names[0]).toBe('Sales');
  });

  it('names the entry index when a chart is not editable', async () => {
    const error = await expectExportError(
      exportChartsToWorkbook([{ chart: render(F.simpleLine) }, { chart: render(F.polarChart) }]),
      'CHART_NOT_EDITABLE',
    );
    expect(error.message).toContain('entry 1');
  });
});
