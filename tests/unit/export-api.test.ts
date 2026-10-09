import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import 'highcharts/highcharts-more';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  ExportError,
  XLSX_MIME_TYPE,
  analyzeChartCompatibility,
  downloadHighchartsAsXlsx,
  exportChartsToWorkbook,
  exportHighchartsOptionsToXlsx,
  exportHighchartsToXlsx,
} from '../../src/index';
import type { Diagnostic } from '../../src/types/diagnostics';
import {
  resolveExportOptions,
  type ExportOptions,
  type ExportProgress,
  type MultiChartExportEntry,
} from '../../src/types/public-api';
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
    await exportHighchartsToXlsx(chart, {
      themeOverrides: { colors: ['#ff0000'] },
      hooks: { transformModel: (m) => ({ ...m }) },
    });
    expect(JSON.stringify(chart.userOptions)).toBe(before);
  });

  it('throws INVALID_CHART for non-charts', async () => {
    await expectExportError(exportHighchartsToXlsx({ foo: 1 }), 'INVALID_CHART');
    await expectExportError(exportHighchartsToXlsx(null), 'INVALID_CHART');
  });

  it('throws CHART_NOT_EDITABLE for a polar column chart even without strictMode', async () => {
    const chart = render(F.polarColumnChart);
    const error = await expectExportError(exportHighchartsToXlsx(chart), 'CHART_NOT_EDITABLE');
    expect(error.details.diagnostics?.some((d) => d.code === 'UNSUPPORTED_POLAR')).toBe(true);
    expect(error.message).toContain('UNSUPPORTED_POLAR');
  });

  it('strictMode throws CHART_NOT_EDITABLE for polar column and unknown-only charts, not for a plain line chart', async () => {
    await expectExportError(
      exportHighchartsToXlsx(render(F.polarColumnChart), { strictMode: true }),
      'CHART_NOT_EDITABLE',
    );
    await expectExportError(
      exportHighchartsToXlsx(render(F.unsupportedType), { strictMode: true }),
      'CHART_NOT_EDITABLE',
    );
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
    const result = await exportHighchartsToXlsx(render(F.simpleLine), {
      onWarning,
      themeOverrides: { chartBackground: 'not-a-color' },
    });
    const unresolved = result.warnings.filter((d) => d.code === 'UNRESOLVED_COLOR');
    expect(unresolved).toHaveLength(1);
    expect(onWarning).toHaveBeenCalledWith(unresolved[0]);
  });

  it('includeReferenceImage outside a browser reports WRITER_LIMITATION and still exports', async () => {
    const result = await exportHighchartsToXlsx(render(F.simpleLine), {
      includeReferenceImage: true,
      strictMode: true,
    });
    const d = result.warnings.find((w) => w.code === 'WRITER_LIMITATION');
    expect(d?.message).toMatch(/requires a browser/);
    expect((await inspectXlsx(result.bytes)).hasImages()).toBe(false);
  });

  it('hooks.transformModel can rename a series (data sheet header changes)', async () => {
    const chart = render(F.simpleLine);
    const result = await exportHighchartsToXlsx(chart, {
      hooks: {
        transformModel: (m) => ({ ...m, series: m.series.map((s, i) => (i === 0 ? { ...s, name: 'Renamed' } : s)) }),
      },
    });
    const x = await inspectXlsx(result.bytes);
    const data = x.sheetPath('Data');
    expect(x.cellValue(data, 'B1')).toBe('Renamed');
    expect(chart.series[0]?.name).toBe('Sales');
  });

  it('wraps unknown writer failures in WRITER_FAILURE', async () => {
    const writer = await import('../../src/excel/ooxml-writer');
    // Restored automatically before the next test (vitest.config.ts: restoreMocks).
    vi.spyOn(writer.OoxmlExcelWriter.prototype, 'write').mockRejectedValueOnce(new Error('disk full'));
    const error = await expectExportError(exportHighchartsToXlsx(render(F.simpleLine)), 'WRITER_FAILURE');
    expect((error.cause as Error).message).toBe('disk full');
    // A11: the cause lives on Error#cause (non-enumerable), not duplicated in details or own keys.
    expect(error.details).not.toHaveProperty('cause');
    expect(Object.keys(error)).not.toContain('cause');
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

  it('reports a polar line chart as an editable Excel radar chart', async () => {
    const chart = render(F.polarChart);
    const report = analyzeChartCompatibility(chart);
    expect(report.editable).toBe(true);
    expect(report.excelChartType).toBe('radar');
    expect(report.warnings).toContainEqual(
      expect.objectContaining({ code: 'APPROXIMATED_CHART_TYPE', property: 'chart.polar', severity: 'info' }),
    );
    const result = await exportHighchartsToXlsx(chart, { strictMode: true });
    expect(result.report.excelChartType).toBe('radar');
  });

  it('reports polar columns as not editable without throwing', () => {
    const report = analyzeChartCompatibility(render(F.polarColumnChart));
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
      exportChartsToWorkbook([{ chart: render(F.simpleLine) }, { chart: render(F.polarColumnChart) }]),
      'CHART_NOT_EDITABLE',
    );
    expect(error.message).toContain('entry 1');
  });

  it('A5 checks blocking before rendering the reference image (multi-chart path)', async () => {
    const onWarning = vi.fn();
    await expectExportError(
      exportChartsToWorkbook([
        { chart: render(F.polarColumnChart), options: { includeReferenceImage: true, onWarning } },
      ]),
      'CHART_NOT_EDITABLE',
    );
    expect(onWarning.mock.calls.map(([d]) => (d as Diagnostic).code)).not.toContain('WRITER_LIMITATION');
  });

  it('A6 entry options cannot carry a writer or includeModel (type level)', () => {
    const writer = { name: 'test-writer', write: async () => new Uint8Array() };
    // @ts-expect-error the writer is workbook-wide: pass it in exportChartsToWorkbook's second argument
    const a: MultiChartExportEntry = { chart: null, options: { writer } };
    // @ts-expect-error includeModel has no per-entry result to attach to
    const b: MultiChartExportEntry = { chart: null, options: { includeModel: true } };
    expect([a, b]).toHaveLength(2);
  });
});

describe('export pipeline audit fixes (Highcharts 13)', () => {
  afterEach(() => destroyAll());

  it('A3 yields to the event loop between phases in a browser-like environment', async () => {
    const order: string[] = [];
    const chart = render(F.simpleLine);
    setTimeout(() => order.push('timer'), 0);
    await exportHighchartsToXlsx(chart);
    order.push('done');
    expect(order).toEqual(['timer', 'done']);
  });

  it('A5 checks blocking before rendering the reference image and times the image separately', async () => {
    const onWarning = vi.fn();
    await expectExportError(
      exportHighchartsToXlsx(render(F.polarColumnChart), { includeReferenceImage: true, onWarning }),
      'CHART_NOT_EDITABLE',
    );
    expect(onWarning.mock.calls.map(([d]) => (d as Diagnostic).code)).not.toContain('WRITER_LIMITATION');
    const plain = await exportHighchartsToXlsx(render(F.simpleLine));
    expect(plain.timings.imageMs).toBe(0);
    const withImage = await exportHighchartsToXlsx(render(F.simpleLine), { includeReferenceImage: true });
    expect(typeof withImage.timings.imageMs).toBe('number');
    expect(withImage.warnings.some((w) => w.code === 'WRITER_LIMITATION')).toBe(true);
  });
});

describe('resolveExportOptions validation (A2)', () => {
  const writer = { name: 'test-writer', write: async () => new Uint8Array() };

  it.each([
    ['fidelity', { fidelity: 'max' }],
    ['dataMode', { dataMode: 'all' }],
    ['seriesVisibility', { seriesVisibility: 'hidden' }],
    ['chartWidth', { chartWidth: '800' }],
    ['chartWidth', { chartWidth: 10 }],
    ['chartHeight', { chartHeight: Number.NaN }],
    ['chartHeight', { chartHeight: 30_000 }],
    ['chartSheetName', { chartSheetName: 5 }],
    ['dataSheetName', { dataSheetName: null }],
    ['filename', { filename: 42 }],
    ['onWarning', { onWarning: 'log' }],
    ['hooks', { hooks: 'nope' }],
    ['hooks.transformModel', { hooks: { transformModel: 1 } }],
    ['writer', { writer: {} }],
  ])('rejects an invalid %s', (name, bad) => {
    let error: unknown;
    try {
      resolveExportOptions(bad as unknown as ExportOptions);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExportError);
    expect((error as ExportError).code).toBe('INVALID_OPTIONS');
    expect((error as ExportError).message).toContain(name);
    expect((error as ExportError).details.property).toBe(name);
  });

  it('rejects a non-object options argument', () => {
    expect(() => resolveExportOptions(42 as unknown as ExportOptions)).toThrow(
      expect.objectContaining({ code: 'INVALID_OPTIONS' }),
    );
  });

  it('accepts valid options and the documented bounds', () => {
    expect(() =>
      resolveExportOptions({
        fidelity: 'minimal',
        dataMode: 'raw',
        seriesVisibility: 'all',
        chartWidth: 50,
        chartHeight: 20_000,
        chartSheetName: 'C',
        dataSheetName: 'D',
        filename: 'f',
        onWarning: () => undefined,
        hooks: { transformModel: (m) => m },
        writer,
      }),
    ).not.toThrow();
    expect(resolveExportOptions().fidelity).toBe('best-effort');
  });

  it('export entry points reject invalid options with INVALID_OPTIONS', async () => {
    await expectExportError(
      exportHighchartsOptionsToXlsx(F.simpleLine, { chartWidth: '800' as unknown as number }),
      'INVALID_OPTIONS',
    );
    await expectExportError(
      exportChartsToWorkbook([{ chart: null }], { writer: {} as unknown as typeof writer }),
      'INVALID_OPTIONS',
    );
  });
});

describe('cancellation, progress and timings (Highcharts 13)', () => {
  afterEach(() => destroyAll());

  /** Plain options with `n` points per series: big enough for several write chunks, no rendering. */
  const bigOptions = (n = 12_000, seriesCount = 2) => ({
    title: { text: 'Big' },
    series: Array.from({ length: seriesCount }, (_, s) => ({
      type: 'line',
      name: `S${s}`,
      data: Array.from({ length: n }, (_, i) => (i * (s + 1)) % 97),
    })),
  });

  it('rejects with ABORTED (cause = signal.reason) when the signal is already aborted, before any work', async () => {
    const writer = await import('../../src/excel/ooxml-writer');
    const write = vi.spyOn(writer.OoxmlExcelWriter.prototype, 'write');
    const reason = new Error('user cancelled');
    const onProgress = vi.fn();
    const error = await expectExportError(
      exportHighchartsToXlsx(render(F.simpleLine), { signal: AbortSignal.abort(reason), onProgress }),
      'ABORTED',
    );
    expect(error.cause).toBe(reason);
    expect(write).not.toHaveBeenCalled();
    expect(onProgress).not.toHaveBeenCalled();
  });

  it('aborts between write chunks: ABORTED, no zip phase, no further progress', async () => {
    const controller = new AbortController();
    const phases: string[] = [];
    const error = await expectExportError(
      exportHighchartsOptionsToXlsx(bigOptions(), {
        signal: controller.signal,
        onProgress: (p) => {
          phases.push(p.phase);
          if (p.phase === 'write' && phases.filter((x) => x === 'write').length === 3) controller.abort();
        },
      }),
      'ABORTED',
    );
    expect((error.cause as DOMException).name).toBe('AbortError');
    expect(phases.filter((x) => x === 'write')).toHaveLength(3);
    expect(phases).not.toContain('zip');
    expect(phases).not.toContain('done');
  });

  it('checks the signal after a custom writer that ignores it', async () => {
    const controller = new AbortController();
    const writer = {
      name: 'ignores-signal',
      write: async () => {
        controller.abort('late');
        return new Uint8Array([1]);
      },
    };
    const error = await expectExportError(
      exportHighchartsToXlsx(render(F.simpleLine), { signal: controller.signal, writer }),
      'ABORTED',
    );
    expect(error.cause).toBe('late');
  });

  it('passes the signal and a progress callback to the writer', async () => {
    const seen: unknown[] = [];
    const writer = {
      name: 'spy',
      write: async (_spec: unknown, ctx?: unknown) => {
        seen.push(ctx);
        return new Uint8Array([1]);
      },
    };
    const controller = new AbortController();
    await exportHighchartsToXlsx(render(F.simpleLine), { signal: controller.signal, writer });
    expect(seen[0]).toMatchObject({ signal: controller.signal });
    expect(typeof (seen[0] as { onProgress?: unknown }).onProgress).toBe('function');
  });

  it('emits progress for every phase with a monotonic overall fraction from 0 to 1', async () => {
    const events: ExportProgress[] = [];
    const result = await exportHighchartsOptionsToXlsx(bigOptions(), { onProgress: (p) => events.push(p) });
    expect(result.bytes.byteLength).toBeGreaterThan(0);
    const phases = events.map((e) => e.phase);
    expect(phases[0]).toBe('extract');
    expect(events[0]!.fraction).toBe(0);
    expect([...new Set(phases)]).toEqual(['extract', 'translate', 'write', 'zip', 'done']);
    expect(phases.filter((p) => p === 'write').length).toBeGreaterThan(3);
    for (let i = 1; i < events.length; i++) expect(events[i]!.fraction).toBeGreaterThanOrEqual(events[i - 1]!.fraction);
    expect(events.at(-1)).toMatchObject({ phase: 'done', fraction: 1 });
    expect(events.filter((e) => e.phase === 'done')).toHaveLength(1);
  });

  it('reports phase boundaries for a custom writer that reports nothing', async () => {
    const events: ExportProgress[] = [];
    const writer = { name: 'quiet', write: async () => new Uint8Array([1]) };
    await exportHighchartsToXlsx(render(F.simpleLine), { writer, onProgress: (p) => events.push(p) });
    expect([...new Set(events.map((e) => e.phase))]).toEqual(['extract', 'translate', 'write', 'done']);
  });

  it('timings include zipMs, part of writeMs', async () => {
    const { timings } = await exportHighchartsOptionsToXlsx(bigOptions(2_000, 1));
    expect(timings.zipMs).toBeGreaterThan(0);
    expect(timings.zipMs).toBeLessThanOrEqual(timings.writeMs);
    expect(timings.imageMs).toBe(0);
    const quiet = await exportHighchartsToXlsx(render(F.simpleLine), {
      writer: { name: 'quiet', write: async () => new Uint8Array([1]) },
    });
    expect(quiet.timings.zipMs).toBe(0);
  });

  it('errors thrown by onProgress propagate unwrapped', async () => {
    const boom = new Error('progress callback failed');
    const thrown = await exportHighchartsOptionsToXlsx(bigOptions(), {
      onProgress: (p) => {
        if (p.phase === 'write' && p.fraction > 0.3) throw boom;
      },
    }).catch((e: unknown) => e);
    expect(thrown).toBe(boom);
  });

  it('validates signal and onProgress', async () => {
    for (const [property, value] of [
      ['signal', { aborted: 'no' }],
      ['signal', 'abort'],
      ['onProgress', 42],
    ] as const) {
      const error = await expectExportError(
        exportHighchartsOptionsToXlsx(F.simpleLine, { [property]: value } as unknown as ExportOptions),
        'INVALID_OPTIONS',
      );
      expect(error.details.property).toBe(property);
    }
  });

  it('exportChartsToWorkbook honours signal and onProgress', async () => {
    const entries = [{ chart: render(F.simpleLine) }, { chart: render(F.columnChart) }];
    const events: ExportProgress[] = [];
    const result = await exportChartsToWorkbook(entries, { onProgress: (p) => events.push(p) });
    expect(result.bytes.byteLength).toBeGreaterThan(0);
    expect(events.at(-1)).toMatchObject({ phase: 'done', fraction: 1 });
    expect(events.filter((e) => e.phase === 'extract')).toHaveLength(2);
    for (let i = 1; i < events.length; i++) expect(events[i]!.fraction).toBeGreaterThanOrEqual(events[i - 1]!.fraction);

    const controller = new AbortController();
    const error = await expectExportError(
      exportChartsToWorkbook(entries, {
        signal: controller.signal,
        onProgress: (p) => {
          if (p.phase === 'translate') controller.abort('stop');
        },
      }),
      'ABORTED',
    );
    expect(error.cause).toBe('stop');
  });

  it('downloadHighchartsAsXlsx does not download after an abort', async () => {
    const click = vi.spyOn(HTMLAnchorElement.prototype, 'click');
    const controller = new AbortController();
    const error = await expectExportError(
      downloadHighchartsAsXlsx(render(F.simpleLine), {
        signal: controller.signal,
        onProgress: (p) => {
          if (p.phase === 'done') controller.abort('too late');
        },
      }),
      'ABORTED',
    );
    expect(error.cause).toBe('too late');
    expect(click).not.toHaveBeenCalled();
  });
});
