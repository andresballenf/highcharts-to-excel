import { describe, it, expect, vi, afterEach } from 'vitest';
import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { installHighchartsExcelExport } from '../../../src/highcharts/install-export-menu';
import { exportHighchartsToXlsx, exportHighchartsOptionsToXlsx, analyzeChartCompatibility, exportChartsToWorkbook } from '../../../src/api/export';
import { DEFAULT_MENU_ITEM_KEY } from '../../../src/types/public-api';
import { renderChart, destroyAll } from '../../helpers/render-chart';
import { unzipSync, strFromU8 } from 'fflate';

const H = Highcharts as any;
import { appendFileSync } from 'node:fs';
const log = (...a: unknown[]) => appendFileSync('/tmp/claude-0/-home-user-highcharts-to-excel/7c299b7a-8c1c-5eca-81a1-ef08ceb87c2f/scratchpad/audit-a3/out.txt', a.map((x) => typeof x === 'string' ? x : JSON.stringify(x)).join(' ') + '\n');
afterEach(() => destroyAll());

describe('audit a3', () => {
  it('onclick: onExport throwing -> unhandled rejection', async () => {
    vi.stubGlobal('URL', Object.assign(Object.create(URL) as object, { createObjectURL: vi.fn(() => 'blob:x'), revokeObjectURL: vi.fn() }));
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
    const unhandled: unknown[] = [];
    const handler = (r: unknown) => unhandled.push(r);
    process.on('unhandledRejection', handler);
    const inst = installHighchartsExcelExport(H, {
      onExport: () => { throw new Error('onExport boom'); },
      onError: (e) => log('onError called with', (e as Error).message),
    });
    const chart = renderChart(H, { title: { text: 'T' }, series: [{ type: 'line', data: [1, 2, 3] }] } as any);
    const def = H.getOptions().exporting.menuItemDefinitions[DEFAULT_MENU_ITEM_KEY];
    def.onclick.call(chart);
    await new Promise((r) => setTimeout(r, 300));
    log('unhandled rejections:', unhandled.map((u) => (u as Error).message));
    process.off('unhandledRejection', handler);
    inst.uninstall();
    vi.unstubAllGlobals();
  });

  it('onWarning / transformModel throwing are not wrapped', async () => {
    const chart = renderChart(H, { title: { text: 'T' }, series: [{ type: 'line', data: [1, null, 3] }], tooltip: { formatter() { return 'x'; } } } as any);
    for (const opts of [
      { onWarning: () => { throw new TypeError('cb boom'); } },
      { hooks: { transformModel: () => { throw new RangeError('hook boom'); } } },
    ]) {
      try { await exportHighchartsToXlsx(chart, opts as any); log('no throw'); } catch (e: any) { log('threw', e.name, e.code, e.message); }
    }
    try { analyzeChartCompatibility(chart, { hooks: { transformModel: () => null as any } }); } catch (e: any) { log('analyze threw', e.code); }
  });

  it('onWarning count vs warnings', async () => {
    const chart = renderChart(H, { title: { text: 'T' }, series: [{ type: 'line', data: [1, null, 3] }, { type: 'line', visible: false, data: [1, 2, 3] }], tooltip: { formatter() { return 'x'; } } } as any);
    const calls: string[] = [];
    const r = await exportHighchartsToXlsx(chart, { onWarning: (d) => calls.push(d.code + '|' + d.property) });
    log('onWarning calls', calls.length, 'warnings', r.warnings.length, 'dupes', calls.length - new Set(calls).size);
  });

  it('filename derivation', async () => {
    for (const t of ['Sales <b>2024</b>', 'Line1<br>Line2', 'R&amp;D report', 'a/b\\c', 'Report.XLSX', '   ', 'CON', 'Ünïcödé 📈 chart', 'x'.repeat(200) + '📈']) {
      const r = await exportHighchartsOptionsToXlsx({ title: { text: t }, series: [{ type: 'line', data: [1, 2] }] });
      log(JSON.stringify(t.slice(0, 30)), '->', JSON.stringify(r.filename), r.filename.length);
    }
  });

  it('options validation', async () => {
    for (const o of [{ chartWidth: NaN }, { chartWidth: -100 }, { chartWidth: 1e7 }, { fidelity: 'bogus' }, { dataMode: 'bogus' }, { chartSheetName: 42 }]) {
      try { const r = await exportHighchartsOptionsToXlsx({ series: [{ type: 'line', data: [1, 2] }] }, o as any);
        const z = unzipSync(r.bytes); const d = Object.keys(z).find((k) => k.includes('drawings/drawing1.xml'));
        log(JSON.stringify(o, (k, v) => (Number.isNaN(v) ? 'NaN' : v)), 'ok', d ? strFromU8(z[d]!).match(/<xdr:ext[^>]*>|<xdr:to>.*?<\/xdr:to>/)?.[0] : '');
      } catch (e: any) { log(JSON.stringify(o), 'threw', e.code ?? e.name, e.message.slice(0, 100)); }
    }
  });

  it('multi: same chart/data sheet name, per-entry writer ignored', async () => {
    const chart = renderChart(H, { series: [{ type: 'line', data: [1, 2] }] } as any);
    let used = 0;
    const writer = { name: 'x', write: async () => { used++; return new Uint8Array([1]); } };
    const r = await exportChartsToWorkbook([
      { chart, options: { chartSheetName: 'Same', dataSheetName: 'same', writer } as any },
      { chart, options: { chartSheetName: 'Same' } },
    ]);
    log('sheets', r.charts.map((c) => `${c.chartSheetName}/${c.dataSheetName}`), 'per-entry writer used', used);
  });

  it('prototype pollution via editableExcel', async () => {
    const cfg = JSON.parse('{"__proto__":{"polluted":"yes"},"filename":"p"}');
    const merged = { ...{}, ...cfg };
    log('polluted?', ({} as any).polluted, 'merged own __proto__?', Object.prototype.hasOwnProperty.call(merged, '__proto__'));
  });

  it('perf: big workbook timings', async () => {
    const n = 100_000;
    const series = Array.from({ length: 5 }, (_, s) => ({ type: 'line', data: Array.from({ length: n }, (_, i) => Math.sin(i / 100 + s) * 100) }));
    const r = await exportHighchartsOptionsToXlsx({ series } as any);
    log('bytes', r.bytes.length, 'timings', JSON.stringify(r.timings));
  });
});
