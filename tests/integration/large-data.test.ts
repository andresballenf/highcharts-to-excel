/**
 * Large-data behaviour: the 3 × 1,000 fixture, a 1 × 40,000-point series that exceeds Excel's
 * 32,000-points-per-series guidance (warning, still exported), and the hard worksheet limits
 * (1,048,576 rows / 16,384 columns), which are tested through `checkLimits` directly because
 * rendering a million-point chart in jsdom is impractical.
 */

import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { afterEach, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx } from '../../src/index';
import { checkLimits } from '../../src/core/data-layout';
import { EXCEL_MAX_COLUMNS, EXCEL_MAX_POINTS_PER_SERIES, EXCEL_MAX_ROWS } from '../../src/excel/writer-interface';
import * as F from '../fixtures/highcharts-options';
import { inspectXlsx } from '../helpers/inspect-xlsx';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';
import { valuesAtRange } from './export-fixtures.shared';

const H = Highcharts as unknown as HighchartsLike;

describe('large data', () => {
  afterEach(() => destroyAll());

  it('largeChart (3 × 1,000) exports 1,001 rows incl. header, without blocking', async () => {
    const chart = renderChart(H, F.largeChart);
    const result = await exportHighchartsToXlsx(chart);
    expect(result.report.editable).toBe(true);
    expect(result.warnings.filter((w) => w.outcome === 'blocking')).toEqual([]);
    expect(result.warnings.map((w) => w.code)).not.toContain('ROW_LIMIT_EXCEEDED');
    const x = await inspectXlsx(result.bytes);
    x.assertWellFormed();
    const data = x.sheetPath('Data');
    expect(x.text(data)).toContain('<dimension ref="A1:D1001"/>');
    expect(x.cellValue(data, 'B1001')).not.toBeNull();
    expect(x.cellValue(data, 'A1002')).toBeNull();
    const fs = x.seriesFormulas(x.chartPaths()[0]!);
    expect(fs.map((f) => f.val)).toEqual(['Data!$B$2:$B$1001', 'Data!$C$2:$C$1001', 'Data!$D$2:$D$1001']);
    const expected0 = (F.largeChart.series![0] as { data: number[] }).data;
    expect(valuesAtRange(x, fs[0]!.val!)).toEqual(expected0);
  });

  it('1 × 40,000 points: ROW_LIMIT_EXCEEDED approximated (Excel 32,000/series guidance), still exported', async () => {
    const n = 40_000;
    const data = Array.from({ length: n }, (_, i) => Math.round(Math.sin(i / 100) * 1000) / 10);
    const chart = renderChart(H, { title: { text: 'Forty thousand' }, series: [{ type: 'line', name: 'Big', data }] });
    const result = await exportHighchartsToXlsx(chart);
    const limit = result.warnings.filter((w) => w.code === 'ROW_LIMIT_EXCEEDED');
    expect(limit).toHaveLength(1);
    expect(limit[0]!.outcome).toBe('approximated');
    expect(limit[0]!.details).toMatchObject({ points: n, limit: EXCEL_MAX_POINTS_PER_SERIES });
    expect(result.report.editable).toBe(true);
    expect(result.report.blocking).toEqual([]);
    const x = await inspectXlsx(result.bytes);
    const fs = x.seriesFormulas(x.chartPaths()[0]!);
    expect(fs[0]!.val).toBe(`Data!$B$2:$B$${n + 1}`);
    const sheet = x.sheetPath('Data');
    expect(x.cellValue(sheet, `B${n + 1}`)).toBe(data[n - 1]);
  }, 60_000);

  describe('checkLimits (worksheet hard limits)', () => {
    it('allows exactly 1,048,576 rows including the header', () => {
      const r = checkLimits(EXCEL_MAX_ROWS - 1, 2, 1000);
      expect(r.blocking).toBe(false);
      expect(r.diagnostics).toEqual([]);
    });

    it('blocks one row more', () => {
      const r = checkLimits(EXCEL_MAX_ROWS, 2, 1000);
      expect(r.blocking).toBe(true);
      expect(r.diagnostics.map((d) => [d.code, d.outcome])).toEqual([['ROW_LIMIT_EXCEEDED', 'blocking']]);
      expect(r.diagnostics[0]!.details).toMatchObject({ rows: EXCEL_MAX_ROWS + 1, limit: EXCEL_MAX_ROWS });
    });

    it('blocks more than 16,384 columns', () => {
      expect(checkLimits(10, EXCEL_MAX_COLUMNS, 10).blocking).toBe(false);
      const r = checkLimits(10, EXCEL_MAX_COLUMNS + 1, 10);
      expect(r.blocking).toBe(true);
      expect(r.diagnostics.map((d) => d.code)).toEqual(['COLUMN_LIMIT_EXCEEDED']);
    });

    it('warns (not blocks) above 32,000 points per series, naming the series', () => {
      expect(checkLimits(EXCEL_MAX_POINTS_PER_SERIES, 2, EXCEL_MAX_POINTS_PER_SERIES).diagnostics).toEqual([]);
      const r = checkLimits(EXCEL_MAX_POINTS_PER_SERIES + 1, 2, EXCEL_MAX_POINTS_PER_SERIES + 1, 3);
      expect(r.blocking).toBe(false);
      expect(r.diagnostics).toHaveLength(1);
      expect(r.diagnostics[0]).toMatchObject({
        code: 'ROW_LIMIT_EXCEEDED',
        outcome: 'approximated',
        seriesIndex: 3,
        property: 'series[3].data',
      });
    });
  });
});
