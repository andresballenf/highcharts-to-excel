import { describe, expect, it, vi } from 'vitest';
import { resolveChartType } from '../../src/core/chart-type-registry';
import { buildDataLayout, checkLimits, type DataLayout } from '../../src/core/data-layout';
import type { ChartModel } from '../../src/types/chart-model';
import type { CellValue, SheetSpec } from '../../src/excel/writer-interface';
import * as F from '../fixtures/chart-models';

function layoutOf(model: ChartModel, sheetName = 'Data', dateFormatCode = () => 'yyyy-mm-dd'): DataLayout {
  return buildDataLayout(model, resolveChartType(model), { sheetName, hidden: false, dateFormatCode });
}

/** Cell value at (col0,row0); undefined when no cell was written (blank). */
function cell(sheet: SheetSpec, col0: number, row0: number): CellValue | undefined {
  return sheet.rows.find((r) => r.row0 === row0)?.cells.find((c) => c.col0 === col0)?.value;
}

function headerRow(sheet: SheetSpec): string[] {
  return (sheet.rows.find((r) => r.row0 === 0)?.cells ?? []).map((c) => (c.value.type === 'string' ? c.value.value : '?'));
}

describe('category layout', () => {
  it('writes categories in column A and one column per series', () => {
    const l = layoutOf(F.lineModel());
    expect(l.blocking).toBe(false);
    expect(l.rowCount).toBe(4);
    expect(headerRow(l.sheet)).toEqual(['Category', 'Sales']);
    expect(l.sheet.rows[0]!.cells[0]!.style).toEqual({ bold: true, fillHex: 'F2F2F2' });
    expect(l.sheet.freezeHeaderRow).toBe(true);
    expect(cell(l.sheet, 0, 1)).toEqual({ type: 'string', value: 'Jan' });
    expect(cell(l.sheet, 1, 4)).toEqual({ type: 'number', value: 4 });
    const r = l.ranges[0]!;
    expect(r.nameRef).toEqual({ formula: 'Data!$B$1', cache: 'Sales' });
    expect(r.categories).toEqual({ formula: 'Data!$A$2:$A$5', kind: 'str', cache: ['Jan', 'Feb', 'Mar', 'Apr'] });
    expect(r.values).toEqual({ formula: 'Data!$B$2:$B$5', cache: [1, 2, 3, 4], formatCode: 'General' });
    expect(r.bubbleSizes).toBeNull();
    expect(r.pointOffsets).toEqual([0, 1, 2, 3]);
  });

  it('quotes sheet names that need it', () => {
    const l = layoutOf(F.lineModel(), 'My Data');
    expect(l.ranges[0]!.values.formula).toBe("'My Data'!$B$2:$B$5");
    expect(l.ranges[0]!.nameRef.formula).toBe("'My Data'!$B$1");
  });

  it('builds the sorted union of numeric x values with blanks for missing keys', () => {
    const l = layoutOf(F.multiSeriesLineModel());
    expect(l.rowCount).toBe(4);
    expect(l.ranges[0]!.categories).toMatchObject({ kind: 'num', cache: [1, 2, 3, 4] });
    expect(l.ranges[0]!.values.cache).toEqual([10, 11, null, 12]);
    expect(l.ranges[1]!.values.cache).toEqual([null, 20, 21, 22]);
    expect(cell(l.sheet, 1, 3)).toBeUndefined();
    expect(cell(l.sheet, 2, 1)).toBeUndefined();
    const unaligned = l.diagnostics.filter((d) => d.code === 'UNALIGNED_X_VALUES');
    expect(unaligned.map((d) => [d.seriesIndex, d.outcome, d.severity])).toEqual([
      [0, 'approximated', 'info'],
      [1, 'approximated', 'info'],
    ]);
  });

  it('builds the union of category labels when points are named instead of indexed', () => {
    const m = F.lineModel();
    m.xAxes = [F.axis({ index: 0, kind: 'category' })];
    m.series = [
      F.series({ kind: 'line', index: 0, name: 'A', points: ['x', 'y'].map((name, i) => F.point({ name, y: i })) }),
      F.series({ kind: 'line', index: 1, name: 'B', points: ['y', 'z'].map((name, i) => F.point({ name, y: 10 + i })) }),
    ];
    const l = layoutOf(m);
    expect(l.ranges[0]!.categories!.cache).toEqual(['x', 'y', 'z']);
    expect(l.ranges[0]!.values.cache).toEqual([0, 1, null]);
    expect(l.ranges[1]!.values.cache).toEqual([null, 10, 11]);
  });

  it('writes nulls as blank cells and keeps negative numbers', () => {
    const l = layoutOf(F.nullNegativeModel());
    expect(l.ranges[0]!.values.cache).toEqual([5, null, -3, 7]);
    expect(cell(l.sheet, 1, 2)).toBeUndefined();
    expect(cell(l.sheet, 1, 3)).toEqual({ type: 'number', value: -3 });
    const d = l.diagnostics.find((x) => x.code === 'NULL_VALUES')!;
    expect(d).toMatchObject({ severity: 'info', property: 'series[0].data', seriesIndex: 0 });
    expect(d.message).toContain('1 null values exported as empty cells');
  });

  it('converts datetime x values to Excel serials with the provided date format', () => {
    const fmt = vi.fn(() => 'yyyy-mm-dd');
    const l = layoutOf(F.datetimeModel(), 'Data', fmt);
    expect(fmt).toHaveBeenCalledWith(expect.objectContaining({ kind: 'datetime' }), 3 * F.DAY);
    expect(headerRow(l.sheet)[0]).toBe('Date');
    expect(cell(l.sheet, 0, 1)).toEqual({ type: 'number', value: 45292 });
    expect(l.sheet.rows[1]!.cells[0]!.style).toEqual({ numberFormat: 'yyyy-mm-dd' });
    expect(l.ranges[0]!.categories).toEqual({
      formula: 'Data!$A$2:$A$5',
      kind: 'num',
      cache: [45292, 45293, 45294, 45295],
      formatCode: 'yyyy-mm-dd',
    });
    expect(l.x).toMatchObject({ kind: 'datetime', dateFormatCode: 'yyyy-mm-dd', min: Date.UTC(2024, 0, 1) });
  });

  it('uses the series y format as value format code and cell style', () => {
    const m = F.lineModel();
    m.series[0] = { ...m.series[0]!, yFormat: { kind: 'excel', code: '0.00' } };
    const l = layoutOf(m);
    expect(l.ranges[0]!.values.formatCode).toBe('0.00');
    expect(l.sheet.rows[1]!.cells[1]!.style).toEqual({ numberFormat: '0.00' });
  });

  it('reports hidden series that reach the layout', () => {
    const l = layoutOf(F.hiddenSeriesModel());
    expect(headerRow(l.sheet)).toEqual(['Category', 'Shown', 'Hidden']);
    expect(l.diagnostics.find((d) => d.code === 'HIDDEN_SERIES_INCLUDED')).toMatchObject({ seriesIndex: 1, property: 'series[1].visible' });
  });

  it('sizes columns between 10 and 40 characters', () => {
    const m = F.lineModel();
    m.series[0] = { ...m.series[0]!, name: 'A very long series name that goes on and on and on' };
    const l = layoutOf(m);
    expect(l.sheet.columns).toEqual([
      { col0: 0, widthChars: 10 },
      { col0: 1, widthChars: 40 },
    ]);
  });
});

describe('pie / doughnut layout', () => {
  it('writes point names and values', () => {
    const l = layoutOf(F.pieModel());
    expect(headerRow(l.sheet)).toEqual(['Category', 'Share']);
    expect(l.ranges[0]!.categories).toEqual({ formula: 'Data!$A$2:$A$5', kind: 'str', cache: ['Chrome', 'Edge', 'Firefox', 'Safari'] });
    expect(l.ranges[0]!.values.cache).toEqual([60, 15, 15, 10]);
  });

  it('falls back to "Point N" for unnamed slices', () => {
    const m = F.pieModel();
    m.series[0] = { ...m.series[0]!, points: [F.point({ y: 1 }), F.point({ y: 2 })] };
    expect(layoutOf(m).ranges[0]!.categories!.cache).toEqual(['Point 1', 'Point 2']);
  });

  it('unions ring categories and maps every point to its row', () => {
    const l = layoutOf(F.doughnutModel());
    expect(headerRow(l.sheet)).toEqual(['Category', 'Inner', 'Outer']);
    expect(l.ranges[0]!.categories!.cache).toEqual(['A', 'B', 'C']);
    expect(l.ranges[0]!.values.cache).toEqual([1, 2, null]);
    expect(l.ranges[1]!.values.cache).toEqual([5, null, 6]);
    expect(l.ranges[1]!.pointOffsets).toEqual([0, 2]);
  });
});

describe('scatter / bubble layout', () => {
  it('writes one X/Y block per series with independent rows', () => {
    const l = layoutOf(F.scatterModel());
    expect(headerRow(l.sheet)).toEqual(['Men X', 'Men Y', 'Women X', 'Women Y']);
    expect(l.rowCount).toBe(3);
    const [men, women] = l.ranges;
    expect(men!.categories).toEqual({ formula: 'Data!$A$2:$A$4', kind: 'num', cache: [1, 3, 5] });
    expect(men!.values).toEqual({ formula: 'Data!$B$2:$B$4', cache: [2, 4, 6], formatCode: 'General' });
    expect(women!.categories!.formula).toBe('Data!$C$2:$C$2');
    expect(women!.values.cache).toEqual([2.5]);
    expect(women!.literalName).toBe('Women');
    expect(cell(l.sheet, 2, 2)).toBeUndefined();
  });

  it('adds a Size column for bubbles', () => {
    const l = layoutOf(F.bubbleModel());
    expect(headerRow(l.sheet)).toEqual(['Countries X', 'Countries Y', 'Countries Size']);
    expect(l.ranges[0]!.bubbleSizes).toEqual({ formula: 'Data!$C$2:$C$3', cache: [10, 20], formatCode: 'General' });
  });

  it('converts datetime scatter x values to serials', () => {
    const m = F.scatterModel();
    m.xAxes = [F.axis({ index: 0, kind: 'datetime' })];
    m.series = [F.series({ kind: 'scatter', index: 0, name: 'S', points: [F.point({ x: Date.UTC(2024, 0, 1), y: 1 })] })];
    const l = layoutOf(m, 'Data', () => 'dd/mm/yyyy');
    expect(l.ranges[0]!.categories).toEqual({ formula: 'Data!$A$2:$A$2', kind: 'num', cache: [45292], formatCode: 'dd/mm/yyyy' });
  });

  it('puts scatter blocks to the right of the shared category columns in a combo', () => {
    const m = F.comboModel();
    m.series.push(F.series({ kind: 'scatter', index: 2, name: 'Pts', points: [F.point({ x: 0.5, y: 3 })] }));
    const l = layoutOf(m);
    expect(headerRow(l.sheet)).toEqual(['Category', 'Rainfall', 'Average', 'Pts X', 'Pts Y']);
    expect(l.ranges.map((r) => r.values.formula)).toEqual(['Data!$B$2:$B$5', 'Data!$C$2:$C$5', 'Data!$E$2:$E$2']);
  });
});

describe('text safety', () => {
  it('keeps formula-like strings as string cells and reports once', () => {
    const l = layoutOf(F.formulaLikeModel());
    expect(cell(l.sheet, 0, 1)).toEqual({ type: 'string', value: '=SUM(A1)' });
    expect(cell(l.sheet, 1, 0)).toEqual({ type: 'string', value: '+foo' });
    const d = l.diagnostics.filter((x) => x.code === 'FORMULA_LIKE_TEXT_ESCAPED');
    expect(d).toHaveLength(1);
    expect(d[0]!.severity).toBe('info');
  });

  it('strips control characters from text', () => {
    const m = F.lineModel();
    m.series[0] = { ...m.series[0]!, name: 'Sa\u0007les' };
    expect(layoutOf(m).ranges[0]!.nameRef.cache).toBe('Sales');
  });
});

describe('limits', () => {
  it('blocks when rows exceed the worksheet', () => {
    const r = checkLimits(1_048_576, 2, 0);
    expect(r.blocking).toBe(true);
    expect(r.diagnostics).toEqual([expect.objectContaining({ code: 'ROW_LIMIT_EXCEEDED', outcome: 'blocking' })]);
    expect(checkLimits(1_048_575, 2, 0).blocking).toBe(false);
  });

  it('blocks when columns exceed the worksheet', () => {
    const r = checkLimits(10, 16_385, 0);
    expect(r.blocking).toBe(true);
    expect(r.diagnostics[0]!.code).toBe('COLUMN_LIMIT_EXCEEDED');
  });

  it('warns (not blocking) above 32,000 points per series', () => {
    const r = checkLimits(40_000, 2, 40_000, 3);
    expect(r.blocking).toBe(false);
    expect(r.diagnostics).toEqual([
      expect.objectContaining({ code: 'ROW_LIMIT_EXCEEDED', outcome: 'approximated', seriesIndex: 3, property: 'series[3].data' }),
    ]);
    expect(r.diagnostics[0]!.message).toContain('32,000');
  });

  it('lays out 40,000 points with a per-series warning', () => {
    const l = layoutOf(F.hugeModel());
    expect(l.blocking).toBe(false);
    expect(l.rowCount).toBe(40_000);
    expect(l.ranges[0]!.values.formula).toBe('Data!$B$2:$B$40001');
    expect(l.ranges[0]!.values.cache).toHaveLength(40_000);
    expect(l.diagnostics.find((d) => d.code === 'ROW_LIMIT_EXCEEDED')).toMatchObject({ outcome: 'approximated', seriesIndex: 0 });
  });
});
