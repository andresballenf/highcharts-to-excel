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
  return (sheet.rows.find((r) => r.row0 === 0)?.cells ?? []).map((c) =>
    c.value.type === 'string' ? c.value.value : '?',
  );
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
    // Alignment gaps (no point at that x) are #N/A cells so Excel lines stay connected.
    expect(cell(l.sheet, 1, 3)).toEqual({ type: 'error', value: '#N/A' });
    expect(cell(l.sheet, 2, 1)).toEqual({ type: 'error', value: '#N/A' });
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
      F.series({
        kind: 'line',
        index: 1,
        name: 'B',
        points: ['y', 'z'].map((name, i) => F.point({ name, y: 10 + i })),
      }),
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
    expect(l.diagnostics.find((d) => d.code === 'HIDDEN_SERIES_INCLUDED')).toMatchObject({
      seriesIndex: 1,
      property: 'series[1].visible',
    });
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
    expect(l.ranges[0]!.categories).toEqual({
      formula: 'Data!$A$2:$A$5',
      kind: 'str',
      cache: ['Chrome', 'Edge', 'Firefox', 'Safari'],
    });
    expect(l.ranges[0]!.values.cache).toEqual([60, 15, 15, 10]);
  });

  it('falls back to "Point N" for unnamed slices', () => {
    const m = F.pieModel();
    m.series[0] = { ...m.series[0]!, points: [F.point({ y: 1 }), F.point({ y: 2 })] };
    expect(layoutOf(m).ranges[0]!.categories!.cache).toEqual(['Point 1', 'Point 2']);
  });

  it('gives rings with different slices their own category columns', () => {
    const l = layoutOf(F.doughnutModel());
    expect(headerRow(l.sheet)).toEqual(['Category 1', 'Inner', 'Category 2', 'Outer']);
    expect(l.ranges[0]!.categories!.cache).toEqual(['A', 'B']);
    expect(l.ranges[1]!.categories!.cache).toEqual(['A', 'C']);
    expect(l.ranges[0]!.values.cache).toEqual([1, 2]);
    expect(l.ranges[1]!.values.cache).toEqual([5, 6]);
    expect(l.ranges[1]!.pointOffsets).toEqual([0, 1]);
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
    m.series = [
      F.series({ kind: 'scatter', index: 0, name: 'S', points: [F.point({ x: Date.UTC(2024, 0, 1), y: 1 })] }),
    ];
    const l = layoutOf(m, 'Data', () => 'dd/mm/yyyy');
    expect(l.ranges[0]!.categories).toEqual({
      formula: 'Data!$A$2:$A$2',
      kind: 'num',
      cache: [45292],
      formatCode: 'dd/mm/yyyy',
    });
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
      expect.objectContaining({
        code: 'ROW_LIMIT_EXCEEDED',
        outcome: 'approximated',
        seriesIndex: 3,
        property: 'series[3].data',
      }),
    ]);
    expect(r.diagnostics[0]!.message).toContain('32,000');
  });

  it('lays out 40,000 points with a per-series warning', () => {
    const l = layoutOf(F.hugeModel());
    expect(l.blocking).toBe(false);
    expect(l.rowCount).toBe(40_000);
    expect(l.ranges[0]!.values.formula).toBe('Data!$B$2:$B$40001');
    expect(l.ranges[0]!.values.cache).toHaveLength(40_000);
    expect(l.diagnostics.find((d) => d.code === 'ROW_LIMIT_EXCEEDED')).toMatchObject({
      outcome: 'approximated',
      seriesIndex: 0,
    });
  });
});

describe('audit fixes: alignment, label mode, duplicates, pre-1900 dates, doughnut rings', () => {
  const NA: CellValue = { type: 'error', value: '#N/A' };

  it('C1: unaligned x → #N/A cells (null in the chart cache); a real null stays blank', () => {
    const m = F.baseModel();
    m.xAxes = [F.axis({ index: 0, kind: 'linear' })];
    m.series = [
      F.series({
        kind: 'line',
        index: 0,
        points: [0, 2, 4, 6].map((x) => F.point({ x, y: x === 4 ? null : x, isNull: x === 4 })),
      }),
      F.series({ kind: 'line', index: 1, points: [1, 3, 5].map((x) => F.point({ x, y: x })) }),
    ];
    const l = layoutOf(m);
    expect(l.x.keys).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(cell(l.sheet, 1, 2)).toEqual(NA); // B3: series 0 has no point at x=1
    expect(cell(l.sheet, 1, 5)).toBeUndefined(); // B6: real null at x=4 stays blank (gap)
    expect(cell(l.sheet, 2, 1)).toEqual(NA); // C2: series 1 has no point at x=0
    expect(l.ranges[0]!.values.cache).toEqual([0, null, 2, null, null, null, 6]);
    expect(l.ranges[1]!.values.cache).toEqual([null, 1, null, 3, null, 5, null]);
    const d = l.diagnostics.find((x) => x.code === 'UNALIGNED_X_VALUES' && x.seriesIndex === 0)!;
    expect(d.message).toContain('written as #N/A so lines stay connected');
    expect(l.diagnostics.find((x) => x.code === 'NULL_VALUES')!.message).toContain('empty cells');
  });

  it('C2: label mode never maps a numeric x onto a category by array position', () => {
    const m = F.baseModel();
    m.xAxes = [F.axis({ index: 0, kind: 'category', categories: ['A', 'B', 'C'] })];
    m.series = [
      F.series({ kind: 'line', index: 0, points: [0, 1, 2].map((x) => F.point({ x, y: 10 + x })) }),
      F.series({
        kind: 'line',
        index: 1,
        points: [F.point({ x: -1, y: 99 }), F.point({ x: 0.5, y: 77 }), F.point({ x: 2, y: 55 })],
      }),
    ];
    const l = layoutOf(m);
    expect(l.x.keys).toEqual(['A', 'B', 'C', '-1', '0.5']);
    expect(l.ranges[1]!.values.cache).toEqual([null, null, 55, 99, 77]);
    expect(l.ranges[0]!.values.cache).toEqual([10, 11, 12, null, null]);
    const d = l.diagnostics.find(
      (x) => x.code === 'UNALIGNED_X_VALUES' && x.seriesIndex === 1 && x.severity === 'warning',
    )!;
    expect(d).toMatchObject({ outcome: 'approximated', severity: 'warning' });
    expect(d.message).toContain('-1');
    expect(d.message).toContain('0.5');
  });

  it('C2: x beyond the categories becomes its own row', () => {
    const m = F.baseModel();
    m.xAxes = [F.axis({ index: 0, kind: 'category', categories: ['A', 'B'] })];
    m.series = [
      F.series({
        kind: 'column',
        index: 0,
        points: [F.point({ x: 0, y: 1 }), F.point({ x: 1, y: 2 }), F.point({ x: 3, y: 4 })],
      }),
      F.series({ kind: 'column', index: 1, points: [F.point({ x: 0.5, y: 5 })] }),
    ];
    const l = layoutOf(m);
    expect(l.x.keys).toEqual(['A', 'B', '3', '0.5']);
    expect(l.ranges[0]!.values.cache).toEqual([1, 2, 4, null]);
    expect(l.ranges[1]!.values.cache).toEqual([null, null, null, 5]);
  });

  it('C3: duplicate x within one series keeps the first value and is reported', () => {
    const m = F.baseModel();
    m.series = [
      F.series({
        kind: 'line',
        index: 0,
        points: [F.point({ x: 0, y: 1 }), F.point({ x: 0, y: 2 }), F.point({ x: 1, y: 3 })],
      }),
    ];
    const l = layoutOf(m);
    expect(l.ranges[0]!.values.cache).toEqual([1, 3]);
    expect(l.ranges[0]!.pointOffsets).toEqual([0, -1, 1]);
    const d = l.diagnostics.find((x) => x.code === 'UNALIGNED_X_VALUES' && x.details?.duplicates !== undefined)!;
    expect(d).toMatchObject({
      outcome: 'approximated',
      severity: 'warning',
      seriesIndex: 0,
      property: 'series[0].data[1]',
      details: { duplicates: 1 },
    });
  });

  it('C6: datetime x before 1899-12-31 falls back to ISO date category strings', () => {
    const m = F.baseModel();
    m.xAxes = [F.axis({ index: 0, kind: 'datetime' })];
    m.series = [
      F.series({
        kind: 'line',
        index: 0,
        points: [Date.UTC(1899, 0, 1), Date.UTC(1900, 0, 1), Date.UTC(1900, 1, 28), Date.UTC(1900, 2, 1)].map((x, i) =>
          F.point({ x, y: i }),
        ),
      }),
    ];
    const l = layoutOf(m);
    expect(l.x.kind).toBe('category');
    expect(cell(l.sheet, 0, 1)).toEqual({ type: 'string', value: '1899-01-01' });
    expect(l.ranges[0]!.categories).toMatchObject({
      kind: 'str',
      cache: ['1899-01-01', '1900-01-01', '1900-02-28', '1900-03-01'],
    });
    expect(l.diagnostics.find((x) => x.code === 'APPROXIMATED_DATETIME')).toMatchObject({
      outcome: 'approximated',
      severity: 'warning',
    });
  });

  it('C6: 1900 dates after the epoch keep Excel serials (leap-year bug applied)', () => {
    const m = F.baseModel();
    m.xAxes = [F.axis({ index: 0, kind: 'datetime' })];
    m.series = [
      F.series({
        kind: 'line',
        index: 0,
        points: [Date.UTC(1900, 0, 1), Date.UTC(1900, 1, 28), Date.UTC(1900, 2, 1)].map((x, i) => F.point({ x, y: i })),
      }),
    ];
    const l = layoutOf(m);
    expect(l.x.kind).toBe('datetime');
    expect(l.ranges[0]!.categories!.cache).toEqual([1, 59, 61]);
    expect(l.diagnostics.some((x) => x.code === 'APPROXIMATED_DATETIME')).toBe(false);
  });

  it('C6: pre-1900 datetime scatter keeps day numbers (no date format) on a value axis', () => {
    const m = F.baseModel();
    m.xAxes = [F.axis({ index: 0, kind: 'datetime' })];
    m.series = [
      F.series({
        kind: 'scatter',
        index: 0,
        points: [Date.UTC(1899, 11, 30), Date.UTC(1900, 0, 1)].map((x, i) => F.point({ x, y: i })),
      }),
    ];
    const l = layoutOf(m);
    expect(l.x.kind).toBe('linear');
    expect(cell(l.sheet, 0, 1)).toEqual({ type: 'number', value: -1 });
    expect(l.ranges[0]!.categories).toEqual({ formula: 'Data!$A$2:$A$3', kind: 'num', cache: [-1, 1] });
  });

  it('C9: doughnut rings with different key sets keep their own slices (no union, no diagnostic)', () => {
    const m = F.baseModel();
    m.xAxes = [];
    m.yAxes = [];
    m.series = [
      F.series({ kind: 'doughnut', index: 0, points: [F.point({ name: 'A', y: 1 }), F.point({ name: 'B', y: 2 })] }),
      F.series({
        kind: 'doughnut',
        index: 1,
        points: [F.point({ name: 'B1', y: 1 }), F.point({ name: 'A', y: 5 }), F.point({ name: 'C', y: 2 })],
      }),
    ];
    const l = layoutOf(m);
    expect(l.ranges[0]!.categories!.cache).toEqual(['A', 'B']);
    expect(l.ranges[1]!.categories!.cache).toEqual(['B1', 'A', 'C']);
    expect(l.ranges[1]!.values.cache).toEqual([1, 5, 2]);
    expect(l.rowCount).toBe(3);
    expect(l.diagnostics.some((x) => x.code === 'APPROXIMATED_LAYOUT')).toBe(false);
    expect(layoutOf(F.doughnutModel()).diagnostics.some((x) => x.code === 'APPROXIMATED_LAYOUT')).toBe(false);
  });
});
