/**
 * Data sheet layout: decides where every value lives on the worksheet and produces the
 * sheet-qualified absolute references the chart series point at.
 */

import type { AxisKind, AxisModel, ChartModel, PointModel, SeriesModel } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import type { CellSpec, CellStyleSpec, CellValue, RowSpec, SheetSpec } from '../excel/writer-interface';
import { EXCEL_MAX_COLUMNS, EXCEL_MAX_POINTS_PER_SERIES, EXCEL_MAX_ROWS } from '../excel/writer-interface';
import { msToExcelSerial } from '../utils/dates';
import { cellRef, quoteSheetNameForFormula, rangeRef } from '../utils/filenames';
import { isFormulaLike, stripControlChars } from '../utils/text';
import { isCategoryGroupKind, type ChartTypeResolution } from './chart-type-registry';

export interface SeriesRange {
  /** Position of the series in `model.series`. */
  seriesIndex: number;
  nameRef: { formula: string; cache: string };
  categories: { formula: string; kind: 'str' | 'num'; cache: (string | number | null)[]; formatCode?: string } | null;
  values: { formula: string; cache: (number | null)[]; formatCode: string };
  bubbleSizes: { formula: string; cache: (number | null)[]; formatCode: string } | null;
  /**
   * For every point (by position in `series.points`), its offset inside `values.cache`
   * (= the Excel data point index), or -1 when the point was not written.
   */
  pointOffsets: number[];
  /**
   * Scatter/bubble blocks have no cell holding the bare series name (headers are "<name> X",
   * "<name> Y"), so the chart uses this literal name instead of `nameRef`. Null elsewhere.
   */
  literalName: string | null;
}

export interface DataLayoutXInfo {
  /** Kind of the x axis used for the shared category column ('none' when there is no category column). */
  kind: AxisKind | 'none';
  /** Raw (source) keys of the category rows: ms timestamps / numbers for numeric axes, labels otherwise. */
  keys: Array<string | number>;
  /** Excel date format used for datetime x values (category column and scatter X columns). */
  dateFormatCode: string | null;
  /** Min/max x in source units (ms for datetime) over every exported series. */
  min: number | null;
  max: number | null;
}

export interface DataLayout {
  sheet: SheetSpec;
  ranges: SeriesRange[];
  /** Number of data rows (header excluded). */
  rowCount: number;
  diagnostics: Diagnostic[];
  blocking: boolean;
  x: DataLayoutXInfo;
}

export interface DataLayoutOptions {
  sheetName: string;
  hidden: boolean;
  /** Excel date format for a datetime x axis spanning `spanMs`. */
  dateFormatCode: (axis: AxisModel, spanMs: number) => string;
}

const HEADER_STYLE: CellStyleSpec = { bold: true, fillHex: 'F2F2F2' };

/**
 * Checks the Excel worksheet/chart limits. `rowCount` excludes the header row.
 * Exceeding rows/columns is blocking; a series longer than Excel's per-series guidance is a warning.
 */
export function checkLimits(
  rowCount: number,
  colCount: number,
  maxPointsPerSeries: number,
  seriesIndex?: number,
): { diagnostics: Diagnostic[]; blocking: boolean } {
  const diagnostics: Diagnostic[] = [];
  let blocking = false;
  if (rowCount + 1 > EXCEL_MAX_ROWS) {
    blocking = true;
    diagnostics.push(
      createDiagnostic('ROW_LIMIT_EXCEEDED', 'blocking', 'series.data', `The data needs ${rowCount + 1} rows; an Excel worksheet holds at most ${EXCEL_MAX_ROWS}.`, {
        details: { rows: rowCount + 1, limit: EXCEL_MAX_ROWS },
      }),
    );
  }
  if (colCount > EXCEL_MAX_COLUMNS) {
    blocking = true;
    diagnostics.push(
      createDiagnostic('COLUMN_LIMIT_EXCEEDED', 'blocking', 'series', `The data needs ${colCount} columns; an Excel worksheet holds at most ${EXCEL_MAX_COLUMNS}.`, {
        details: { columns: colCount, limit: EXCEL_MAX_COLUMNS },
      }),
    );
  }
  if (maxPointsPerSeries > EXCEL_MAX_POINTS_PER_SERIES) {
    const extra: Parameters<typeof createDiagnostic>[4] = { details: { points: maxPointsPerSeries, limit: EXCEL_MAX_POINTS_PER_SERIES } };
    if (seriesIndex !== undefined) extra.seriesIndex = seriesIndex;
    diagnostics.push(
      createDiagnostic(
        'ROW_LIMIT_EXCEEDED',
        'approximated',
        seriesIndex !== undefined ? `series[${seriesIndex}].data` : 'series.data',
        `A series references ${maxPointsPerSeries} points; Excel recommends at most 32,000 points per chart series and may truncate or render slowly (all values are still on the data sheet).`,
        extra,
      ),
    );
  }
  return { diagnostics, blocking };
}

/** Text written to cells: control characters stripped. */
function cleanText(s: string): string {
  return stripControlChars(s);
}

function finiteOrNull(v: number | null): number | null {
  return v !== null && Number.isFinite(v) ? v : null;
}

interface ColumnWriter {
  col0: number;
  maxLen: number;
}

class SheetBuilder {
  private readonly rows = new Map<number, CellSpec[]>();
  private readonly widths = new Map<number, number>();

  set(col0: number, row0: number, value: CellValue, style?: CellStyleSpec, displayLen?: number): void {
    let row = this.rows.get(row0);
    if (!row) {
      row = [];
      this.rows.set(row0, row);
    }
    const cell: CellSpec = { col0, row0, value };
    if (style) cell.style = style;
    row.push(cell);
    const len =
      displayLen ??
      (value.type === 'string' ? value.value.length : value.type === 'number' ? String(value.value).length : value.type === 'boolean' ? 5 : 0);
    this.widths.set(col0, Math.max(this.widths.get(col0) ?? 0, len));
  }

  touch(col: ColumnWriter): void {
    this.widths.set(col.col0, Math.max(this.widths.get(col.col0) ?? 0, col.maxLen));
  }

  build(name: string, hidden: boolean): SheetSpec {
    const rows: RowSpec[] = [...this.rows.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([row0, cells]) => ({ row0, cells: cells.sort((a, b) => a.col0 - b.col0) }));
    const columns = [...this.widths.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([col0, len]) => ({ col0, widthChars: Math.max(10, Math.min(40, len + 2)) }));
    return { name, hidden, columns, rows, freezeHeaderRow: true, drawings: [] };
  }
}

/** Builds the data sheet and the cell references for every series in the plan. */
export function buildDataLayout(model: ChartModel, plan: ChartTypeResolution, opts: DataLayoutOptions): DataLayout {
  const diagnostics: Diagnostic[] = [];
  const sheetName = opts.sheetName;
  const qSheet = quoteSheetNameForFormula(sheetName);
  const builder = new SheetBuilder();
  const ranges: SeriesRange[] = [];
  let formulaLikeReported = false;
  const noteFormulaLike = (s: string, property: string, seriesIndex?: number): void => {
    if (formulaLikeReported || !isFormulaLike(s)) return;
    formulaLikeReported = true;
    const extra: Parameters<typeof createDiagnostic>[4] = { severity: 'info', details: { example: s } };
    if (seriesIndex !== undefined) extra.seriesIndex = seriesIndex;
    diagnostics.push(
      createDiagnostic(
        'FORMULA_LIKE_TEXT_ESCAPED',
        'translated',
        property,
        `Text that looks like a formula (e.g. "${s}") is written as a plain string cell and is never evaluated.`,
        extra,
      ),
    );
  };
  const writeHeader = (col0: number, text: string, property: string, seriesIndex?: number): string => {
    const clean = cleanText(text);
    noteFormulaLike(clean, property, seriesIndex);
    builder.set(col0, 0, { type: 'string', value: clean }, HEADER_STYLE);
    return clean;
  };
  const seriesName = (s: SeriesModel, pos: number): string => {
    const n = cleanText(s.name ?? '').trim();
    return n === '' ? `Series ${pos + 1}` : n;
  };

  // --- Which series go where ------------------------------------------------------------------
  const categoryPositions: number[] = [];
  const scatterPositions: number[] = [];
  const piePositions: number[] = [];
  let hasPieGroup = false;
  for (const g of plan.groups) {
    if (isCategoryGroupKind(g.kind)) categoryPositions.push(...g.seriesIndices);
    else if (g.kind === 'scatter' || g.kind === 'bubble') scatterPositions.push(...g.seriesIndices);
    else hasPieGroup = true;
  }
  if (hasPieGroup) {
    // Every pie-family series (including extra pie series not drawn by Excel) keeps its data on the sheet.
    model.series.forEach((s, pos) => {
      if (s.kind === 'pie' || s.kind === 'doughnut') piePositions.push(pos);
    });
  }
  categoryPositions.sort((a, b) => a - b);
  scatterPositions.sort((a, b) => a - b);

  const xAxis0 = model.xAxes[0] ?? null;
  const xKind: AxisKind = xAxis0?.kind ?? inferXKind(model, categoryPositions.length > 0 ? categoryPositions : scatterPositions);
  const isDatetime = xKind === 'datetime';

  // x extent over every exported series (for date format choice).
  let xMin: number | null = null;
  let xMax: number | null = null;
  if (isDatetime) {
    for (const pos of [...categoryPositions, ...scatterPositions]) {
      for (const p of model.series[pos]?.points ?? []) {
        if (p.x === null || !Number.isFinite(p.x)) continue;
        xMin = xMin === null ? p.x : Math.min(xMin, p.x);
        xMax = xMax === null ? p.x : Math.max(xMax, p.x);
      }
    }
  }
  const dateFormatCode =
    isDatetime && xAxis0 ? opts.dateFormatCode(xAxis0, xMin !== null && xMax !== null ? xMax - xMin : 0) : isDatetime ? 'yyyy-mm-dd' : null;

  // --- Plan the shapes (row/column counts) before writing anything ------------------------------
  let rowKeys: Array<string | number> = [];
  let pieKeys: string[] = [];
  let categoryMode: 'label' | 'index' | 'numeric' = 'numeric';
  const categoryLabelByIndex = new Map<number, string>();

  if (categoryPositions.length > 0) {
    if (xKind === 'category') {
      const allIndexed = categoryPositions.every((pos) =>
        (model.series[pos]?.points ?? []).every((p) => p.x !== null && Number.isInteger(p.x) && p.x >= 0),
      );
      if (allIndexed) {
        categoryMode = 'index';
        const set = new Set<number>();
        for (const pos of categoryPositions) {
          const s = model.series[pos]!;
          const cats = (model.xAxes[s.xAxisIndex] ?? xAxis0)?.categories ?? null;
          for (const p of s.points) {
            const x = p.x as number;
            set.add(x);
            if (!categoryLabelByIndex.has(x)) {
              const label = cats?.[x] ?? p.name ?? null;
              if (label !== null) categoryLabelByIndex.set(x, label);
            }
          }
        }
        rowKeys = [...set].sort((a, b) => a - b);
      } else {
        categoryMode = 'label';
        const seen = new Set<string>();
        for (const pos of categoryPositions) {
          const s = model.series[pos]!;
          s.points.forEach((p, j) => {
            const key = categoryLabel(p, j, (model.xAxes[s.xAxisIndex] ?? xAxis0)?.categories ?? null);
            if (!seen.has(key)) {
              seen.add(key);
              rowKeys.push(key);
            }
          });
        }
      }
    } else {
      categoryMode = 'numeric';
      const set = new Set<number>();
      for (const pos of categoryPositions) {
        model.series[pos]!.points.forEach((p, j) => set.add(numericX(p, j)));
      }
      rowKeys = [...set].sort((a, b) => a - b);
    }
  }

  if (piePositions.length > 0) {
    const seen = new Set<string>();
    for (const pos of piePositions) {
      const s = model.series[pos]!;
      pieKeysOf(s).forEach((k) => {
        if (!seen.has(k)) {
          seen.add(k);
          pieKeys.push(k);
        }
      });
    }
  }

  const categoryRows = categoryPositions.length > 0 ? rowKeys.length : 0;
  const pieRows = pieKeys.length;
  const scatterRows = scatterPositions.reduce((m, pos) => Math.max(m, model.series[pos]?.points.length ?? 0), 0);
  const rowCount = Math.max(categoryRows, pieRows, scatterRows);
  const colCount =
    (categoryPositions.length > 0 ? 1 + categoryPositions.length : 0) +
    (piePositions.length > 0 ? 1 + piePositions.length : 0) +
    scatterPositions.reduce((n, pos) => n + (model.series[pos]?.kind === 'bubble' ? 3 : 2), 0);

  const limits = checkLimits(rowCount, colCount, 0);
  diagnostics.push(...limits.diagnostics);
  if (limits.blocking) {
    return {
      sheet: builder.build(sheetName, opts.hidden),
      ranges: [],
      rowCount,
      diagnostics,
      blocking: true,
      x: { kind: xKind, keys: [], dateFormatCode, min: xMin, max: xMax },
    };
  }

  const nameRefOf = (col0: number, text: string): SeriesRange['nameRef'] => ({ formula: `${qSheet}!${cellRef(col0, 0)}`, cache: text });
  /** Rows 1..max(1,n): a chart reference always spans at least one cell. */
  const refTo = (n: number): number => Math.max(1, n);
  const padCache = <T>(cache: T[], n: number, fill: T): T[] => (cache.length === 0 && n === 0 ? [fill] : cache);

  const valueFormatOf = (s: SeriesModel): string => (s.yFormat?.kind === 'excel' && s.yFormat.code ? s.yFormat.code : 'General');
  const valueStyle = (code: string): CellStyleSpec | undefined => (code !== 'General' ? { numberFormat: code } : undefined);

  const reportSeriesData = (s: SeriesModel, pos: number, nulls: number, nonNumeric: number, missing: number): void => {
    const path = `series[${s.index}].data`;
    if (nulls > 0) {
      diagnostics.push(
        createDiagnostic('NULL_VALUES', 'translated', path, `${nulls} null values exported as empty cells (shown as gaps).`, {
          severity: 'info',
          seriesIndex: s.index,
          details: { count: nulls },
        }),
      );
    }
    if (nonNumeric > 0) {
      diagnostics.push(
        createDiagnostic('NON_NUMERIC_VALUE', 'approximated', path, `${nonNumeric} non-numeric values exported as empty cells.`, {
          seriesIndex: s.index,
          details: { count: nonNumeric },
        }),
      );
    }
    if (missing > 0) {
      diagnostics.push(
        createDiagnostic(
          'UNALIGNED_X_VALUES',
          'approximated',
          path,
          `Series has no value for ${missing} of the shared x values; those cells are left empty.`,
          { severity: 'info', seriesIndex: s.index, details: { missing } },
        ),
      );
    }
    if (s.points.length === 0) {
      diagnostics.push(
        createDiagnostic('EMPTY_SERIES', 'translated', path, 'Series has no data points.', { severity: 'info', seriesIndex: s.index }),
      );
    }
    if (!s.visible) {
      diagnostics.push(
        createDiagnostic('HIDDEN_SERIES_INCLUDED', 'approximated', `series[${s.index}].visible`, 'Hidden series is exported and appears visible in Excel.', {
          severity: 'info',
          seriesIndex: s.index,
        }),
      );
    }
    void pos;
  };

  const yValue = (p: PointModel, counters: { nulls: number; nonNumeric: number }): number | null => {
    if (p.isNull || p.y === null) {
      counters.nulls++;
      return null;
    }
    if (!Number.isFinite(p.y)) {
      counters.nonNumeric++;
      return null;
    }
    return p.y;
  };

  let nextCol = 0;

  // --- Category layout ------------------------------------------------------------------------------
  if (categoryPositions.length > 0) {
    const catCol = nextCol;
    const xTitle = xAxis0?.title?.text?.trim();
    writeHeader(catCol, xTitle ? xTitle : isDatetime ? 'Date' : 'Category', xTitle ? 'xAxis[0].title.text' : 'xAxis[0].categories');
    const n = rowKeys.length;
    const catCache: (string | number | null)[] = [];
    const catIsStr = categoryMode !== 'numeric';
    const catFormat = isDatetime ? (dateFormatCode ?? 'yyyy-mm-dd') : 'General';
    rowKeys.forEach((key, r) => {
      if (catIsStr) {
        const label = cleanText(categoryMode === 'index' ? (categoryLabelByIndex.get(key as number) ?? String(key)) : String(key));
        noteFormulaLike(label, 'xAxis[0].categories');
        builder.set(catCol, r + 1, { type: 'string', value: label });
        catCache.push(label);
      } else {
        const v = isDatetime ? msToExcelSerial(key as number) : (key as number);
        builder.set(catCol, r + 1, { type: 'number', value: v }, isDatetime ? { numberFormat: catFormat } : undefined, isDatetime ? catFormat.length + 2 : undefined);
        catCache.push(v);
      }
    });
    const categories: NonNullable<SeriesRange['categories']> = {
      formula: rangeRef(sheetName, catCol, 1, refTo(n)),
      kind: catIsStr ? 'str' : 'num',
      cache: padCache(catCache, n, null),
    };
    if (!catIsStr) categories.formatCode = catFormat;
    nextCol++;

    const rowOfKey = new Map<string | number, number>();
    rowKeys.forEach((k, r) => rowOfKey.set(k, r));

    for (const pos of categoryPositions) {
      const s = model.series[pos]!;
      const col = nextCol++;
      const name = writeHeader(col, seriesName(s, pos), `series[${s.index}].name`, s.index);
      const formatCode = valueFormatOf(s);
      const style = valueStyle(formatCode);
      const cache: (number | null)[] = new Array<number | null>(n).fill(null);
      const filled = new Uint8Array(n);
      const pointOffsets: number[] = [];
      const counters = { nulls: 0, nonNumeric: 0 };
      const cats = (model.xAxes[s.xAxisIndex] ?? xAxis0)?.categories ?? null;
      s.points.forEach((p, j) => {
        const key =
          categoryMode === 'label' ? categoryLabel(p, j, cats) : categoryMode === 'index' ? (p.x as number) : numericX(p, j);
        const r = rowOfKey.get(key);
        if (r === undefined || filled[r] === 1) {
          pointOffsets.push(-1);
          return;
        }
        filled[r] = 1;
        pointOffsets.push(r);
        const v = yValue(p, counters);
        cache[r] = v;
        if (v !== null) builder.set(col, r + 1, { type: 'number', value: v }, style);
      });
      let missing = 0;
      for (let r = 0; r < n; r++) if (filled[r] === 0) missing++;
      reportSeriesData(s, pos, counters.nulls, counters.nonNumeric, missing);
      const lim = checkLimits(0, 0, n, s.index);
      diagnostics.push(...lim.diagnostics);
      ranges.push({
        seriesIndex: pos,
        nameRef: nameRefOf(col, name),
        categories: { ...categories, cache: categories.cache.slice() },
        values: { formula: rangeRef(sheetName, col, 1, refTo(n)), cache: padCache(cache, n, null), formatCode },
        bubbleSizes: null,
        pointOffsets,
        literalName: null,
      });
    }
  }

  // --- Pie / doughnut layout --------------------------------------------------------------------------
  if (piePositions.length > 0) {
    const catCol = nextCol++;
    writeHeader(catCol, 'Category', 'series.data.name');
    const n = pieKeys.length;
    const catCache: string[] = [];
    pieKeys.forEach((key, r) => {
      const label = cleanText(pieLabelOfKey(key));
      noteFormulaLike(label, 'series.data.name');
      builder.set(catCol, r + 1, { type: 'string', value: label });
      catCache.push(label);
    });
    const rowOfKey = new Map<string, number>();
    pieKeys.forEach((k, r) => rowOfKey.set(k, r));
    for (const pos of piePositions) {
      const s = model.series[pos]!;
      const col = nextCol++;
      const name = writeHeader(col, seriesName(s, pos), `series[${s.index}].name`, s.index);
      const formatCode = valueFormatOf(s);
      const style = valueStyle(formatCode);
      const cache: (number | null)[] = new Array<number | null>(n).fill(null);
      const counters = { nulls: 0, nonNumeric: 0 };
      const keys = pieKeysOf(s);
      const pointOffsets: number[] = [];
      s.points.forEach((p, j) => {
        const r = rowOfKey.get(keys[j]!)!;
        pointOffsets.push(r);
        const v = yValue(p, counters);
        cache[r] = v;
        if (v !== null) builder.set(col, r + 1, { type: 'number', value: v }, style);
      });
      reportSeriesData(s, pos, counters.nulls, counters.nonNumeric, 0);
      ranges.push({
        seriesIndex: pos,
        nameRef: nameRefOf(col, name),
        categories: { formula: rangeRef(sheetName, catCol, 1, refTo(n)), kind: 'str', cache: padCache(catCache.slice() as (string | null)[], n, null) },
        values: { formula: rangeRef(sheetName, col, 1, refTo(n)), cache: padCache(cache, n, null), formatCode },
        bubbleSizes: null,
        pointOffsets,
        literalName: null,
      });
    }
  }

  // --- Scatter / bubble blocks --------------------------------------------------------------------------
  for (const pos of scatterPositions) {
    const s = model.series[pos]!;
    const isBubble = s.kind === 'bubble';
    const xCol = nextCol;
    const yCol = nextCol + 1;
    const zCol = nextCol + 2;
    nextCol += isBubble ? 3 : 2;
    const base = seriesName(s, pos);
    writeHeader(xCol, `${base} X`, `series[${s.index}].name`, s.index);
    const name = writeHeader(yCol, `${base} Y`, `series[${s.index}].name`, s.index);
    if (isBubble) writeHeader(zCol, `${base} Size`, `series[${s.index}].name`, s.index);
    const formatCode = valueFormatOf(s);
    const style = valueStyle(formatCode);
    const xFormat = isDatetime ? (dateFormatCode ?? 'yyyy-mm-dd') : 'General';
    const xs: (number | null)[] = [];
    const ys: (number | null)[] = [];
    const zs: (number | null)[] = [];
    const counters = { nulls: 0, nonNumeric: 0 };
    const pointOffsets: number[] = [];
    s.points.forEach((p, j) => {
      const r = j + 1;
      pointOffsets.push(j);
      const rawX = numericX(p, j);
      const x = isDatetime ? msToExcelSerial(rawX) : rawX;
      xs.push(x);
      builder.set(xCol, r, { type: 'number', value: x }, isDatetime ? { numberFormat: xFormat } : undefined, isDatetime ? xFormat.length + 2 : undefined);
      const y = yValue(p, counters);
      ys.push(y);
      if (y !== null) builder.set(yCol, r, { type: 'number', value: y }, style);
      if (isBubble) {
        const z = finiteOrNull(p.z);
        zs.push(z);
        if (z !== null) builder.set(zCol, r, { type: 'number', value: z });
      }
    });
    const n = s.points.length;
    reportSeriesData(s, pos, counters.nulls, counters.nonNumeric, 0);
    diagnostics.push(...checkLimits(0, 0, n, s.index).diagnostics);
    const categories: NonNullable<SeriesRange['categories']> = {
      formula: rangeRef(sheetName, xCol, 1, refTo(n)),
      kind: 'num',
      cache: padCache(xs as (string | number | null)[], n, null),
    };
    if (isDatetime) categories.formatCode = xFormat;
    ranges.push({
      seriesIndex: pos,
      nameRef: nameRefOf(yCol, name),
      literalName: base,
      categories,
      values: { formula: rangeRef(sheetName, yCol, 1, refTo(n)), cache: padCache(ys, n, null), formatCode },
      bubbleSizes: isBubble ? { formula: rangeRef(sheetName, zCol, 1, refTo(n)), cache: padCache(zs, n, null), formatCode: 'General' } : null,
      pointOffsets,
    });
  }

  return {
    sheet: builder.build(sheetName, opts.hidden),
    ranges,
    rowCount,
    diagnostics,
    blocking: false,
    x: {
      kind: categoryPositions.length > 0 || scatterPositions.length > 0 ? xKind : 'none',
      keys: rowKeys,
      dateFormatCode,
      min: xMin,
      max: xMax,
    },
  };
}

/** Infers the x axis kind when the model has no x axis. */
function inferXKind(model: ChartModel, positions: number[]): AxisKind {
  const pts = positions.flatMap((pos) => model.series[pos]?.points ?? []);
  if (pts.length > 0 && pts.every((p) => p.name !== null)) return 'category';
  return pts.some((p) => p.x !== null) ? 'linear' : 'category';
}

function numericX(p: PointModel, j: number): number {
  return p.x !== null && Number.isFinite(p.x) ? p.x : j;
}

function categoryLabel(p: PointModel, j: number, cats: string[] | null): string {
  if (p.x !== null && Number.isInteger(p.x) && cats && p.x >= 0 && p.x < cats.length) return cats[p.x]!;
  if (p.name !== null && p.name !== '') return p.name;
  if (cats && j < cats.length) return cats[j]!;
  return String(p.x ?? j);
}

const PIE_KEY_SEP = '\u0000';

/** Unique row keys for the points of a pie series (duplicate names get an occurrence suffix). */
function pieKeysOf(s: SeriesModel): string[] {
  const occurrences = new Map<string, number>();
  return s.points.map((p, j) => {
    const label = p.name !== null && p.name !== '' ? p.name : `Point ${j + 1}`;
    const n = (occurrences.get(label) ?? 0) + 1;
    occurrences.set(label, n);
    return n === 1 ? label : `${label}${PIE_KEY_SEP}${n}`;
  });
}

function pieLabelOfKey(key: string): string {
  const i = key.indexOf(PIE_KEY_SEP);
  return i === -1 ? key : key.slice(0, i);
}
