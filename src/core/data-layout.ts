/**
 * Data sheet layout: decides where every value lives on the worksheet and produces the
 * sheet-qualified absolute references the chart series point at.
 */

import type { AxisKind, AxisModel, ChartModel, PointModel, SeriesModel } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import type { CellSpec, CellStyleSpec, CellValue, RowSpec, SheetSpec } from '../excel/writer-interface';
import { EXCEL_MAX_COLUMNS, EXCEL_MAX_POINTS_PER_SERIES, EXCEL_MAX_ROWS } from '../excel/writer-interface';
import { msToExcelSerial } from '../utils/dates';
import { isValidExcelFormatCode } from '../utils/format-code';
import { cellRef, quoteSheetNameForFormula, rangeRef } from '../utils/filenames';
import { isFormulaLike, stripControlChars } from '../utils/text';
import { isCategoryGroupKind, isRangeKind, rangeNeedsSplit, type ChartTypeResolution } from './chart-type-registry';

/** A worksheet column referenced by a chart (absolute range formula + cache). */
export interface ColumnRef {
  formula: string;
  cache: (number | null)[];
  formatCode: string;
}

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
   * "<name> Y"), so the chart uses this literal name instead of `nameRef`. Range series too (their
   * columns are "<name> Low/High/Range"). Null elsewhere.
   */
  literalName: string | null;
  /** Custom error bar lengths (+err / -err columns) from the errorbar series at `seriesIndex`. */
  errorBars?: { seriesIndex: number; plus: ColumnRef; minus: ColumnRef } | null;
  /**
   * Range series only: the hidden base series stacked under `values` (the Low column, or the Base
   * helper column of a column range crossing zero) and, for such a split range, the Down column
   * stacked below the axis.
   */
  rangeParts?: { base: ColumnRef & { name: { formula: string; cache: string } }; down: ColumnRef | null } | null;
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
      createDiagnostic(
        'ROW_LIMIT_EXCEEDED',
        'blocking',
        'series.data',
        `The data needs ${rowCount + 1} rows; an Excel worksheet holds at most ${EXCEL_MAX_ROWS}.`,
        {
          details: { rows: rowCount + 1, limit: EXCEL_MAX_ROWS },
        },
      ),
    );
  }
  if (colCount > EXCEL_MAX_COLUMNS) {
    blocking = true;
    diagnostics.push(
      createDiagnostic(
        'COLUMN_LIMIT_EXCEEDED',
        'blocking',
        'series',
        `The data needs ${colCount} columns; an Excel worksheet holds at most ${EXCEL_MAX_COLUMNS}.`,
        {
          details: { columns: colCount, limit: EXCEL_MAX_COLUMNS },
        },
      ),
    );
  }
  if (maxPointsPerSeries > EXCEL_MAX_POINTS_PER_SERIES) {
    const extra: Parameters<typeof createDiagnostic>[4] = {
      details: { points: maxPointsPerSeries, limit: EXCEL_MAX_POINTS_PER_SERIES },
    };
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
      (value.type === 'string'
        ? value.value.length
        : value.type === 'number'
          ? String(value.value).length
          : value.type === 'boolean'
            ? 5
            : value.type === 'error'
              ? value.value.length
              : value.type === 'formula'
                ? value.cached === null
                  ? 0
                  : String(value.cached).length
                : 0);
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
  const errorBarOf = new Map<number, number>((plan.errorBars ?? []).map((e) => [e.parent, e.errorBar]));

  const xAxis0 = model.xAxes[0] ?? null;
  const sourceXKind: AxisKind =
    xAxis0?.kind ?? inferXKind(model, categoryPositions.length > 0 ? categoryPositions : scatterPositions);

  // x extent over every exported series (for date format choice).
  let xMin: number | null = null;
  let xMax: number | null = null;
  if (sourceXKind === 'datetime') {
    for (const pos of [...categoryPositions, ...scatterPositions]) {
      for (const p of model.series[pos]?.points ?? []) {
        if (p.x === null || !Number.isFinite(p.x)) continue;
        xMin = xMin === null ? p.x : Math.min(xMin, p.x);
        xMax = xMax === null ? p.x : Math.max(xMax, p.x);
      }
    }
  }
  // Excel has no dates before 1899-12-31 (negative serials). Such data is written as ISO date text in a
  // category column (scatter X columns keep plain day numbers so the geometry survives).
  const pre1900 = sourceXKind === 'datetime' && xMin !== null && msToExcelSerial(xMin) < 0;
  if (pre1900) {
    diagnostics.push(
      createDiagnostic(
        'APPROXIMATED_DATETIME',
        'approximated',
        'xAxis[0].type',
        `Excel dates start on 1900-01-01; x values from ${isoDateText([xMin!])[0]} are written as date text on a category axis instead of a date axis.`,
        { details: { min: xMin } },
      ),
    );
  }
  const xKind: AxisKind = pre1900 ? (categoryPositions.length > 0 ? 'category' : 'linear') : sourceXKind;
  const isDatetime = xKind === 'datetime';
  const dateFormatCode =
    isDatetime && xAxis0
      ? opts.dateFormatCode(xAxis0, xMin !== null && xMax !== null ? xMax - xMin : 0)
      : isDatetime
        ? 'yyyy-mm-dd'
        : null;

  // --- Plan the shapes (row/column counts) before writing anything ------------------------------
  let rowKeys: Array<string | number> = [];
  const pieKeys: string[] = [];
  let perRingCategories = false;
  let categoryMode: 'label' | 'index' | 'numeric' = 'numeric';
  const categoryLabelByIndex = new Map<number, string>();

  if (categoryPositions.length > 0) {
    if (xKind === 'category' && !pre1900) {
      const allIndexed =
        !pre1900 &&
        categoryPositions.every((pos) =>
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
            const { key } = categoryLabel(p, j, (model.xAxes[s.xAxisIndex] ?? xAxis0)?.categories ?? null);
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
        model.series[pos]!.points.forEach((p, j) => {
          set.add(numericX(p, j));
        });
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
    // Doughnut rings whose slices differ (names or order) each get their own category column: Excel
    // takes c:cat per series, so every ring keeps its own slices in its own order.
    const ringPositions = plan.groups.filter((g) => g.kind === 'doughnut').flatMap((g) => g.seriesIndices);
    if (ringPositions.length > 1) {
      const sequence = (pos: number): string => pieKeysOf(model.series[pos]!).join(PIE_KEY_SEP + PIE_KEY_SEP);
      const first = sequence(ringPositions[0]!);
      perRingCategories = ringPositions.slice(1).some((pos) => sequence(pos) !== first);
    }
  }

  const categoryRows = categoryPositions.length > 0 ? rowKeys.length : 0;
  const pieRows = perRingCategories
    ? piePositions.reduce((m, pos) => Math.max(m, model.series[pos]?.points.length ?? 0), 0)
    : pieKeys.length;
  const scatterRows = scatterPositions.reduce((m, pos) => Math.max(m, model.series[pos]?.points.length ?? 0), 0);
  const rowCount = Math.max(categoryRows, pieRows, scatterRows);
  const errorColumns = (pos: number): number => (errorBarOf.has(pos) ? 2 : 0);
  const categoryColumns = (pos: number): number => {
    const s = model.series[pos]!;
    return (isRangeKind(s.kind) ? (rangeNeedsSplit(s) ? 5 : 3) : 1) + errorColumns(pos);
  };
  const colCount =
    (categoryPositions.length > 0 ? 1 + categoryPositions.reduce((n, pos) => n + categoryColumns(pos), 0) : 0) +
    (piePositions.length > 0 ? (perRingCategories ? 2 : 1) * piePositions.length + (perRingCategories ? 0 : 1) : 0) +
    scatterPositions.reduce((n, pos) => n + (model.series[pos]?.kind === 'bubble' ? 3 : 2) + errorColumns(pos), 0);

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

  const nameRefOf = (col0: number, text: string): SeriesRange['nameRef'] => ({
    formula: `${qSheet}!${cellRef(col0, 0)}`,
    cache: text,
  });
  /** Rows 1..max(1,n): a chart reference always spans at least one cell. */
  const refTo = (n: number): number => Math.max(1, n);
  const padCache = <T>(cache: T[], n: number, fill: T): T[] => (cache.length === 0 && n === 0 ? [fill] : cache);

  const valueFormatOf = (s: SeriesModel): string =>
    s.yFormat?.kind === 'excel' && s.yFormat.code && isValidExcelFormatCode(s.yFormat.code)
      ? s.yFormat.code
      : 'General';
  const valueStyle = (code: string): CellStyleSpec | undefined =>
    code !== 'General' ? { numberFormat: code } : undefined;

  const reportSeriesData = (s: SeriesModel, pos: number, nulls: number, nonNumeric: number, missing: number): void => {
    const path = `series[${s.index}].data`;
    if (nulls > 0) {
      diagnostics.push(
        createDiagnostic(
          'NULL_VALUES',
          'translated',
          path,
          `${nulls} null values exported as empty cells (shown as gaps).`,
          {
            severity: 'info',
            seriesIndex: s.index,
            details: { count: nulls },
          },
        ),
      );
    }
    if (nonNumeric > 0) {
      diagnostics.push(
        createDiagnostic(
          'NON_NUMERIC_VALUE',
          'approximated',
          path,
          `${nonNumeric} non-numeric values exported as empty cells.`,
          {
            seriesIndex: s.index,
            details: { count: nonNumeric },
          },
        ),
      );
    }
    if (missing > 0) {
      diagnostics.push(
        createDiagnostic(
          'UNALIGNED_X_VALUES',
          'approximated',
          path,
          `Series has no point at ${missing} of the shared x values; those cells are written as #N/A so lines stay connected (as in the source chart).`,
          { severity: 'info', seriesIndex: s.index, details: { missing } },
        ),
      );
    }
    if (s.points.length === 0) {
      diagnostics.push(
        createDiagnostic('EMPTY_SERIES', 'translated', path, 'Series has no data points.', {
          severity: 'info',
          seriesIndex: s.index,
        }),
      );
    }
    if (!s.visible) {
      diagnostics.push(
        createDiagnostic(
          'HIDDEN_SERIES_INCLUDED',
          'approximated',
          `series[${s.index}].visible`,
          'Hidden series is exported and appears visible in Excel.',
          {
            severity: 'info',
            seriesIndex: s.index,
          },
        ),
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

  const columnRef = (col0: number, cache: (number | null)[], n: number, formatCode: string): ColumnRef => ({
    formula: rangeRef(sheetName, col0, 1, refTo(n)),
    cache: padCache(cache, n, null),
    formatCode,
  });
  /** -0 → 0 (Excel and the chart cache have no negative zero). */
  const noNegZero = (v: number): number => (v === 0 ? 0 : v);

  /**
   * Error bar columns ("<parent> +err", "<parent> -err") at the next two columns: for every errorbar
   * point on a parent row, plus = high − y and minus = y − low. Points without a parent row, a
   * parent value, a low or a high get no bar (empty cells); nothing is fabricated.
   */
  const writeErrorBarColumns = (
    parent: SeriesModel,
    parentPos: number,
    e: SeriesModel,
    ebPos: number,
    n: number,
    rowOf: (p: PointModel, j: number) => number | undefined,
    parentValues: ReadonlyArray<number | null>,
    formatCode: string,
  ): NonNullable<SeriesRange['errorBars']> => {
    const base = seriesName(parent, parentPos);
    const plusCol = nextCol++;
    const minusCol = nextCol++;
    writeHeader(plusCol, `${base} +err`, `series[${e.index}].name`, e.index);
    writeHeader(minusCol, `${base} -err`, `series[${e.index}].name`, e.index);
    const style = valueStyle(formatCode);
    const plus: (number | null)[] = new Array<number | null>(n).fill(null);
    const minus: (number | null)[] = new Array<number | null>(n).fill(null);
    const used = new Uint8Array(n);
    let unaligned = 0;
    e.points.forEach((p, j) => {
      const r = rowOf(p, j);
      if (r === undefined || r < 0 || r >= n || used[r] === 1) {
        unaligned++;
        return;
      }
      used[r] = 1;
      const y = parentValues[r] ?? null;
      const low = finiteOrNull(p.low ?? null);
      const high = finiteOrNull(p.high ?? null);
      if (y === null || low === null || high === null || p.isNull) return;
      plus[r] = noNegZero(high - y);
      minus[r] = noNegZero(y - low);
      builder.set(plusCol, r + 1, { type: 'number', value: plus[r]! }, style);
      builder.set(minusCol, r + 1, { type: 'number', value: minus[r]! }, style);
    });
    if (unaligned > 0) {
      diagnostics.push(
        createDiagnostic(
          'UNALIGNED_X_VALUES',
          'approximated',
          `series[${e.index}].data`,
          `${unaligned} error bar point(s) have no matching point in series "${base}"; they are not exported.`,
          { seriesIndex: e.index, details: { unaligned } },
        ),
      );
    }
    return {
      seriesIndex: ebPos,
      plus: columnRef(plusCol, plus, n, formatCode),
      minus: columnRef(minusCol, minus, n, formatCode),
    };
  };

  /**
   * A range series in the category layout: "<name> Low" | "<name> High" | "<name> Range" where
   * Range is the formula =High−Low (so editing Low/High updates the chart), drawn as a hidden Low
   * base with Range stacked on top. A column range with a negative low cannot use that (Excel
   * stacks negative values below the axis), so it gets "<name> Base" = MAX(0,Low)+MIN(0,High),
   * "<name> Up" = MAX(0,High)−MAX(0,Low) and "<name> Down" = MIN(0,Low)−MIN(0,High) instead.
   */
  const writeRangeSeries = (
    s: SeriesModel,
    pos: number,
    placed: { rows: number[]; filled: Uint8Array; missing: number },
    categories: NonNullable<SeriesRange['categories']>,
  ): void => {
    const n = placed.filled.length;
    const split = rangeNeedsSplit(s);
    const base = seriesName(s, pos);
    const lowCol = nextCol;
    const highCol = nextCol + 1;
    const helperNames = split ? ['Base', 'Up', 'Down'] : ['Range'];
    const helperCols = helperNames.map((_, i) => nextCol + 2 + i);
    nextCol += 2 + helperNames.length;
    const property = `series[${s.index}].name`;
    writeHeader(lowCol, `${base} Low`, property, s.index);
    writeHeader(highCol, `${base} High`, property, s.index);
    const helperHeaders = helperNames.map((h, i) => writeHeader(helperCols[i]!, `${base} ${h}`, property, s.index));
    const formatCode = valueFormatOf(s);
    const style = valueStyle(formatCode);
    const lows: (number | null)[] = new Array<number | null>(n).fill(null);
    const helpers: (number | null)[][] = helperNames.map(() => new Array<number | null>(n).fill(null));
    const counters = { nulls: 0, nonNumeric: 0 };
    s.points.forEach((p, j) => {
      const r = placed.rows[j]!;
      if (r < 0) return;
      const rawLow = p.low ?? null;
      const rawHigh = p.high ?? null;
      const low = finiteOrNull(rawLow);
      const high = finiteOrNull(rawHigh);
      if ((rawLow !== null && low === null) || (rawHigh !== null && high === null)) counters.nonNumeric++;
      if (low !== null) builder.set(lowCol, r + 1, { type: 'number', value: low }, style);
      if (high !== null) builder.set(highCol, r + 1, { type: 'number', value: high }, style);
      if (low === null || high === null || p.isNull) {
        // No range without both ends: the helper cells stay empty.
        if (rawLow === null || rawHigh === null) counters.nulls++;
        return;
      }
      lows[r] = low;
      const L = cellRef(lowCol, r + 1, false);
      const H = cellRef(highCol, r + 1, false);
      const cells: Array<{ formula: string; value: number }> = split
        ? [
            { formula: `MAX(0,${L})+MIN(0,${H})`, value: Math.max(0, low) + Math.min(0, high) },
            { formula: `MAX(0,${H})-MAX(0,${L})`, value: Math.max(0, high) - Math.max(0, low) },
            { formula: `MIN(0,${L})-MIN(0,${H})`, value: Math.min(0, low) - Math.min(0, high) },
          ]
        : [{ formula: `${H}-${L}`, value: high - low }];
      cells.forEach((c, i) => {
        const v = noNegZero(c.value);
        helpers[i]![r] = v;
        builder.set(helperCols[i]!, r + 1, { type: 'formula', formula: c.formula, cached: v }, style);
      });
    });
    for (let r = 0; r < n; r++) {
      if (placed.filled[r] === 1) continue;
      // No point at this x: #N/A, like the other category series.
      for (const col of [lowCol, highCol, ...helperCols]) builder.set(col, r + 1, { type: 'error', value: '#N/A' });
    }
    reportSeriesData(s, pos, counters.nulls, counters.nonNumeric, placed.missing);
    diagnostics.push(...checkLimits(0, 0, n, s.index).diagnostics);
    const baseCol = split ? helperCols[0]! : lowCol;
    const mainIdx = split ? 1 : 0;
    ranges.push({
      seriesIndex: pos,
      nameRef: nameRefOf(helperCols[mainIdx]!, helperHeaders[mainIdx]!),
      literalName: base,
      categories,
      values: columnRef(helperCols[mainIdx]!, helpers[mainIdx]!, n, formatCode),
      bubbleSizes: null,
      pointOffsets: placed.rows,
      rangeParts: {
        base: {
          ...columnRef(baseCol, split ? helpers[0]! : lows, n, formatCode),
          name: nameRefOf(baseCol, split ? helperHeaders[0]! : `${base} Low`),
        },
        down: split ? columnRef(helperCols[2]!, helpers[2]!, n, formatCode) : null,
      },
    });
  };

  // --- Category layout ------------------------------------------------------------------------------
  if (categoryPositions.length > 0) {
    const catCol = nextCol;
    const xTitle = xAxis0?.title?.text?.trim();
    writeHeader(
      catCol,
      xTitle ? xTitle : isDatetime ? 'Date' : 'Category',
      xTitle ? 'xAxis[0].title.text' : 'xAxis[0].categories',
    );
    const n = rowKeys.length;
    const catCache: (string | number | null)[] = [];
    const catIsStr = categoryMode !== 'numeric' || pre1900;
    const catFormat = isDatetime ? (dateFormatCode ?? 'yyyy-mm-dd') : 'General';
    const isoLabels = pre1900 ? isoDateText(rowKeys as number[]) : null;
    rowKeys.forEach((key, r) => {
      if (catIsStr) {
        const label = cleanText(
          isoLabels
            ? isoLabels[r]!
            : categoryMode === 'index'
              ? (categoryLabelByIndex.get(key as number) ?? String(key))
              : String(key),
        );
        noteFormulaLike(label, 'xAxis[0].categories');
        builder.set(catCol, r + 1, { type: 'string', value: label });
        catCache.push(label);
      } else {
        const v = isDatetime ? msToExcelSerial(key as number) : (key as number);
        builder.set(
          catCol,
          r + 1,
          { type: 'number', value: v },
          isDatetime ? { numberFormat: catFormat } : undefined,
          isDatetime ? catFormat.length + 2 : undefined,
        );
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
    rowKeys.forEach((k, r) => {
      rowOfKey.set(k, r);
    });

    /** Row key of a point in the category layout. */
    const keyOfPoint = (
      p: PointModel,
      j: number,
      cats: string[] | null,
    ): { key: string | number | null; offCategory: boolean } => {
      if (categoryMode === 'label') return categoryLabel(p, j, cats);
      if (categoryMode === 'index') {
        return {
          key: p.x !== null && Number.isInteger(p.x) && p.x >= 0 ? p.x : null,
          offCategory: false,
        };
      }
      return { key: numericX(p, j), offCategory: false };
    };

    /**
     * Places every point of `s` on its category row: the row per point (-1 when not written: an x the
     * layout has no row for, or a repeat of an x already used by the series) and the filled rows.
     * Reports off-category x values and repeated x values.
     */
    const placePoints = (s: SeriesModel): { rows: number[]; filled: Uint8Array; missing: number } => {
      const filled = new Uint8Array(n);
      const rows: number[] = [];
      const cats = (model.xAxes[s.xAxisIndex] ?? xAxis0)?.categories ?? null;
      const offCategory: string[] = [];
      let firstOffCategory = -1;
      let duplicates = 0;
      let firstDuplicate = -1;
      s.points.forEach((p, j) => {
        const lab = keyOfPoint(p, j, cats);
        if (lab.offCategory) {
          if (firstOffCategory < 0) firstOffCategory = j;
          if (offCategory.length < 10) offCategory.push(String(lab.key));
        }
        const r = lab.key === null ? undefined : rowOfKey.get(lab.key);
        if (r === undefined || filled[r] === 1) {
          // Two points of one series at the same x: Excel has one cell per x, the first point wins.
          if (r !== undefined) {
            duplicates++;
            if (firstDuplicate < 0) firstDuplicate = j;
          }
          rows.push(-1);
          return;
        }
        filled[r] = 1;
        rows.push(r);
      });
      if (firstOffCategory >= 0) {
        diagnostics.push(
          createDiagnostic(
            'UNALIGNED_X_VALUES',
            'approximated',
            `series[${s.index}].data[${firstOffCategory}].x`,
            `Points at x = ${offCategory.join(', ')} do not fall on a category of the x axis; each such x is written as its own category row.`,
            { seriesIndex: s.index, details: { xs: offCategory } },
          ),
        );
      }
      if (duplicates > 0) {
        diagnostics.push(
          createDiagnostic(
            'UNALIGNED_X_VALUES',
            'approximated',
            `series[${s.index}].data[${firstDuplicate}]`,
            `${duplicates} point(s) repeat an x value already used by this series; Excel holds one value per x, so the first point is kept and the repeats are not exported to the chart.`,
            { seriesIndex: s.index, details: { duplicates } },
          ),
        );
      }
      let missing = 0;
      for (let r = 0; r < n; r++) if (filled[r] !== 1) missing++;
      return { rows, filled, missing };
    };

    /** +err / -err columns of the error bars of the series whose values (by row) are `parentValues`. */
    const writeCategoryErrorBars = (
      parent: SeriesModel,
      parentPos: number,
      parentValues: ReadonlyArray<number | null>,
      formatCode: string,
    ): SeriesRange['errorBars'] => {
      const ebPos = errorBarOf.get(parentPos);
      if (ebPos === undefined) return null;
      const e = model.series[ebPos]!;
      const cats = (model.xAxes[e.xAxisIndex] ?? xAxis0)?.categories ?? null;
      const rowOf = (p: PointModel, j: number): number | undefined => {
        const { key } = keyOfPoint(p, j, cats);
        return key === null ? undefined : rowOfKey.get(key);
      };
      return writeErrorBarColumns(parent, parentPos, e, ebPos, n, rowOf, parentValues, formatCode);
    };

    for (const pos of categoryPositions) {
      const s = model.series[pos]!;
      if (isRangeKind(s.kind)) {
        writeRangeSeries(s, pos, placePoints(s), { ...categories, cache: categories.cache.slice() });
        continue;
      }
      const col = nextCol++;
      const name = writeHeader(col, seriesName(s, pos), `series[${s.index}].name`, s.index);
      const formatCode = valueFormatOf(s);
      const style = valueStyle(formatCode);
      const cache: (number | null)[] = new Array<number | null>(n).fill(null);
      const counters = { nulls: 0, nonNumeric: 0 };
      const { rows, filled, missing } = placePoints(s);
      s.points.forEach((p, j) => {
        const r = rows[j]!;
        if (r < 0) return;
        const v = yValue(p, counters);
        cache[r] = v;
        if (v !== null) builder.set(col, r + 1, { type: 'number', value: v }, style);
      });
      for (let r = 0; r < n; r++) {
        if (filled[r] === 1) continue;
        // No point at this x (alignment gap, not a null point): #N/A, which line/area charts skip
        // while connecting the neighbours. The chart cache keeps null for it.
        builder.set(col, r + 1, { type: 'error', value: '#N/A' });
      }
      reportSeriesData(s, pos, counters.nulls, counters.nonNumeric, missing);
      const lim = checkLimits(0, 0, n, s.index);
      diagnostics.push(...lim.diagnostics);
      const range: SeriesRange = {
        seriesIndex: pos,
        nameRef: nameRefOf(col, name),
        categories: { ...categories, cache: categories.cache.slice() },
        values: { formula: rangeRef(sheetName, col, 1, refTo(n)), cache: padCache(cache, n, null), formatCode },
        bubbleSizes: null,
        pointOffsets: rows,
        literalName: null,
      };
      const errorBars = writeCategoryErrorBars(s, pos, cache, formatCode);
      if (errorBars) range.errorBars = errorBars;
      ranges.push(range);
    }
  }

  // --- Pie / doughnut layout --------------------------------------------------------------------------
  if (piePositions.length > 0) {
    /** Writes the category column of `keys` at `catCol` and returns its reference. */
    const writePieCategories = (catCol: number, header: string, keys: readonly string[], n: number) => {
      writeHeader(catCol, header, 'series.data.name');
      const catCache: string[] = [];
      keys.forEach((key, r) => {
        const label = cleanText(pieLabelOfKey(key));
        noteFormulaLike(label, 'series.data.name');
        builder.set(catCol, r + 1, { type: 'string', value: label });
        catCache.push(label);
      });
      return {
        formula: rangeRef(sheetName, catCol, 1, refTo(n)),
        kind: 'str' as const,
        cache: padCache(catCache.slice() as (string | null)[], n, null),
      };
    };
    const shared = perRingCategories
      ? null
      : (() => {
          const catCol = nextCol++;
          const n = pieKeys.length;
          const rowOfKey = new Map<string, number>();
          pieKeys.forEach((k, r) => {
            rowOfKey.set(k, r);
          });
          return { n, rowOfKey, categories: writePieCategories(catCol, 'Category', pieKeys, n) };
        })();
    piePositions.forEach((pos, ring) => {
      const s = model.series[pos]!;
      const keys = pieKeysOf(s);
      // Per-ring layout: "Category <ring>" | "<ring name>", one row per point in the ring's own order.
      const own = shared
        ? null
        : { n: keys.length, categories: writePieCategories(nextCol++, `Category ${ring + 1}`, keys, keys.length) };
      const n = shared ? shared.n : own!.n;
      const col = nextCol++;
      const name = writeHeader(col, seriesName(s, pos), `series[${s.index}].name`, s.index);
      const formatCode = valueFormatOf(s);
      const style = valueStyle(formatCode);
      const cache: (number | null)[] = new Array<number | null>(n).fill(null);
      const counters = { nulls: 0, nonNumeric: 0 };
      const pointOffsets: number[] = [];
      s.points.forEach((p, j) => {
        const r = shared ? shared.rowOfKey.get(keys[j]!)! : j;
        pointOffsets.push(r);
        const v = yValue(p, counters);
        cache[r] = v;
        if (v !== null) builder.set(col, r + 1, { type: 'number', value: v }, style);
      });
      reportSeriesData(s, pos, counters.nulls, counters.nonNumeric, 0);
      const categories = shared ? shared.categories : own!.categories;
      ranges.push({
        seriesIndex: pos,
        nameRef: nameRefOf(col, name),
        categories: { ...categories, cache: categories.cache.slice() },
        values: { formula: rangeRef(sheetName, col, 1, refTo(n)), cache: padCache(cache, n, null), formatCode },
        bubbleSizes: null,
        pointOffsets,
        literalName: null,
      });
    });
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
      // Pre-1900 datetime x: plain (possibly negative) day numbers, no date format.
      const x = isDatetime || pre1900 ? msToExcelSerial(rawX) : rawX;
      xs.push(x);
      builder.set(
        xCol,
        r,
        { type: 'number', value: x },
        isDatetime ? { numberFormat: xFormat } : undefined,
        isDatetime ? xFormat.length + 2 : undefined,
      );
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
    const range: SeriesRange = {
      seriesIndex: pos,
      nameRef: nameRefOf(yCol, name),
      literalName: base,
      categories,
      values: { formula: rangeRef(sheetName, yCol, 1, refTo(n)), cache: padCache(ys, n, null), formatCode },
      bubbleSizes: isBubble
        ? { formula: rangeRef(sheetName, zCol, 1, refTo(n)), cache: padCache(zs, n, null), formatCode: 'General' }
        : null,
      pointOffsets,
    };
    const ebPos = errorBarOf.get(pos);
    if (ebPos !== undefined) {
      // Error bar points pair with the parent point at the same x (first unused one).
      const rowsByX = new Map<number, number[]>();
      s.points.forEach((p, j) => {
        const x = numericX(p, j);
        rowsByX.set(x, [...(rowsByX.get(x) ?? []), j]);
      });
      range.errorBars = writeErrorBarColumns(
        s,
        pos,
        model.series[ebPos]!,
        ebPos,
        n,
        (p, j) => rowsByX.get(numericX(p, j))?.shift(),
        ys,
        formatCode,
      );
    }
    ranges.push(range);
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

/**
 * Row label of a point on a category axis in label mode. A numeric x is the point's position on the
 * axis, so it is never replaced by `cats[j]` (its array position): it maps to the category at index x
 * when that exists, else (named point on an axis without that category) to its name, else to `String(x)`
 * as its own row (`offCategory`). Only points without x fall back to their name or position.
 */
function categoryLabel(p: PointModel, j: number, cats: string[] | null): { key: string; offCategory: boolean } {
  const named = p.name !== null && p.name !== '';
  if (p.x !== null && Number.isFinite(p.x)) {
    const isIndex = Number.isInteger(p.x) && p.x >= 0;
    if (isIndex && cats && p.x < cats.length) return { key: cats[p.x]!, offCategory: false };
    if (isIndex && named) return { key: p.name!, offCategory: false };
    return { key: String(p.x), offCategory: true };
  }
  if (named) return { key: p.name!, offCategory: false };
  if (cats && j < cats.length) return { key: cats[j]!, offCategory: false };
  return { key: String(j), offCategory: false };
}

/** ISO text for ms timestamps: date only when every value is at midnight UTC. */
function isoDateText(ms: readonly number[]): string[] {
  const dayOnly = ms.every((v) => ((v % 86_400_000) + 86_400_000) % 86_400_000 === 0);
  return ms.map((v) => {
    const iso = new Date(v).toISOString();
    return dayOnly ? iso.slice(0, iso.indexOf('T')) : iso.slice(0, iso.indexOf('.')).replace('T', ' ');
  });
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
