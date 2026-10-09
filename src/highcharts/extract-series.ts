/**
 * Series and point extraction.
 */

import type {
  Color,
  DataLabelPosition,
  DataLabelStyle,
  Fill,
  MarkerStyle,
  NumberFormat,
  PointModel,
  SeriesDataSemantics,
  SeriesKind,
  SeriesModel,
  Stacking,
  Stroke,
} from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import { fillToSolidColor, resolveSeriesColor, toFill } from '../translators/color-translator';
import { translateFormatString } from '../translators/number-format-translator';
import { markerSymbolFromHighcharts, strokeFromOptions } from '../translators/style-translator';
import { toFont, type CssStyleLike } from '../translators/typography-translator';
import { parseColor } from '../utils/colors';
import { readEffectiveStyle } from './css-resolver';
import { axisCategories, axisKind } from './extract-axes';
import { borderStroke, fillAt, fontFor, langSeparators, seriesShowsInLegend } from './extract-styles';
import { arr, bool, deepMerge, firstRec, get, num, rec, str, type Rec } from './guards';
import { isBarLike, isPieLike, isRangePointKind, kindOf, styledColorCount } from './series-types';
import type { ChartView, ExtractContext, HcPointLike, HcSeriesLike, SeriesView } from './types';

const MARKER_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>([
  'line',
  'spline',
  'area',
  'areaspline',
  'scatter',
  'bubble',
]);
const LINE_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>([
  'line',
  'spline',
  'area',
  'areaspline',
  'scatter',
  'arearange',
]);
const AREA_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>(['area', 'areaspline', 'arearange']);

/** `pointArrayMap` of common non-core types, used when no live series tells us. */
const POINT_ARRAY_MAPS: Readonly<Record<string, readonly string[]>> = {
  bubble: ['y', 'z'],
  arearange: ['low', 'high'],
  areasplinerange: ['low', 'high'],
  columnrange: ['low', 'high'],
  errorbar: ['low', 'high'],
  boxplot: ['low', 'q1', 'median', 'q3', 'high'],
  ohlc: ['open', 'high', 'low', 'close'],
  candlestick: ['open', 'high', 'low', 'close'],
  hlc: ['high', 'low', 'close'],
};

const HUGE_PX = 1e6;

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function percentOrNumber(x: unknown, relativeTo: number): number | undefined {
  const n = num(x);
  if (n !== undefined) return n;
  if (typeof x !== 'string') return undefined;
  const m = /^\s*(-?\d*\.?\d+)\s*(%|px)?\s*$/.exec(x);
  if (!m) return undefined;
  const v = Number(m[1]);
  return m[2] === '%' ? (v / 100) * relativeTo : v;
}

function solidColor(input: unknown, view: ChartView): Color | null {
  if (input === undefined || input === null) return null;
  return fillToSolidColor(toFill(input, view.resolver).fill);
}

/** Number of points of a series view (live points, else the source data). */
function pointCountOf(s: SeriesView): number {
  const live = s.rt?.points?.length;
  if (typeof live === 'number' && live > 0) return live;
  return rawData(s)?.length ?? 0;
}

function seriesKind(view: ChartView, s: SeriesView): SeriesKind {
  if (s.type === 'pie') return pieInnerSize(view, s) > 0 ? 'doughnut' : 'pie';
  return kindOf(s.type);
}

/** Inner size of a pie as a fraction of its outer diameter (0 for a plain pie). */
function pieInnerSize(view: ChartView, s: SeriesView): number {
  const center = s.rt?.center;
  const outerPx = num(center?.[2]);
  const innerPx = num(center?.[3]);
  if (outerPx !== undefined && outerPx > 0 && innerPx !== undefined) return clamp01(innerPx / outerPx);
  const inner = s.opts.innerSize;
  if (typeof inner === 'string' && inner.trim().endsWith('%')) return clamp01(percentOrNumber(inner, 1) ?? 0);
  const box = view.plotBox ?? { width: view.width, height: view.height };
  const plotMin = Math.min(box.width, box.height);
  const outer = percentOrNumber(s.opts.size, plotMin) ?? plotMin;
  const px = percentOrNumber(inner, outer);
  return px !== undefined && outer > 0 ? clamp01(px / outer) : 0;
}

function clamp01(n: number): number {
  return Number.isFinite(n) ? Math.min(1, Math.max(0, n)) : 0;
}

// ---------------------------------------------------------------------------
// Series-level styling
// ---------------------------------------------------------------------------

function styledModeColor(ctx: ExtractContext, s: SeriesView, kind: SeriesKind): Color | null {
  const { view } = ctx;
  // Explicit CSS variable overrides win (headless and browser): the series class color variable.
  const count = styledColorCount(view.opts);
  const variable = (i: number): Color | null => {
    const name = `--highcharts-color-${((i % count) + count) % count}`;
    return typeof view.cssVariables?.[name] === 'string' ? parseColor(`var(${name})`, view.resolver) : null;
  };
  const own = variable(s.colorIndex);
  if (own) {
    // Point-colored series (pies) need a variable per slice color too.
    const slices = isPieLike(kind) || s.opts.colorByPoint === true ? Math.min(count, pointCountOf(s)) : 0;
    let allResolved = true;
    for (let i = 0; i < slices && allResolved; i++) allResolved = variable(i) !== null;
    if (allResolved) return own;
  }
  if (view.browser && s.rt) {
    const lineLike = LINE_KINDS.has(kind);
    const firstGraphic = (s.rt.points ?? []).find((p) => p?.graphic?.element)?.graphic?.element;
    const element = lineLike
      ? (s.rt.graph?.element ?? s.rt.area?.element ?? firstGraphic)
      : (firstGraphic ?? s.rt.area?.element);
    const prop = lineLike && s.rt.graph?.element ? 'stroke' : 'fill';
    const computed = readEffectiveStyle(element, [prop]);
    const c = parseColor(computed[prop], view.resolver);
    if (c) return c;
  }
  ctx.diagnostics.report(
    'STYLED_MODE_FALLBACK',
    'approximated',
    'chart.styledMode',
    'Styled mode colors come from CSS that is not readable here; the default palette is used by color index.',
  );
  return view.palette.length > 0 ? (view.palette[s.colorIndex % view.palette.length] ?? null) : null;
}

/**
 * Whether Highcharts draws markers for the series.
 *
 * Highcharts: `marker.enabled ?? (radial axis ? true : closestPointRangePx >= enabledThreshold * radius)`.
 * HEURISTIC when `enabled` is undefined: in a real browser use the rendered marker graphics /
 * measured `closestPointRangePx`; headless (jsdom never lays out markers) estimate the point
 * spacing as plot length / (pointCount - 1) and apply the same threshold (default 2 × 4px).
 */
function markerEnabled(view: ChartView, s: SeriesView, kind: SeriesKind, m: Rec, pointCount: number): boolean {
  const explicit = bool(m.enabled);
  if (explicit !== undefined) return explicit;
  if (kind === 'scatter' || kind === 'bubble' || view.polar) return true;
  const radius = num(m.radius) ?? 4;
  const need = (num(m.enabledThreshold) ?? 2) * radius;
  if (view.browser && s.rt) {
    if ((s.rt.points ?? []).some((p) => p?.graphic)) return true;
    const px = num(s.rt.closestPointRangePx);
    if (px !== undefined && px < HUGE_PX) return px >= need;
  }
  if (pointCount <= 1) return true;
  const box = view.plotBox;
  const length = box ? (view.inverted ? box.height : box.width) : (view.inverted ? view.height : view.width) * 0.85;
  return length / (pointCount - 1) >= need;
}

function seriesMarker(
  view: ChartView,
  s: SeriesView,
  kind: SeriesKind,
  color: Color | null,
  pointCount: number,
): MarkerStyle | null {
  if (!MARKER_KINDS.has(kind)) return null;
  const m = rec(s.opts.marker) ?? {};
  const isBubble = kind === 'bubble';
  return {
    enabled: markerEnabled(view, s, kind, m, pointCount),
    symbol: markerSymbolFromHighcharts(m.symbol ?? s.symbol),
    radius: num(m.radius) ?? 4,
    fill: solidColor(m.fillColor, view) ?? color,
    stroke: solidColor(m.lineColor, view) ?? (isBubble ? color : parseColor('#ffffff')),
    strokeWidth: num(m.lineWidth) ?? (isBubble ? 1 : 0),
  };
}

function dataLabelPosition(kind: SeriesKind, dl: Rec, stacking: Stacking): DataLabelPosition {
  if (isPieLike(kind)) return (num(dl.distance) ?? 30) < 0 ? 'insideEnd' : 'outsideEnd';
  if (isBarLike(kind)) {
    // Highcharts: `inside` defaults to true for stacked columns; inside labels default to the middle.
    const inside = bool(dl.inside) ?? stacking !== null;
    if (!inside) return 'outsideEnd';
    if (dl.verticalAlign === 'top') return 'insideEnd';
    if (dl.verticalAlign === 'bottom') return 'insideBase';
    return 'center';
  }
  if (MARKER_KINDS.has(kind)) {
    const va = dl.verticalAlign ?? (kind === 'bubble' ? 'middle' : 'bottom');
    if (va === 'top') return 'below';
    if (va === 'middle') return 'center';
    return 'above';
  }
  return 'auto';
}

function seriesDataLabels(ctx: ExtractContext, s: SeriesView, kind: SeriesKind, stacking: Stacking): DataLabelStyle {
  const { view, diagnostics } = ctx;
  const dl = firstRec(s.opts.dataLabels) ?? {};
  const userDl = firstRec(s.userOpts.dataLabels) ?? {};
  const isPie = isPieLike(kind);
  const enabled = bool(dl.enabled) ?? (isPie && !s.rt);
  const base: DataLabelStyle = {
    enabled,
    font: fontFor(view, 'dataLabels', dl.style),
    format: null,
    position: dataLabelPosition(kind, dl, stacking),
    showValue: !isPie,
    showCategoryName: isPie,
    showSeriesName: false,
    showPercentage: false,
    background: fillAt(dl.backgroundColor, view, diagnostics, `${s.path}.dataLabels.backgroundColor`),
    border: borderStroke(view, dl.borderColor, dl.borderWidth, { color: '#333333', width: 0 }),
  };
  if (!enabled) return base;
  // Highcharts' own default formatter lives in the merged options; only a developer-supplied one counts.
  const formatter = typeof userDl.formatter === 'function' ? userDl.formatter : undefined;
  const format = str(dl.format);
  // Pie default label is the point name (built-in formatter).
  if (isPie && format === undefined && formatter === undefined) return base;
  const t = translateFormatString(format, formatter, {
    kind: 'dataLabel',
    property: `${s.path}.dataLabels`,
    ...langSeparators(view),
  });
  diagnostics.addAll(t.diagnostics);
  return {
    ...base,
    format: t.format,
    showValue: t.showValue,
    showCategoryName: t.showCategoryName,
    showSeriesName: t.showSeriesName,
    showPercentage: t.showPercentage,
  };
}

function seriesYFormat(ctx: ExtractContext, s: SeriesView): NumberFormat | null {
  const tooltip = rec(s.rt?.tooltipOptions) ?? deepMerge(ctx.view.opts.tooltip, s.opts.tooltip);
  const valueDecimals = num(tooltip.valueDecimals);
  const valuePrefix = str(tooltip.valuePrefix);
  const valueSuffix = str(tooltip.valueSuffix);
  if (valueDecimals === undefined && valuePrefix === undefined && valueSuffix === undefined) return null;
  const t = translateFormatString(undefined, undefined, {
    kind: 'value',
    property: `${s.path}.tooltip`,
    ...langSeparators(ctx.view),
    ...(valueDecimals !== undefined ? { valueDecimals } : {}),
    ...(valuePrefix !== undefined ? { valuePrefix } : {}),
    ...(valueSuffix !== undefined ? { valueSuffix } : {}),
  });
  ctx.diagnostics.addAll(t.diagnostics);
  return t.format;
}

// ---------------------------------------------------------------------------
// Point-level overrides
// ---------------------------------------------------------------------------

function pointBorder(view: ChartView, po: Rec): Stroke | null {
  if (po.borderColor === undefined && po.borderWidth === undefined) return null;
  return strokeFromOptions({ color: po.borderColor, width: po.borderWidth }, { width: 1 }, view.resolver);
}

function pointMarker(view: ChartView, po: Rec): Partial<MarkerStyle> | null {
  const m = rec(po.marker);
  if (!m) return null;
  const out: Partial<MarkerStyle> = {};
  const enabled = bool(m.enabled);
  if (enabled !== undefined) out.enabled = enabled;
  if (m.symbol !== undefined) out.symbol = markerSymbolFromHighcharts(m.symbol);
  const radius = num(m.radius);
  if (radius !== undefined) out.radius = radius;
  const fill = solidColor(m.fillColor, view);
  if (fill) out.fill = fill;
  const stroke = solidColor(m.lineColor, view);
  if (stroke) out.stroke = stroke;
  const lw = num(m.lineWidth);
  if (lw !== undefined) out.strokeWidth = lw;
  return Object.keys(out).length > 0 ? out : null;
}

function pointDataLabels(pc: PointContext, po: Rec, pointIndex: number): Partial<DataLabelStyle> | null {
  const dl = firstRec(po.dataLabels);
  if (!dl) return null;
  const out: Partial<DataLabelStyle> = {};
  const enabled = bool(dl.enabled);
  if (enabled !== undefined) out.enabled = enabled;
  if (rec(dl.style)) out.font = toFont(dl.style as CssStyleLike, null, pc.ctx.view.resolver);
  if (typeof dl.format === 'string' || typeof dl.formatter === 'function') {
    const t = translateFormatString(dl.format, dl.formatter, {
      kind: 'dataLabel',
      property: `${pc.s.path}.data[${pointIndex}].dataLabels`,
      ...langSeparators(pc.ctx.view),
    });
    collectPointIssues(pc, t.diagnostics, pointIndex);
    out.format = t.format;
    out.showValue = t.showValue;
    out.showCategoryName = t.showCategoryName;
    out.showSeriesName = t.showSeriesName;
    out.showPercentage = t.showPercentage;
  }
  return Object.keys(out).length > 0 ? out : null;
}

/** Per-point diagnostics of one kind, reported once per series (first occurrence + count). */
interface PointIssue {
  first: Diagnostic;
  count: number;
  firstIndices: number[];
}

const MAX_LISTED_INDICES = 5;

interface PointContext {
  ctx: ExtractContext;
  s: SeriesView;
  kind: SeriesKind;
  categories: string[] | null;
  /** The series is bound to a datetime x axis: x values move to the displayed wall-clock time. */
  datetime: boolean;
  /** Per-point palette colors are recorded (pie slices, colorByPoint). */
  byPoint: boolean;
  pointPalette: Color[];
  /** Aggregated per-point diagnostics, keyed by code. */
  issues: Map<string, PointIssue>;
}

function collectPointIssues(pc: PointContext, diagnostics: readonly Diagnostic[], pointIndex: number): void {
  for (const d of diagnostics) {
    const issue = pc.issues.get(d.code);
    if (!issue) {
      pc.issues.set(d.code, { first: d, count: 1, firstIndices: [pointIndex] });
      continue;
    }
    issue.count++;
    if (issue.firstIndices.length < MAX_LISTED_INDICES) issue.firstIndices.push(pointIndex);
  }
}

/** Reports each aggregated per-point issue once: the first occurrence, with `count` and `firstIndices`. */
function flushPointIssues(pc: PointContext): void {
  for (const { first, count, firstIndices } of pc.issues.values()) {
    const message =
      first.code === 'NON_NUMERIC_VALUE' && count > 1
        ? `${count} data points have a shape or value that is not numeric; they are skipped.`
        : first.message;
    pc.ctx.diagnostics.add({ ...first, message, details: { ...(first.details ?? {}), count, firstIndices } });
  }
  pc.issues.clear();
}

function categoryName(pc: PointContext, x: number | null): string | null {
  if (pc.categories === null || x === null || !Number.isInteger(x)) return null;
  return pc.categories[x] ?? null;
}

function slicedOffset(pc: PointContext, sliced: unknown): number | null {
  return sliced === true ? (num(pc.s.opts.slicedOffset) ?? 10) : null;
}

/** x of a live point: datetime values move to the displayed wall-clock time. */
function runtimeX(pc: PointContext, x: unknown): number | null {
  const n = num(x);
  if (n === undefined) return null;
  return pc.datetime ? pc.ctx.time.point(n) : n;
}

/** PointModel from a live Highcharts Point. */
function pointFromRuntime(pc: PointContext, p: HcPointLike, i: number): PointModel {
  const { view } = pc.ctx;
  const po = rec(p.options) ?? {};
  const rawX = num(p.x) ?? null;
  const x = runtimeX(pc, rawX);
  const ranged = isRangePointKind(pc.kind);
  // Range points (errorbar, columnrange, arearange) carry low/high; their y is not a data value.
  const y = ranged ? null : (num(p.y) ?? null);
  const low = ranged ? (num(p.low) ?? null) : null;
  const high = ranged ? (num(p.high) ?? null) : null;
  let color: Color | null = solidColor(po.color, view);
  if (color === null && pc.byPoint) {
    color = view.styledMode ? null : solidColor(p.color, view);
    if (color === null && pc.pointPalette.length > 0) {
      const ci = num(p.colorIndex) ?? num(p.index) ?? i;
      color = pc.pointPalette[ci % pc.pointPalette.length] ?? null;
    }
  }
  return {
    x,
    name: str(p.name) ?? categoryName(pc, rawX),
    y,
    z: num(p.z) ?? null,
    isNull: p.isNull === true || (ranged ? low === null || high === null : y === null),
    color,
    border: pointBorder(view, po),
    marker: pointMarker(view, po),
    sliced: slicedOffset(pc, p.sliced),
    dataLabels: pointDataLabels(pc, po, num(p.index) ?? i),
    visible: p.visible !== false,
    ...(ranged ? { low, high } : {}),
  };
}

/**
 * x of the k-th auto-incremented point (`pointStart` + k × `pointInterval`, in `pointIntervalUnit`
 * calendar steps when set), already converted to wall-clock time on datetime axes. `k` may be a
 * relative x value (`relativeXValue`).
 */
function xSequence(pc: PointContext): (k: number) => number {
  const { ctx, s } = pc;
  const o = s.opts;
  const interval = num(o.pointInterval) ?? 1;
  const unit = str(o.pointIntervalUnit);
  let start = num(o.pointStart) ?? 0;
  // A date string start (Highcharts 12+) is a wall-clock time already.
  let startIsWall = false;
  if (pc.datetime && typeof o.pointStart === 'string') {
    const parsed = ctx.time.parse(o.pointStart, false);
    if (parsed === undefined) {
      ctx.diagnostics.report(
        'NON_NUMERIC_VALUE',
        'approximated',
        `${s.path}.pointStart`,
        'pointStart is not a date; 0 is used.',
        {
          seriesIndex: s.index,
          details: { valueType: 'string' },
        },
      );
    } else {
      start = parsed;
      startIsWall = true;
    }
  }
  if (unit === 'day' || unit === 'month' || unit === 'year') {
    ctx.diagnostics.report(
      'APPROXIMATED_DATETIME',
      'approximated',
      `${s.path}.pointIntervalUnit`,
      'Calendar point intervals are stepped from pointStart on the dates the chart displays; month-end dates may differ from Highcharts, which steps from the previous point.',
      { severity: 'info', seriesIndex: s.index, details: { pointIntervalUnit: unit } },
    );
    // Calendar steps happen on the displayed (wall-clock) calendar, as Highcharts does in the chart time zone.
    const base = pc.datetime && !startIsWall ? ctx.time.bound(start) : start;
    const d = new Date(base);
    return (k) => {
      const steps = interval * k;
      const v = Date.UTC(
        d.getUTCFullYear() + (unit === 'year' ? steps : 0),
        d.getUTCMonth() + (unit === 'month' ? steps : 0),
        d.getUTCDate() + (unit === 'day' ? steps : 0),
        d.getUTCHours(),
        d.getUTCMinutes(),
        d.getUTCSeconds(),
        d.getUTCMilliseconds(),
      );
      return pc.datetime ? ctx.time.wallPoint(v) : v;
    };
  }
  return (k) => {
    const v = start + k * interval;
    if (!pc.datetime) return v;
    return startIsWall ? ctx.time.wallPoint(v) : ctx.time.point(v);
  };
}

function nonNumeric(pc: PointContext, i: number, value: unknown): void {
  const first = createDiagnostic(
    'NON_NUMERIC_VALUE',
    'approximated',
    `${pc.s.path}.data[${i}]`,
    'A data point has a shape or value that is not numeric; it is skipped.',
    {
      seriesIndex: pc.s.index,
      details: { valueType: Array.isArray(value) ? 'array' : typeof value },
    },
  );
  collectPointIssues(pc, [first], i);
}

/** Sets `obj[a][b]` for a dotted `series.keys` entry such as "marker.radius". */
function setPath(obj: Rec, path: string, value: unknown): void {
  const keys = path.split('.');
  let cur = obj;
  keys.forEach((k, j) => {
    if (j === keys.length - 1) {
      cur[k] = value;
      return;
    }
    const next = rec(cur[k]) ?? {};
    cur[k] = next;
    cur = next;
  });
}

/**
 * Parses `series.options.data` the way Highcharts' `Point.optionsToObject` does: numbers, nulls,
 * `[x, y]`, `[x, y, z]`, `[name, y]` (per `pointArrayMap`), arrays mapped through `series.keys`
 * and point objects. `relativeXValue` x values count `pointInterval`s from `pointStart`. On datetime
 * axes, date strings are parsed. Unknown shapes are skipped with NON_NUMERIC_VALUE; nothing is
 * fabricated.
 */
function pointsFromRaw(pc: PointContext, data: readonly unknown[]): PointModel[] {
  const { view } = pc.ctx;
  const o = pc.s.opts;
  const map = pc.s.rt?.pointArrayMap ?? POINT_ARRAY_MAPS[pc.s.type] ?? ['y'];
  const hasY = map.includes('y');
  const ranged = isRangePointKind(pc.kind);
  if (!hasY && !ranged) {
    pc.ctx.diagnostics.report(
      'NON_NUMERIC_VALUE',
      'approximated',
      `${pc.s.path}.data`,
      `Points of type "${pc.s.type}" have no single y value; y is left empty.`,
      {
        seriesIndex: pc.s.index,
        details: { pointArrayMap: [...map] },
      },
    );
  }
  const keys = (arr(o.keys) ?? []).filter((k): k is string => typeof k === 'string');
  const relative = o.relativeXValue === true;
  const xAt = xSequence(pc);
  let autoIndex = 0;
  let colorCounter = 0;
  const out: PointModel[] = [];

  /** Final x for an explicit x option; undefined when it is not usable. */
  const explicitX = (v: unknown): number | undefined => {
    if (typeof v === 'string' && pc.datetime) {
      const wall = pc.ctx.time.parse(v);
      return wall;
    }
    const n = num(v);
    if (n === undefined) return undefined;
    if (relative) return xAt(n);
    return pc.datetime ? pc.ctx.time.point(n) : n;
  };

  for (let i = 0; i < data.length; i++) {
    const item = data[i];
    let x: number | undefined;
    let hasExplicitX = false;
    let name: string | undefined;
    const values: Rec = {};
    let po: Rec = {};
    let source: unknown = item;
    if (Array.isArray(item) && keys.length > 0) {
      // series.keys: positional values become named point options.
      const mapped: Rec = {};
      keys.forEach((k, j) => {
        if (j < item.length) setPath(mapped, k, item[j]);
      });
      source = mapped;
    }
    if (source === null || typeof source === 'number') {
      if (typeof source === 'number' && !Number.isFinite(source)) {
        nonNumeric(pc, i, item);
        continue;
      }
      values[map[0] ?? 'y'] = source;
    } else if (Array.isArray(source)) {
      let k = 0;
      if (source.length > map.length) {
        const first: unknown = source[0];
        if (typeof first === 'string' && pc.datetime) {
          x = explicitX(first);
          if (x === undefined) {
            nonNumeric(pc, i, item);
            continue;
          }
          hasExplicitX = true;
        } else if (typeof first === 'string') {
          name = first;
        } else if (typeof first === 'number') {
          x = explicitX(first);
          if (x === undefined) {
            nonNumeric(pc, i, item);
            continue;
          }
          hasExplicitX = true;
        } else if (first !== null) {
          nonNumeric(pc, i, item);
          continue;
        }
        k = 1;
      }
      map.forEach((key, j) => {
        values[key] = source[k + j];
      });
    } else if (rec(source)) {
      po = source as Rec;
      for (const key of map) values[key] = po[key];
      if (po.z !== undefined) values.z = po.z;
      if (po.x !== undefined && po.x !== null) {
        x = explicitX(po.x);
        if (x === undefined) {
          nonNumeric(pc, i, item);
          continue;
        }
        hasExplicitX = true;
      }
      if (typeof po.name === 'string') name = po.name;
    } else {
      nonNumeric(pc, i, item);
      continue;
    }
    const yRaw = hasY && !ranged ? values.y : null;
    if (yRaw !== undefined && yRaw !== null && num(yRaw) === undefined) {
      nonNumeric(pc, i, item);
      continue;
    }
    const lowRaw = ranged ? values.low : null;
    const highRaw = ranged ? values.high : null;
    if (
      (lowRaw !== undefined && lowRaw !== null && num(lowRaw) === undefined) ||
      (highRaw !== undefined && highRaw !== null && num(highRaw) === undefined)
    ) {
      nonNumeric(pc, i, item);
      continue;
    }
    const low = num(lowRaw) ?? null;
    const high = num(highRaw) ?? null;
    const zRaw = values.z;
    const xValue = hasExplicitX && x !== undefined ? x : xAt(autoIndex++);
    const y = num(yRaw) ?? null;
    let color = solidColor(po.color, view);
    if (pc.byPoint) {
      const ci = colorCounter++;
      if (color === null && pc.pointPalette.length > 0) {
        const explicitIndex = num(po.colorIndex);
        color = pc.pointPalette[(explicitIndex ?? ci) % pc.pointPalette.length] ?? null;
      }
    }
    out.push({
      x: xValue,
      name: name ?? (isPieLike(pc.kind) ? 'Slice' : categoryName(pc, xValue)),
      y,
      z: num(zRaw) ?? null,
      isNull: ranged ? low === null || high === null : y === null,
      color,
      border: pointBorder(view, po),
      marker: pointMarker(view, po),
      sliced: slicedOffset(pc, po.sliced),
      dataLabels: pointDataLabels(pc, po, i),
      visible: po.visible !== false,
      ...(ranged ? { low, high } : {}),
    });
  }
  return out;
}

function columnLength(rt: HcSeriesLike | undefined): number | undefined {
  if (!rt || typeof rt.getColumn !== 'function') return undefined;
  try {
    const col = rt.getColumn('x');
    return col && typeof col.length === 'number' ? col.length : undefined;
  } catch {
    return undefined;
  }
}

/** A processed data column (`getColumn(name, true)`, or `processed*Data` on Highcharts 11). */
function processedColumn(rt: HcSeriesLike, name: 'x' | 'y' | 'z'): ArrayLike<unknown> | undefined {
  try {
    const col = typeof rt.getColumn === 'function' ? rt.getColumn(name, true) : undefined;
    if (col && typeof col.length === 'number' && col.length > 0) return col;
  } catch {
    // Fall through to the Highcharts 11 arrays.
  }
  const legacy = (rt as Rec)[`processed${name.toUpperCase()}Data`];
  return Array.isArray(legacy) && legacy.length > 0 ? legacy : undefined;
}

/**
 * Points of a boosted series from its processed data columns. The boost module replaces
 * `series.points` with pixel pseudo-points (no y), so they cannot be used. Null when the columns
 * are unavailable.
 */
function pointsFromColumns(pc: PointContext, rt: HcSeriesLike): PointModel[] | null {
  const ys = processedColumn(rt, 'y');
  if (!ys) return null;
  const xs = processedColumn(rt, 'x');
  const zs = processedColumn(rt, 'z');
  const out: PointModel[] = [];
  for (let i = 0; i < ys.length; i++) {
    const rawX = num(xs?.[i]) ?? i;
    const y = num(ys[i]) ?? null;
    let color: Color | null = null;
    if (pc.byPoint && pc.pointPalette.length > 0) color = pc.pointPalette[i % pc.pointPalette.length] ?? null;
    out.push({
      x: runtimeX(pc, rawX),
      name: categoryName(pc, rawX),
      y,
      z: num(zs?.[i]) ?? null,
      isNull: y === null,
      color,
      border: null,
      marker: null,
      sliced: null,
      dataLabels: null,
      visible: true,
    });
  }
  return out;
}

/** `series.options.data` as an array; typed arrays (y-only data) are accepted. */
function asDataArray(x: unknown): unknown[] | undefined {
  if (Array.isArray(x)) return x;
  if (ArrayBuffer.isView(x) && !(x instanceof DataView)) return Array.from(x as unknown as ArrayLike<unknown>);
  return undefined;
}

function rawData(s: SeriesView): unknown[] | undefined {
  return asDataArray(s.opts.data) ?? asDataArray(get(s.rt?.userOptions, 'data'));
}

function extractPoints(pc: PointContext): { points: PointModel[]; semantics: SeriesDataSemantics } {
  const { ctx, s } = pc;
  const rt = s.rt;
  const raw = rawData(s);
  const boosted = rt?.boosted === true;
  const runtimePoints = rt && !boosted ? (rt.points ?? []).filter((p): p is HcPointLike => !!p) : [];
  const runtimeData =
    rt && !boosted && runtimePoints.length === 0 ? (rt.data ?? []).filter((p): p is HcPointLike => !!p) : [];
  const rendered = runtimePoints.length > 0 ? runtimePoints : runtimeData;
  const columnPoints = rt && boosted ? pointsFromColumns(pc, rt) : null;
  const liveCount = columnPoints ? columnPoints.length : rendered.length;
  const grouped = rt?.hasGroupedData === true;
  const sourcePointCount = raw?.length ?? columnLength(rt) ?? liveCount;
  const renderedPointCount = rt ? liveCount : (raw?.length ?? 0);
  const cropped = rt?.cropped === true || (!grouped && liveCount > 0 && liveCount < sourcePointCount);
  const livePoints = (): PointModel[] => columnPoints ?? rendered.map((p, i) => pointFromRuntime(pc, p, i));

  let mode: 'rendered' | 'raw' = ctx.dataMode;
  let points: PointModel[];
  if (!rt) {
    // Options-only extraction: the source data is all there is.
    mode = 'raw';
    if (raw) {
      points = pointsFromRaw(pc, raw);
    } else {
      points = [];
      const data = s.opts.data;
      if (data !== undefined && data !== null) {
        ctx.diagnostics.report(
          'NON_NUMERIC_VALUE',
          'approximated',
          `${s.path}.data`,
          'The series data is not an array; the series has no points.',
          {
            seriesIndex: s.index,
            details: { valueType: typeof data },
          },
        );
      }
    }
  } else if (mode === 'raw') {
    if (raw) {
      points = pointsFromRaw(pc, raw);
    } else {
      ctx.diagnostics.report(
        'DATA_MODE_FALLBACK',
        'approximated',
        `${s.path}.data`,
        'The series source data is not available (e.g. loaded through a data table); the rendered points are exported.',
        { seriesIndex: s.index },
      );
      mode = 'rendered';
      points = livePoints();
    }
  } else if (boosted) {
    ctx.diagnostics.report(
      'DATA_MODE_FALLBACK',
      'approximated',
      `${s.path}.boostThreshold`,
      'Highcharts boost module active; values read from series data, not from the boosted pixel points.',
      {
        severity: 'info',
        seriesIndex: s.index,
        details: { source: columnPoints ? 'processed columns' : 'options.data' },
      },
    );
    if (columnPoints) {
      points = columnPoints;
    } else {
      mode = 'raw';
      points = raw ? pointsFromRaw(pc, raw) : [];
    }
  } else if (rendered.length > 0) {
    points = livePoints();
  } else if (raw && raw.length > 0) {
    // Never-rendered series (e.g. hidden from the start): its source data is what it would show.
    mode = 'raw';
    points = pointsFromRaw(pc, raw);
    ctx.diagnostics.report(
      'DATA_MODE_FALLBACK',
      'approximated',
      `${s.path}.data`,
      'The series has no rendered points (e.g. hidden since it was created); its source data is exported.',
      { severity: 'info', seriesIndex: s.index },
    );
  } else {
    points = [];
  }
  flushPointIssues(pc);

  if (mode === 'rendered' && rt) {
    if (grouped) {
      ctx.diagnostics.report(
        'DATA_GROUPED',
        'approximated',
        `${s.path}.dataGrouping`,
        'The chart shows grouped data; the grouped points are exported.',
        {
          severity: 'info',
          seriesIndex: s.index,
          details: { sourcePointCount, renderedPointCount },
        },
      );
    } else if (cropped) {
      ctx.diagnostics.report(
        'DATA_CROPPED',
        'approximated',
        `${s.path}.data`,
        'Only the points in the visible range are exported.',
        {
          severity: 'info',
          seriesIndex: s.index,
          details: { sourcePointCount, renderedPointCount },
        },
      );
    }
  }

  return { points, semantics: { mode, grouped, cropped, sourcePointCount, renderedPointCount } };
}

// ---------------------------------------------------------------------------
// Series
// ---------------------------------------------------------------------------

function extractOne(ctx: ExtractContext, s: SeriesView): SeriesModel {
  const { view, diagnostics } = ctx;
  const o = s.opts;
  const kind = seriesKind(view, s);
  const isPie = isPieLike(kind);

  const color = view.styledMode
    ? styledModeColor(ctx, s, kind)
    : resolveSeriesColor(s.explicitColor, s.colorIndex, view.palette, view.resolver);
  const stacking: Stacking = o.stacking === 'normal' || o.stacking === 'percent' ? o.stacking : null;

  const seriesPalette = (arr(o.colors) ?? [])
    .map((c) => parseColor(c, view.resolver))
    .filter((c): c is Color => c !== null);
  const xAxisView = isPie ? undefined : view.xAxes[s.xAxis];
  const pc: PointContext = {
    ctx,
    s,
    kind,
    categories: xAxisView ? axisCategories(xAxisView) : null,
    datetime: xAxisView !== undefined && axisKind(xAxisView) === 'datetime',
    byPoint: isPie || o.colorByPoint === true,
    pointPalette: seriesPalette.length > 0 ? seriesPalette : view.palette,
    issues: new Map(),
  };
  const { points, semantics } = extractPoints(pc);
  const linked = linkedParentId(ctx, s);

  let line: Stroke | null = null;
  if (LINE_KINDS.has(kind) || (kind === 'unknown' && num(o.lineWidth) !== undefined)) {
    const width = num(o.lineWidth) ?? (kind === 'scatter' ? 0 : kind === 'arearange' ? 1 : 2);
    const lineColor = AREA_KINDS.has(kind) ? o.lineColor : undefined;
    line = strokeFromOptions(
      { color: lineColor, width, dashStyle: o.dashStyle },
      { color, width, dash: 'solid' },
      view.resolver,
    );
  } else if (kind === 'errorbar') {
    // The stem (and whiskers) of an error bar: stemWidth/stemColor, else the series line.
    const width = num(o.stemWidth) ?? num(o.lineWidth) ?? 1;
    line = strokeFromOptions(
      { color: o.stemColor ?? undefined, width, dashStyle: o.stemDashStyle ?? o.dashStyle },
      { color, width, dash: 'solid' },
      view.resolver,
    );
  }

  let fill: Fill | null = null;
  let fillOpacity = 1;
  const explicitFill = view.styledMode ? undefined : s.explicitColor;
  const seriesFill = (): Fill | null =>
    explicitFill !== undefined && explicitFill !== null && typeof explicitFill === 'object'
      ? fillAt(explicitFill, view, diagnostics, `${s.path}.color`)
      : color
        ? { type: 'solid', color }
        : null;
  if (AREA_KINDS.has(kind)) {
    const fillColor = view.styledMode ? undefined : o.fillColor;
    if (fillColor !== undefined && fillColor !== null) {
      fill = fillAt(fillColor, view, diagnostics, `${s.path}.fillColor`);
      fillOpacity = 1; // Highcharts applies fillOpacity only to the derived fill (#18939).
    } else {
      fill = seriesFill();
      fillOpacity = num(o.fillOpacity) ?? 0.75;
    }
  } else if (isBarLike(kind) || isPie || kind === 'unknown') {
    fill = seriesFill();
  } else if (kind === 'bubble') {
    fill = seriesFill();
    fillOpacity = num(get(o, 'marker', 'fillOpacity')) ?? 0.5;
  }

  const border =
    isBarLike(kind) || isPie ? borderStroke(view, o.borderColor, o.borderWidth, { color: '#ffffff', width: 1 }) : null;

  let bars: SeriesModel['bars'] = null;
  if (isBarLike(kind)) {
    // Highcharts rounds bar corners by 3px by default; only an explicit setting is worth a diagnostic,
    // so read the user's own options (series, then plotOptions) rather than the merged defaults.
    const uo = s.userOpts;
    const plot = get(view.userOpts, 'plotOptions') as Record<string, unknown> | undefined;
    const br = uo.borderRadius ?? get(plot, s.type, 'borderRadius') ?? get(plot, 'series', 'borderRadius');
    const radius = num(br) ?? num(get(br, 'radius')) ?? 0;
    bars = { pointPadding: num(o.pointPadding) ?? 0.1, groupPadding: num(o.groupPadding) ?? 0.2, borderRadius: radius };
  }

  if (o.negativeColor !== undefined && o.negativeColor !== null) {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'approximated',
      `${s.path}.negativeColor`,
      'Negative colors are not reproduced; negative values use the series color.',
      {
        seriesIndex: s.index,
      },
    );
  }
  const zones = arr(o.zones);
  if (zones && zones.length > 0) {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'approximated',
      `${s.path}.zones`,
      'Zones are not reproduced; the whole series uses its base color.',
      {
        seriesIndex: s.index,
        details: { zones: zones.length },
      },
    );
  }

  if (points.length === 0) {
    diagnostics.report('EMPTY_SERIES', 'translated', `${s.path}.data`, `Series "${s.name}" has no data points.`, {
      severity: 'info',
      seriesIndex: s.index,
    });
  }

  return {
    id: str(o.id) ?? `series-${s.index}`,
    index: s.index,
    name: s.name,
    kind,
    sourceType: s.type,
    visible: s.visible,
    showInLegend: seriesShowsInLegend(s),
    xAxisIndex: s.xAxis,
    yAxisIndex: s.yAxis,
    color,
    stacking,
    stackGroup: o.stack === undefined || o.stack === null ? null : String(o.stack),
    line,
    fill,
    fillOpacity,
    border,
    marker: seriesMarker(view, s, kind, color, points.length),
    dataLabels: seriesDataLabels(ctx, s, kind, stacking),
    smooth: kind === 'spline' || kind === 'areaspline',
    bars,
    pie: isPie
      ? { innerSize: pieInnerSize(view, s), startAngle: num(o.startAngle) ?? 0, endAngle: num(o.endAngle) ?? null }
      : null,
    yFormat: seriesYFormat(ctx, s),
    points,
    dataSemantics: semantics,
    ...(linked !== null ? { linkedTo: linked } : {}),
  };
}

/**
 * Model id of the series `s` is linked to: the live `linkedParent`, else `linkedTo` (a series id or
 * `':previous'`, the default of errorbar series). Null when not linked or the parent is unknown.
 */
function linkedParentId(ctx: ExtractContext, s: SeriesView): string | null {
  const views = ctx.view.series;
  const idOf = (v: SeriesView): string => str(v.opts.id) ?? `series-${v.index}`;
  const lp = s.rt?.linkedParent;
  if (lp) {
    const v = views.find((x) => x.rt === lp);
    if (v) return idOf(v);
  }
  const linkedTo = s.opts.linkedTo === undefined && s.type === 'errorbar' ? ':previous' : s.opts.linkedTo;
  if (typeof linkedTo !== 'string' || linkedTo === '') return null;
  if (linkedTo === ':previous') {
    const at = views.indexOf(s);
    return at > 0 ? idOf(views[at - 1]!) : null;
  }
  const v = views.find((x) => x !== s && str(x.opts.id) === linkedTo);
  return v ? idOf(v) : null;
}

/** Extracts the series of the view, honoring `seriesVisibility`. */
export function extractSeries(ctx: ExtractContext): SeriesModel[] {
  const out: SeriesModel[] = [];
  for (const s of ctx.view.series) {
    if (!s.visible) {
      if (ctx.seriesVisibility === 'visible') {
        ctx.diagnostics.report(
          'HIDDEN_SERIES_EXCLUDED',
          'translated',
          `${s.path}.visible`,
          `Hidden series "${s.name}" is not exported.`,
          {
            severity: 'info',
            seriesIndex: s.index,
          },
        );
        continue;
      }
      ctx.diagnostics.report(
        'HIDDEN_SERIES_INCLUDED',
        'approximated',
        `${s.path}.visible`,
        `Hidden series "${s.name}" is exported and will be visible in Excel.`,
        {
          severity: 'info',
          seriesIndex: s.index,
        },
      );
    }
    out.push(extractOne(ctx, s));
  }
  return out;
}
