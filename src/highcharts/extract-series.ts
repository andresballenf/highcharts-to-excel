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
import { fillToSolidColor, resolveSeriesColor, toFill } from '../translators/color-translator';
import { translateFormatString } from '../translators/number-format-translator';
import { markerSymbolFromHighcharts, strokeFromOptions } from '../translators/style-translator';
import { toFont, type CssStyleLike } from '../translators/typography-translator';
import { parseColor } from '../utils/colors';
import { readEffectiveStyle } from './css-resolver';
import { axisCategories } from './extract-axes';
import { borderStroke, fillAt, fontFor } from './extract-styles';
import { arr, bool, deepMerge, firstRec, get, num, rec, str, type Rec } from './guards';
import type { ChartView, ExtractContext, HcPointLike, HcSeriesLike, SeriesView } from './types';

const DIRECT_KINDS: Readonly<Record<string, SeriesKind>> = {
  line: 'line',
  spline: 'spline',
  area: 'area',
  areaspline: 'areaspline',
  column: 'column',
  bar: 'bar',
  scatter: 'scatter',
  bubble: 'bubble',
};

const MARKER_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>(['line', 'spline', 'area', 'areaspline', 'scatter', 'bubble']);
const LINE_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>(['line', 'spline', 'area', 'areaspline', 'scatter']);
const AREA_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>(['area', 'areaspline']);
const BAR_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>(['column', 'bar']);
const PIE_KINDS: ReadonlySet<SeriesKind> = new Set<SeriesKind>(['pie', 'doughnut']);

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

const MS_PER_UNIT: Readonly<Record<string, number>> = { day: 86_400_000 };

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

function seriesKind(view: ChartView, s: SeriesView): SeriesKind {
  if (s.type === 'pie') return pieInnerSize(view, s) > 0 ? 'doughnut' : 'pie';
  return DIRECT_KINDS[s.type] ?? 'unknown';
}

/** Inner size of a pie as a fraction of its outer diameter (0 for a plain pie). */
function pieInnerSize(view: ChartView, s: SeriesView): number {
  const center = s.rt?.center;
  const outerPx = num(center?.[2]);
  const innerPx = num(center?.[3]);
  if (outerPx !== undefined && outerPx > 0 && innerPx !== undefined) return clamp01(innerPx / outerPx);
  const inner = s.opts.innerSize;
  if (typeof inner === 'string' && inner.trim().endsWith('%')) return clamp01((percentOrNumber(inner, 1) ?? 0));
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
  if (view.browser && s.rt) {
    const lineLike = LINE_KINDS.has(kind);
    const firstGraphic = (s.rt.points ?? []).find((p) => p?.graphic?.element)?.graphic?.element;
    const element = lineLike ? (s.rt.graph?.element ?? s.rt.area?.element ?? firstGraphic) : (firstGraphic ?? s.rt.area?.element);
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

function seriesMarker(view: ChartView, s: SeriesView, kind: SeriesKind, color: Color | null, pointCount: number): MarkerStyle | null {
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
  if (PIE_KINDS.has(kind)) return (num(dl.distance) ?? 30) < 0 ? 'insideEnd' : 'outsideEnd';
  if (BAR_KINDS.has(kind)) {
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
  const isPie = PIE_KINDS.has(kind);
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
  const t = translateFormatString(format, formatter, { kind: 'dataLabel', property: `${s.path}.dataLabels` });
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

function pointDataLabels(ctx: ExtractContext, s: SeriesView, po: Rec, pointIndex: number): Partial<DataLabelStyle> | null {
  const dl = firstRec(po.dataLabels);
  if (!dl) return null;
  const out: Partial<DataLabelStyle> = {};
  const enabled = bool(dl.enabled);
  if (enabled !== undefined) out.enabled = enabled;
  if (rec(dl.style)) out.font = toFont(dl.style as CssStyleLike, null, ctx.view.resolver);
  if (typeof dl.format === 'string' || typeof dl.formatter === 'function') {
    const t = translateFormatString(dl.format, dl.formatter, { kind: 'dataLabel', property: `${s.path}.data[${pointIndex}].dataLabels` });
    ctx.diagnostics.addAll(t.diagnostics);
    out.format = t.format;
    out.showValue = t.showValue;
    out.showCategoryName = t.showCategoryName;
    out.showSeriesName = t.showSeriesName;
    out.showPercentage = t.showPercentage;
  }
  return Object.keys(out).length > 0 ? out : null;
}

interface PointContext {
  ctx: ExtractContext;
  s: SeriesView;
  kind: SeriesKind;
  categories: string[] | null;
  /** Per-point palette colors are recorded (pie slices, colorByPoint). */
  byPoint: boolean;
  pointPalette: Color[];
}

function categoryName(pc: PointContext, x: number | null): string | null {
  if (pc.categories === null || x === null || !Number.isInteger(x)) return null;
  return pc.categories[x] ?? null;
}

function slicedOffset(pc: PointContext, sliced: unknown): number | null {
  return sliced === true ? (num(pc.s.opts.slicedOffset) ?? 10) : null;
}

/** PointModel from a live Highcharts Point. */
function pointFromRuntime(pc: PointContext, p: HcPointLike, i: number): PointModel {
  const { view } = pc.ctx;
  const po = rec(p.options) ?? {};
  const x = num(p.x) ?? null;
  const y = num(p.y) ?? null;
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
    name: str(p.name) ?? categoryName(pc, x),
    y,
    z: num(p.z) ?? null,
    isNull: p.isNull === true || y === null,
    color,
    border: pointBorder(view, po),
    marker: pointMarker(view, po),
    sliced: slicedOffset(pc, p.sliced),
    dataLabels: pointDataLabels(pc.ctx, pc.s, po, num(p.index) ?? i),
    visible: p.visible !== false,
  };
}

/** Creates the auto-incrementing x generator (`pointStart` + n × `pointInterval`). */
function xIncrementer(pc: PointContext): () => number {
  const o = pc.s.opts;
  const start = num(o.pointStart) ?? 0;
  const interval = num(o.pointInterval) ?? 1;
  const unit = str(o.pointIntervalUnit);
  let n = 0;
  if (unit !== undefined) {
    pc.ctx.diagnostics.report(
      'APPROXIMATED_DATETIME',
      'approximated',
      `${pc.s.path}.pointIntervalUnit`,
      'Calendar point intervals are computed in UTC; a chart time zone offset is not applied.',
      { severity: 'info', seriesIndex: pc.s.index, details: { pointIntervalUnit: unit } },
    );
  }
  return () => {
    const k = n++;
    if (unit === 'month' || unit === 'year') {
      const d = new Date(start);
      const months = (unit === 'year' ? 12 : 1) * interval * k;
      return Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + months, d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds(), d.getUTCMilliseconds());
    }
    const unitMs = unit !== undefined ? (MS_PER_UNIT[unit] ?? 1) : 1;
    return start + k * interval * unitMs;
  };
}

function nonNumeric(pc: PointContext, i: number, value: unknown): void {
  pc.ctx.diagnostics.report('NON_NUMERIC_VALUE', 'approximated', `${pc.s.path}.data[${i}]`, 'A data point has a shape or value that is not numeric; it is skipped.', {
    seriesIndex: pc.s.index,
    details: { valueType: Array.isArray(value) ? 'array' : typeof value },
  });
}

/**
 * Parses `series.options.data` the way Highcharts' `Point.optionsToObject` does: numbers, nulls,
 * `[x, y]`, `[x, y, z]`, `[name, y]` (per `pointArrayMap`) and point objects. Unknown shapes are
 * skipped with NON_NUMERIC_VALUE; nothing is fabricated.
 */
function pointsFromRaw(pc: PointContext, data: readonly unknown[]): PointModel[] {
  const { view } = pc.ctx;
  const map = pc.s.rt?.pointArrayMap ?? POINT_ARRAY_MAPS[pc.s.type] ?? ['y'];
  const hasY = map.includes('y');
  if (!hasY) {
    pc.ctx.diagnostics.report('NON_NUMERIC_VALUE', 'approximated', `${pc.s.path}.data`, `Points of type "${pc.s.type}" have no single y value; y is left empty.`, {
      seriesIndex: pc.s.index,
      details: { pointArrayMap: [...map] },
    });
  }
  const nextX = xIncrementer(pc);
  let colorCounter = 0;
  const out: PointModel[] = [];
  data.forEach((item, i) => {
    let x: number | null | undefined;
    let name: string | undefined;
    const values: Rec = {};
    let po: Rec = {};
    if (item === null || typeof item === 'number') {
      if (typeof item === 'number' && !Number.isFinite(item)) return nonNumeric(pc, i, item);
      values[map[0] ?? 'y'] = item;
    } else if (Array.isArray(item)) {
      let k = 0;
      if (item.length > map.length) {
        const first: unknown = item[0];
        if (typeof first === 'string') name = first;
        else if (typeof first === 'number') x = first;
        else if (first !== null) return nonNumeric(pc, i, item);
        k = 1;
      }
      map.forEach((key, j) => {
        values[key] = item[k + j];
      });
    } else if (rec(item)) {
      po = item as Rec;
      for (const key of map) values[key] = po[key];
      if (po.z !== undefined) values.z = po.z;
      if (po.x !== undefined && po.x !== null) {
        if (num(po.x) === undefined) return nonNumeric(pc, i, item);
        x = num(po.x);
      }
      if (typeof po.name === 'string') name = po.name;
    } else {
      return nonNumeric(pc, i, item);
    }
    const yRaw = hasY ? values.y : null;
    if (yRaw !== undefined && yRaw !== null && num(yRaw) === undefined) return nonNumeric(pc, i, item);
    const zRaw = values.z;
    const xValue = x ?? nextX();
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
      name: name ?? (PIE_KINDS.has(pc.kind) ? 'Slice' : categoryName(pc, xValue)),
      y,
      z: num(zRaw) ?? null,
      isNull: y === null,
      color,
      border: pointBorder(view, po),
      marker: pointMarker(view, po),
      sliced: slicedOffset(pc, po.sliced),
      dataLabels: pointDataLabels(pc.ctx, pc.s, po, i),
      visible: po.visible !== false,
    });
  });
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

function rawData(s: SeriesView): unknown[] | undefined {
  return arr(s.opts.data) ?? arr(get(s.rt?.userOptions, 'data'));
}

function extractPoints(pc: PointContext): { points: PointModel[]; semantics: SeriesDataSemantics } {
  const { ctx, s } = pc;
  const rt = s.rt;
  const raw = rawData(s);
  const runtimePoints = rt ? (rt.points ?? []).filter((p): p is HcPointLike => !!p) : [];
  const runtimeData = rt && runtimePoints.length === 0 ? (rt.data ?? []).filter((p): p is HcPointLike => !!p) : [];
  const rendered = runtimePoints.length > 0 ? runtimePoints : runtimeData;
  const grouped = rt?.hasGroupedData === true;
  const sourcePointCount = raw?.length ?? columnLength(rt) ?? rendered.length;
  const renderedPointCount = rt ? rendered.length : (raw?.length ?? 0);
  const cropped = rt?.cropped === true || (!grouped && rendered.length > 0 && rendered.length < sourcePointCount);

  let mode: 'rendered' | 'raw' = ctx.dataMode;
  let points: PointModel[];
  if (mode === 'raw' || !rt) {
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
      points = rendered.map((p, i) => pointFromRuntime(pc, p, i));
    }
  } else if (rendered.length > 0) {
    points = rendered.map((p, i) => pointFromRuntime(pc, p, i));
  } else {
    // Never-rendered series (e.g. hidden from the start): its source data is what it would show.
    points = raw ? pointsFromRaw(pc, raw) : [];
  }

  if (mode === 'rendered' && rt) {
    if (grouped) {
      ctx.diagnostics.report('DATA_GROUPED', 'approximated', `${s.path}.dataGrouping`, 'The chart shows grouped data; the grouped points are exported.', {
        severity: 'info',
        seriesIndex: s.index,
        details: { sourcePointCount, renderedPointCount },
      });
    } else if (cropped) {
      ctx.diagnostics.report('DATA_CROPPED', 'approximated', `${s.path}.data`, 'Only the points in the visible range are exported.', {
        severity: 'info',
        seriesIndex: s.index,
        details: { sourcePointCount, renderedPointCount },
      });
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
  const isPie = PIE_KINDS.has(kind);

  const color = view.styledMode ? styledModeColor(ctx, s, kind) : resolveSeriesColor(s.explicitColor, s.colorIndex, view.palette, view.resolver);
  const stacking: Stacking = o.stacking === 'normal' || o.stacking === 'percent' ? o.stacking : null;

  const seriesPalette = (arr(o.colors) ?? []).map((c) => parseColor(c, view.resolver)).filter((c): c is Color => c !== null);
  const pc: PointContext = {
    ctx,
    s,
    kind,
    categories: view.xAxes[s.xAxis] ? axisCategories(view.xAxes[s.xAxis]!) : null,
    byPoint: isPie || o.colorByPoint === true,
    pointPalette: seriesPalette.length > 0 ? seriesPalette : view.palette,
  };
  const { points, semantics } = extractPoints(pc);

  let line: Stroke | null = null;
  if (LINE_KINDS.has(kind) || (kind === 'unknown' && num(o.lineWidth) !== undefined)) {
    const width = num(o.lineWidth) ?? (kind === 'scatter' ? 0 : 2);
    const lineColor = AREA_KINDS.has(kind) ? o.lineColor : undefined;
    line = strokeFromOptions({ color: lineColor, width, dashStyle: o.dashStyle }, { color, width, dash: 'solid' }, view.resolver);
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
  } else if (BAR_KINDS.has(kind) || isPie || kind === 'unknown') {
    fill = seriesFill();
  } else if (kind === 'bubble') {
    fill = seriesFill();
    fillOpacity = num(get(o, 'marker', 'fillOpacity')) ?? 0.5;
  }

  const border = BAR_KINDS.has(kind) || isPie ? borderStroke(view, o.borderColor, o.borderWidth, { color: '#ffffff', width: 1 }) : null;

  let bars: SeriesModel['bars'] = null;
  if (BAR_KINDS.has(kind)) {
    const br = o.borderRadius;
    const radius = num(br) ?? num(get(br, 'radius')) ?? (br === undefined ? 3 : 0);
    bars = { pointPadding: num(o.pointPadding) ?? 0.1, groupPadding: num(o.groupPadding) ?? 0.2, borderRadius: radius };
  }

  if (o.negativeColor !== undefined && o.negativeColor !== null) {
    diagnostics.report('UNSUPPORTED_STYLE', 'approximated', `${s.path}.negativeColor`, 'Negative colors are not reproduced; negative values use the series color.', {
      seriesIndex: s.index,
    });
  }
  const zones = arr(o.zones);
  if (zones && zones.length > 0) {
    diagnostics.report('UNSUPPORTED_STYLE', 'approximated', `${s.path}.zones`, 'Zones are not reproduced; the whole series uses its base color.', {
      seriesIndex: s.index,
      details: { zones: zones.length },
    });
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
    pie: isPie ? { innerSize: pieInnerSize(view, s), startAngle: num(o.startAngle) ?? 0, endAngle: num(o.endAngle) ?? null } : null,
    yFormat: seriesYFormat(ctx, s),
    points,
    dataSemantics: semantics,
  };
}

/** Extracts the series of the view, honoring `seriesVisibility`. */
export function extractSeries(ctx: ExtractContext): SeriesModel[] {
  const out: SeriesModel[] = [];
  for (const s of ctx.view.series) {
    if (!s.visible) {
      if (ctx.seriesVisibility === 'visible') {
        ctx.diagnostics.report('HIDDEN_SERIES_EXCLUDED', 'translated', `${s.path}.visible`, `Hidden series "${s.name}" is not exported.`, {
          severity: 'info',
          seriesIndex: s.index,
        });
        continue;
      }
      ctx.diagnostics.report('HIDDEN_SERIES_INCLUDED', 'approximated', `${s.path}.visible`, `Hidden series "${s.name}" is exported and will be visible in Excel.`, {
        severity: 'info',
        seriesIndex: s.index,
      });
    }
    out.push(extractOne(ctx, s));
  }
  return out;
}
