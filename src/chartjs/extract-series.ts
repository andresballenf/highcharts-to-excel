/**
 * Chart.js datasets → IR series and points.
 */

import type {
  AxisKind,
  Color,
  Fill,
  MarkerStyle,
  NumberFormat,
  PointModel,
  SeriesKind,
  SeriesModel,
  Stacking,
  Stroke,
} from '../types/chart-model';
import { isColorEqual } from '../utils/colors';
import { labelText } from './extract-axes';
import {
  colorOf,
  dashFromBorderDash,
  elementKindsOf,
  pointStyleToSymbol,
  resolveOption,
  solid,
} from './extract-styles';
import { get, num, pick, rec, resolveKey, str } from './guards';
import type { WallClock } from './time';
import type { ChartJsView, DatasetView, ExtractContext, Rec, ScaleView } from './types';

const PIE_TYPES: ReadonlySet<string> = new Set(['pie', 'doughnut']);
const POLAR_TYPES: ReadonlySet<string> = new Set(['radar', 'polarArea']);

/** What the series extractor needs to know about the dataset's scales. */
export interface SeriesAxes {
  xAxisIndex: number;
  yAxisIndex: number;
  indexKind: AxisKind;
  indexScale: ScaleView | undefined;
  valueScale: ScaleView | undefined;
  /** Category labels of the index scale. */
  categories: unknown[];
  yFormat: NumberFormat | null;
}

interface RawValue {
  index: unknown;
  value: unknown;
  r: unknown;
  /** The index came from the data itself (not the dataset position). */
  explicitIndex: boolean;
}

interface Issues {
  nonNumeric: number;
  floating: number;
  badTime: number;
}

function toNumber(v: unknown): { value: number | null; bad: boolean } {
  if (v === null || v === undefined) return { value: null, bad: false };
  if (typeof v === 'number') return { value: Number.isFinite(v) ? v : null, bad: false };
  if (typeof v === 'string') {
    if (v.trim() === '') return { value: null, bad: true };
    const n = Number(v);
    return Number.isFinite(n) ? { value: n, bad: false } : { value: null, bad: true };
  }
  return { value: null, bad: true };
}

/** The dataset's data array, honoring `dataMode` when the decimation plugin replaced it. */
function sourceData(
  ctx: ExtractContext,
  dv: DatasetView,
): { items: unknown[]; grouped: boolean; sourceCount: number; renderedCount: number } {
  const ds = dv.ds as Rec & { _decimated?: unknown; _data?: unknown };
  const raw = ds.data;
  let items: unknown[];
  if (Array.isArray(raw)) items = raw;
  else if (rec(raw)) {
    // Object data ({ Jan: 10 }) is parsed by Chart.js as { [indexAxis]: key, [valueAxis]: value }.
    const valueAxis = dv.indexAxis === 'x' ? 'y' : 'x';
    items = Object.entries(rec(raw)!).map(([k, v]) => ({ [dv.indexAxis]: k, [valueAxis]: v }));
  } else items = [];
  const original = Array.isArray(ds._data) ? ds._data : null;
  if (ds._decimated === true && original) {
    if (ctx.dataMode === 'raw') {
      return { items: original, grouped: false, sourceCount: original.length, renderedCount: items.length };
    }
    return { items, grouped: true, sourceCount: original.length, renderedCount: items.length };
  }
  return { items, grouped: false, sourceCount: items.length, renderedCount: items.length };
}

function parsingKeys(view: ChartJsView, dv: DatasetView): { x: string; y: string; key: string } {
  const p = pick(dv.ds.parsing, get(view.options, 'datasets', dv.type, 'parsing'), view.options.parsing);
  const r = rec(p) ?? {};
  return { x: str(r.xAxisKey) ?? 'x', y: str(r.yAxisKey) ?? 'y', key: str(r.key) ?? 'value' };
}

/** Reads index/value/radius of every point (live parsed values first, else the raw data). */
function rawValues(view: ChartJsView, dv: DatasetView, items: unknown[], useParsed: boolean): RawValue[] {
  const indexAxis = dv.indexAxis;
  const valueAxis = indexAxis === 'x' ? 'y' : 'x';
  const parsed = useParsed ? dv.meta?._parsed : undefined;
  const isPie = PIE_TYPES.has(dv.type);
  if (parsed && parsed.length === items.length) {
    return parsed.map((p, j) => {
      if (isPie) return { index: j, value: p, r: undefined, explicitIndex: false };
      const o = rec(p) ?? {};
      const custom = o._custom;
      const floating = rec(custom) && 'barStart' in (custom as Rec);
      return {
        index: o[indexAxis],
        value: floating ? [get(custom, 'start'), get(custom, 'end')] : o[valueAxis],
        r: dv.type === 'bubble' ? custom : undefined,
        explicitIndex: true,
      };
    });
  }
  const keys = parsingKeys(view, dv);
  const indexKey = indexAxis === 'x' ? keys.x : keys.y;
  const valueKey = indexAxis === 'x' ? keys.y : keys.x;
  const parsingOff = dv.ds.parsing === false || view.options.parsing === false;
  // Chart.js picks one parser for the whole dataset from its first item (array, object or primitive).
  const first = items[0];
  const mode = isPie ? 'pie' : Array.isArray(first) ? 'array' : rec(first) ? 'object' : 'primitive';
  return items.map((item, j): RawValue => {
    if (mode === 'pie') {
      const o = rec(item);
      return { index: j, value: o ? resolveKey(o, keys.key) : item, r: undefined, explicitIndex: false };
    }
    if (mode === 'array') {
      const a = Array.isArray(item) ? item : [];
      if (dv.type === 'bar') return { index: undefined, value: item, r: undefined, explicitIndex: false };
      return { index: a[0], value: a[1], r: a[2], explicitIndex: true };
    }
    if (mode === 'object') {
      const o = rec(item) ?? {};
      const index = resolveKey(o, parsingOff ? indexAxis : indexKey);
      const value = resolveKey(o, parsingOff ? valueAxis : valueKey);
      return { index, value, r: o.r, explicitIndex: true };
    }
    return { index: undefined, value: item, r: undefined, explicitIndex: false };
  });
}

/** Resolves a point's x (and category name) on the dataset's index scale. */
function pointX(
  raw: RawValue,
  j: number,
  axes: SeriesAxes,
  time: WallClock,
  live: boolean,
  issues: Issues,
): { x: number | null; name: string | null } {
  const labels = axes.categories;
  const indexValue = raw.explicitIndex ? raw.index : axes.indexKind === 'category' ? j : labels[j];
  switch (axes.indexKind) {
    case 'category': {
      if (typeof indexValue === 'number' && Number.isInteger(indexValue) && indexValue >= 0) {
        const label = labels[indexValue];
        return { x: indexValue, name: label !== undefined ? labelText(label) : null };
      }
      if (typeof indexValue === 'string') {
        const at = labels.findIndex((l) => labelText(l) === indexValue);
        return { x: at >= 0 ? at : null, name: indexValue };
      }
      const label = labels[j];
      return { x: j, name: label !== undefined ? labelText(label) : null };
    }
    case 'datetime': {
      if (indexValue === undefined || indexValue === null) return { x: null, name: null };
      // Live parsed values are adapter-parsed instants.
      const x =
        live && raw.explicitIndex && typeof indexValue === 'number'
          ? time.fromInstant(indexValue)
          : time.parse(indexValue);
      if (x === null) issues.badTime++;
      return { x, name: null };
    }
    default: {
      const { value } = toNumber(indexValue);
      return { x: value, name: null };
    }
  }
}

function emptyPoint(): PointModel {
  return {
    x: null,
    name: null,
    y: null,
    z: null,
    isNull: false,
    color: null,
    border: null,
    marker: null,
    sliced: null,
    dataLabels: null,
    visible: true,
  };
}

function sameColor(a: Color | null, b: Color | null): boolean {
  if (a === null || b === null) return a === b;
  return isColorEqual(a, b);
}

/** Chart.js `cutout` (px number or "NN%") → fraction of the outer radius. */
function cutoutFraction(ctx: ExtractContext, dv: DatasetView, cutout: unknown): number {
  if (typeof cutout === 'string') {
    const m = /^\s*(\d*\.?\d+)\s*%\s*$/.exec(cutout);
    if (m) return Math.min(0.95, Math.max(0, Number(m[1]) / 100));
    const n = Number(cutout);
    if (Number.isFinite(n)) return cutoutFraction(ctx, dv, n);
    return 0;
  }
  const px = num(cutout);
  if (px === undefined || px <= 0) return 0;
  const { view } = ctx;
  const area = view.live?.chartArea;
  const w = num(area?.width) ?? view.width;
  const h = num(area?.height) ?? view.height;
  const radius = Math.max(1, Math.min(w, h) / 2);
  ctx.diagnostics.report(
    'APPROXIMATED_LAYOUT',
    'approximated',
    `${dv.path}.cutout`,
    'A pixel cutout is converted to a hole size relative to the chart size.',
    { severity: 'info', seriesIndex: dv.index, details: { cutoutPx: px, radiusPx: radius } },
  );
  return Math.min(0.95, px / radius);
}

function datasetOpt(ctx: ExtractContext, dv: DatasetView, key: string, extra: { prefix?: string } = {}): unknown {
  const kinds = elementKindsOf(dv.type);
  const typeOpts = rec(get(ctx.view.options, 'datasets', dv.type)) ?? {};
  if (kinds.dataset === null) return pick(dv.ds[key], typeOpts[key], ctx.view.options[key]);
  return resolveOption(
    ctx.view,
    dv,
    { element: kinds.dataset, key, index: dv.index, liveElement: 'dataset', ...extra, indexable: key !== 'borderDash' },
    ctx.diagnostics,
  );
}

function controllerOpt(ctx: ExtractContext, dv: DatasetView, key: string, fallback: unknown): unknown {
  const typeOpts = rec(get(ctx.view.options, 'datasets', dv.type)) ?? {};
  return pick(dv.ds[key], typeOpts[key], ctx.view.options[key], fallback);
}

function pointOpt(ctx: ExtractContext, dv: DatasetView, key: string, j: number): unknown {
  const kinds = elementKindsOf(dv.type);
  return resolveOption(
    ctx.view,
    dv,
    {
      element: kinds.data,
      key,
      index: j,
      liveElement: 'data',
      ...(kinds.data === 'point' ? { prefix: 'point' } : {}),
    },
    ctx.diagnostics,
  );
}

interface PointStyle {
  bg: Color | null;
  border: Color | null;
  borderWidth: number;
  radius: number;
  symbol: MarkerStyle['symbol'];
  offset: number;
}

function pointStyleAt(ctx: ExtractContext, dv: DatasetView, j: number): PointStyle {
  const p = `${dv.path}`;
  return {
    bg: colorOf(pointOpt(ctx, dv, 'backgroundColor', j), `${p}.backgroundColor`, ctx.diagnostics, dv.index),
    border: colorOf(pointOpt(ctx, dv, 'borderColor', j), `${p}.borderColor`, ctx.diagnostics, dv.index),
    borderWidth: num(pointOpt(ctx, dv, 'borderWidth', j)) ?? 0,
    radius: num(pointOpt(ctx, dv, 'radius', j)) ?? 3,
    symbol: pointStyleToSymbol(pointOpt(ctx, dv, 'pointStyle', j)),
    offset: num(pointOpt(ctx, dv, 'offset', j)) ?? 0,
  };
}

interface FillTarget {
  filled: boolean;
  /** Fill semantics Excel reproduces (to the axis, or between stacked series). */
  exact: boolean;
  raw: unknown;
}

function fillTarget(fill: unknown, stacked: boolean): FillTarget {
  const target = rec(fill) && 'target' in (fill as Rec) ? (fill as Rec).target : fill;
  if (target === undefined || target === null || target === false || target === 'false' || target === 'shape') {
    return { filled: false, exact: true, raw: fill };
  }
  const toAxis = target === true || target === 'origin' || target === 'start';
  const betweenStacked = stacked && (target === '-1' || target === 'stack');
  const exact = (toAxis || betweenStacked) && !(rec(fill) && ('above' in (fill as Rec) || 'below' in (fill as Rec)));
  return { filled: true, exact, raw: fill };
}

function seriesKindOf(dv: DatasetView, filled: boolean, smooth: boolean, innerSize: number): SeriesKind {
  switch (dv.type) {
    case 'bar':
      return dv.indexAxis === 'y' ? 'bar' : 'column';
    case 'line':
      return filled ? (smooth ? 'areaspline' : 'area') : smooth ? 'spline' : 'line';
    case 'scatter':
      return 'scatter';
    case 'bubble':
      return 'bubble';
    case 'pie':
    case 'doughnut':
      return innerSize > 0 ? 'doughnut' : 'pie';
    default:
      return 'unknown';
  }
}

/** Extracts one dataset. */
export function extractDataset(ctx: ExtractContext, dv: DatasetView, axes: SeriesAxes, time: WallClock): SeriesModel {
  const { view, diagnostics } = ctx;
  const path = dv.path;
  const isPie = PIE_TYPES.has(dv.type);
  const kinds = elementKindsOf(dv.type);

  if (POLAR_TYPES.has(dv.type)) {
    diagnostics.report(
      'UNSUPPORTED_SERIES_TYPE',
      'unsupported',
      `${path}.type`,
      `Chart.js "${dv.type}" charts have no editable Excel equivalent; the dataset is not exported to the chart.`,
      { seriesIndex: dv.index, details: { sourceType: dv.type } },
    );
  }

  // --- Values ------------------------------------------------------------------------------------
  const { items, grouped, sourceCount, renderedCount } = sourceData(ctx, dv);
  if (grouped) {
    diagnostics.report(
      'DATA_GROUPED',
      'approximated',
      `${path}.data`,
      'The decimation plugin reduced this dataset; the decimated points are exported (use dataMode "raw" for all points).',
      { seriesIndex: dv.index, details: { sourcePoints: sourceCount, renderedPoints: renderedCount } },
    );
  }
  const live = dv.meta?._parsed !== undefined && dv.meta._parsed.length === items.length;
  const raws = rawValues(view, dv, items, live);
  const issues: Issues = { nonNumeric: 0, floating: 0, badTime: 0 };
  const dataVisibility = isPie && view.live?.getDataVisibility ? view.live.getDataVisibility.bind(view.live) : null;

  const points: PointModel[] = raws.map((raw, j) => {
    const p = emptyPoint();
    if (isPie) {
      const label = view.labels[j];
      p.name = label !== undefined ? labelText(label) : null;
      if (dataVisibility) {
        try {
          p.visible = dataVisibility(j) !== false;
        } catch {
          // keep visible
        }
      }
    } else {
      const { x, name } = pointX(raw, j, axes, time, live, issues);
      p.x = x;
      p.name = name;
    }
    if (Array.isArray(raw.value)) {
      issues.floating++;
      p.isNull = true;
      return p;
    }
    const { value, bad } = toNumber(raw.value);
    if (bad) issues.nonNumeric++;
    p.y = value;
    p.isNull = value === null;
    if (dv.type === 'bubble') {
      // Chart.js draws a missing or zero "r" with the dataset's radius option.
      const r = toNumber(raw.r).value;
      p.z = r !== null && r !== 0 ? r : (num(pointOpt(ctx, dv, 'radius', j)) ?? 3);
    }
    return p;
  });

  if (issues.floating > 0) {
    diagnostics.report(
      'NON_NUMERIC_VALUE',
      'unsupported',
      `${path}.data`,
      'Floating bars ([start, end] values) have no Excel equivalent; their cells are left empty.',
      { seriesIndex: dv.index, details: { count: issues.floating } },
    );
  }
  if (issues.nonNumeric > 0) {
    diagnostics.report(
      'NON_NUMERIC_VALUE',
      'approximated',
      `${path}.data`,
      'Some values are not numbers; their cells are left empty.',
      { seriesIndex: dv.index, details: { count: issues.nonNumeric } },
    );
  }
  if (issues.badTime > 0) {
    diagnostics.report(
      'NON_NUMERIC_VALUE',
      'approximated',
      `${path}.data`,
      'Some time values could not be parsed without the chart date adapter; their x cells are left empty.',
      { seriesIndex: dv.index, details: { count: issues.badTime } },
    );
  }
  if (dv.type === 'bubble' && points.length > 0) {
    diagnostics.report(
      'APPROXIMATED_LAYOUT',
      'approximated',
      `${path}.data`,
      'Chart.js bubble "r" is a radius in pixels; Excel sizes bubbles relative to the largest "r" value.',
      { severity: 'info', seriesIndex: dv.index },
    );
  }

  // --- Geometry and kind ---------------------------------------------------------------------------
  const valueStacked = axes.valueScale ? pick(axes.valueScale.opts.stacked, false) : false;
  const stacked = valueStacked === true || valueStacked === 'single';
  const fill = kinds.dataset === 'line' ? fillTarget(datasetOpt(ctx, dv, 'fill'), stacked) : fillTarget(false, false);
  const tension = num(datasetOpt(ctx, dv, 'tension')) ?? 0;
  const smooth = dv.type !== 'scatter' && (tension > 0 || datasetOpt(ctx, dv, 'cubicInterpolationMode') === 'monotone');

  let innerSize = 0;
  let pie: SeriesModel['pie'] = null;
  if (isPie) {
    innerSize = cutoutFraction(ctx, dv, controllerOpt(ctx, dv, 'cutout', dv.type === 'doughnut' ? '50%' : 0));
    const rotation = num(controllerOpt(ctx, dv, 'rotation', 0)) ?? 0;
    const circumference = num(controllerOpt(ctx, dv, 'circumference', 360)) ?? 360;
    pie = { innerSize, startAngle: rotation, endAngle: circumference < 360 ? rotation + circumference : null };
  }
  const kind = seriesKindOf(dv, fill.filled && dv.type === 'line', smooth, innerSize);

  if (fill.filled && dv.type === 'line' && !fill.exact) {
    diagnostics.report(
      'APPROXIMATED_CHART_TYPE',
      'approximated',
      `${path}.fill`,
      `Fill mode ${JSON.stringify(fill.raw)} is drawn as an area filled to the axis.`,
      { seriesIndex: dv.index, details: { fill: fill.raw } },
    );
  }
  if (fill.filled && dv.type === 'scatter') {
    diagnostics.report('UNSUPPORTED_STYLE', 'unsupported', `${path}.fill`, 'Excel scatter charts cannot be filled.', {
      seriesIndex: dv.index,
    });
  }
  const stepped = kinds.dataset === 'line' ? datasetOpt(ctx, dv, 'stepped') : false;
  if (stepped !== undefined && stepped !== false) {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'approximated',
      `${path}.stepped`,
      'Excel has no stepped lines; straight segments are drawn.',
      { seriesIndex: dv.index },
    );
  }
  if (kinds.dataset === 'line' && points.some((p) => p.isNull)) {
    const spanGaps = pick(dv.ds.spanGaps, get(view.options, 'datasets', dv.type, 'spanGaps'), view.options.spanGaps);
    if (spanGaps !== undefined && spanGaps !== false) {
      diagnostics.report(
        'UNSUPPORTED_STYLE',
        'approximated',
        `${path}.spanGaps`,
        'spanGaps is not reproduced: Excel leaves a gap at empty values.',
        { severity: 'info', seriesIndex: dv.index },
      );
    }
  }
  if (dv.ds.datalabels !== undefined) {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'unsupported',
      `${path}.datalabels`,
      'chartjs-plugin-datalabels options are not exported (the plugin is not part of Chart.js).',
      { severity: 'info', seriesIndex: dv.index },
    );
  }

  // --- Styles --------------------------------------------------------------------------------------
  const styles = points.length > 0 ? points.map((_, j) => pointStyleAt(ctx, dv, j)) : [pointStyleAt(ctx, dv, 0)];
  const first = styles[0]!;

  let color: Color | null = null;
  let line: Stroke | null = null;
  let fillModel: Fill | null = null;
  let border: Stroke | null = null;
  let marker: MarkerStyle | null = null;

  if (kinds.dataset === 'line') {
    const lineColor = colorOf(datasetOpt(ctx, dv, 'borderColor'), `${path}.borderColor`, diagnostics, dv.index);
    const width = num(datasetOpt(ctx, dv, 'borderWidth')) ?? 3;
    const showLine = controllerOpt(ctx, dv, 'showLine', dv.type !== 'scatter') !== false;
    const dashRaw = datasetOpt(ctx, dv, 'borderDash');
    const dash = showLine ? dashFromBorderDash(dashRaw, width, `${path}.borderDash`, diagnostics, dv.index) : 'solid';
    line = { color: lineColor, width: showLine ? width : 0, dash };
    color = lineColor ?? first.bg;
    if (fill.filled && dv.type === 'line') {
      fillModel = solid(
        colorOf(datasetOpt(ctx, dv, 'backgroundColor'), `${path}.backgroundColor`, diagnostics, dv.index),
      );
    }
    marker = {
      enabled: first.radius > 0 && first.symbol !== 'none',
      symbol: first.symbol === 'none' ? 'circle' : first.symbol,
      radius: first.radius,
      fill: first.bg,
      stroke: first.border,
      strokeWidth: first.borderWidth,
    };
    if (dv.type === 'scatter') color = first.border ?? first.bg ?? lineColor;
  } else if (dv.type === 'bubble') {
    color = first.bg ?? first.border;
    fillModel = solid(first.bg);
    border = { color: first.border, width: first.borderWidth, dash: 'solid' };
    marker = {
      enabled: true,
      symbol: first.symbol === 'none' ? 'circle' : first.symbol,
      radius: first.radius,
      fill: first.bg,
      stroke: first.border,
      strokeWidth: first.borderWidth,
    };
  } else {
    // bar, pie, doughnut, polarArea
    color = first.bg;
    fillModel = solid(first.bg);
    border = { color: first.border, width: first.borderWidth, dash: 'solid' };
  }

  // Per-point overrides (only where they differ from the series; pie slices always carry their color).
  points.forEach((p, j) => {
    const s = styles[j];
    if (!s) return;
    if (isPie) {
      p.color = s.bg;
      if (!sameColor(s.border, first.border) || s.borderWidth !== first.borderWidth) {
        p.border = { color: s.border, width: s.borderWidth, dash: 'solid' };
      }
      if (s.offset > 0) p.sliced = s.offset;
      return;
    }
    if (kinds.data === 'point') {
      const m: Partial<MarkerStyle> = {};
      if (!sameColor(s.bg, first.bg)) m.fill = s.bg;
      if (!sameColor(s.border, first.border)) m.stroke = s.border;
      if (s.radius !== first.radius) {
        m.radius = s.radius;
        if (s.radius <= 0) m.enabled = false;
      }
      if (s.symbol !== first.symbol) m.symbol = s.symbol === 'none' ? 'circle' : s.symbol;
      if (Object.keys(m).length > 0) p.marker = m;
      if (m.fill !== undefined && dv.type !== 'line') p.color = s.bg;
      return;
    }
    if (!sameColor(s.bg, first.bg)) p.color = s.bg;
    if (!sameColor(s.border, first.border) || s.borderWidth !== first.borderWidth) {
      p.border = { color: s.border, width: s.borderWidth, dash: 'solid' };
    }
  });

  let bars: SeriesModel['bars'] = null;
  if (dv.type === 'bar') {
    const categoryPercentage = num(controllerOpt(ctx, dv, 'categoryPercentage', 0.8)) ?? 0.8;
    const barPercentage = num(controllerOpt(ctx, dv, 'barPercentage', 0.9)) ?? 0.9;
    const radius = pointOpt(ctx, dv, 'borderRadius', 0);
    const radiusPx = num(radius) ?? Math.max(0, ...Object.values(rec(radius) ?? {}).map((v) => num(v) ?? 0));
    bars = {
      pointPadding: Math.max(0, Math.round(((1 - barPercentage) / 2) * 1e6) / 1e6),
      groupPadding: Math.max(0, Math.round(((1 - categoryPercentage) / 2) * 1e6) / 1e6),
      borderRadius: radiusPx,
    };
    if (controllerOpt(ctx, dv, 'barThickness', undefined) !== undefined) {
      diagnostics.report(
        'APPROXIMATED_LAYOUT',
        'approximated',
        `${path}.barThickness`,
        'A fixed bar thickness is not reproduced; Excel sizes bars from the gap width.',
        { severity: 'info', seriesIndex: dv.index },
      );
    }
  }

  const stacking: Stacking = stacked && !isPie && dv.type !== 'scatter' && dv.type !== 'bubble' ? 'normal' : null;
  if (points.length === 0) {
    diagnostics.report('EMPTY_SERIES', 'translated', `${path}.data`, `Dataset "${dv.label}" has no data points.`, {
      severity: 'info',
      seriesIndex: dv.index,
    });
  }

  return {
    id: str(dv.ds.id) ?? `dataset-${dv.index}`,
    index: dv.index,
    name: dv.label,
    kind,
    sourceType: dv.type,
    visible: dv.visible,
    showInLegend: true,
    xAxisIndex: axes.xAxisIndex,
    yAxisIndex: axes.yAxisIndex,
    color,
    stacking,
    stackGroup: stacking && dv.ds.stack !== undefined && dv.ds.stack !== null ? String(dv.ds.stack) : null,
    line,
    fill: fillModel,
    fillOpacity: 1,
    border,
    marker,
    dataLabels: null,
    smooth,
    bars,
    pie,
    yFormat: axes.yFormat,
    points,
    dataSemantics: {
      mode: ctx.dataMode,
      grouped,
      cropped: false,
      sourcePointCount: sourceCount,
      renderedPointCount: renderedCount,
    },
  };
}
