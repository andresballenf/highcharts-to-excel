/**
 * Axis extraction (x and y). Internal/navigator axes are filtered out when the view is built
 * (see `isInternalAxis`); `AxisView.index` is therefore the IR index.
 */

import type { AxisKind, AxisModel, Stroke, TextBlock } from '../types/chart-model';
import { translateFormatString } from '../translators/number-format-translator';
import { strokeFromOptions } from '../translators/style-translator';
import { parseColor } from '../utils/colors';
import { fontFor, borderStroke } from './extract-styles';
import { arr, num, plainText, rec, str } from './guards';
import type { AxisView, ChartView, ExtractContext, HcAxisLike } from './types';

/** Navigator / scrollbar / other internal axes (Highcharts Stock). */
export function isInternalAxis(axis: HcAxisLike | Record<string, unknown> | undefined, opts?: Record<string, unknown>): boolean {
  if (!axis) return true;
  const a = axis as HcAxisLike;
  const o = opts ?? rec(a.options) ?? {};
  if (a.isInternal === true || o.isInternal === true) return true;
  const cls = str(o.className);
  if (cls !== undefined && cls.includes('navigator')) return true;
  return a.coll === 'navigatorXAxis';
}

export function axisKind(view: AxisView): AxisKind {
  const t = str(view.opts.type) ?? str(view.rt?.type);
  if (t === 'datetime') return 'datetime';
  if (t === 'logarithmic') return 'logarithmic';
  if (t === 'category') return 'category';
  if (axisCategories(view) !== null) return 'category';
  return 'linear';
}

/** Explicit categories (or, for `type: 'category'`, the point names Highcharts collected). */
export function axisCategories(view: AxisView): string[] | null {
  const list = arr(view.rt?.categories) ?? arr(view.opts.categories) ?? (view.opts.type === 'category' ? arr(view.rt?.names) : undefined);
  if (!list || list.length === 0) return null;
  return list.map((c) => plainText(c) ?? (c === null || c === undefined ? '' : String(c)));
}

/** `{value:%b %e}` → `%b %e`; a bare strftime string is returned as is. */
function explicitDateFormat(format: unknown): string | null {
  if (typeof format !== 'string' || !format.includes('%')) return null;
  const m = /\{value:([^}]*%[^}]*)\}/.exec(format);
  if (m?.[1]) return m[1];
  return format.includes('{') ? null : format;
}

function axisTitle(view: ChartView, axis: AxisView): TextBlock | null {
  const t = rec(axis.opts.title) ?? {};
  // Highcharts' default y-axis title is "Values" (already merged into live chart options).
  const raw = t.text === undefined && axis.which === 'y' && !axis.rt ? 'Values' : t.text;
  const text = plainText(raw);
  if (text === null) return null;
  const align = t.align === 'low' ? 'left' : t.align === 'high' ? 'right' : 'center';
  return { text, font: fontFor(view, 'axisTitle', t.style), align, verticalAlign: 'middle' };
}

function lineStroke(view: ChartView, color: unknown, width: unknown, defaults: { color: string; width: number }, dashStyle?: unknown): Stroke {
  const w = num(width) ?? defaults.width;
  return strokeFromOptions({ color, width: w, dashStyle }, { color: parseColor(defaults.color, view.resolver), width: w, dash: 'solid' }, view.resolver);
}

function extractAxis(ctx: ExtractContext, axis: AxisView): AxisModel {
  const { view, diagnostics } = ctx;
  const o = axis.opts;
  const isX = axis.which === 'x';
  const kind = axisKind(axis);
  const labels = rec(o.labels) ?? {};
  const userLabels = rec(axis.userOpts.labels) ?? {};

  const translated = translateFormatString(labels.format, userLabels.formatter, {
    kind: 'axisLabel',
    axisType: kind,
    property: `${axis.path}.labels`,
  });
  diagnostics.addAll(translated.diagnostics);

  let min = num(o.min) ?? null;
  let max = num(o.max) ?? null;
  const userMin = num(axis.rt?.userMin);
  const userMax = num(axis.rt?.userMax);
  if ((min === null && userMin !== undefined) || (max === null && userMax !== undefined)) {
    // The user zoomed: export the window they are looking at.
    if (min === null && userMin !== undefined) min = userMin;
    if (max === null && userMax !== undefined) max = userMax;
    diagnostics.report('APPROXIMATED_AXIS_SCALE', 'approximated', `${axis.path}.min`, 'The chart is zoomed; the zoomed range is exported as fixed axis bounds.', {
      severity: 'info',
      details: { min, max },
    });
  }

  let dataMin: number | null = null;
  let dataMax: number | null = null;
  try {
    const ext = typeof axis.rt?.getExtremes === 'function' ? axis.rt.getExtremes() : undefined;
    dataMin = num(ext?.dataMin) ?? null;
    dataMax = num(ext?.dataMax) ?? null;
  } catch {
    // Extremes unavailable (axis not yet laid out); filled from points by the caller.
  }

  for (const key of ['plotBands', 'plotLines'] as const) {
    const list = arr(o[key]);
    if (list && list.length > 0) {
      diagnostics.report('UNSUPPORTED_PLOT_BAND', 'unsupported', `${axis.path}.${key}`, `Axis ${key} have no Excel chart equivalent and are omitted.`, {
        details: { count: list.length },
      });
    }
  }

  const gridWidth = num(o.gridLineWidth) ?? (isX ? 0 : 1);
  const minorEnabled = o.minorTicks === true || (o.minorTickInterval !== undefined && o.minorTickInterval !== null);
  const minorWidth = num(o.minorGridLineWidth) ?? 1;

  const reversed = typeof axis.rt?.reversed === 'boolean' ? axis.rt.reversed : o.reversed === true || (isX && view.inverted && o.reversed === undefined);
  const opposite = typeof axis.rt?.opposite === 'boolean' ? axis.rt.opposite : o.opposite === true;

  return {
    index: axis.index,
    id: str(o.id) ?? null,
    kind,
    categories: kind === 'category' ? axisCategories(axis) : null,
    title: axisTitle(view, axis),
    labels: {
      enabled: labels.enabled !== false,
      font: fontFor(view, 'axisLabels', labels.style),
      format: translated.format,
      rotation: num(labels.rotation) ?? 0,
    },
    min,
    max,
    dataMin,
    dataMax,
    tickInterval: num(o.tickInterval) ?? null,
    minorTickInterval: num(o.minorTickInterval) ?? null,
    reversed,
    opposite,
    visible: o.visible !== false,
    gridLines: gridWidth > 0 ? borderStroke(view, o.gridLineColor, gridWidth, { color: '#e6e6e6', width: 1 }, o.gridLineDashStyle) : null,
    minorGridLines:
      minorEnabled && minorWidth > 0
        ? borderStroke(view, o.minorGridLineColor, minorWidth, { color: '#f2f2f2', width: 1 }, o.minorGridLineDashStyle)
        : null,
    axisLine: lineStroke(view, o.lineColor, o.lineWidth, { color: '#333333', width: isX ? 1 : 0 }),
    tickMarks: lineStroke(view, o.tickColor, o.tickWidth, { color: '#333333', width: isX ? 1 : 0 }),
    crossing: num(o.crossing) ?? null,
    logBase: kind === 'logarithmic' ? 10 : null,
    dateFormat: kind === 'datetime' ? explicitDateFormat(labels.format) : null,
  };
}

/** Extracts the x or y axes of the view. */
export function extractAxes(ctx: ExtractContext, which: 'x' | 'y'): AxisModel[] {
  const list = which === 'x' ? ctx.view.xAxes : ctx.view.yAxes;
  return list.map((a) => extractAxis(ctx, a));
}
