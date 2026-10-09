/**
 * The extractor's single table of Highcharts series types: the neutral kind each type maps to and
 * how Highcharts treats it (cartesian axes, per-point colors, per-point legend entries). Every
 * series-type decision in `src/highcharts` goes through the helpers below; do not add parallel sets.
 */

import type { SeriesKind } from '../types/chart-model';

export interface SeriesTypeInfo {
  /** Neutral kind (`pie` becomes `doughnut` when it has an inner size; see `seriesKind`). */
  kind: SeriesKind;
  /** Plotted on cartesian x/y axes. */
  cartesian: boolean;
  /** Colored by point: the series does not consume a palette slot and its runtime color is a placeholder. */
  byPoint: boolean;
  /** The legend lists the points rather than the series (`showInLegend` defaults to false). */
  legendPerPoint: boolean;
}

const cartesian = (kind: SeriesKind): SeriesTypeInfo => ({
  kind,
  cartesian: true,
  byPoint: false,
  legendPerPoint: false,
});
/** Pie-like: no axes, colored by point, legend per point. */
const pieLike = (kind: SeriesKind): SeriesTypeInfo => ({ kind, cartesian: false, byPoint: true, legendPerPoint: true });
/** Other types without cartesian axes (not exported as a chart: kind `unknown`). */
const nonCartesian: SeriesTypeInfo = { kind: 'unknown', cartesian: false, byPoint: false, legendPerPoint: false };

/**
 * Known Highcharts series types. Types missing from the table (areasplinerange, boxplot, ohlc…) are
 * cartesian series of kind `unknown`.
 */
export const SERIES_TYPE_TABLE: Readonly<Record<string, Readonly<SeriesTypeInfo>>> = Object.freeze({
  line: cartesian('line'),
  spline: cartesian('spline'),
  area: cartesian('area'),
  areaspline: cartesian('areaspline'),
  column: cartesian('column'),
  bar: cartesian('bar'),
  scatter: cartesian('scatter'),
  bubble: cartesian('bubble'),
  errorbar: cartesian('errorbar'),
  columnrange: cartesian('columnrange'),
  arearange: cartesian('arearange'),
  pie: pieLike('pie'),
  variablepie: pieLike('unknown'),
  funnel: pieLike('unknown'),
  pyramid: pieLike('unknown'),
  item: nonCartesian,
  sunburst: nonCartesian,
  treemap: nonCartesian,
  treegraph: nonCartesian,
  networkgraph: nonCartesian,
  packedbubble: nonCartesian,
  organization: nonCartesian,
  sankey: nonCartesian,
  dependencywheel: nonCartesian,
  venn: nonCartesian,
  wordcloud: nonCartesian,
  timeline: nonCartesian,
  solidgauge: nonCartesian,
});

/** Table entry of a Highcharts series type (own properties only), or undefined when unknown. */
export function seriesTypeInfo(type: string): Readonly<SeriesTypeInfo> | undefined {
  return Object.hasOwn(SERIES_TYPE_TABLE, type) ? SERIES_TYPE_TABLE[type] : undefined;
}

/** Neutral kind of a Highcharts series type, ignoring pie inner size (`unknown` for unknown types). */
export function kindOf(type: string): SeriesKind {
  return seriesTypeInfo(type)?.kind ?? 'unknown';
}

/** True unless the type is known to have no cartesian axes. */
export function isCartesian(type: string): boolean {
  return seriesTypeInfo(type)?.cartesian ?? true;
}

/** True for types colored by point (pie, variablepie, funnel, pyramid). */
export function isColoredByPoint(type: string): boolean {
  return seriesTypeInfo(type)?.byPoint ?? false;
}

/** True for types whose legend lists points instead of the series. */
export function showsPointsInLegend(type: string): boolean {
  return seriesTypeInfo(type)?.legendPerPoint ?? false;
}

/** Number of styled-mode colors (`chart.colorCount`, default 10): CSS classes cycle through them. */
export function styledColorCount(opts: Record<string, unknown>): number {
  const chart = opts.chart;
  const n = typeof chart === 'object' && chart !== null ? (chart as Record<string, unknown>).colorCount : undefined;
  return typeof n === 'number' && Number.isInteger(n) && n > 0 ? n : 10;
}

/** Pie or doughnut kind. */
export function isPieLike(kind: SeriesKind): boolean {
  return kind === 'pie' || kind === 'doughnut';
}

/** Vertical (column) or horizontal (bar) bar kind, including floating column ranges. */
export function isBarLike(kind: SeriesKind): boolean {
  return kind === 'column' || kind === 'bar' || kind === 'columnrange';
}

/** Kinds whose points carry `low`/`high` instead of a single y. */
export function isRangePointKind(kind: SeriesKind): boolean {
  return kind === 'errorbar' || kind === 'columnrange' || kind === 'arearange';
}

/** Default series colors Highcharts sets in plotOptions (they do not consume a palette slot). */
export const DEFAULT_TYPE_COLORS: Readonly<Record<string, string>> = Object.freeze({ errorbar: '#000000' });
