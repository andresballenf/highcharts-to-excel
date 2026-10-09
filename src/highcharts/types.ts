/**
 * Minimal STRUCTURAL types for the Highcharts runtime objects the extractor reads.
 *
 * Every field is optional: the extractor must work on Highcharts 11-13 (and on whatever a
 * developer hands us), so nothing here is trusted. Option bags are typed `unknown` and read
 * through the defensive accessors in `./guards.ts`.
 */

import type { Color, Font } from '../types/chart-model';
import type { DataMode, SeriesVisibilityMode } from '../types/public-api';
import type { DiagnosticCollector } from '../types/diagnostics';
import type { CssVariableResolver } from '../utils/colors';
import type { DatetimeShifter } from './extract-time';

/** An SVG/HTML wrapper (`SVGElement` in Highcharts' renderer). */
export interface HcWrapperLike {
  element?: Element;
  textStr?: string;
  styles?: Record<string, unknown>;
}

export interface HcPointLike {
  x?: number | null;
  y?: number | null;
  z?: number | null;
  name?: string | null;
  index?: number;
  isNull?: boolean;
  color?: unknown;
  colorIndex?: number;
  sliced?: boolean;
  visible?: boolean;
  options?: unknown;
  graphic?: HcWrapperLike;
}

export interface HcExtremes {
  min?: number | null;
  max?: number | null;
  dataMin?: number | null;
  dataMax?: number | null;
  userMin?: number | null;
  userMax?: number | null;
}

export interface HcAxisLike {
  index?: number;
  coll?: string;
  isInternal?: boolean;
  /** Merged axis options (defaults + user options). */
  options?: unknown;
  userOptions?: unknown;
  categories?: unknown;
  /** Names collected for `type: 'category'` axes without explicit categories. */
  names?: unknown;
  type?: string;
  reversed?: boolean;
  opposite?: boolean;
  userMin?: number | null;
  userMax?: number | null;
  getExtremes?: () => HcExtremes;
}

export interface HcSeriesLike {
  index?: number;
  name?: string;
  type?: string;
  visible?: boolean;
  color?: unknown;
  colorIndex?: number;
  symbol?: unknown;
  /** Merged series options (plotOptions + user options). */
  options?: unknown;
  userOptions?: unknown;
  tooltipOptions?: unknown;
  points?: Array<HcPointLike | undefined | null>;
  data?: Array<HcPointLike | undefined | null>;
  xAxis?: HcAxisLike;
  yAxis?: HcAxisLike;
  hasGroupedData?: boolean;
  cropped?: boolean;
  /** Set by the boost module: `points` then holds pixel pseudo-points without values. */
  boosted?: boolean;
  /** Present on navigator series (Highcharts Stock). */
  baseSeries?: unknown;
  /** Pie geometry in px: [centerX, centerY, diameter, innerDiameter]. */
  center?: Array<number | null | undefined>;
  closestPointRangePx?: number | null;
  pointArrayMap?: string[];
  graph?: HcWrapperLike;
  area?: HcWrapperLike;
  getColumn?: (name: string, processed?: boolean) => ArrayLike<unknown> | undefined;
}

export interface HcLegendLike {
  allItems?: unknown[];
  display?: boolean;
}

export interface HcChartLike {
  /** Merged chart options (defaults + user options). */
  options: unknown;
  userOptions?: unknown;
  series: HcSeriesLike[];
  xAxis: HcAxisLike[];
  yAxis: HcAxisLike[];
  renderTo?: HTMLElement | { id?: string };
  container?: HTMLElement | { id?: string };
  chartWidth?: number;
  chartHeight?: number;
  plotLeft?: number;
  plotTop?: number;
  plotWidth?: number;
  plotHeight?: number;
  inverted?: boolean;
  polar?: boolean;
  styledMode?: boolean;
  hasCartesianSeries?: boolean;
  title?: HcWrapperLike;
  subtitle?: HcWrapperLike;
  chartBackground?: HcWrapperLike;
  plotBackground?: HcWrapperLike;
  legend?: HcLegendLike;
  renderer?: { style?: Record<string, unknown> };
  /** `Highcharts.Time` instance (`getTimezoneOffset`, `dateFormat`). */
  time?: unknown;
}

/** A plain (JSON-like) Highcharts options object as supplied by a developer. */
export type HcOptionsLike = Record<string, unknown>;

// ---------------------------------------------------------------------------
// Internal normalized "view" shared by the instance and the plain-options extractors.
// ---------------------------------------------------------------------------

export interface AxisView {
  which: 'x' | 'y';
  /** Position in the filtered (non-internal) axis array; this is the IR index. */
  index: number;
  /** Option path prefix used in diagnostics, e.g. `yAxis[1]`. */
  path: string;
  /** Effective options (merged with defaults for a live chart; as supplied otherwise). */
  opts: Record<string, unknown>;
  /** Options as the developer supplied them (used to tell user callbacks from built-in ones). */
  userOpts: Record<string, unknown>;
  rt?: HcAxisLike;
}

export interface SeriesView {
  /** Source index (`series.index` / position in `options.series`). */
  index: number;
  /** Option path prefix used in diagnostics, e.g. `series[2]`. */
  path: string;
  type: string;
  name: string;
  visible: boolean;
  opts: Record<string, unknown>;
  userOpts: Record<string, unknown>;
  rt?: HcSeriesLike;
  xAxis: number;
  yAxis: number;
  /** Palette index Highcharts assigned (or would assign). */
  colorIndex: number;
  /** Explicit series color (option or runtime `series.color`); undefined in styled mode. */
  explicitColor: unknown;
  /** Marker symbol Highcharts assigned (runtime `series.symbol` or emulated cycling). */
  symbol: unknown;
}

export interface ChartView {
  rt?: HcChartLike;
  opts: Record<string, unknown>;
  userOpts: Record<string, unknown>;
  styledMode: boolean;
  /** Computed styles may be read (real browser and a live chart). */
  browser: boolean;
  width: number;
  height: number;
  plotBox: { left: number; top: number; width: number; height: number } | null;
  inverted: boolean;
  polar: boolean;
  resolver: CssVariableResolver;
  palette: Color[];
  baseFont: Font;
  xAxes: AxisView[];
  yAxes: AxisView[];
  series: SeriesView[];
  cartesian: boolean;
}

export interface ExtractContext {
  view: ChartView;
  diagnostics: DiagnosticCollector;
  dataMode: DataMode;
  seriesVisibility: SeriesVisibilityMode;
  /** Moves datetime x values to the wall-clock time the chart displays. */
  time: DatetimeShifter;
}
