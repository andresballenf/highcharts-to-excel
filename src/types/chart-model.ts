/**
 * Normalized, source-library-independent chart representation (the "Chart IR").
 *
 * Everything the Excel writer needs is expressed here in neutral vocabulary.
 * Highcharts-specific details stay in `src/highcharts/`; Excel/OOXML details
 * stay in `src/excel/`. The model is plain JSON-serializable data so it can be
 * snapshotted in tests and dumped for debugging.
 */

import type { Diagnostic } from './diagnostics';

/** Resolved RGBA color. Channels are 0-255, alpha is 0-1. */
export interface Color {
  r: number;
  g: number;
  b: number;
  a: number;
  /** Original string the color was parsed from, kept for diagnostics only. */
  source?: string;
}

export type Fill =
  | { type: 'none' }
  | { type: 'solid'; color: Color }
  | {
      type: 'gradient';
      stops: Array<{ offset: number; color: Color }>;
      /** Angle in degrees, 0 = left-to-right, 90 = top-to-bottom. */
      angle: number;
    };

export type DashStyle =
  | 'solid'
  | 'dash'
  | 'dot'
  | 'dashdot'
  | 'longdash'
  | 'longdashdot'
  | 'longdashdotdot'
  | 'shortdash'
  | 'shortdot'
  | 'shortdashdot'
  | 'shortdashdotdot';

export interface Stroke {
  color: Color | null;
  /** Width in CSS pixels. 0 means hidden. */
  width: number;
  dash: DashStyle;
}

export interface Font {
  family: string | null;
  /** Size in CSS pixels. */
  size: number | null;
  bold: boolean;
  italic: boolean;
  color: Color | null;
}

export type HorizontalAlign = 'left' | 'center' | 'right';
export type VerticalAlign = 'top' | 'middle' | 'bottom';

export interface TextBlock {
  text: string;
  font: Font;
  align: HorizontalAlign;
  verticalAlign: VerticalAlign;
}

/**
 * A number/date format in neutral terms. `excel` carries an Excel format code
 * that the translator already produced; `unsupported` records that the source
 * used a formatter the library cannot (and must not) execute.
 */
export type NumberFormat =
  | { kind: 'excel'; code: string; /** e.g. "{value}%" */ source?: string }
  | { kind: 'unsupported'; reason: string; source?: string };

export type MarkerSymbol = 'circle' | 'square' | 'diamond' | 'triangle' | 'triangle-down' | 'none' | 'other';

export interface MarkerStyle {
  enabled: boolean;
  symbol: MarkerSymbol;
  /** Radius in CSS pixels. */
  radius: number;
  fill: Color | null;
  stroke: Color | null;
  strokeWidth: number;
}

export type DataLabelPosition = 'auto' | 'center' | 'insideEnd' | 'insideBase' | 'outsideEnd' | 'above' | 'below' | 'left' | 'right' | 'bestFit';

export interface DataLabelStyle {
  enabled: boolean;
  font: Font | null;
  format: NumberFormat | null;
  position: DataLabelPosition;
  /** Pie-style content switches. For non-pie charts only `showValue` is meaningful. */
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  showPercentage: boolean;
  background: Fill | null;
  border: Stroke | null;
}

/** Neutral series kind. Unknown source types map to `unknown` and keep `sourceType`. */
export type SeriesKind =
  | 'line'
  | 'spline'
  | 'area'
  | 'areaspline'
  | 'column'
  | 'bar'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'bubble'
  | 'unknown';

export type Stacking = 'normal' | 'percent' | null;

export interface PointModel {
  /** Numeric x value (timestamp in ms for datetime axes, index for category axes). */
  x: number | null;
  /** Category or point name when the x axis is categorical or the point is named (pie slices). */
  name: string | null;
  y: number | null;
  /** Third dimension (bubble). */
  z: number | null;
  /** True when the source point is an explicit null (gap). Missing points are absent from the array. */
  isNull: boolean;
  /** Per-point overrides. Null means inherit from series. */
  color: Color | null;
  border: Stroke | null;
  marker: Partial<MarkerStyle> | null;
  /** Pie slice explode offset in pixels; null when not sliced. */
  sliced: number | null;
  /** Per-point data label override. */
  dataLabels: Partial<DataLabelStyle> | null;
  /** Point is selected/hidden in the source (pie legend toggles). */
  visible: boolean;
}

export interface SeriesModel {
  /** Stable id within the model (source id or generated). */
  id: string;
  /** Order index in the source. */
  index: number;
  name: string;
  kind: SeriesKind;
  /** Original source type string, e.g. "areaspline", "treemap". */
  sourceType: string;
  visible: boolean;
  xAxisIndex: number;
  yAxisIndex: number;
  color: Color | null;
  stacking: Stacking;
  /** Series in the same group stack together. Null = default group. */
  stackGroup: string | null;
  /** Line stroke for line/spline/area outlines and scatter connecting lines. */
  line: Stroke | null;
  /** Area / column / bar / pie fill. */
  fill: Fill | null;
  /** Opacity applied to `fill` (0-1). Area charts default to 0.75 in Highcharts. */
  fillOpacity: number;
  /** Column/bar/pie slice border. */
  border: Stroke | null;
  marker: MarkerStyle | null;
  dataLabels: DataLabelStyle | null;
  /** Smooth curve requested (spline family). */
  smooth: boolean;
  /** Column/bar geometry. Values are fractions as in Highcharts (0-1). */
  bars: { pointPadding: number; groupPadding: number; borderRadius: number } | null;
  /** Pie/doughnut geometry. innerSize is a fraction 0-1. Angles in degrees. */
  pie: { innerSize: number; startAngle: number; endAngle: number | null } | null;
  /** Number format applied to y values (tooltip valueDecimals / valueSuffix / dataLabels.format). */
  yFormat: NumberFormat | null;
  points: PointModel[];
  dataSemantics: SeriesDataSemantics;
}

export interface SeriesDataSemantics {
  /** Which data the points represent. */
  mode: 'rendered' | 'raw';
  /** Points are the result of data grouping (Highcharts Stock). */
  grouped: boolean;
  /** Points are cropped to the visible range (zoom/navigator). */
  cropped: boolean;
  /** Number of points in the raw source data. */
  sourcePointCount: number;
  /** Number of points after processing in the source library. */
  renderedPointCount: number;
}

export type AxisKind = 'category' | 'linear' | 'datetime' | 'logarithmic';

export interface AxisLabelStyle {
  enabled: boolean;
  font: Font | null;
  format: NumberFormat | null;
  /** Rotation in degrees. */
  rotation: number;
}

export interface AxisModel {
  index: number;
  id: string | null;
  kind: AxisKind;
  /** For category axes. Null for numeric/datetime axes. */
  categories: string[] | null;
  title: TextBlock | null;
  labels: AxisLabelStyle;
  /** Explicit min/max from options (null = auto). */
  min: number | null;
  max: number | null;
  /** Effective extremes currently displayed (always populated when extractable). */
  dataMin: number | null;
  dataMax: number | null;
  tickInterval: number | null;
  minorTickInterval: number | null;
  reversed: boolean;
  /** Rendered on the opposite side (right for y, top for x). */
  opposite: boolean;
  visible: boolean;
  gridLines: Stroke | null;
  minorGridLines: Stroke | null;
  axisLine: Stroke | null;
  tickMarks: Stroke | null;
  /** Where this axis crosses the other axis, if explicitly set. */
  crossing: number | null;
  /** Log base for logarithmic axes. */
  logBase: number | null;
  /** Explicit datetime label format (Highcharts dateTimeLabelFormats / labels.format). */
  dateFormat: string | null;
}

export type LegendPosition = 'top' | 'bottom' | 'left' | 'right' | 'topRight' | 'topLeft' | 'bottomRight' | 'bottomLeft';

export interface LegendModel {
  enabled: boolean;
  position: LegendPosition;
  layout: 'horizontal' | 'vertical';
  font: Font | null;
  background: Fill | null;
  border: Stroke | null;
  /** Floating legends overlay the plot area. */
  overlay: boolean;
  reversed: boolean;
}

export interface PlotAreaModel {
  background: Fill | null;
  border: Stroke | null;
  /** Plot box in CSS pixels relative to the chart box (as rendered). Null when not measurable. */
  box: { left: number; top: number; width: number; height: number } | null;
}

export interface ChartMeta {
  sourceLibrary: 'highcharts';
  sourceVersion: string | null;
  /** Default chart type from options (chart.type) or first series type. */
  sourceChartType: string;
  /** How styles were obtained. */
  extraction: 'browser' | 'headless';
  styledMode: boolean;
  /** The chart's id/renderTo id if any; used in diagnostics. */
  chartId: string | null;
}

export interface ChartModel {
  meta: ChartMeta;
  /** Chart box in CSS pixels. */
  width: number;
  height: number;
  /** Bar charts in Highcharts are inverted column charts. */
  inverted: boolean;
  polar: boolean;
  background: Fill | null;
  border: Stroke | null;
  plotArea: PlotAreaModel;
  title: TextBlock | null;
  subtitle: TextBlock | null;
  legend: LegendModel;
  xAxes: AxisModel[];
  yAxes: AxisModel[];
  series: SeriesModel[];
  /** Palette in effect (chart.options.colors), resolved. */
  colors: Color[];
  /** Diagnostics raised during extraction/normalization. Translators append theirs. */
  warnings: Diagnostic[];
}

/** Creates an empty-but-valid model skeleton; used by extractors and tests. */
export function createEmptyChartModel(meta: ChartMeta): ChartModel {
  return {
    meta,
    width: 600,
    height: 400,
    inverted: false,
    polar: false,
    background: null,
    border: null,
    plotArea: { background: null, border: null, box: null },
    title: null,
    subtitle: null,
    legend: {
      enabled: true,
      position: 'bottom',
      layout: 'horizontal',
      font: null,
      background: null,
      border: null,
      overlay: false,
      reversed: false,
    },
    xAxes: [],
    yAxes: [],
    series: [],
    colors: [],
    warnings: [],
  };
}
