/**
 * Structural types for the Chart.js adapter. Nothing here imports `chart.js`: the adapter reads a
 * live chart or a plain configuration through these minimal shapes.
 */

import type { DiagnosticCollector } from '../types/diagnostics';
import type { DataMode, SeriesVisibilityMode } from '../types/public-api';

export type Rec = Record<string, unknown>;

/** A Chart.js configuration: `new Chart(canvas, config)`'s second argument. */
export interface ChartJsConfigLike {
  type?: string;
  data: { labels?: unknown[]; datasets: Rec[]; xLabels?: unknown[]; yLabels?: unknown[] } & Rec;
  options?: Rec;
  plugins?: unknown[];
}

/** The parts of a Chart.js `Element` (point, bar, arc, line) the adapter reads. */
export interface ChartJsElementLike {
  options?: Rec;
}

/** `chart.getDatasetMeta(i)`. */
export interface ChartJsMetaLike {
  type?: string;
  hidden?: boolean | null;
  xAxisID?: string;
  yAxisID?: string;
  indexAxis?: string;
  data?: ChartJsElementLike[];
  dataset?: ChartJsElementLike | null;
  /** Values parsed by Chart.js (category index, adapter-parsed time, parsing keys applied). */
  _parsed?: unknown[];
  controller?: { innerRadius?: unknown; outerRadius?: unknown } & Rec;
}

/** A live Chart.js scale (`chart.scales[id]`). */
export interface ChartJsScaleLike {
  id?: string;
  type?: string;
  axis?: string;
  position?: string;
  options?: Rec;
  min?: number;
  max?: number;
  getLabels?: () => unknown[];
}

/** A live Chart.js 4 `Chart` instance. */
export interface ChartJsChartLike {
  config: { type?: string; data?: Rec; options?: Rec } & Rec;
  data?: Rec;
  options?: Rec;
  scales: Record<string, ChartJsScaleLike>;
  width?: number;
  height?: number;
  canvas?: unknown;
  chartArea?: { left?: number; top?: number; width?: number; height?: number } & Rec;
  getDatasetMeta(index: number): ChartJsMetaLike;
  isDatasetVisible(index: number): boolean;
  getDataVisibility?(index: number): boolean;
}

/** Options of `extractChartJsModel`. Every field is optional. */
export interface ChartJsExtractOptions {
  /** 'rendered' (default) exports decimated data when the decimation plugin ran; 'raw' the source data. */
  dataMode?: DataMode;
  /** 'visible' (default) skips hidden datasets; 'all' exports them too. */
  seriesVisibility?: SeriesVisibilityMode;
  /** Receives the diagnostics; a fresh collector is used when omitted. */
  diagnostics?: DiagnosticCollector;
  chartWidth?: number;
  chartHeight?: number;
}

export type AxisLetter = 'x' | 'y';

/** A Chart.js scale, normalized. */
export interface ScaleView {
  id: string;
  axis: AxisLetter | 'r';
  /** Chart.js scale type: 'category', 'linear', 'logarithmic', 'time', 'timeseries', 'radialLinear'. */
  type: string;
  /** The scale options (user options, merged by Chart.js for a live chart). */
  opts: Rec;
  /** Option path used in diagnostics, e.g. `options.scales.y1`. */
  path: string;
  live?: ChartJsScaleLike;
}

/** A dataset, normalized: its effective type, axes and source objects. */
export interface DatasetView {
  index: number;
  /** Option path used in diagnostics, e.g. `data.datasets[1]`. */
  path: string;
  /** Effective type: `dataset.type ?? config.type`. */
  type: string;
  /** The dataset object (never mutated; colors from the colors plugin live in `autoColors`). */
  ds: Rec;
  label: string;
  visible: boolean;
  /** Category/index axis letter ('y' for horizontal bars). */
  indexAxis: AxisLetter;
  indexScaleId: string;
  valueScaleId: string;
  /** Colors the Chart.js colors plugin would assign (config input only). */
  autoColors: { backgroundColor?: unknown; borderColor?: unknown } | null;
  /** True when the colors plugin runs with `forceOverride` (its colors beat the dataset's). */
  autoColorsForce: boolean;
  meta?: ChartJsMetaLike;
}

/** Everything the extractors need, from either input kind. */
export interface ChartJsView {
  type: string;
  /** `config.data` (labels, xLabels, yLabels, datasets). */
  data: Rec;
  labels: unknown[];
  datasets: DatasetView[];
  options: Rec;
  scales: ScaleView[];
  /** `options.indexAxis` ('x' unless horizontal). */
  indexAxis: AxisLetter;
  width: number;
  height: number;
  polar: boolean;
  /** Present for a live chart. */
  live?: ChartJsChartLike;
  /** `Chart.defaults` when reachable from a live chart. */
  defaults?: Rec;
}

export interface ExtractContext {
  view: ChartJsView;
  diagnostics: DiagnosticCollector;
  dataMode: DataMode;
  seriesVisibility: SeriesVisibilityMode;
}
