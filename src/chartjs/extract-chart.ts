/**
 * Entry point: Chart.js chart instance (or plain configuration) → ChartModel IR.
 *
 * Both inputs are normalized into a `ChartJsView` (configuration + live runtime objects when
 * available) and share every parsing step after that. Nothing here mutates the chart or the
 * configuration.
 */

import {
  createEmptyChartModel,
  type AxisModel,
  type ChartMeta,
  type ChartModel,
  type SeriesModel,
} from '../types/chart-model';
import { DiagnosticCollector } from '../types/diagnostics';
import { ExportError } from '../types/public-api';
import { parseColor } from '../utils/colors';
import {
  CHARTJS_COLORS_PLUGIN_BACKGROUND,
  CHARTJS_COLORS_PLUGIN_BORDER,
  CHARTJS_DEFAULT_ELEMENT_COLOR,
  defaultAspectRatio,
} from './defaults';
import { extractAxis, scaleLabels, scalesFromChart, scalesFromConfig, type ExtractedAxis } from './extract-axes';
import { extractDataset, type SeriesAxes } from './extract-series';
import { baseFont, extractBackground, extractLegend, extractTitle } from './extract-styles';
import {
  arr,
  get,
  getChartJsDefaults,
  getChartJsVersion,
  isChartJsChart,
  isChartJsConfig,
  isRealBrowser,
  num,
  rec,
  str,
} from './guards';
import { WallClock } from './time';
import type {
  AxisLetter,
  ChartJsChartLike,
  ChartJsConfigLike,
  ChartJsExtractOptions,
  ChartJsView,
  DatasetView,
  ExtractContext,
  Rec,
} from './types';

const PIE_TYPES: ReadonlySet<string> = new Set(['pie', 'doughnut']);
const POLAR_TYPES: ReadonlySet<string> = new Set(['radar', 'polarArea']);

/** Chart.js `getIndexAxis`: `options.datasets[type].indexAxis`, then `options.indexAxis`, else 'x'. */
function indexAxisOf(type: string, options: Rec, ds?: Rec): AxisLetter {
  const v = ds?.indexAxis ?? get(options, 'datasets', type, 'indexAxis') ?? options.indexAxis;
  return v === 'y' ? 'y' : 'x';
}

function datasetLabel(ds: Rec, i: number): string {
  const label = ds.label;
  return typeof label === 'string' || typeof label === 'number' ? String(label) : `Dataset ${i + 1}`;
}

function chartSize(
  type: string,
  options: Rec,
  extract: ChartJsExtractOptions,
  live?: { width: unknown; height: unknown },
): { width: number; height: number } {
  const liveW = num(live?.width);
  const liveH = num(live?.height);
  const width = num(extract.chartWidth) ?? (liveW !== undefined && liveW > 0 ? liveW : 600);
  const aspect = num(options.aspectRatio) ?? defaultAspectRatio(type);
  const height =
    num(extract.chartHeight) ??
    (liveH !== undefined && liveH > 0 && extract.chartWidth === undefined
      ? liveH
      : Math.round(width / (aspect > 0 ? aspect : 2)));
  return { width, height };
}

/** Chart.js colors plugin: is it active for this configuration (it would color the datasets)? */
function colorsPluginActive(options: Rec, datasets: readonly Rec[]): { active: boolean; force: boolean } {
  const plugin = rec(get(options, 'plugins', 'colors')) ?? {};
  if (plugin.enabled === false) return { active: false, force: false };
  const force = plugin.forceOverride === true;
  const hasColor = (o: unknown): boolean => {
    const r = rec(o);
    return r !== undefined && (Boolean(r.borderColor) || Boolean(r.backgroundColor));
  };
  const defined =
    datasets.some(hasColor) || hasColor(options) || Object.values(rec(options.elements) ?? {}).some(hasColor);
  return { active: force || !defined, force };
}

/** Assigns the colors the Chart.js colors plugin would (copying, never mutating the datasets). */
function applyColorsPlugin(datasets: DatasetView[], force: boolean): void {
  const n = CHARTJS_COLORS_PLUGIN_BORDER.length;
  let i = 0;
  for (const dv of datasets) {
    const count = arr(dv.ds.data)?.length ?? Object.keys(rec(dv.ds.data) ?? {}).length;
    if (PIE_TYPES.has(dv.type)) {
      dv.autoColors = { backgroundColor: Array.from({ length: count }, () => CHARTJS_COLORS_PLUGIN_BORDER[i++ % n]) };
    } else if (dv.type === 'polarArea') {
      dv.autoColors = {
        backgroundColor: Array.from({ length: count }, () => CHARTJS_COLORS_PLUGIN_BACKGROUND[i++ % n]),
      };
    } else {
      dv.autoColors = {
        borderColor: CHARTJS_COLORS_PLUGIN_BORDER[i % n],
        backgroundColor: CHARTJS_COLORS_PLUGIN_BACKGROUND[i % n],
      };
      i++;
    }
    dv.autoColorsForce = force;
  }
}

function viewFromConfig(config: ChartJsConfigLike, extract: ChartJsExtractOptions): ChartJsView {
  const options = rec(config.options) ?? {};
  const data = rec(config.data) ?? {};
  const rawDatasets = (arr(data.datasets) ?? []).map((d) => rec(d) ?? {});
  const type = str(config.type) ?? str(rawDatasets[0]?.type) ?? 'line';
  const datasets: DatasetView[] = rawDatasets.map((ds, i) => {
    const dtype = str(ds.type) ?? type;
    return {
      index: i,
      path: `data.datasets[${i}]`,
      type: dtype,
      ds,
      label: datasetLabel(ds, i),
      visible: ds.hidden !== true,
      indexAxis: indexAxisOf(dtype, options, ds),
      indexScaleId: '',
      valueScaleId: '',
      autoColors: null,
      autoColorsForce: false,
    };
  });
  const colors = colorsPluginActive(options, rawDatasets);
  if (colors.active) applyColorsPlugin(datasets, colors.force);
  const scales = scalesFromConfig({ type, options }, datasets);
  return {
    type,
    data,
    labels: arr(data.labels) ?? [],
    datasets,
    options,
    scales,
    indexAxis: indexAxisOf(type, options),
    ...chartSize(type, options, extract),
    polar: POLAR_TYPES.has(type) || datasets.some((d) => POLAR_TYPES.has(d.type)),
  };
}

function viewFromChart(chart: ChartJsChartLike, extract: ChartJsExtractOptions): ChartJsView {
  const cfg = rec(chart.config) ?? {};
  const options = rec(cfg.options) ?? rec(chart.options) ?? {};
  const data = rec(cfg.data) ?? rec(chart.data) ?? {};
  const rawDatasets = (arr(data.datasets) ?? []).map((d) => rec(d) ?? {});
  const type = str(cfg.type) ?? str(rawDatasets[0]?.type) ?? 'line';
  const datasets: DatasetView[] = rawDatasets.map((ds, i) => {
    let meta: ReturnType<ChartJsChartLike['getDatasetMeta']> | undefined;
    try {
      meta = chart.getDatasetMeta(i);
    } catch {
      meta = undefined;
    }
    const dtype = str(meta?.type) ?? str(ds.type) ?? type;
    const indexAxis: AxisLetter =
      meta?.indexAxis === 'x' || meta?.indexAxis === 'y' ? meta.indexAxis : indexAxisOf(dtype, options, ds);
    let visible = ds.hidden !== true;
    try {
      visible = chart.isDatasetVisible(i);
    } catch {
      // keep the configured visibility
    }
    const ids = { x: str(meta?.xAxisID) ?? '', y: str(meta?.yAxisID) ?? '' };
    return {
      index: i,
      path: `data.datasets[${i}]`,
      type: dtype,
      ds,
      label: datasetLabel(ds, i),
      visible,
      indexAxis,
      indexScaleId: ids[indexAxis],
      valueScaleId: ids[indexAxis === 'x' ? 'y' : 'x'],
      autoColors: null,
      autoColorsForce: false,
      ...(meta ? { meta } : {}),
    };
  });
  const scales = scalesFromChart(chart);
  const firstOf = (axis: string): string => scales.find((s) => s.axis === axis)?.id ?? '';
  for (const dv of datasets) {
    if (dv.indexScaleId === '') dv.indexScaleId = firstOf(dv.indexAxis);
    if (dv.valueScaleId === '') dv.valueScaleId = firstOf(dv.indexAxis === 'x' ? 'y' : 'x');
  }
  const defaults = getChartJsDefaults(chart);
  return {
    type,
    data,
    labels: arr(data.labels) ?? [],
    datasets,
    options,
    scales,
    indexAxis: indexAxisOf(type, options),
    ...chartSize(type, options, extract, { width: chart.width, height: chart.height }),
    polar: POLAR_TYPES.has(type) || datasets.some((d) => POLAR_TYPES.has(d.type)),
    live: chart,
    ...(defaults ? { defaults } : {}),
  };
}

/** Fills missing data extremes from the extracted points of the visible series bound to the axis. */
function fillExtremesFromPoints(axes: AxisModel[], series: SeriesModel[], which: 'x' | 'y'): void {
  for (const axis of axes) {
    let min = Infinity;
    let max = -Infinity;
    for (const s of series) {
      if ((which === 'x' ? s.xAxisIndex : s.yAxisIndex) !== axis.index) continue;
      for (const p of s.points) {
        const v = which === 'x' ? (axis.kind === 'category' ? null : p.x) : p.y;
        if (v === null || !Number.isFinite(v)) continue;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
    if (min <= max) {
      axis.dataMin = min;
      axis.dataMax = max;
    }
  }
}

/** Chart.js `suggestedMin/Max` extend the range; `beginAtZero` pins zero into it. */
function applyBoundHints(axes: ExtractedAxis[]): void {
  for (const { model: a, hints } of axes) {
    if (a.kind === 'category') continue;
    if (a.min === null && a.dataMin !== null) {
      if (hints.suggestedMin !== null && hints.suggestedMin < a.dataMin) a.min = hints.suggestedMin;
      else if (hints.beginAtZero && a.kind !== 'logarithmic' && a.dataMin > 0) a.min = 0;
    }
    if (a.max === null && a.dataMax !== null) {
      if (hints.suggestedMax !== null && hints.suggestedMax > a.dataMax) a.max = hints.suggestedMax;
      else if (hints.beginAtZero && a.kind !== 'logarithmic' && a.dataMax < 0) a.max = 0;
    }
  }
}

function chartLevelDiagnostics(view: ChartJsView, diagnostics: DiagnosticCollector): void {
  const plugins = rec(view.options.plugins) ?? {};
  const tooltip = rec(plugins.tooltip) ?? {};
  const callbacks = rec(tooltip.callbacks) ?? {};
  const callbackKey = Object.keys(callbacks).find((k) => typeof callbacks[k] === 'function');
  const tooltipProperty =
    typeof tooltip.external === 'function'
      ? 'options.plugins.tooltip.external'
      : callbackKey !== undefined
        ? `options.plugins.tooltip.callbacks.${callbackKey}`
        : null;
  if (tooltipProperty !== null) {
    diagnostics.report(
      'UNSUPPORTED_TOOLTIP',
      'unsupported',
      tooltipProperty,
      'Tooltip callbacks are interactive and are not exported.',
      { severity: 'info' },
    );
  }
  const annotations = get(plugins, 'annotation', 'annotations');
  const count = arr(annotations)?.length ?? Object.keys(rec(annotations) ?? {}).length;
  if (count > 0) {
    diagnostics.report(
      'UNSUPPORTED_ANNOTATION',
      'unsupported',
      'options.plugins.annotation',
      'Annotations (chartjs-plugin-annotation) have no Excel chart equivalent and are omitted.',
      { details: { count } },
    );
  }
  if (plugins.datalabels !== undefined && plugins.datalabels !== false) {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'unsupported',
      'options.plugins.datalabels',
      'chartjs-plugin-datalabels options are not exported (the plugin is not part of Chart.js).',
      { severity: 'info' },
    );
  }
}

function buildModel(
  view: ChartJsView,
  extract: Required<Pick<ChartJsExtractOptions, 'dataMode' | 'seriesVisibility'>> & {
    diagnostics: DiagnosticCollector;
  },
): ChartModel {
  const { diagnostics } = extract;
  const ctx: ExtractContext = {
    view,
    diagnostics,
    dataMode: extract.dataMode,
    seriesVisibility: extract.seriesVisibility,
  };
  const time = new WallClock();
  const base = baseFont(view);
  const background = extractBackground(view, diagnostics);
  const title = extractTitle(view, 'title', base, diagnostics);
  const subtitle = extractTitle(view, 'subtitle', base, diagnostics);
  const legend = extractLegend(view, base, diagnostics);

  const cartesian =
    view.datasets.length === 0 || view.datasets.some((d) => !PIE_TYPES.has(d.type) && !POLAR_TYPES.has(d.type));
  const valueAxis: AxisLetter = view.indexAxis === 'x' ? 'y' : 'x';
  const indexScales = cartesian && !view.polar ? view.scales.filter((s) => s.axis === view.indexAxis) : [];
  const valueScales = cartesian && !view.polar ? view.scales.filter((s) => s.axis === valueAxis) : [];
  const xAxes = indexScales.map((s, i) => extractAxis(view, s, i, base, time, diagnostics));
  const yAxes = valueScales.map((s, i) => extractAxis(view, s, i, base, time, diagnostics));

  let series: SeriesModel[] = [];
  for (const dv of view.datasets) {
    if (!dv.visible) {
      if (extract.seriesVisibility === 'visible') {
        diagnostics.report(
          'HIDDEN_SERIES_EXCLUDED',
          'translated',
          `${dv.path}.hidden`,
          `Hidden dataset "${dv.label}" is not exported.`,
          { severity: 'info', seriesIndex: dv.index },
        );
        continue;
      }
      diagnostics.report(
        'HIDDEN_SERIES_INCLUDED',
        'approximated',
        `${dv.path}.hidden`,
        `Hidden dataset "${dv.label}" is exported and will be visible in Excel.`,
        { severity: 'info', seriesIndex: dv.index },
      );
    }
    const xi = Math.max(
      0,
      indexScales.findIndex((s) => s.id === dv.indexScaleId),
    );
    const yi = Math.max(
      0,
      valueScales.findIndex((s) => s.id === dv.valueScaleId),
    );
    const xAxis = xAxes[xi];
    const yAxis = yAxes[yi];
    const scaleOfIndex = view.scales.find((s) => s.id === dv.indexScaleId);
    const axes: SeriesAxes = {
      xAxisIndex: xi,
      yAxisIndex: yi,
      indexKind: xAxis?.model.kind ?? (dv.type === 'scatter' || dv.type === 'bubble' ? 'linear' : 'category'),
      indexScale: scaleOfIndex,
      valueScale: view.scales.find((s) => s.id === dv.valueScaleId),
      categories: scaleOfIndex ? scaleLabels(view, scaleOfIndex) : view.labels,
      yFormat: yAxis?.model.labels.format?.kind === 'excel' ? yAxis.model.labels.format : null,
    };
    series.push(extractDataset(ctx, dv, axes, time));
  }

  // Chart.js draws the first doughnut dataset as the OUTER ring; Excel draws the first series inside.
  const rings = series.filter((s) => s.kind === 'doughnut');
  if (rings.length > 1 && rings.length === series.length) {
    series = [...series].reverse();
    diagnostics.report(
      'APPROXIMATED_LAYOUT',
      'approximated',
      'data.datasets',
      'Doughnut rings are written in reverse order so the first dataset stays the outer ring in Excel.',
      { severity: 'info', details: { rings: rings.length } },
    );
  }

  fillExtremesFromPoints(
    xAxes.map((a) => a.model),
    series,
    'x',
  );
  fillExtremesFromPoints(
    yAxes.map((a) => a.model),
    series,
    'y',
  );
  applyBoundHints([...xAxes, ...yAxes]);
  chartLevelDiagnostics(view, diagnostics);

  if (time.shifted) {
    const scale = xAxes.find((a) => a.model.kind === 'datetime')?.scale;
    diagnostics.report(
      'APPROXIMATED_DATETIME',
      'approximated',
      scale ? `${scale.path}.time` : 'options.scales',
      "Excel has no time zones: datetime values are exported as the wall-clock times the chart shows in the browser's time zone.",
      { severity: 'info', details: { offsetMinutes: time.offsetMinutes() } },
    );
  }

  if (series.length === 0 || series.every((s) => s.points.length === 0)) {
    diagnostics.report(
      'EMPTY_CHART',
      'blocking',
      'data.datasets',
      view.datasets.length === 0 ? 'The chart has no datasets.' : 'The chart has no data to export.',
      { details: { sourceDatasets: view.datasets.length, exportedSeries: series.length } },
    );
  }

  const canvasId = str(get(view.live?.canvas, 'id'));
  const meta: ChartMeta = {
    sourceLibrary: 'chartjs',
    sourceVersion: view.live ? getChartJsVersion(view.live) : null,
    datetimeOffsetMinutes: time.offsetMinutes(),
    sourceChartType: view.type,
    extraction: view.live && isRealBrowser() ? 'browser' : 'headless',
    styledMode: false,
    chartId: canvasId ?? null,
  };
  const model = createEmptyChartModel(meta);
  model.width = view.width;
  model.height = view.height;
  model.inverted = cartesian && view.indexAxis === 'y';
  model.polar = view.polar;
  model.background = background;
  const area = view.live?.chartArea;
  const [left, top, width, height] = [num(area?.left), num(area?.top), num(area?.width), num(area?.height)];
  model.plotArea = {
    background: null,
    border: null,
    box:
      left !== undefined && top !== undefined && width !== undefined && height !== undefined
        ? { left, top, width, height }
        : null,
  };
  model.title = title;
  model.subtitle = subtitle;
  model.legend = { ...legend, enabled: legend.enabled && series.some((s) => s.showInLegend) };
  model.xAxes = xAxes.map((a) => a.model);
  model.yAxes = yAxes.map((a) => a.model);
  model.series = series;
  const pluginOn = get(view.options, 'plugins', 'colors', 'enabled') !== false;
  model.colors = (pluginOn ? CHARTJS_COLORS_PLUGIN_BORDER : [CHARTJS_DEFAULT_ELEMENT_COLOR])
    .map((c) => parseColor(c))
    .filter((c) => c !== null);
  model.warnings = [...diagnostics.items];
  return model;
}

/**
 * Extracts the IR from a Chart.js 4 chart instance or a plain Chart.js configuration
 * (`{ type, data, options }`). Never mutates the chart or the configuration.
 *
 * @throws ExportError INVALID_CHART when the input is neither.
 * @experimental The IR and this entry point may change in minor versions.
 */
export function extractChartJsModel(chartOrConfig: unknown, options: ChartJsExtractOptions = {}): ChartModel {
  const diagnostics = options.diagnostics ?? new DiagnosticCollector();
  let view: ChartJsView;
  if (isChartJsChart(chartOrConfig)) view = viewFromChart(chartOrConfig, options);
  else if (isChartJsConfig(chartOrConfig)) view = viewFromConfig(chartOrConfig, options);
  else {
    throw new ExportError(
      'INVALID_CHART',
      'Expected a Chart.js chart instance or a Chart.js configuration object ({ type, data: { datasets }, options }).',
    );
  }
  return buildModel(view, {
    dataMode: options.dataMode ?? 'rendered',
    seriesVisibility: options.seriesVisibility ?? 'visible',
    diagnostics,
  });
}
