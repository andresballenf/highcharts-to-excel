/**
 * Entry points: Highcharts chart instance (or plain options object) → ChartModel IR.
 *
 * Both entry points normalize their input into a `ChartView` (series/axis options plus the live
 * runtime objects when available) and share every parsing step after that.
 */

import {
  createEmptyChartModel,
  type AxisModel,
  type ChartMeta,
  type ChartModel,
  type SeriesModel,
} from '../types/chart-model';
import type { DiagnosticCollector } from '../types/diagnostics';
import { ExportError, type DataMode, type SeriesVisibilityMode } from '../types/public-api';
import { createCssVariableResolver } from './css-resolver';
import { extractAxes, isInternalAxis } from './extract-axes';
import { computeBaseFont, extractChartStyles, extractPalette } from './extract-styles';
import { extractSeries } from './extract-series';
import { DatetimeShifter, timeZoneFromChart, timeZoneFromOptions } from './extract-time';
import {
  arr,
  deepMerge,
  get,
  getHighchartsVersion,
  isHighchartsChart,
  isRealBrowser,
  num,
  rec,
  str,
  type Rec,
} from './guards';
import type { AxisView, ChartView, ExtractContext, HcAxisLike, HcChartLike, HcSeriesLike, SeriesView } from './types';

export interface ExtractOptions {
  dataMode: DataMode;
  seriesVisibility: SeriesVisibilityMode;
  diagnostics: DiagnosticCollector;
  chartWidth?: number;
  chartHeight?: number;
  /** Highcharts version; when omitted it is read from a global `Highcharts` if present. */
  highchartsVersion?: string | null;
}

/** Series types without cartesian axes. */
const NON_CARTESIAN_TYPES: ReadonlySet<string> = new Set([
  'pie',
  'variablepie',
  'funnel',
  'pyramid',
  'item',
  'sunburst',
  'treemap',
  'treegraph',
  'networkgraph',
  'packedbubble',
  'organization',
  'sankey',
  'dependencywheel',
  'venn',
  'wordcloud',
  'timeline',
  'solidgauge',
]);

/** Highcharts' default marker symbol cycle (`chart.options.symbols`). */
const DEFAULT_SYMBOLS = ['circle', 'diamond', 'square', 'triangle', 'triangle-down'];
/** Series types that consume a symbol from the cycle (`getSymbol` is a no-op for the others). */
const SYMBOL_TYPES: ReadonlySet<string> = new Set(['line', 'spline', 'area', 'areaspline', 'scatter']);
/** Series types colored by point (they do not consume a series color). */
const BY_POINT_TYPES: ReadonlySet<string> = new Set(['pie', 'variablepie', 'funnel', 'pyramid']);

function isInternalSeries(s: HcSeriesLike | Rec | undefined): boolean {
  if (!s) return true;
  const rt = s as HcSeriesLike;
  const o = rec(rt.options) ?? (s as Rec);
  if (o.isInternal === true) return true;
  const cls = str(o.className);
  if (cls?.includes('navigator')) return true;
  return rt.baseSeries !== undefined && rt.baseSeries !== null;
}

function axisPath(which: 'x' | 'y', i: number): string {
  return `${which}Axis[${i}]`;
}

function buildPlaceholderView(opts: Rec, userOpts: Rec, rt: HcChartLike | undefined): ChartView {
  const resolver = createCssVariableResolver(rt);
  return {
    ...(rt ? { rt } : {}),
    opts,
    userOpts,
    styledMode: rt?.styledMode === true || get(opts, 'chart', 'styledMode') === true,
    browser: rt !== undefined && isRealBrowser(),
    width: 600,
    height: 400,
    plotBox: null,
    inverted: false,
    polar: false,
    resolver,
    palette: extractPalette(opts, resolver),
    baseFont: computeBaseFont(rt, opts, resolver),
    xAxes: [],
    yAxes: [],
    series: [],
    cartesian: true,
  };
}

// ---------------------------------------------------------------------------
// Live chart → view
// ---------------------------------------------------------------------------

function viewFromChart(chart: HcChartLike, extract: ExtractOptions): ChartView {
  const opts = rec(chart.options) ?? {};
  const userOpts = rec(chart.userOptions) ?? {};
  const view = buildPlaceholderView(opts, userOpts, chart);

  view.width = num(extract.chartWidth) ?? num(chart.chartWidth) ?? 600;
  view.height = num(extract.chartHeight) ?? num(chart.chartHeight) ?? 400;
  const [left, top, width, height] = [
    num(chart.plotLeft),
    num(chart.plotTop),
    num(chart.plotWidth),
    num(chart.plotHeight),
  ];
  view.plotBox =
    left !== undefined && top !== undefined && width !== undefined && height !== undefined
      ? { left, top, width, height }
      : null;
  view.inverted = chart.inverted === true;
  view.polar = chart.polar === true;

  const series = chart.series.filter((s) => !isInternalSeries(s));
  view.cartesian =
    typeof chart.hasCartesianSeries === 'boolean'
      ? chart.hasCartesianSeries
      : series.some((s) => !NON_CARTESIAN_TYPES.has(str(s.type) ?? 'line'));

  const axisViews = (which: 'x' | 'y', list: HcAxisLike[]): AxisView[] =>
    view.cartesian
      ? list
          .filter((a) => !isInternalAxis(a))
          .map((a, i) => ({
            which,
            index: i,
            path: axisPath(which, num(a.index) ?? i),
            opts: rec(a.options) ?? {},
            userOpts: rec(a.userOptions) ?? {},
            rt: a,
          }))
      : [];
  view.xAxes = axisViews('x', chart.xAxis);
  view.yAxes = axisViews('y', chart.yAxis);

  const chartType = str(get(userOpts, 'chart', 'type')) ?? str(get(opts, 'chart', 'type')) ?? 'line';
  view.series = series.map((s, position): SeriesView => {
    const index = num(s.index) ?? position;
    const type = str(s.type) ?? chartType;
    const sOpts = rec(s.options) ?? {};
    const xIdx = view.xAxes.findIndex((a) => a.rt === s.xAxis);
    const yIdx = view.yAxes.findIndex((a) => a.rt === s.yAxis);
    return {
      index,
      path: `series[${index}]`,
      type,
      name: str(s.name) ?? str(sOpts.name) ?? `Series ${index + 1}`,
      visible: s.visible !== false,
      opts: sOpts,
      userOpts: deepMerge(get(userOpts, 'plotOptions', 'series'), get(userOpts, 'plotOptions', type), s.userOptions),
      rt: s,
      xAxis: Math.max(0, xIdx),
      yAxis: Math.max(0, yIdx),
      colorIndex: num(s.colorIndex) ?? num(sOpts.colorIndex) ?? index,
      // Pie-like series keep a placeholder runtime color (#cccccc for empty pies); only an explicit one counts.
      explicitColor: view.styledMode
        ? undefined
        : BY_POINT_TYPES.has(type)
          ? get(s.userOptions, 'color')
          : (s.color ?? sOpts.color),
      symbol: s.symbol,
    };
  });
  return view;
}

// ---------------------------------------------------------------------------
// Plain options → view
// ---------------------------------------------------------------------------

function axisList(x: unknown): Rec[] {
  const list = arr(x);
  if (list) return list.map((a) => rec(a) ?? {});
  return [rec(x) ?? {}];
}

function resolveAxisRef(ref: unknown, axes: AxisView[]): number {
  if (typeof ref === 'number' && Number.isInteger(ref)) {
    const i = axes.findIndex((a) => a.path.endsWith(`[${ref}]`));
    return Math.max(0, i);
  }
  if (typeof ref === 'string') {
    const i = axes.findIndex((a) => a.opts.id === ref);
    return Math.max(0, i);
  }
  return 0;
}

function viewFromOptions(options: Rec, extract: ExtractOptions): ChartView {
  const view = buildPlaceholderView(options, options, undefined);
  const c = rec(options.chart) ?? {};
  const width = num(extract.chartWidth) ?? num(c.width) ?? 600;
  view.width = width;
  // chart.height may be a percentage of the width.
  const h = c.height;
  const pct = typeof h === 'string' ? /^\s*(\d*\.?\d+)\s*%\s*$/.exec(h) : null;
  view.height =
    num(extract.chartHeight) ??
    num(h) ??
    (pct ? (Number(pct[1]) / 100) * width : undefined) ??
    (typeof h === 'string' ? num(parseFloat(h)) : undefined) ??
    400;
  view.polar = c.polar === true;

  const chartType = str(c.type) ?? 'line';
  const rawSeries = (arr(options.series) ?? []).map((s) => rec(s) ?? {}).filter((s) => !isInternalSeries(s));
  const types = rawSeries.map((s) => str(s.type) ?? chartType);
  view.inverted = c.inverted === true || types.includes('bar');
  view.cartesian = rawSeries.length === 0 || types.some((t) => !NON_CARTESIAN_TYPES.has(t));

  const axisViews = (which: 'x' | 'y', list: Rec[]): AxisView[] =>
    view.cartesian
      ? list
          .map((o, i) => ({ o, i }))
          .filter(({ o }) => !isInternalAxis(o, o))
          .map(({ o, i }, position) => ({ which, index: position, path: axisPath(which, i), opts: o, userOpts: o }))
      : [];
  view.xAxes = axisViews('x', axisList(options.xAxis));
  view.yAxes = axisViews('y', axisList(options.yAxis));

  const plotOptions = rec(options.plotOptions) ?? {};
  let colorCounter = 0;
  let symbolCounter = 0;
  view.series = rawSeries.map((raw, index): SeriesView => {
    const type = types[index] ?? chartType;
    const merged = deepMerge(plotOptions.series, plotOptions[type], raw);
    let colorIndex = num(merged.colorIndex);
    if (colorIndex === undefined) {
      colorIndex = colorCounter;
      // Highcharts' getCyclic: an explicit color does not consume a palette slot.
      if (merged.color === undefined && !BY_POINT_TYPES.has(type)) colorCounter++;
    }
    let symbol: unknown = get(merged, 'marker', 'symbol');
    if (symbol === undefined && SYMBOL_TYPES.has(type)) {
      symbol = DEFAULT_SYMBOLS[symbolCounter % DEFAULT_SYMBOLS.length];
      symbolCounter++;
    }
    return {
      index,
      path: `series[${index}]`,
      type,
      name: str(merged.name) ?? `Series ${index + 1}`,
      visible: merged.visible !== false,
      opts: merged,
      userOpts: merged,
      xAxis: resolveAxisRef(merged.xAxis, view.xAxes),
      yAxis: resolveAxisRef(merged.yAxis, view.yAxes),
      colorIndex,
      explicitColor: view.styledMode ? undefined : merged.color,
      symbol,
    };
  });
  return view;
}

// ---------------------------------------------------------------------------
// Shared model assembly
// ---------------------------------------------------------------------------

/** Fills missing data extremes from the extracted points of the visible series bound to the axis. */
function fillExtremesFromPoints(axes: AxisModel[], series: SeriesModel[], which: 'x' | 'y'): void {
  for (const axis of axes) {
    if (axis.dataMin !== null && axis.dataMax !== null) continue;
    let min = Infinity;
    let max = -Infinity;
    for (const s of series) {
      if (!s.visible || (which === 'x' ? s.xAxisIndex : s.yAxisIndex) !== axis.index) continue;
      for (const p of s.points) {
        const v = which === 'x' ? p.x : p.y;
        if (v === null) continue;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
    if (min <= max) {
      axis.dataMin ??= min;
      axis.dataMax ??= max;
    }
  }
}

function hasUserFunction(x: unknown): boolean {
  return typeof x === 'function';
}

function chartLevelDiagnostics(view: ChartView, series: SeriesView[], diagnostics: DiagnosticCollector): void {
  if (get(view.opts, 'chart', 'options3d', 'enabled') === true) {
    diagnostics.report(
      'UNSUPPORTED_3D',
      'approximated',
      'chart.options3d',
      '3D charts are exported as flat 2D charts.',
    );
  }

  // Only developer-supplied callbacks count (Highcharts' built-in defaults live in merged options).
  const tooltip = rec(view.userOpts.tooltip) ?? {};
  let tooltipProperty: string | null = null;
  if (hasUserFunction(tooltip.formatter)) tooltipProperty = 'tooltip.formatter';
  else if (hasUserFunction(tooltip.pointFormatter)) tooltipProperty = 'tooltip.pointFormatter';
  else {
    const s = series.find(
      (sv) =>
        hasUserFunction(get(sv.userOpts, 'tooltip', 'pointFormatter')) ||
        hasUserFunction(get(sv.userOpts, 'tooltip', 'formatter')),
    );
    if (s) tooltipProperty = `${s.path}.tooltip.pointFormatter`;
  }
  if (tooltipProperty !== null) {
    diagnostics.report(
      'UNSUPPORTED_TOOLTIP',
      'unsupported',
      tooltipProperty,
      'Tooltip formatters are interactive and are not exported.',
      { severity: 'info' },
    );
  }

  const annotations = view.opts.annotations;
  const count = arr(annotations)?.length ?? (rec(annotations) ? 1 : 0);
  if (count > 0) {
    diagnostics.report(
      'UNSUPPORTED_ANNOTATION',
      'unsupported',
      'annotations',
      'Annotations have no Excel chart equivalent and are omitted.',
      {
        details: { count },
      },
    );
  }
}

/** Reports how datetime x values were moved to the chart's displayed time zone. */
function reportTimeZone(time: DatetimeShifter, diagnostics: DiagnosticCollector): void {
  const offsetMinutes = time.offsetMinutes();
  if (time.unresolved) {
    diagnostics.report(
      'APPROXIMATED_DATETIME',
      'approximated',
      'time.timezone',
      `The chart time zone "${time.zone.label}" could not be resolved; datetime values are exported as UTC.`,
      { details: { timezone: time.zone.label } },
    );
  } else if (time.shifted) {
    diagnostics.report(
      'APPROXIMATED_DATETIME',
      'approximated',
      'time.timezone',
      `Excel has no time zones: datetime values are exported as the wall-clock times the chart shows in "${time.zone.label}".`,
      { severity: 'info', details: { timezone: time.zone.label, offsetMinutes } },
    );
  }
}

function buildModel(
  view: ChartView,
  baseMeta: Omit<ChartMeta, 'datetimeOffsetMinutes'>,
  extract: ExtractOptions,
): ChartModel {
  const { diagnostics } = extract;
  const time = new DatetimeShifter(view.rt ? timeZoneFromChart(view.rt) : timeZoneFromOptions(view.opts));
  const ctx: ExtractContext = {
    view,
    diagnostics,
    dataMode: extract.dataMode,
    seriesVisibility: extract.seriesVisibility,
    time,
  };

  const styles = extractChartStyles(view, diagnostics);
  const xAxes = extractAxes(ctx, 'x');
  const yAxes = extractAxes(ctx, 'y');
  const series = extractSeries(ctx);
  fillExtremesFromPoints(xAxes, series, 'x');
  fillExtremesFromPoints(yAxes, series, 'y');
  chartLevelDiagnostics(view, view.series, diagnostics);
  reportTimeZone(time, diagnostics);
  const meta: ChartMeta = { ...baseMeta, datetimeOffsetMinutes: time.offsetMinutes() };

  if (series.length === 0 || series.every((s) => s.points.length === 0)) {
    diagnostics.report(
      'EMPTY_CHART',
      'blocking',
      'series',
      view.series.length === 0 ? 'The chart has no series.' : 'The chart has no data to export.',
      {
        details: { sourceSeries: view.series.length, exportedSeries: series.length },
      },
    );
  }

  const model = createEmptyChartModel(meta);
  model.width = view.width;
  model.height = view.height;
  model.inverted = view.inverted;
  model.polar = view.polar;
  model.background = styles.background;
  model.border = styles.border;
  model.plotArea = styles.plotArea;
  model.title = styles.title;
  model.subtitle = styles.subtitle;
  // The legend is drawn only when an exported series has a legend item.
  model.legend = { ...styles.legend, enabled: styles.legend.enabled && series.some((s) => s.showInLegend) };
  model.xAxes = xAxes;
  model.yAxes = yAxes;
  model.series = series;
  model.colors = view.palette;
  model.warnings = [...diagnostics.items];
  return model;
}

/**
 * Extracts the IR from a live (rendered) Highcharts chart. Never mutates the chart.
 *
 * @experimental The IR and this entry point may change in minor versions.
 */
export function extractChartModel(chart: unknown, options: ExtractOptions): ChartModel {
  if (!isHighchartsChart(chart)) {
    throw new ExportError(
      'INVALID_CHART',
      'Expected a rendered Highcharts chart instance (with series, axes, options and a container).',
    );
  }
  const view = viewFromChart(chart, options);
  const renderToId = str(get(chart.renderTo, 'id'));
  const containerId = str(get(chart.container, 'id'));
  const meta: Omit<ChartMeta, 'datetimeOffsetMinutes'> = {
    sourceLibrary: 'highcharts',
    sourceVersion: options.highchartsVersion !== undefined ? options.highchartsVersion : getHighchartsVersion(chart),
    sourceChartType: str(get(view.userOpts, 'chart', 'type')) ?? view.series[0]?.type ?? 'line',
    extraction: isRealBrowser() ? 'browser' : 'headless',
    styledMode: view.styledMode,
    chartId: renderToId ?? containerId ?? null,
  };
  return buildModel(view, meta, options);
}

/**
 * Server-side path: extracts the IR from a plain Highcharts options object (no rendering).
 * Data is the raw `series[i].data`; sizes come from `chart.width/height` (default 600×400).
 *
 * @experimental The IR and this entry point may change in minor versions.
 */
export function extractChartModelFromOptions(options: object, extract: ExtractOptions): ChartModel {
  const opts = rec(options);
  if (!opts) throw new ExportError('INVALID_OPTIONS', 'Expected a Highcharts options object.');
  const view = viewFromOptions(opts, extract);
  const meta: Omit<ChartMeta, 'datetimeOffsetMinutes'> = {
    sourceLibrary: 'highcharts',
    sourceVersion:
      extract.highchartsVersion !== undefined ? extract.highchartsVersion : getHighchartsVersion(undefined),
    sourceChartType: str(get(opts, 'chart', 'type')) ?? view.series[0]?.type ?? 'line',
    extraction: 'headless',
    styledMode: view.styledMode,
    chartId: str(get(opts, 'chart', 'renderTo')) ?? null,
  };
  return buildModel(view, meta, { ...extract, dataMode: 'raw' });
}
