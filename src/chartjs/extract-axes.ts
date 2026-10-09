/**
 * Chart.js scales → IR axes.
 *
 * The IR's x axes are the index (category) scales and its y axes the value scales: for a
 * horizontal bar chart (`indexAxis: 'y'`) Chart.js's `y` scales become IR x axes, as Highcharts
 * models an inverted chart.
 */

import type { AxisKind, AxisModel, Font, Stroke, TextBlock } from '../types/chart-model';
import type { DiagnosticCollector } from '../types/diagnostics';
import { parseColor } from '../utils/colors';
import { CHARTJS_DEFAULT_ELEMENT_COLOR, CHARTJS_DEFAULT_SCALES } from './defaults';
import { componentFont, dashFromBorderDash } from './extract-styles';
import { arr, get, num, pick, rec, str } from './guards';
import { dateDisplayFormatToExcel, intlFormatToExcel } from './number-format';
import type { WallClock } from './time';
import type { AxisLetter, ChartJsChartLike, ChartJsView, DatasetView, Rec, ScaleView } from './types';

function idMatchesAxis(id: string): 'x' | 'y' | 'r' | undefined {
  return id === 'x' || id === 'y' || id === 'r' ? id : undefined;
}

function axisFromPosition(position: unknown): 'x' | 'y' | undefined {
  if (position === 'top' || position === 'bottom') return 'x';
  if (position === 'left' || position === 'right') return 'y';
  return undefined;
}

/** Chart.js `determineAxis`: the id itself, `axis`, the position, a dataset binding, the id's first letter. */
function determineAxis(id: string, opts: Rec, datasets: readonly Rec[]): 'x' | 'y' | 'r' {
  const fromId = idMatchesAxis(id);
  if (fromId) return fromId;
  const explicit = opts.axis === 'x' || opts.axis === 'y' || opts.axis === 'r' ? opts.axis : undefined;
  if (explicit) return explicit;
  const fromPosition = axisFromPosition(opts.position);
  if (fromPosition) return fromPosition;
  const bound = datasets.find((d) => d.xAxisID === id || d.yAxisID === id);
  if (bound) return bound.xAxisID === id ? 'x' : 'y';
  return idMatchesAxis(id.charAt(0).toLowerCase()) ?? 'y';
}

function otherAxis(a: AxisLetter): AxisLetter {
  return a === 'x' ? 'y' : 'x';
}

/**
 * Builds the scale list of a configuration the way Chart.js's `mergeScaleConfig` does (declared
 * scales first, then the default scales each dataset type implies) and binds every dataset to
 * its index and value scales.
 */
export function scalesFromConfig(view: Pick<ChartJsView, 'type' | 'options'>, datasets: DatasetView[]): ScaleView[] {
  const declared = rec(view.options.scales) ?? {};
  const rawDatasets = datasets.map((d) => d.ds);
  const scales = new Map<string, ScaleView & { defaultType?: string; defaults: Rec }>();
  for (const [id, value] of Object.entries(declared)) {
    const opts = rec(value);
    if (!opts) continue;
    scales.set(id, {
      id,
      axis: determineAxis(id, opts, rawDatasets),
      type: '',
      opts,
      path: `options.scales.${id}`,
      defaults: {},
    });
  }
  const chartDefaults = CHARTJS_DEFAULT_SCALES[view.type] ?? {};
  const chartIndex: AxisLetter = view.options.indexAxis === 'y' ? 'y' : 'x';
  for (const s of scales.values()) {
    if (s.axis === 'r') continue;
    const d = rec(chartDefaults[s.axis]) ?? rec(chartDefaults[s.axis === chartIndex ? '_index_' : '_value_']) ?? {};
    s.defaults = d;
  }
  const firstOf = (axis: string): string | undefined => [...scales.values()].find((s) => s.axis === axis)?.id;

  for (const dv of datasets) {
    const typeDefaults = CHARTJS_DEFAULT_SCALES[dv.type] ?? {};
    for (const [defaultId, defaultsValue] of Object.entries(typeDefaults)) {
      const axis =
        defaultId === '_index_' ? dv.indexAxis : defaultId === '_value_' ? otherAxis(dv.indexAxis) : defaultId;
      const explicitId = str(dv.ds[`${axis}AxisID`]);
      const id = explicitId ?? axis;
      let s = scales.get(id);
      if (!s) {
        s = {
          id,
          axis: axis as 'x' | 'y' | 'r',
          type: '',
          opts: {},
          path: `options.scales.${id}`,
          defaults: {},
        };
        scales.set(id, s);
      }
      s.defaults = { ...(rec(defaultsValue) ?? {}), ...s.defaults };
    }
  }
  for (const dv of datasets) {
    if (dv.indexScaleId === '') dv.indexScaleId = str(dv.ds[`${dv.indexAxis}AxisID`]) ?? firstOf(dv.indexAxis) ?? '';
    const v = otherAxis(dv.indexAxis);
    if (dv.valueScaleId === '') dv.valueScaleId = str(dv.ds[`${v}AxisID`]) ?? firstOf(v) ?? '';
  }
  return [...scales.values()].map(({ defaults, ...s }) => ({
    ...s,
    type: str(s.opts.type) ?? str(defaults.type) ?? (s.axis === 'r' ? 'radialLinear' : 'linear'),
    opts: { ...defaults, ...s.opts },
  }));
}

/** The scales of a live chart, with Chart.js's merged scale options. */
export function scalesFromChart(chart: ChartJsChartLike): ScaleView[] {
  const merged = rec(get(chart.config.options, 'scales')) ?? {};
  return Object.entries(chart.scales).map(([id, scale]) => {
    const opts = rec(merged[id]) ?? rec(scale.options) ?? {};
    const axis =
      scale.axis === 'x' || scale.axis === 'y' || scale.axis === 'r' ? scale.axis : determineAxis(id, opts, []);
    return {
      id,
      axis,
      type: str(scale.type) ?? str(opts.type) ?? 'linear',
      opts,
      path: `options.scales.${id}`,
      live: scale,
    };
  });
}

export function axisKindOf(scaleType: string): AxisKind {
  switch (scaleType) {
    case 'category':
      return 'category';
    case 'logarithmic':
      return 'logarithmic';
    case 'time':
    case 'timeseries':
      return 'datetime';
    default:
      return 'linear';
  }
}

/** Label text of a category: multi-line labels (arrays) are joined with spaces. */
export function labelText(label: unknown): string {
  if (Array.isArray(label)) return label.map(labelText).join(' ');
  if (label === null || label === undefined) return '';
  if (label instanceof Date) return label.toISOString();
  return String(label);
}

/** The category labels a category scale shows. */
export function scaleLabels(view: ChartJsView, scale: ScaleView): unknown[] {
  if (scale.live && typeof scale.live.getLabels === 'function') {
    try {
      const labels = scale.live.getLabels();
      if (Array.isArray(labels)) return labels;
    } catch {
      // fall through to the configuration
    }
  }
  const own = arr(scale.opts.labels);
  if (own) return own;
  const axisLabels =
    scale.axis === 'x' ? arr(view.data.xLabels) : scale.axis === 'y' ? arr(view.data.yLabels) : undefined;
  return axisLabels ?? view.labels;
}

/** date-fns adapter default display formats by unit (Chart.js `chartjs-adapter-date-fns`). */
const DEFAULT_DISPLAY_FORMATS: Readonly<Record<string, string>> = {
  millisecond: 'h:mm:ss.SSS a',
  second: 'h:mm:ss a',
  minute: 'h:mm a',
  hour: 'ha',
  day: 'MMM d',
  month: 'MMM yyyy',
  year: 'yyyy',
};

/** Extra bound hints applied once the data extremes are known. */
export interface AxisHints {
  suggestedMin: number | null;
  suggestedMax: number | null;
  beginAtZero: boolean;
}

export interface ExtractedAxis {
  model: AxisModel;
  hints: AxisHints;
  scale: ScaleView;
}

function isDefaultCallback(view: ChartJsView, scale: ScaleView, fn: unknown): boolean {
  if (!view.defaults) return false;
  return (
    fn === get(view.defaults, 'scales', scale.type, 'ticks', 'callback') ||
    fn === get(view.defaults, 'scale', 'ticks', 'callback')
  );
}

function scaleColor(value: unknown, property: string, diagnostics: DiagnosticCollector): string | undefined {
  if (typeof value === 'string') return value;
  if (Array.isArray(value) && typeof value[0] === 'string') {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'approximated',
      property,
      'Per-tick colors are not reproduced; the first color is used for every line.',
      { severity: 'info' },
    );
    return value[0];
  }
  if (typeof value === 'function') {
    diagnostics.report(
      'UNSUPPORTED_STYLE',
      'approximated',
      property,
      'Scriptable grid/border colors are not evaluated; the default color is used.',
      { severity: 'info' },
    );
  }
  return undefined;
}

function strokeOf(
  color: string | undefined,
  width: unknown,
  fallbackColor: string,
  dash: Stroke['dash'] = 'solid',
): Stroke {
  return {
    color: parseColor(color ?? fallbackColor) ?? parseColor(CHARTJS_DEFAULT_ELEMENT_COLOR),
    width: num(Array.isArray(width) ? width[0] : width) ?? 1,
    dash,
  };
}

function axisTitle(
  view: ChartJsView,
  scale: ScaleView,
  base: Font,
  diagnostics: DiagnosticCollector,
): TextBlock | null {
  const t = rec(scale.opts.title);
  if (t?.display !== true) return null;
  const raw = Array.isArray(t.text) ? t.text.map(labelText).join('\n') : t.text;
  if (typeof raw !== 'string' && typeof raw !== 'number') return null;
  const text = String(raw).trim();
  if (text === '') return null;
  const font = componentFont(
    base,
    [
      {
        font: get(view.defaults, 'scale', 'title', 'font'),
        color: get(view.defaults, 'scale', 'title', 'color'),
      },
      { font: t.font, color: t.color },
    ],
    `${scale.path}.title`,
    diagnostics,
  );
  const align = t.align === 'start' ? 'left' : t.align === 'end' ? 'right' : 'center';
  return { text, font, align, verticalAlign: 'middle' };
}

function boundValue(value: unknown, kind: AxisKind, time: WallClock): number | null {
  if (kind === 'datetime') return value === undefined || value === null ? null : time.parse(value);
  if (kind === 'category') return null;
  return num(value) ?? (typeof value === 'string' && value.trim() !== '' ? (num(Number(value)) ?? null) : null);
}

/** One Chart.js scale → AxisModel (index = position in the IR list). */
export function extractAxis(
  view: ChartJsView,
  scale: ScaleView,
  index: number,
  base: Font,
  time: WallClock,
  diagnostics: DiagnosticCollector,
): ExtractedAxis {
  const o = scale.opts;
  const kind = axisKindOf(scale.type);
  const ticks = rec(o.ticks) ?? {};
  const grid = rec(o.grid) ?? {};
  const border = rec(o.border) ?? {};
  const defaultLine =
    typeof view.defaults?.borderColor === 'string' ? view.defaults.borderColor : CHARTJS_DEFAULT_ELEMENT_COLOR;

  let format: AxisModel['labels']['format'] = null;
  if (typeof ticks.callback === 'function' && !isDefaultCallback(view, scale, ticks.callback)) {
    diagnostics.report(
      'UNSUPPORTED_FORMATTER',
      'unsupported',
      `${scale.path}.ticks.callback`,
      'Tick callbacks are not executed; Excel uses its default formatting.',
    );
    format = { kind: 'unsupported', reason: 'ticks.callback' };
  } else if (kind === 'datetime') {
    const t = rec(o.time) ?? {};
    const unit = str(t.unit);
    const fmt = unit ? (get(t, 'displayFormats', unit) ?? DEFAULT_DISPLAY_FORMATS[unit]) : undefined;
    const code = dateDisplayFormatToExcel(fmt);
    if (code) format = { kind: 'excel', code, source: String(fmt) };
    else if (fmt !== undefined) {
      diagnostics.report(
        'APPROXIMATED_NUMBER_FORMAT',
        'approximated',
        `${scale.path}.time.displayFormats.${unit}`,
        `Date format "${String(fmt)}" has no Excel equivalent; Excel picks a date format.`,
        { severity: 'info' },
      );
    }
  } else if (ticks.format !== undefined) {
    const r = intlFormatToExcel(ticks.format, `${scale.path}.ticks.format`);
    format = r.format;
    if (r.diagnostic) diagnostics.add(r.diagnostic);
  }

  if (kind === 'datetime') {
    const t = rec(o.time) ?? {};
    if (t.parser !== undefined) {
      diagnostics.report(
        'APPROXIMATED_DATETIME',
        'approximated',
        `${scale.path}.time.parser`,
        'Custom time parsers are not run; time values are parsed as ISO strings or timestamps.',
        { severity: 'info' },
      );
    }
    if (get(o, 'adapters', 'date', 'zone') !== undefined) {
      diagnostics.report(
        'APPROXIMATED_DATETIME',
        'approximated',
        `${scale.path}.adapters.date.zone`,
        "The adapter time zone is not applied; datetime values are exported in the browser's local time.",
      );
    }
    if (scale.type === 'timeseries') {
      diagnostics.report(
        'APPROXIMATED_AXIS_SCALE',
        'approximated',
        `${scale.path}.type`,
        'A timeseries scale spaces points evenly; Excel date axes space them by time.',
        { severity: 'info' },
      );
    }
  }
  if (o.grace !== undefined && o.grace !== 0 && o.grace !== '0%') {
    diagnostics.report(
      'APPROXIMATED_AXIS_SCALE',
      'approximated',
      `${scale.path}.grace`,
      'Scale grace is not reproduced; Excel picks its own padding.',
      { severity: 'info' },
    );
  }

  let rotation = 0;
  const minRot = num(ticks.minRotation);
  const maxRot = num(ticks.maxRotation);
  if (minRot !== undefined && minRot !== 0 && (maxRot === undefined || maxRot === minRot)) rotation = -minRot;

  const position = o.position;
  const posValue = rec(position);
  let crossing: number | null = null;
  if (posValue) {
    const v = Object.values(posValue)[0];
    crossing = num(v) ?? null;
  } else if (position === 'center') {
    diagnostics.report(
      'UNSUPPORTED_AXIS_FEATURE',
      'approximated',
      `${scale.path}.position`,
      'A centered axis is drawn at the plot edge in Excel.',
      { severity: 'info' },
    );
  }
  const livePosition = scale.live?.position;
  const effectivePosition = typeof livePosition === 'string' ? livePosition : position;
  const opposite = scale.axis === 'x' ? effectivePosition === 'top' : effectivePosition === 'right';

  const gridVisible = grid.display !== false;
  const gridColor = scaleColor(grid.color, `${scale.path}.grid.color`, diagnostics);
  const borderColor = scaleColor(border.color, `${scale.path}.border.color`, diagnostics);
  const borderWidth = num(border.width) ?? 1;
  const axisDash =
    arr(border.dash) && (border.dash as unknown[]).length > 0
      ? dashFromBorderDash(border.dash, borderWidth, `${scale.path}.border.dash`, diagnostics)
      : 'solid';

  const model: AxisModel = {
    index,
    id: scale.id,
    kind,
    categories: kind === 'category' ? scaleLabels(view, scale).map(labelText) : null,
    title: axisTitle(view, scale, base, diagnostics),
    labels: {
      enabled: ticks.display !== false,
      font: componentFont(
        base,
        [
          { font: get(view.defaults, 'scale', 'ticks', 'font'), color: get(view.defaults, 'scale', 'ticks', 'color') },
          { font: ticks.font, color: typeof ticks.color === 'string' ? ticks.color : undefined },
        ],
        `${scale.path}.ticks`,
        diagnostics,
      ),
      format,
      rotation,
    },
    min: boundValue(o.min, kind, time),
    max: boundValue(o.max, kind, time),
    dataMin: null,
    dataMax: null,
    tickInterval: kind === 'datetime' ? null : (num(ticks.stepSize) ?? null),
    minorTickInterval: null,
    reversed: o.reverse === true,
    opposite,
    visible: o.display !== false,
    gridLines: gridVisible ? strokeOf(gridColor, grid.lineWidth, defaultLine) : null,
    minorGridLines: null,
    axisLine:
      border.display !== false && o.display !== false
        ? strokeOf(borderColor, borderWidth, defaultLine, axisDash)
        : { color: null, width: 0, dash: 'solid' },
    tickMarks:
      gridVisible && grid.drawTicks !== false
        ? strokeOf(
            scaleColor(grid.tickColor, `${scale.path}.grid.tickColor`, diagnostics) ?? gridColor,
            pick(grid.tickWidth, grid.lineWidth),
            defaultLine,
          )
        : { color: null, width: 0, dash: 'solid' },
    crossing,
    logBase: kind === 'logarithmic' ? 10 : null,
    dateFormat: null,
  };
  return {
    model,
    scale,
    hints: {
      suggestedMin: boundValue(o.suggestedMin, kind, time),
      suggestedMax: boundValue(o.suggestedMax, kind, time),
      beginAtZero: o.beginAtZero === true,
    },
  };
}
