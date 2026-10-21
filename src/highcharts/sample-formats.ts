/**
 * Collects `{ value, text }` samples from a live chart so `inferNumberFormat` can turn formatter
 * output into Excel number formats:
 * - axis tick labels and data labels are read as rendered (no code runs);
 * - tooltip text is produced the way Highcharts produces it on hover: the chart `tooltip.formatter`,
 *   else the series `pointFormatter` / `pointFormat`, called on a few sample points. Formatters
 *   are called with a context object and never change the chart; a throwing formatter yields no
 *   sample.
 */

import { type FormatSample, type InferOptions, inferNumberFormat } from '../translators/format-inference';
import type { NumberFormat } from '../types/chart-model';
import { arr, rec, str } from './guards';
import type {
  AxisView,
  ChartView,
  ExtractContext,
  HcPointLike,
  HcSeriesLike,
  HcWrapperLike,
  SeriesView,
} from './types';

/** Points sampled per series for tooltips and data labels. */
const MAX_POINT_SAMPLES = 8;

function wrapperText(w: HcWrapperLike | undefined): string | null {
  if (!w) return null;
  if (typeof w.textStr === 'string') return w.textStr;
  const t = w.element?.textContent;
  return typeof t === 'string' ? t : null;
}

/** Separators and numeric symbols the chart renders with (`lang`). */
export function inferOptionsFor(view: ChartView): InferOptions {
  const lang = rec(view.opts.lang) ?? {};
  const out: InferOptions = {};
  const d = str(lang.decimalPoint);
  const t = str(lang.thousandsSep);
  if (d !== undefined && d.length === 1) out.decimalPoint = d;
  if (t !== undefined) out.thousandsSep = t;
  const syms = arr(lang.numericSymbols);
  if (syms) out.numericSymbols = syms.map((x) => (typeof x === 'string' ? x : null));
  return out;
}

/** Rendered tick labels of a live axis. */
export function sampleAxisLabels(axis: AxisView): FormatSample[] {
  const ticks = axis.rt?.ticks;
  const positions = axis.rt?.tickPositions;
  if (!ticks || !Array.isArray(positions)) return [];
  const out: FormatSample[] = [];
  for (const pos of positions) {
    if (typeof pos !== 'number' || !Number.isFinite(pos)) continue;
    const text = wrapperText(ticks[String(pos)]?.label);
    if (text !== null && text.trim() !== '') out.push({ value: pos, text });
  }
  return out;
}

function pointValue(p: HcPointLike): number | null {
  if (typeof p.y === 'number' && Number.isFinite(p.y)) return p.y;
  if (typeof p.high === 'number' && Number.isFinite(p.high)) return p.high;
  return null;
}

/** Up to `max` points with values, spread over the series, always including the largest |value|. */
function samplePoints(rt: HcSeriesLike, max = MAX_POINT_SAMPLES): HcPointLike[] {
  const pts = (rt.points ?? []).filter((p): p is HcPointLike => !!p && p.isNull !== true && pointValue(p) !== null);
  if (pts.length <= max) return pts;
  const picked = new Set<HcPointLike>();
  let largest = pts[0]!;
  for (const p of pts) if (Math.abs(pointValue(p)!) > Math.abs(pointValue(largest)!)) largest = p;
  picked.add(largest);
  const step = (pts.length - 1) / (max - 2);
  for (let i = 0; i <= max - 2 && picked.size < max; i++) picked.add(pts[Math.round(i * step)]!);
  return [...picked];
}

/** Rendered data labels of a live series. */
export function sampleDataLabels(s: SeriesView): FormatSample[] {
  if (!s.rt) return [];
  const out: FormatSample[] = [];
  for (const p of samplePoints(s.rt, Number.POSITIVE_INFINITY)) {
    const text = wrapperText(p.dataLabel ?? p.dataLabels?.[0]);
    const value = pointValue(p);
    if (text !== null && text.trim() !== '' && value !== null) out.push({ value, text });
    if (out.length >= MAX_POINT_SAMPLES) break;
  }
  return out;
}

function textOf(result: unknown): string | null {
  if (typeof result === 'string') return result;
  if (Array.isArray(result)) return result.filter((x) => typeof x === 'string').join('<br/>');
  return null;
}

/** The points a shared tooltip shows together with `p`: one per visible series at the same x. */
function sharedPoints(view: ChartView, p: HcPointLike, own: HcSeriesLike): HcPointLike[] {
  const out: HcPointLike[] = [];
  for (const sv of view.series) {
    const rt = sv.rt;
    if (!rt || !sv.visible || rt.visible === false) continue;
    if (rt === own) {
      out.push(p);
      continue;
    }
    const q = (rt.points ?? []).find((x) => !!x && x.x === p.x && x.isNull !== true);
    if (q) out.push(q);
  }
  return out;
}

/**
 * Tooltip text Highcharts would show for a few points of `s`, paired with the point values.
 * Empty when the tooltip is disabled or the chart is not live.
 */
export function sampleTooltip(view: ChartView, s: SeriesView): FormatSample[] {
  const chart = view.rt;
  const rt = s.rt;
  if (!chart || !rt) return [];
  const tt = rec(view.opts.tooltip) ?? {};
  if (tt.enabled === false || s.opts.enableMouseTracking === false) return [];
  const sTt = rec(rt.tooltipOptions) ?? {};
  const chartFormatter = typeof tt.formatter === 'function' ? (tt.formatter as (...a: unknown[]) => unknown) : null;
  // A chart-level `tooltip.format` string needs Highcharts' `format`; the string translator covers it.
  if (!chartFormatter && typeof tt.format === 'string' && tt.format !== '') return [];
  const shared = tt.shared === true && tt.split !== true;
  const pointFormatter =
    typeof sTt.pointFormatter === 'function' ? (sTt.pointFormatter as (...a: unknown[]) => unknown) : null;
  const pointFormat = str(sTt.pointFormat) ?? '';

  const out: FormatSample[] = [];
  for (const p of samplePoints(rt)) {
    const value = pointValue(p)!;
    let result: unknown;
    try {
      if (chartFormatter) {
        const group = shared ? sharedPoints(view, p, rt) : [p];
        if (typeof p.getLabelConfig === 'function') {
          // Highcharts 11: `this` is the label config, with `points` when shared.
          const ctx = p.getLabelConfig();
          if (shared) ctx.points = group.map((q) => (typeof q.getLabelConfig === 'function' ? q.getLabelConfig() : q));
          result = chartFormatter.call(ctx, chart.tooltip);
        } else {
          // Highcharts 12+: `this` is the point, with `points` when shared (set on a child object).
          const ctx = shared ? Object.create(p, { points: { value: group, configurable: true } }) : p;
          result = chartFormatter.call(ctx, chart.tooltip, ctx);
        }
      } else if (pointFormatter) {
        result = pointFormatter.call(p, pointFormat, p);
      } else if (typeof p.tooltipFormatter === 'function') {
        result = p.tooltipFormatter.call(p, pointFormat);
      }
    } catch {
      continue;
    }
    const text = textOf(result);
    if (text !== null && text.trim() !== '') out.push({ value, text });
  }
  return out;
}

/** The inferred format as an IR NumberFormat, plus an info diagnostic naming it. `samples` runs only when inference is on. */
export function inferredFormat(
  ctx: ExtractContext,
  samples: () => FormatSample[],
  property: string,
  what: string,
  opts: { plainFormats?: boolean; seriesIndex?: number } = {},
): NumberFormat | null {
  if (!ctx.inferFormats) return null;
  const infer = inferOptionsFor(ctx.view);
  if (opts.plainFormats === false) infer.plainFormats = false;
  const r = inferNumberFormat(samples(), infer);
  if (!r) return null;
  ctx.diagnostics.report(
    'APPROXIMATED_NUMBER_FORMAT',
    'approximated',
    property,
    `Excel number format "${r.code}" was inferred from the ${what} the chart renders.`,
    {
      severity: 'info',
      ...(opts.seriesIndex !== undefined ? { seriesIndex: opts.seriesIndex } : {}),
      details: { code: r.code, example: r.example.slice(0, 200), matched: r.matched, sampled: r.total },
    },
  );
  return { kind: 'excel', code: r.code, source: `inferred from ${what}` };
}
