/**
 * Chart-level styling: background, border, plot area, title/subtitle, legend, palette, fonts.
 */

import type {
  Color,
  Fill,
  Font,
  HorizontalAlign,
  LegendModel,
  LegendPosition,
  PlotAreaModel,
  Stroke,
  TextBlock,
  VerticalAlign,
} from '../types/chart-model';
import type { DiagnosticCollector } from '../types/diagnostics';
import { toFill } from '../translators/color-translator';
import { strokeFromOptions } from '../translators/style-translator';
import { DEFAULT_HIGHCHARTS_FONT, toFont, type CssStyleLike } from '../translators/typography-translator';
import { HIGHCHARTS_DEFAULT_PALETTE, parseColor, type CssVariableResolver } from '../utils/colors';
import { readEffectiveStyle } from './css-resolver';
import { arr, bool, deepMerge, get, num, plainText, rec, str, type Rec } from './guards';
import { showsPointsInLegend } from './series-types';
import type { ChartView, HcChartLike, SeriesView } from './types';

/**
 * Highcharts 11-13 component text defaults. They are relative to the chart font (`chart.style`,
 * default `1rem` = 16px), so `em` sizes resolve against the base font size: title 1.2em = 19.2px,
 * subtitle/axis/legend 0.8em = 12.8px, data labels 0.7em = 11.2px.
 *
 * ASSUMPTION: headless, `em` is resolved against the base font (16px unless `chart.style.fontSize`
 * says otherwise), which is what the browser does for Highcharts' SVG text. In a real browser the
 * computed font-size of the rendered element wins.
 * Note: legend `itemStyle` has no `fontWeight` in Highcharts 12/13 sources, so legend text is normal weight.
 */
export const TEXT_DEFAULTS = {
  title: { fontSize: '1.2em', fontWeight: 'bold', color: '#333333' },
  subtitle: { fontSize: '0.8em', color: '#666666' },
  axisLabels: { fontSize: '0.8em', color: '#333333' },
  axisTitle: { fontSize: '0.8em', color: '#666666' },
  legend: { fontSize: '0.8em', color: '#333333' },
  dataLabels: { fontSize: '0.7em', fontWeight: 'bold' },
} as const;

export type TextKind = keyof typeof TEXT_DEFAULTS;

/** Base chart font: DEFAULT_HIGHCHARTS_FONT overridden by the renderer style and `chart.style`. */
export function computeBaseFont(rt: HcChartLike | undefined, opts: Rec, resolver: CssVariableResolver): Font {
  const style = deepMerge(rt?.renderer?.style, get(opts, 'chart', 'style'));
  // `fontSize: '1rem'` resolves to 16px in toFont; an `em` here is relative to the 16px root.
  return toFont(style as CssStyleLike, DEFAULT_HIGHCHARTS_FONT, resolver);
}

/** Font for a text component: defaults ← option style, resolved against the base font. */
export function fontFor(view: ChartView, kind: TextKind, style: unknown, element?: Element): Font {
  const merged = deepMerge(TEXT_DEFAULTS[kind], rec(style));
  const font = toFont(merged as CssStyleLike, view.baseFont, view.resolver);
  if (view.browser && element) {
    const computed = readEffectiveStyle(element, ['font-size', 'font-family', 'font-weight', 'font-style', 'fill']);
    const size = /^(\d*\.?\d+)px$/.exec(computed['font-size'] ?? '');
    if (size) font.size = Math.round(Number(size[1]) * 1000) / 1000;
    if (computed['font-family']) font.family = computed['font-family'];
    const weight = computed['font-weight'];
    if (weight) font.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
    if (computed['font-style']) font.italic = /italic|oblique/.test(computed['font-style']);
    const fill = parseColor(computed.fill, view.resolver);
    if (fill) font.color = fill;
  }
  return font;
}

/** The palette in effect (`colors` option), resolved. Empty or unresolvable → Highcharts defaults. */
export function extractPalette(opts: Rec, resolver: CssVariableResolver): Color[] {
  const list = arr(opts.colors) ?? [];
  const colors = list.map((c) => parseColor(c, resolver)).filter((c): c is Color => c !== null);
  if (colors.length > 0) return colors;
  return HIGHCHARTS_DEFAULT_PALETTE.map((c) => parseColor(c, resolver)).filter((c): c is Color => c !== null);
}

/** toFill that reports its diagnostic under `property`. */
export function fillAt(
  input: unknown,
  view: ChartView,
  diagnostics: DiagnosticCollector,
  property: string,
): Fill | null {
  const { fill, diagnostic } = toFill(input, view.resolver, property);
  if (diagnostic) diagnostics.add(diagnostic);
  return fill;
}

/** A border/outline stroke, or null when its width is 0 / unset. */
export function borderStroke(
  view: ChartView,
  color: unknown,
  width: unknown,
  defaults: { color: string; width: number },
  dashStyle?: unknown,
): Stroke | null {
  const w = num(width) ?? defaults.width;
  if (w <= 0) return null;
  const fallbackColor = parseColor(defaults.color, view.resolver);
  return strokeFromOptions(
    { color, width: w, dashStyle },
    { color: fallbackColor, width: w, dash: 'solid' },
    view.resolver,
  );
}

function hAlign(x: unknown, fallback: HorizontalAlign): HorizontalAlign {
  return x === 'left' || x === 'center' || x === 'right' ? x : fallback;
}

function vAlign(x: unknown, fallback: VerticalAlign): VerticalAlign {
  return x === 'top' || x === 'middle' || x === 'bottom' ? x : fallback;
}

function textBlock(view: ChartView, which: 'title' | 'subtitle'): TextBlock | null {
  const o = rec(view.opts[which]) ?? {};
  const wrapper = view.rt?.[which];
  let text: string | null;
  if (view.rt) {
    text = plainText(wrapper?.textStr ?? o.text);
  } else {
    // Highcharts renders "Chart title" when no title text is configured at all.
    const raw = o.text === undefined && which === 'title' ? 'Chart title' : o.text;
    text = plainText(raw);
  }
  if (text === null) return null;
  return {
    text,
    font: fontFor(view, which, o.style, wrapper?.element),
    align: hAlign(o.align, 'center'),
    verticalAlign: vAlign(o.verticalAlign, 'top'),
  };
}

/**
 * Highcharts `showInLegend`: explicit option, else true except for pie-like series (whose legend,
 * when enabled, lists points rather than the series).
 */
export function seriesShowsInLegend(s: SeriesView): boolean {
  return bool(s.opts.showInLegend) ?? !showsPointsInLegend(s.type);
}

function legendPosition(align: string, verticalAlign: string, diagnostics: DiagnosticCollector): LegendPosition {
  if (verticalAlign === 'bottom') return 'bottom';
  if (verticalAlign === 'top') return align === 'right' ? 'topRight' : align === 'left' ? 'topLeft' : 'top';
  if (align === 'right') return 'right';
  if (align === 'left') return 'left';
  diagnostics.report(
    'APPROXIMATED_LEGEND_POSITION',
    'approximated',
    'legend.align',
    'A legend centered over the plot area is placed at the right.',
    {
      severity: 'info',
      details: { align, verticalAlign },
    },
  );
  return 'right';
}

function extractLegend(view: ChartView, diagnostics: DiagnosticCollector): LegendModel {
  const o = rec(view.opts.legend) ?? {};
  // A legend without items is not drawn; whether the EXPORTED series have items is decided after
  // series extraction (extract-chart.ts), so hidden, excluded series do not keep the legend on.
  const align = str(o.align) ?? 'center';
  const verticalAlign = str(o.verticalAlign) ?? 'bottom';
  return {
    enabled: o.enabled !== false,
    position: legendPosition(align, verticalAlign, diagnostics),
    layout: o.layout === 'vertical' ? 'vertical' : 'horizontal',
    font: fontFor(view, 'legend', o.itemStyle),
    background: fillAt(o.backgroundColor, view, diagnostics, 'legend.backgroundColor'),
    border: borderStroke(view, o.borderColor, o.borderWidth, { color: '#999999', width: 0 }),
    overlay: o.floating === true,
    reversed: o.reversed === true,
  };
}

export interface ChartStyles {
  background: Fill | null;
  border: Stroke | null;
  plotArea: PlotAreaModel;
  title: TextBlock | null;
  subtitle: TextBlock | null;
  legend: LegendModel;
}

export function extractChartStyles(view: ChartView, diagnostics: DiagnosticCollector): ChartStyles {
  const c = rec(view.opts.chart) ?? {};

  let background: Fill | null = null;
  if (view.styledMode) {
    // Styled mode: colors live in CSS. Read the rendered background in a browser; else white.
    const computed = readEffectiveStyle(view.rt?.chartBackground?.element, ['fill']);
    background = computed.fill ? fillAt(computed.fill, view, diagnostics, 'chart.backgroundColor') : null;
    const bgVariable = view.cssVariables?.['--highcharts-background-color'];
    if (!background && typeof bgVariable === 'string') {
      background = fillAt('var(--highcharts-background-color)', view, diagnostics, 'chart.backgroundColor');
    }
    background ??= fillAt('#ffffff', view, diagnostics, 'chart.backgroundColor');
  } else {
    background = fillAt(
      c.backgroundColor === undefined ? '#ffffff' : c.backgroundColor,
      view,
      diagnostics,
      'chart.backgroundColor',
    );
  }

  let plotBackground: Fill | null = null;
  if (view.styledMode) {
    const computed = readEffectiveStyle(view.rt?.plotBackground?.element, ['fill']);
    if (computed.fill) plotBackground = fillAt(computed.fill, view, diagnostics, 'chart.plotBackgroundColor');
  } else {
    plotBackground = fillAt(c.plotBackgroundColor, view, diagnostics, 'chart.plotBackgroundColor');
  }

  return {
    background,
    border: borderStroke(view, c.borderColor, c.borderWidth, { color: '#334eff', width: 0 }),
    plotArea: {
      background: plotBackground,
      border: borderStroke(view, c.plotBorderColor, c.plotBorderWidth, { color: '#cccccc', width: 0 }),
      box: view.plotBox,
    },
    title: textBlock(view, 'title'),
    subtitle: textBlock(view, 'subtitle'),
    legend: extractLegend(view, diagnostics),
  };
}

/**
 * `lang.thousandsSep` / `lang.decimalPoint` of the chart (merged options on a live chart), for
 * `translateFormatString`, which reports APPROXIMATED_NUMBER_FORMAT (once per chart: property
 * `lang.thousandsSep` / `lang.decimalPoint`) when a translated format uses a separator that differs
 * from Excel's locale-driven `,` / `.`. Unset separators (Highcharts 12+ default: locale) are omitted.
 */
export function langSeparators(view: ChartView): { thousandsSep?: string; decimalPoint?: string } {
  const lang = rec(view.opts.lang) ?? {};
  const out: { thousandsSep?: string; decimalPoint?: string } = {};
  const t = str(lang.thousandsSep);
  const d = str(lang.decimalPoint);
  if (t !== undefined) out.thousandsSep = t;
  if (d !== undefined) out.decimalPoint = d;
  return out;
}
