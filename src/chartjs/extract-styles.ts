/**
 * Chart.js styling → neutral IR styles: option resolution (dataset → `options.datasets[type]` →
 * `options.elements` → defaults, as Chart.js resolves them), colors, fonts, title, legend, background.
 */

import type {
  Color,
  DashStyle,
  Fill,
  Font,
  LegendModel,
  LegendPosition,
  MarkerSymbol,
  TextBlock,
} from '../types/chart-model';
import type { DiagnosticCollector } from '../types/diagnostics';
import { toFont, type CssStyleLike } from '../translators/typography-translator';
import { parseColor } from '../utils/colors';
import {
  CHARTJS_CONTROLLER_DEFAULTS,
  CHARTJS_DEFAULT_ELEMENT_COLOR,
  CHARTJS_DEFAULT_FONT,
  CHARTJS_DEFAULT_TEXT_COLOR,
  CHARTJS_ELEMENT_DEFAULTS,
} from './defaults';
import { arr, get, isRealBrowser, pick, rec, str } from './guards';
import type { ChartJsView, DatasetView } from './types';

export type ElementKind = 'line' | 'point' | 'bar' | 'arc';

/** The dataset-level and data-level element kinds of a Chart.js dataset type. */
export function elementKindsOf(type: string): { dataset: 'line' | null; data: ElementKind } {
  switch (type) {
    case 'line':
    case 'radar':
    case 'scatter':
      return { dataset: 'line', data: 'point' };
    case 'bubble':
      return { dataset: null, data: 'point' };
    case 'pie':
    case 'doughnut':
    case 'polarArea':
      return { dataset: null, data: 'arc' };
    default:
      return { dataset: null, data: 'bar' };
  }
}

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

export interface OptionQuery {
  /** Element whose defaults apply. */
  element: ElementKind;
  key: string;
  /** Data index (data elements) or dataset index (dataset elements), for indexable arrays. */
  index: number;
  /** Chart.js data-element prefix: `point` makes `pointRadius` win over `radius`. */
  prefix?: string;
  /** False for options Chart.js never indexes (`borderDash`). Default true. */
  indexable?: boolean;
  /** Read the resolved option from the live element at this data index (`dataset` = the line). */
  liveElement?: 'dataset' | 'data';
}

/**
 * Resolves one element option of a dataset the way Chart.js does. A live chart's resolved element
 * options win (Chart.js already evaluated scriptable options there). In a configuration,
 * scriptable (function) options cannot be evaluated: they are reported and the next fallback is used.
 */
export function resolveOption(
  view: ChartJsView,
  dv: DatasetView,
  q: OptionQuery,
  diagnostics: DiagnosticCollector,
): unknown {
  if (q.liveElement && dv.meta) {
    const el = q.liveElement === 'dataset' ? dv.meta.dataset : dv.meta.data?.[q.index];
    const v = rec(el?.options)?.[q.key];
    if (v !== undefined && typeof v !== 'function') return v;
  }
  const prefixed = q.prefix ? `${q.prefix}${capitalize(q.key)}` : null;
  const typeOpts = rec(get(view.options, 'datasets', dv.type));
  const isColor = q.key === 'backgroundColor' || q.key === 'borderColor';
  const candidates: Array<[unknown, string]> = [];
  if (prefixed) candidates.push([dv.ds[prefixed], `${dv.path}.${prefixed}`]);
  const auto: [unknown, string] | null =
    isColor && dv.autoColors ? [dv.autoColors[q.key as 'backgroundColor'], 'options.plugins.colors'] : null;
  if (auto && dv.autoColorsForce) candidates.push(auto);
  candidates.push([dv.ds[q.key], `${dv.path}.${q.key}`]);
  if (auto && !dv.autoColorsForce) candidates.push(auto);
  if (typeOpts) {
    if (prefixed) candidates.push([typeOpts[prefixed], `options.datasets.${dv.type}.${prefixed}`]);
    candidates.push([typeOpts[q.key], `options.datasets.${dv.type}.${q.key}`]);
  }
  candidates.push([get(view.options, 'elements', q.element, q.key), `options.elements.${q.element}.${q.key}`]);
  if (isColor) candidates.push([view.options[q.key], `options.${q.key}`]);
  candidates.push([get(view.defaults, 'elements', q.element, q.key), `Chart.defaults.elements.${q.element}.${q.key}`]);
  candidates.push([CHARTJS_ELEMENT_DEFAULTS[q.element][q.key], '']);
  candidates.push([CHARTJS_CONTROLLER_DEFAULTS[dv.type]?.[q.key], '']);
  if (isColor) {
    candidates.push([view.defaults?.[q.key], 'Chart.defaults']);
    candidates.push([CHARTJS_DEFAULT_ELEMENT_COLOR, '']);
  }

  for (const [value, path] of candidates) {
    if (value === undefined || value === null) continue;
    if (typeof value === 'function') {
      diagnostics.report(
        'UNSUPPORTED_STYLE',
        'approximated',
        path,
        `Scriptable option "${q.key}" (a function) cannot be evaluated outside a rendered chart; the next fallback is used.`,
        { seriesIndex: dv.index, details: { option: q.key } },
      );
      continue;
    }
    if (Array.isArray(value) && q.indexable !== false) {
      if (value.length === 0) continue;
      const item = value[((q.index % value.length) + value.length) % value.length];
      if (item === undefined || item === null) continue;
      if (typeof item === 'function') continue;
      return item;
    }
    return value;
  }
  return undefined;
}

/** Converts a Chart.js color value to a Color, reporting canvas gradients/patterns and bad strings. */
export function colorOf(
  value: unknown,
  property: string,
  diagnostics: DiagnosticCollector,
  seriesIndex?: number,
): Color | null {
  if (value === undefined || value === null) return null;
  const extra = seriesIndex !== undefined ? { seriesIndex } : {};
  if (typeof value === 'string') {
    const c = parseColor(value);
    if (c) return c;
    diagnostics.report(
      'UNRESOLVED_COLOR',
      'approximated',
      property,
      `Could not resolve color "${value}"; Excel's automatic color is used.`,
      {
        ...extra,
        details: { value },
      },
    );
    return null;
  }
  diagnostics.report(
    'UNSUPPORTED_GRADIENT',
    'approximated',
    property,
    'Canvas gradients and patterns are not exported; Excel picks an automatic color.',
    extra,
  );
  return null;
}

export function solid(color: Color | null): Fill | null {
  return color ? { type: 'solid', color } : null;
}

/**
 * Chart.js `borderDash` (canvas dash lengths in px) → the nearest Excel preset dash. Excel scales
 * presets with the line width, so the pattern is always an approximation and is reported as such.
 */
export function dashFromBorderDash(
  dash: unknown,
  width: number,
  property: string,
  diagnostics: DiagnosticCollector,
  seriesIndex?: number,
): DashStyle {
  const list = (arr(dash) ?? []).map((v) => (typeof v === 'number' && Number.isFinite(v) && v >= 0 ? v : 0));
  if (list.length === 0 || list.every((v) => v === 0)) return 'solid';
  const w = width > 0 ? width : 1;
  const pattern = list.length % 2 === 1 ? [...list, ...list] : list;
  const on = pattern[0]! / w;
  const off = (pattern[1] ?? pattern[0]!) / w;
  let style: DashStyle;
  if (pattern.length >= 6) style = on > 4.5 ? 'longdashdotdot' : 'shortdashdotdot';
  else if (pattern.length >= 4) style = on > 4.5 ? 'longdashdot' : on > 1.5 ? 'dashdot' : 'shortdashdot';
  else if (on <= 1.5) style = off <= 1.5 ? 'shortdot' : 'dot';
  else if (on <= 4.5) style = off <= 1.5 ? 'shortdash' : 'dash';
  else style = 'longdash';
  diagnostics.report(
    'UNSUPPORTED_STYLE',
    'approximated',
    property,
    `Dash pattern [${list.join(', ')}] is drawn with Excel's closest preset dash ("${style}").`,
    { severity: 'info', ...(seriesIndex !== undefined ? { seriesIndex } : {}), details: { borderDash: list, style } },
  );
  return style;
}

/** Chart.js `pointStyle` → MarkerSymbol. */
export function pointStyleToSymbol(style: unknown): MarkerSymbol {
  if (style === false) return 'none';
  if (typeof style !== 'string') return style === undefined || style === null || style === true ? 'circle' : 'other';
  switch (style) {
    case 'circle':
      return 'circle';
    case 'rect':
    case 'rectRounded':
      return 'square';
    case 'rectRot':
      return 'diamond';
    case 'triangle':
      return 'triangle';
    default:
      return 'other';
  }
}

// ---------------------------------------------------------------------------
// Fonts and text
// ---------------------------------------------------------------------------

function fontCss(font: unknown, color: unknown): CssStyleLike {
  const f = rec(font) ?? {};
  const css: CssStyleLike = {};
  if (typeof f.family === 'string') css.fontFamily = f.family;
  if (typeof f.size === 'number' || typeof f.size === 'string') css.fontSize = f.size;
  if (typeof f.weight === 'number' || typeof f.weight === 'string') css.fontWeight = f.weight;
  if (typeof f.style === 'string') css.fontStyle = f.style;
  if (typeof color === 'string') css.color = color;
  return css;
}

/** `Chart.defaults.font` / `Chart.defaults.color` (live) or Chart.js 4's built-in defaults. */
export function baseFont(view: ChartJsView): Font {
  const hard = toFont(fontCss(CHARTJS_DEFAULT_FONT, CHARTJS_DEFAULT_TEXT_COLOR), null);
  const liveFont = get(view.defaults, 'font');
  const liveColor = view.defaults?.color;
  const chartFont = view.options.font;
  const chartColor = view.options.color;
  const withLive = toFont(fontCss(liveFont, liveColor), hard);
  return toFont(fontCss(chartFont, chartColor), withLive);
}

/**
 * A component font: Chart.js merges the component `font` over `Chart.defaults.font` and takes
 * `color` from the component, else `Chart.defaults.color`. Scriptable fonts are reported.
 */
export function componentFont(
  base: Font,
  layers: Array<{ font: unknown; color: unknown }>,
  property: string,
  diagnostics: DiagnosticCollector,
): Font {
  let font = base;
  for (const layer of layers) {
    let f = layer.font;
    let c = layer.color;
    if (typeof f === 'function') {
      diagnostics.report(
        'UNSUPPORTED_STYLE',
        'approximated',
        `${property}.font`,
        'Scriptable fonts are not evaluated.',
        {
          severity: 'info',
        },
      );
      f = undefined;
    }
    if (typeof c === 'function') {
      diagnostics.report(
        'UNSUPPORTED_STYLE',
        'approximated',
        `${property}.color`,
        'Scriptable colors are not evaluated.',
        {
          severity: 'info',
        },
      );
      c = undefined;
    }
    font = toFont(fontCss(f, c), font);
  }
  return font;
}

/** `options.plugins.title` / `options.plugins.subtitle` → TextBlock (null when not displayed). */
export function extractTitle(
  view: ChartJsView,
  which: 'title' | 'subtitle',
  base: Font,
  diagnostics: DiagnosticCollector,
): TextBlock | null {
  const path = `options.plugins.${which}`;
  const o = rec(get(view.options, 'plugins', which)) ?? {};
  const d = rec(get(view.defaults, 'plugins', which)) ?? {};
  if (pick(o.display, d.display, false) !== true) return null;
  const raw = o.text;
  if (typeof raw === 'function') {
    diagnostics.report(
      'UNSUPPORTED_FORMATTER',
      'unsupported',
      `${path}.text`,
      'Scriptable title text is not evaluated.',
    );
    return null;
  }
  const lines = (Array.isArray(raw) ? raw : [raw]).filter(
    (t): t is string | number => typeof t === 'string' || typeof t === 'number',
  );
  const text = lines
    .map((t) => String(t).trim())
    .filter((t) => t !== '')
    .join('\n');
  if (text === '') return null;
  const font = componentFont(
    base,
    [
      { font: which === 'title' ? { weight: 'bold' } : undefined, color: undefined },
      { font: d.font, color: d.color },
      { font: o.font, color: o.color },
    ],
    path,
    diagnostics,
  );
  const position = str(o.position) ?? 'top';
  if (position !== 'top') {
    diagnostics.report(
      'APPROXIMATED_LAYOUT',
      'approximated',
      `${path}.position`,
      `Excel draws the ${which} above the chart; position "${position}" is not reproduced.`,
      { severity: 'info' },
    );
  }
  const align = o.align === 'start' ? 'left' : o.align === 'end' ? 'right' : 'center';
  return { text, font, align, verticalAlign: 'top' };
}

/** `options.plugins.legend` → LegendModel (`enabled` still needs a series with a legend item). */
export function extractLegend(view: ChartJsView, base: Font, diagnostics: DiagnosticCollector): LegendModel {
  const path = 'options.plugins.legend';
  const o = rec(get(view.options, 'plugins', 'legend')) ?? {};
  const d = rec(get(view.defaults, 'plugins', 'legend')) ?? {};
  const labels = rec(o.labels) ?? {};
  const dLabels = rec(d.labels) ?? {};
  const rawPosition = pick(o.position, d.position, 'top');
  let position: LegendPosition = 'top';
  let overlay = false;
  if (rawPosition === 'top' || rawPosition === 'bottom' || rawPosition === 'left' || rawPosition === 'right') {
    position = rawPosition;
  } else {
    position = 'topRight';
    overlay = rawPosition === 'chartArea';
    diagnostics.report(
      'APPROXIMATED_LEGEND_POSITION',
      'approximated',
      `${path}.position`,
      `Legend position ${JSON.stringify(rawPosition)} is approximated by a top-right legend.`,
      { severity: 'info' },
    );
  }
  const align = pick(o.align, d.align, 'center');
  if (align !== 'center') {
    diagnostics.report(
      'APPROXIMATED_LEGEND_POSITION',
      'approximated',
      `${path}.align`,
      `Excel centers legends on their side; align "${String(align)}" is not reproduced.`,
      { severity: 'info' },
    );
  }
  for (const key of ['filter', 'generateLabels', 'sort'] as const) {
    if (typeof labels[key] === 'function') {
      diagnostics.report(
        'UNSUPPORTED_FORMATTER',
        'unsupported',
        `${path}.labels.${key}`,
        `Legend label callbacks (${key}) are not executed; Excel lists every exported series.`,
        { severity: 'info' },
      );
    }
  }
  if (get(o, 'title', 'display') === true) {
    diagnostics.report('UNSUPPORTED_STYLE', 'unsupported', `${path}.title`, 'Excel legends have no title.', {
      severity: 'info',
    });
  }
  const font = componentFont(
    base,
    [
      { font: dLabels.font, color: dLabels.color },
      { font: labels.font, color: labels.color },
    ],
    `${path}.labels`,
    diagnostics,
  );
  return {
    enabled: pick(o.display, d.display, true) !== false,
    position,
    layout: position === 'left' || position === 'right' ? 'vertical' : 'horizontal',
    font,
    background: null,
    border: null,
    overlay,
    reversed: o.reverse === true,
  };
}

/** First non-transparent computed background of the canvas or its ancestors (real browsers only). */
function computedCanvasBackground(canvas: unknown): Color | null {
  if (!isRealBrowser() || typeof getComputedStyle !== 'function') return null;
  let el = canvas as Element | null;
  for (let depth = 0; el && depth < 32; depth++) {
    if (typeof Element !== 'undefined' && !(el instanceof Element)) return null;
    const c = parseColor(getComputedStyle(el).backgroundColor);
    if (c && c.a > 0) return c;
    el = el.parentElement;
  }
  return null;
}

/**
 * Chart.js has no chart background option: a canvas is transparent and shows the page. The
 * documented `customCanvasBackgroundColor` plugin option wins, then (in a real browser) the
 * canvas's computed background, else white.
 */
export function extractBackground(view: ChartJsView, diagnostics: DiagnosticCollector): Fill {
  const custom = get(view.options, 'plugins', 'customCanvasBackgroundColor', 'color');
  if (custom !== undefined) {
    const c = colorOf(custom, 'options.plugins.customCanvasBackgroundColor.color', diagnostics);
    if (c) return { type: 'solid', color: c };
  }
  const computed = view.live ? computedCanvasBackground(view.live.canvas) : null;
  if (computed) return { type: 'solid', color: computed };
  return { type: 'solid', color: { r: 255, g: 255, b: 255, a: 1, source: '#ffffff' } };
}
