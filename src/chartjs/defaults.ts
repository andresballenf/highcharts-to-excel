/**
 * Chart.js 4 defaults the adapter needs when it reads a plain configuration (a live chart exposes
 * `Chart.defaults`, which wins when reachable). Values mirror chart.js 4.5.
 */

import type { Rec } from './types';

/** `Chart.defaults.font`. */
export const CHARTJS_DEFAULT_FONT = Object.freeze({
  family: "'Helvetica Neue', 'Helvetica', 'Arial', sans-serif",
  size: 12,
  style: 'normal',
  weight: null as string | number | null,
});

/** `Chart.defaults.color` (text). */
export const CHARTJS_DEFAULT_TEXT_COLOR = '#666';
/** `Chart.defaults.backgroundColor` / `borderColor` (element colors when nothing is set). */
export const CHARTJS_DEFAULT_ELEMENT_COLOR = 'rgba(0,0,0,0.1)';

/** Colors plugin (`plugins.colors`) palette, in Chart.js order. Background = border at 50% alpha. */
export const CHARTJS_COLORS_PLUGIN_BORDER: readonly string[] = Object.freeze([
  'rgb(54, 162, 235)',
  'rgb(255, 99, 132)',
  'rgb(255, 159, 64)',
  'rgb(255, 205, 86)',
  'rgb(75, 192, 192)',
  'rgb(153, 102, 255)',
  'rgb(201, 203, 207)',
]);
export const CHARTJS_COLORS_PLUGIN_BACKGROUND: readonly string[] = Object.freeze(
  CHARTJS_COLORS_PLUGIN_BORDER.map((c) => c.replace('rgb(', 'rgba(').replace(')', ', 0.5)')),
);

/** `Chart.defaults.elements.*` (colors fall back to the chart-level element color). */
export const CHARTJS_ELEMENT_DEFAULTS: Readonly<Record<'line' | 'point' | 'bar' | 'arc', Rec>> = Object.freeze({
  line: {
    borderWidth: 3,
    borderDash: [],
    fill: false,
    tension: 0,
    stepped: false,
    cubicInterpolationMode: 'default',
    spanGaps: false,
  },
  point: { radius: 3, pointStyle: 'circle', borderWidth: 1, rotation: 0 },
  bar: { borderWidth: 0, borderRadius: 0 },
  arc: { borderWidth: 2, borderColor: '#fff', offset: 0 },
});

/** Per-controller dataset defaults (`Chart.overrides[type]` / controller defaults). */
export const CHARTJS_CONTROLLER_DEFAULTS: Readonly<Record<string, Rec>> = Object.freeze({
  line: { showLine: true, spanGaps: false },
  scatter: { showLine: false, fill: false },
  bubble: {},
  bar: { categoryPercentage: 0.8, barPercentage: 0.9 },
  pie: { cutout: 0, rotation: 0, circumference: 360 },
  doughnut: { cutout: '50%', rotation: 0, circumference: 360 },
  radar: {},
  polarArea: {},
});

/** Default scales per chart type: `_index_` / `_value_` follow `indexAxis`; letters are fixed. */
export const CHARTJS_DEFAULT_SCALES: Readonly<Record<string, Record<string, Rec>>> = Object.freeze({
  line: { _index_: { type: 'category' }, _value_: { type: 'linear' } },
  bar: { _index_: { type: 'category' }, _value_: { type: 'linear', beginAtZero: true } },
  scatter: { x: { type: 'linear' }, y: { type: 'linear' } },
  bubble: { x: { type: 'linear' }, y: { type: 'linear' } },
  radar: { r: { type: 'radialLinear' } },
  polarArea: { r: { type: 'radialLinear' } },
});

/** Default `options.aspectRatio` (width / height). */
export function defaultAspectRatio(type: string): number {
  return type === 'pie' || type === 'doughnut' || type === 'radar' || type === 'polarArea' ? 1 : 2;
}
