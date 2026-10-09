/**
 * Renders Highcharts charts into jsdom containers for tests. Works with any Highcharts namespace
 * (v12 or v13) passed in; never mutates the given options (they are deep-cloned).
 */

import type { Chart, Options } from 'highcharts';

export interface HighchartsLike {
  chart: (renderTo: HTMLElement, options: Options) => unknown;
  stockChart?: (renderTo: HTMLElement, options: Options) => unknown;
  version?: string;
}

const rendered: Array<{ chart: Chart; container: HTMLElement }> = [];
/** Charts this helper has already destroyed (a test may destroy one itself via `destroyChart`). */
const destroyed = new WeakSet<Chart>();
let counter = 0;

/** Deep clone that keeps functions (structuredClone rejects them). */
export function cloneOptions<T>(value: T): T {
  if (Array.isArray(value)) return value.map((v: unknown) => cloneOptions(v)) as T;
  if (value !== null && typeof value === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) out[k] = cloneOptions(v);
    return out as T;
  }
  return value;
}

export function renderChart(
  Highcharts: HighchartsLike,
  options: Options,
  ctor: 'chart' | 'stockChart' = 'chart',
): Chart {
  const container = document.createElement('div');
  container.id = `test-chart-${++counter}`;
  document.body.appendChild(container);
  const o = cloneOptions(options);
  const merged: Options = {
    ...o,
    accessibility: { enabled: false },
    exporting: { enabled: true, ...(o.exporting ?? {}) },
    chart: { ...(o.chart ?? {}), animation: false },
    plotOptions: { ...(o.plotOptions ?? {}), series: { ...(o.plotOptions?.series ?? {}), animation: false } },
  };
  const factory = ctor === 'stockChart' ? Highcharts.stockChart : Highcharts.chart;
  if (!factory) throw new Error(`Highcharts.${ctor} is not available (module not loaded?)`);
  const chart = factory(container, merged) as Chart;
  rendered.push({ chart, container });
  return chart;
}

/** Destroys one chart and marks it so `destroyAll` skips it later. */
export function destroyChart(chart: Chart): void {
  if (destroyed.has(chart)) return;
  chart.destroy();
  destroyed.add(chart);
}

/**
 * Destroys every chart rendered through `renderChart` and removes its container. Charts that are
 * already gone (destroyed through `destroyChart`, or by Highcharts itself, which deletes
 * `renderTo`/`container`) are skipped; any other error from `chart.destroy()` propagates.
 */
export function destroyAll(): void {
  for (const { chart, container } of rendered.splice(0)) {
    const live = chart as Chart & { renderTo?: unknown; container?: unknown };
    if (!destroyed.has(chart) && live.renderTo && live.container) destroyChart(chart);
    container.remove();
  }
}
