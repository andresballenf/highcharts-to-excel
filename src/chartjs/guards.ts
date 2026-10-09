/**
 * Defensive accessors for Chart.js option bags, and input detection.
 *
 * Chart.js is never imported: a live chart is recognized structurally, so the library keeps no
 * runtime dependency on it.
 */

import type { ChartJsChartLike, ChartJsConfigLike, Rec } from './types';

/** `x` when it is a plain object (not an array), else undefined. */
export function rec(x: unknown): Rec | undefined {
  return typeof x === 'object' && x !== null && !Array.isArray(x) ? (x as Rec) : undefined;
}

/** Finite number or undefined. */
export function num(x: unknown): number | undefined {
  return typeof x === 'number' && Number.isFinite(x) ? x : undefined;
}

/** Non-empty string or undefined. */
export function str(x: unknown): string | undefined {
  return typeof x === 'string' && x !== '' ? x : undefined;
}

export function arr(x: unknown): unknown[] | undefined {
  return Array.isArray(x) ? x : undefined;
}

/** Reads a nested path such as `get(o, 'plugins', 'title')`. */
export function get(x: unknown, ...path: string[]): unknown {
  let cur: unknown = x;
  for (const key of path) {
    const r = rec(cur);
    if (!r) return undefined;
    cur = r[key];
  }
  return cur;
}

/** First defined (not undefined) value. */
export function pick(...values: unknown[]): unknown {
  for (const v of values) if (v !== undefined) return v;
  return undefined;
}

/**
 * Resolves a Chart.js object key, including the dotted paths Chart.js accepts in `parsing`
 * (`'data.value'`; a literal key containing a dot wins, as in Chart.js `resolveObjectKey`).
 */
export function resolveKey(obj: Rec, key: string): unknown {
  if (key in obj) return obj[key];
  if (!key.includes('.')) return undefined;
  return get(obj, ...key.split('.'));
}

/**
 * True in a real browser. jsdom provides `window` and `navigator` but no layout or computed style
 * cascade, so it counts as headless.
 */
export function isRealBrowser(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof navigator !== 'undefined' &&
    typeof navigator.userAgent === 'string' &&
    !/jsdom/i.test(navigator.userAgent)
  );
}

/** A live Chart.js 4 `Chart` instance (recognized by shape). */
export function isChartJsChart(x: unknown): x is ChartJsChartLike {
  const c = rec(x);
  if (!c) return false;
  return (
    typeof c.getDatasetMeta === 'function' &&
    typeof c.isDatasetVisible === 'function' &&
    rec(c.config) !== undefined &&
    rec(c.scales) !== undefined
  );
}

/** A plain Chart.js configuration object: `{ type, data: { datasets: [...] }, options }`. */
export function isChartJsConfig(x: unknown): x is ChartJsConfigLike {
  const c = rec(x);
  if (!c) return false;
  const data = rec(c.data);
  return data !== undefined && Array.isArray(data.datasets);
}

/** The Chart.js version of a live chart (`Chart.version`), else null. */
export function getChartJsVersion(chart: unknown): string | null {
  const ctor = rec(chart) ? (chart as { constructor?: unknown }).constructor : undefined;
  const v = (ctor as { version?: unknown } | undefined)?.version;
  return typeof v === 'string' && v !== '' ? v : null;
}

/** `Chart.defaults` of a live chart's constructor, when reachable. */
export function getChartJsDefaults(chart: unknown): Rec | undefined {
  const ctor = rec(chart) ? (chart as { constructor?: unknown }).constructor : undefined;
  return rec((ctor as { defaults?: unknown } | undefined)?.defaults);
}
