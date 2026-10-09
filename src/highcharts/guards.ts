/**
 * Runtime detection helpers. Pure duck-typing: no Highcharts import at runtime.
 */

import type { HcChartLike } from './types';

function isObject(x: unknown): x is Record<string, unknown> {
  return typeof x === 'object' && x !== null;
}

/** True when `x` looks like a rendered Highcharts chart instance. */
export function isHighchartsChart(x: unknown): x is HcChartLike {
  if (!isObject(x)) return false;
  return (
    Array.isArray(x.series) &&
    Array.isArray(x.xAxis) &&
    Array.isArray(x.yAxis) &&
    isObject(x.options) &&
    (isObject(x.renderTo) || isObject(x.container))
  );
}

/**
 * The Highcharts version string. Highcharts does not expose its version on chart instances, so
 * this reads it from the namespace passed in, or from a `Highcharts` global (UMD builds), else null.
 */
export function getHighchartsVersion(_chart: unknown, highcharts?: unknown): string | null {
  const fromArg = isObject(highcharts) ? highcharts.version : undefined;
  if (typeof fromArg === 'string' && fromArg !== '') return fromArg;
  const global = (globalThis as { Highcharts?: unknown }).Highcharts;
  const fromGlobal = isObject(global) ? global.version : undefined;
  return typeof fromGlobal === 'string' && fromGlobal !== '' ? fromGlobal : null;
}

/**
 * True in a real browser. jsdom provides `window` and `navigator` but its computed styles are
 * useless (no stylesheet cascade for SVG presentation), so it counts as headless.
 */
export function isRealBrowser(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof navigator !== 'undefined' &&
    typeof navigator.userAgent === 'string' &&
    !/jsdom/i.test(navigator.userAgent)
  );
}

// ---------------------------------------------------------------------------
// Defensive accessors for untrusted option bags.
// ---------------------------------------------------------------------------

export type Rec = Record<string, unknown>;

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

export function bool(x: unknown): boolean | undefined {
  return typeof x === 'boolean' ? x : undefined;
}

export function arr(x: unknown): unknown[] | undefined {
  return Array.isArray(x) ? x : undefined;
}

/** First element when `x` is an array (Highcharts allows `dataLabels: [{...}]`), else `x` itself as an object. */
export function firstRec(x: unknown): Rec | undefined {
  return Array.isArray(x) ? rec(x[0]) : rec(x);
}

/** Reads a nested path such as `get(o, 'title', 'style')`. */
export function get(x: unknown, ...path: string[]): unknown {
  let cur: unknown = x;
  for (const key of path) {
    const r = rec(cur);
    if (!r) return undefined;
    cur = r[key];
  }
  return cur;
}

/**
 * Deep merge of plain objects (later wins); arrays and other values are taken by reference.
 * Never mutates its inputs.
 */
export function deepMerge(...sources: unknown[]): Rec {
  const out: Rec = {};
  for (const src of sources) {
    const r = rec(src);
    if (!r) continue;
    for (const [k, v] of Object.entries(r)) {
      const prev = rec(out[k]);
      const next = rec(v);
      out[k] = prev && next && !isWrapperLike(next) ? deepMerge(prev, next) : v;
    }
  }
  return out;
}

/** Objects with a prototype other than Object (class instances, DOM nodes) are not merged. */
function isWrapperLike(x: Rec): boolean {
  const proto = Object.getPrototypeOf(x) as unknown;
  return proto !== Object.prototype && proto !== null;
}

/**
 * Strips HTML tags and decodes the handful of entities Highcharts text commonly contains.
 * Only real tags (`<b>`, `</span>`, `<br/>`…) are removed, so literal comparisons such as
 * "Growth < 5% vs > 3% target" survive; `<br>` variants become line breaks.
 */
export function plainText(x: unknown): string | null {
  if (typeof x !== 'string' && typeof x !== 'number') return null;
  const s = String(x)
    .replace(/<br\b[^>]*>/gi, '\n')
    .replace(/<\/?[a-zA-Z][^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .trim();
  return s === '' ? null : s;
}
