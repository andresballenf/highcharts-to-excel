/**
 * CSS custom property / computed style access. Explicit overrides (`ThemeOverrides.cssVariables`)
 * are consulted first, headless and in a browser; otherwise only a real browser resolves variables
 * (headless callers get `undefined` and `parseColor` then applies HIGHCHARTS_CSS_VARIABLE_DEFAULTS).
 */

import type { CssVariableResolver } from '../utils/colors';
import { isRealBrowser } from './guards';
import type { HcChartLike } from './types';

function asElement(x: unknown): Element | undefined {
  return typeof Element !== 'undefined' && x instanceof Element ? x : undefined;
}

/** Non-empty string values of `overrides` keyed by custom property name (`--…` names only). */
function overrideTable(overrides: Readonly<Record<string, string>> | undefined): Map<string, string> {
  const table = new Map<string, string>();
  if (!overrides || typeof overrides !== 'object') return table;
  for (const [name, value] of Object.entries(overrides)) {
    if (name.startsWith('--') && typeof value === 'string' && value.trim() !== '') table.set(name, value.trim());
  }
  return table;
}

/**
 * Resolves `--highcharts-*` (and any other) custom properties: from `overrides` first, then (real
 * browser only) against the chart container.
 */
export function createCssVariableResolver(
  chart: HcChartLike | undefined,
  overrides?: Readonly<Record<string, string>>,
): CssVariableResolver {
  const table = overrideTable(overrides);
  const fromOverrides = (name: string): string | undefined => table.get(name);
  if (!chart || !isRealBrowser()) return fromOverrides;
  const container = asElement(chart.container) ?? asElement(chart.renderTo);
  if (!container) return fromOverrides;
  const cache = new Map<string, string | undefined>();
  return (name: string) => {
    const own = table.get(name);
    if (own !== undefined) return own;
    if (cache.has(name)) return cache.get(name);
    let value: string | undefined;
    try {
      const v = getComputedStyle(container).getPropertyValue(name).trim();
      value = v === '' ? undefined : v;
    } catch {
      value = undefined;
    }
    cache.set(name, value);
    return value;
  };
}

/**
 * Computed style values of `el` for the given CSS properties (kebab-case). Returns an empty
 * object outside a real browser: jsdom's computed `fill`/`stroke` are meaningless defaults.
 */
export function readEffectiveStyle(el: Element | undefined, props: string[]): Record<string, string> {
  const out: Record<string, string> = {};
  const element = asElement(el);
  if (!element || !isRealBrowser()) return out;
  try {
    const cs = getComputedStyle(element);
    for (const p of props) {
      const v = cs.getPropertyValue(p).trim();
      if (v !== '') out[p] = v;
    }
  } catch {
    // Detached or foreign elements: nothing to report.
  }
  return out;
}
