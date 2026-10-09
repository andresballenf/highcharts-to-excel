/**
 * CSS custom property / computed style access. Only active in a real browser; headless callers
 * get `undefined` and `parseColor` then applies HIGHCHARTS_CSS_VARIABLE_DEFAULTS.
 */

import type { CssVariableResolver } from '../utils/colors';
import { isRealBrowser } from './guards';
import type { HcChartLike } from './types';

function asElement(x: unknown): Element | undefined {
  return typeof Element !== 'undefined' && x instanceof Element ? x : undefined;
}

/** Resolves `--highcharts-*` (and any other) custom properties against the chart container. */
export function createCssVariableResolver(chart: HcChartLike | undefined): CssVariableResolver {
  if (!chart || !isRealBrowser()) return () => undefined;
  const container = asElement(chart.container) ?? asElement(chart.renderTo);
  if (!container) return () => undefined;
  const cache = new Map<string, string | undefined>();
  return (name: string) => {
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
