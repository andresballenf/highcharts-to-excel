/**
 * Line dash, marker and stroke translation.
 */

import type { DashStyle, MarkerSymbol, Stroke } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import type { CssVariableResolver } from '../utils/colors';
import { clamp } from '../utils/units';
import { fillToSolidColor, toFill } from './color-translator';

const DASH_STYLES: ReadonlySet<DashStyle> = new Set<DashStyle>([
  'solid',
  'dash',
  'dot',
  'dashdot',
  'longdash',
  'longdashdot',
  'longdashdotdot',
  'shortdash',
  'shortdot',
  'shortdashdot',
  'shortdashdotdot',
]);

/** Highcharts dashStyle name (case-insensitive) → neutral DashStyle; unknown → 'solid'. */
export function dashStyleFromHighcharts(dashStyle: unknown): DashStyle {
  if (typeof dashStyle !== 'string') return 'solid';
  const key = dashStyle.trim().toLowerCase() as DashStyle;
  return DASH_STYLES.has(key) ? key : 'solid';
}

const DASH_TO_OOXML: Readonly<Record<DashStyle, string>> = {
  solid: 'solid',
  dash: 'dash',
  dot: 'dot',
  dashdot: 'dashDot',
  longdash: 'lgDash',
  longdashdot: 'lgDashDot',
  longdashdotdot: 'lgDashDotDot',
  shortdash: 'sysDash',
  shortdot: 'sysDot',
  shortdashdot: 'sysDashDot',
  shortdashdotdot: 'sysDashDotDot',
};

/** DashStyle → OOXML `a:prstDash/@val`. */
export function dashStyleToOoxml(d: DashStyle): string {
  return DASH_TO_OOXML[d] ?? 'solid';
}

const MARKER_SYMBOLS: ReadonlySet<string> = new Set(['circle', 'square', 'diamond', 'triangle', 'triangle-down']);

/** Highcharts marker symbol → MarkerSymbol. `url(...)` and custom symbols → 'other'; undefined → 'circle'. */
export function markerSymbolFromHighcharts(symbol: unknown): MarkerSymbol {
  if (typeof symbol !== 'string') return 'circle';
  const s = symbol.trim().toLowerCase();
  if (s === '') return 'circle';
  if (MARKER_SYMBOLS.has(s)) return s as MarkerSymbol;
  return 'other';
}

/**
 * MarkerSymbol → OOXML `c:symbol/@val`. `property` (optional, default `"marker.symbol"`) is the
 * option path used in diagnostics.
 */
export function markerSymbolToOoxml(
  symbol: MarkerSymbol,
  property = 'marker.symbol',
): { symbol: 'circle' | 'square' | 'diamond' | 'triangle' | 'none' | 'auto'; diagnostic?: Diagnostic } {
  switch (symbol) {
    case 'circle':
    case 'square':
    case 'diamond':
    case 'triangle':
    case 'none':
      return { symbol };
    case 'triangle-down':
      return {
        symbol: 'triangle',
        diagnostic: createDiagnostic('APPROXIMATED_MARKER', 'approximated', property, 'Excel has no downward triangle marker; an upward triangle is used.', {
          details: { symbol },
        }),
      };
    default:
      return {
        symbol: 'circle',
        diagnostic: createDiagnostic('APPROXIMATED_MARKER', 'approximated', property, 'Custom/image marker symbols are replaced by circles.', {
          details: { symbol },
        }),
      };
  }
}

/** Marker radius (CSS px) → OOXML `c:size` (points, 2..72). */
export function markerRadiusToOoxmlSize(radiusPx: number): number {
  if (!Number.isFinite(radiusPx)) return 5;
  return clamp(Math.round(radiusPx * 2 * 0.75), 2, 72);
}

/**
 * Builds a Stroke from Highcharts `{ color, width, dashStyle }`-like options; missing or invalid
 * values come from `fallback` (then `null` color, width 1, solid).
 */
export function strokeFromOptions(
  opts: { color?: unknown; width?: unknown; dashStyle?: unknown },
  fallback: Partial<Stroke>,
  resolveVariable?: CssVariableResolver,
): Stroke {
  const color =
    opts.color !== undefined && opts.color !== null ? fillToSolidColor(toFill(opts.color, resolveVariable).fill) : null;
  let width: number | undefined;
  if (typeof opts.width === 'number' && Number.isFinite(opts.width) && opts.width >= 0) width = opts.width;
  else if (typeof opts.width === 'string' && /^\s*\d*\.?\d+\s*(px)?\s*$/i.test(opts.width)) width = parseFloat(opts.width);
  return {
    color: color ?? fallback.color ?? null,
    width: width ?? fallback.width ?? 1,
    dash: opts.dashStyle !== undefined && opts.dashStyle !== null ? dashStyleFromHighcharts(opts.dashStyle) : (fallback.dash ?? 'solid'),
  };
}
