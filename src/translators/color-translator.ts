/**
 * Translates Highcharts color options (strings, gradient objects) into neutral fills/colors.
 */

import type { Color, Fill } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import { HIGHCHARTS_DEFAULT_PALETTE, parseColor, type CssVariableResolver } from '../utils/colors';

/** The Highcharts gradient color option shape. */
export interface GradientLike {
  linearGradient?: { x1: number; y1: number; x2: number; y2: number } | [number, number, number, number];
  radialGradient?: { cx: number; cy: number; r: number };
  stops: Array<[number, string]>;
}

/**
 * Converts a Highcharts color option into a Fill.
 *
 * `property` (optional, default `"color"`) is the option path used in diagnostics.
 */
export function toFill(
  input: unknown,
  resolveVariable?: CssVariableResolver,
  property = 'color',
): { fill: Fill | null; diagnostic?: Diagnostic } {
  if (input === null || input === undefined) return { fill: null };

  if (typeof input === 'string') {
    const trimmed = input.trim();
    if (trimmed === '' || trimmed.toLowerCase() === 'none') return { fill: null };
    const color = parseColor(input, resolveVariable);
    if (color) return { fill: { type: 'solid', color } };
    return {
      fill: null,
      diagnostic: createDiagnostic(
        'UNRESOLVED_COLOR',
        'approximated',
        property,
        `Could not resolve color "${input}"; Excel's automatic color is used.`,
        {
          details: { value: input },
        },
      ),
    };
  }

  if (isColorObject(input)) {
    return { fill: { type: 'solid', color: { r: input.r, g: input.g, b: input.b, a: input.a } } };
  }

  if (typeof input === 'object' && Array.isArray((input as { stops?: unknown }).stops)) {
    return gradientToFill(input as GradientLike, resolveVariable, property);
  }

  return {
    fill: null,
    diagnostic: createDiagnostic(
      'UNRESOLVED_COLOR',
      'unsupported',
      property,
      'Color option is neither a color string nor a gradient (patterns are not supported).',
      {
        details: { valueType: typeof input },
      },
    ),
  };
}

function isColorObject(v: unknown): v is Color {
  if (typeof v !== 'object' || v === null) return false;
  const o = v as Record<string, unknown>;
  return ['r', 'g', 'b', 'a'].every((k) => typeof o[k] === 'number' && Number.isFinite(o[k]));
}

function gradientToFill(
  g: GradientLike,
  resolveVariable: CssVariableResolver | undefined,
  property: string,
): { fill: Fill | null; diagnostic?: Diagnostic } {
  const stops: Array<{ offset: number; color: Color }> = [];
  const unresolved: string[] = [];
  for (const stop of g.stops) {
    if (!Array.isArray(stop)) continue;
    const [offset, colorStr] = stop as [unknown, unknown];
    const color = parseColor(colorStr, resolveVariable);
    if (!color || typeof offset !== 'number' || !Number.isFinite(offset)) {
      unresolved.push(String(colorStr));
      continue;
    }
    stops.push({ offset: Math.min(1, Math.max(0, offset)), color });
  }
  stops.sort((a, b) => a.offset - b.offset);

  if (stops.length === 0) {
    return {
      fill: null,
      diagnostic: createDiagnostic(
        'UNRESOLVED_COLOR',
        'approximated',
        `${property}.stops`,
        'No gradient stop color could be resolved.',
        {
          details: { stops: unresolved },
        },
      ),
    };
  }

  if (g.radialGradient && !g.linearGradient) {
    return {
      fill: { type: 'gradient', stops, angle: 90 },
      diagnostic: createDiagnostic(
        'UNSUPPORTED_GRADIENT',
        'approximated',
        `${property}.radialGradient`,
        'Radial gradients are approximated as a top-to-bottom linear gradient.',
      ),
    };
  }

  const angle = linearAngle(g.linearGradient);
  const fill: Fill = { type: 'gradient', stops, angle };
  if (unresolved.length > 0) {
    return {
      fill,
      diagnostic: createDiagnostic(
        'UNRESOLVED_COLOR',
        'approximated',
        `${property}.stops`,
        'Some gradient stop colors could not be resolved and were dropped.',
        {
          details: { stops: unresolved },
        },
      ),
    };
  }
  return { fill };
}

/** Angle in degrees of the gradient vector: 0 = left→right, 90 = top→bottom. */
function linearAngle(lg: GradientLike['linearGradient']): number {
  if (!lg) return 90;
  const [x1, y1, x2, y2] = Array.isArray(lg) ? lg : [lg.x1, lg.y1, lg.x2, lg.y2];
  const dx = num(x2) - num(x1);
  const dy = num(y2) - num(y1);
  if (dx === 0 && dy === 0) return 0;
  const deg = (Math.atan2(dy, dx) * 180) / Math.PI;
  return Math.round((((deg % 360) + 360) % 360) * 1000) / 1000;
}

function num(v: unknown): number {
  return typeof v === 'number' && Number.isFinite(v) ? v : 0;
}

/** Solid → its color; gradient → first stop color; none/null → null. */
export function fillToSolidColor(fill: Fill | null): Color | null {
  if (!fill) return null;
  if (fill.type === 'solid') return fill.color;
  if (fill.type === 'gradient') return fill.stops[0]?.color ?? null;
  return null;
}

/**
 * The effective series color: the explicit option when it resolves, else `palette[index % length]`.
 * An empty palette falls back to the Highcharts default palette.
 */
export function resolveSeriesColor(
  explicit: unknown,
  index: number,
  palette: readonly Color[],
  resolveVariable?: CssVariableResolver,
): Color | null {
  if (explicit !== null && explicit !== undefined) {
    const color = fillToSolidColor(toFill(explicit, resolveVariable).fill);
    if (color) return color;
  }
  const i = Number.isInteger(index) && index >= 0 ? index : 0;
  if (palette.length > 0) return palette[i % palette.length] ?? null;
  return parseColor(HIGHCHARTS_DEFAULT_PALETTE[i % HIGHCHARTS_DEFAULT_PALETTE.length], resolveVariable);
}
