/**
 * Translates CSS text styles (Highcharts `style` options) into neutral fonts and OOXML font attributes.
 */

import type { Font } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import { colorToHex, parseColor, type CssVariableResolver } from '../utils/colors';
import { clamp, roundTo } from '../utils/units';

export interface CssStyleLike {
  fontFamily?: unknown;
  fontSize?: unknown;
  fontWeight?: unknown;
  fontStyle?: unknown;
  color?: unknown;
  font?: unknown;
}

/** Highcharts' classic default body text style. */
/**
 * Highcharts 11-13 default `chart.style`: a system font stack at 1rem (16px), color #333333.
 * Component defaults are relative: title 1.2em bold, subtitle/axis labels/legend 0.8em, data labels 0.7em bold.
 */
export const DEFAULT_HIGHCHARTS_FONT: Font = {
  family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
  size: 16,
  bold: false,
  italic: false,
  color: { r: 0x33, g: 0x33, b: 0x33, a: 1, source: '#333333' },
};

const DEFAULT_ROOT_SIZE_PX = 16;

const ABSOLUTE_SIZE_KEYWORDS: Readonly<Record<string, number>> = {
  'xx-small': 9,
  'x-small': 10,
  small: 13,
  medium: 16,
  large: 18,
  'x-large': 24,
  'xx-large': 32,
  'xxx-large': 48,
};

/**
 * Builds a Font from a CSS style object. Missing or unparseable fields inherit from `base`.
 * The `font` shorthand is applied first; longhand properties override it.
 */
export function toFont(
  style: CssStyleLike | undefined | null,
  base: Font | null,
  resolveVariable?: CssVariableResolver,
): Font {
  const font: Font = base
    ? { family: base.family, size: base.size, bold: base.bold, italic: base.italic, color: base.color }
    : { family: null, size: null, bold: false, italic: false, color: null };
  if (!style || typeof style !== 'object') return font;
  const parentSize = base?.size ?? DEFAULT_ROOT_SIZE_PX;

  if (typeof style.font === 'string' && style.font.trim() !== '') {
    applyShorthand(style.font, font, parentSize);
  }

  if (typeof style.fontFamily === 'string' && style.fontFamily.trim() !== '') {
    font.family = style.fontFamily.trim();
  }
  const size = parseFontSize(style.fontSize, parentSize);
  if (size !== null) font.size = size;
  const bold = parseWeight(style.fontWeight);
  if (bold !== null) font.bold = bold;
  const italic = parseFontStyle(style.fontStyle);
  if (italic !== null) font.italic = italic;
  const color = parseColor(style.color, resolveVariable);
  if (color) font.color = color;
  return font;
}

/** Parses a CSS font-size into px. Returns null when unparseable. */
function parseFontSize(value: unknown, parentPx: number): number | null {
  if (typeof value === 'number') return Number.isFinite(value) && value > 0 ? roundTo(value, 3) : null;
  if (typeof value !== 'string') return null;
  const s = value.trim().toLowerCase();
  const keyword = ABSOLUTE_SIZE_KEYWORDS[s];
  if (keyword !== undefined) return keyword;
  if (s === 'smaller') return roundTo(parentPx / 1.2, 3);
  if (s === 'larger') return roundTo(parentPx * 1.2, 3);
  const m = /^(\d*\.?\d+)\s*(px|pt|em|rem|%)?$/.exec(s);
  if (!m) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n) || n <= 0) return null;
  let px: number;
  switch (m[2]) {
    case 'pt':
      px = n / 0.75;
      break;
    case 'em':
      px = n * parentPx;
      break;
    case 'rem':
      px = n * DEFAULT_ROOT_SIZE_PX;
      break;
    case '%':
      px = (n / 100) * parentPx;
      break;
    default:
      px = n;
  }
  return roundTo(px, 3);
}

function parseWeight(value: unknown): boolean | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value >= 600 : null;
  if (typeof value !== 'string') return null;
  const s = value.trim().toLowerCase();
  if (s === 'bold' || s === 'bolder') return true;
  if (s === 'normal' || s === 'lighter') return false;
  if (/^\d+$/.test(s)) return Number(s) >= 600;
  return null;
}

function parseFontStyle(value: unknown): boolean | null {
  if (typeof value !== 'string') return null;
  const s = value.trim().toLowerCase();
  if (s === 'italic' || s.startsWith('oblique')) return true;
  if (s === 'normal') return false;
  return null;
}

const SIZE_TOKEN = String.raw`(?:\d*\.?\d+(?:px|pt|em|rem|%)|xx-small|x-small|small|medium|large|x-large|xx-large|xxx-large|smaller|larger)`;

/** Best-effort parse of `[style] [variant] [weight] size[/line-height] family`. */
function applyShorthand(shorthand: string, font: Font, parentPx: number): void {
  const re = new RegExp(String.raw`(^|\s)(${SIZE_TOKEN})(?:\s*\/\s*\S+)?(?:\s+|$)`, 'i');
  const m = re.exec(shorthand);
  if (!m) return;
  const prefix = shorthand.slice(0, m.index).trim();
  const family = shorthand.slice(m.index + m[0].length).trim();
  for (const token of prefix.split(/\s+/).filter(Boolean)) {
    const italic = parseFontStyle(token);
    if (italic === true) font.italic = true;
    const bold = parseWeight(token);
    if (bold !== null) font.bold = bold;
  }
  const size = parseFontSize(m[2], parentPx);
  if (size !== null) font.size = size;
  if (family) font.family = family;
}

const SYSTEM_FAMILIES = new Set(['-apple-system', 'blinkmacsystemfont', 'system-ui', 'ui-sans-serif']);

const GENERIC_FAMILY_MAP: Readonly<Record<string, string>> = {
  'sans-serif': 'Arial',
  serif: 'Times New Roman',
  monospace: 'Consolas',
  'system-ui': 'Segoe UI',
  '-apple-system': 'Segoe UI',
  blinkmacsystemfont: 'Segoe UI',
  'ui-sans-serif': 'Segoe UI',
  'ui-serif': 'Times New Roman',
  'ui-monospace': 'Consolas',
  cursive: 'Comic Sans MS',
  fantasy: 'Impact',
};

/**
 * The first usable family of a CSS font-family list, unquoted. Leading system-font aliases
 * (`-apple-system`, `BlinkMacSystemFont`, `system-ui`) are skipped in favour of the next entry;
 * generic families map to common Office fonts with `generic: true`.
 */
export function primaryFontFamily(family: string | null): { name: string | null; generic: boolean } {
  if (typeof family !== 'string') return { name: null, generic: false };
  const entries = family
    .split(',')
    .map((f) =>
      f
        .trim()
        .replace(/^(['"])(.*)\1$/, '$2')
        .trim(),
    )
    .filter(Boolean);
  if (entries.length === 0) return { name: null, generic: false };
  const chosen = entries.find((e) => !SYSTEM_FAMILIES.has(e.toLowerCase())) ?? entries[0] ?? '';
  const mapped = GENERIC_FAMILY_MAP[chosen.toLowerCase()];
  if (mapped !== undefined) return { name: mapped, generic: true };
  return { name: chosen, generic: false };
}

/**
 * OOXML text run attributes for a Font. `sizeHundredthsPt` is the `sz` attribute value.
 * `property` (optional, default `"style.fontFamily"`) is the option path used in diagnostics.
 */
export function fontToOoxml(
  font: Font | null,
  property = 'style.fontFamily',
): {
  typeface: string | null;
  sizeHundredthsPt: number | null;
  bold: boolean;
  italic: boolean;
  colorHex: string | null;
  diagnostic?: Diagnostic;
} {
  if (!font) return { typeface: null, sizeHundredthsPt: null, bold: false, italic: false, colorHex: null };
  const { name, generic } = primaryFontFamily(font.family);
  const sizeHundredthsPt =
    font.size !== null && Number.isFinite(font.size) ? clamp(Math.round(font.size * 0.75 * 100), 100, 40000) : null;
  const result: ReturnType<typeof fontToOoxml> = {
    typeface: name,
    sizeHundredthsPt,
    bold: font.bold,
    italic: font.italic,
    colorHex: font.color ? colorToHex(font.color) : null,
  };
  if (generic && name !== null) {
    result.diagnostic = createDiagnostic(
      'APPROXIMATED_FONT',
      'approximated',
      property,
      `Generic/system font family "${font.family ?? ''}" is mapped to "${name}".`,
      { details: { family: font.family, typeface: name } },
    );
  }
  return result;
}
