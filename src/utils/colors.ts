/**
 * CSS color parsing and color math. Pure functions, browser-safe.
 */

import type { Color } from '../types/chart-model';

/** Highcharts 11-13 default series palette. */
export const HIGHCHARTS_DEFAULT_PALETTE: readonly string[] = Object.freeze([
  '#2caffe',
  '#544fc5',
  '#00e272',
  '#fe6a35',
  '#6b8abc',
  '#d568fb',
  '#2ee0ca',
  '#fa4b42',
  '#feb56a',
  '#91e8e1',
]);

/** Highcharts 13 CSS variable defaults (light scheme). */
export const HIGHCHARTS_CSS_VARIABLE_DEFAULTS: Readonly<Record<string, string>> = Object.freeze({
  ...Object.fromEntries(HIGHCHARTS_DEFAULT_PALETTE.map((c, i) => [`--highcharts-color-${i}`, c])),
  '--highcharts-background-color': '#ffffff',
  '--highcharts-neutral-color-100': '#000000',
  '--highcharts-neutral-color-80': '#333333',
  '--highcharts-neutral-color-60': '#666666',
  '--highcharts-neutral-color-40': '#999999',
  '--highcharts-neutral-color-20': '#cccccc',
  '--highcharts-neutral-color-10': '#e6e6e6',
  '--highcharts-neutral-color-5': '#f2f2f2',
  '--highcharts-neutral-color-3': '#f7f7f7',
  '--highcharts-neutral-color-0': '#ffffff',
  '--highcharts-highlight-color-100': '#0022ff',
  '--highcharts-highlight-color-80': '#334eff',
  '--highcharts-highlight-color-60': '#667aff',
  '--highcharts-highlight-color-20': '#ccd3ff',
  '--highcharts-highlight-color-10': '#e6e9ff',
});

/** Resolves a CSS custom property (e.g. via getComputedStyle). Return null/undefined when unknown. */
export type CssVariableResolver = (variableName: string) => string | null | undefined;

/** The 148 CSS Color Module Level 4 named colors (excluding `transparent`). */
const NAMED_COLORS: Readonly<Record<string, string>> = {
  aliceblue: 'f0f8ff',
  antiquewhite: 'faebd7',
  aqua: '00ffff',
  aquamarine: '7fffd4',
  azure: 'f0ffff',
  beige: 'f5f5dc',
  bisque: 'ffe4c4',
  black: '000000',
  blanchedalmond: 'ffebcd',
  blue: '0000ff',
  blueviolet: '8a2be2',
  brown: 'a52a2a',
  burlywood: 'deb887',
  cadetblue: '5f9ea0',
  chartreuse: '7fff00',
  chocolate: 'd2691e',
  coral: 'ff7f50',
  cornflowerblue: '6495ed',
  cornsilk: 'fff8dc',
  crimson: 'dc143c',
  cyan: '00ffff',
  darkblue: '00008b',
  darkcyan: '008b8b',
  darkgoldenrod: 'b8860b',
  darkgray: 'a9a9a9',
  darkgreen: '006400',
  darkgrey: 'a9a9a9',
  darkkhaki: 'bdb76b',
  darkmagenta: '8b008b',
  darkolivegreen: '556b2f',
  darkorange: 'ff8c00',
  darkorchid: '9932cc',
  darkred: '8b0000',
  darksalmon: 'e9967a',
  darkseagreen: '8fbc8f',
  darkslateblue: '483d8b',
  darkslategray: '2f4f4f',
  darkslategrey: '2f4f4f',
  darkturquoise: '00ced1',
  darkviolet: '9400d3',
  deeppink: 'ff1493',
  deepskyblue: '00bfff',
  dimgray: '696969',
  dimgrey: '696969',
  dodgerblue: '1e90ff',
  firebrick: 'b22222',
  floralwhite: 'fffaf0',
  forestgreen: '228b22',
  fuchsia: 'ff00ff',
  gainsboro: 'dcdcdc',
  ghostwhite: 'f8f8ff',
  gold: 'ffd700',
  goldenrod: 'daa520',
  gray: '808080',
  green: '008000',
  greenyellow: 'adff2f',
  grey: '808080',
  honeydew: 'f0fff0',
  hotpink: 'ff69b4',
  indianred: 'cd5c5c',
  indigo: '4b0082',
  ivory: 'fffff0',
  khaki: 'f0e68c',
  lavender: 'e6e6fa',
  lavenderblush: 'fff0f5',
  lawngreen: '7cfc00',
  lemonchiffon: 'fffacd',
  lightblue: 'add8e6',
  lightcoral: 'f08080',
  lightcyan: 'e0ffff',
  lightgoldenrodyellow: 'fafad2',
  lightgray: 'd3d3d3',
  lightgreen: '90ee90',
  lightgrey: 'd3d3d3',
  lightpink: 'ffb6c1',
  lightsalmon: 'ffa07a',
  lightseagreen: '20b2aa',
  lightskyblue: '87cefa',
  lightslategray: '778899',
  lightslategrey: '778899',
  lightsteelblue: 'b0c4de',
  lightyellow: 'ffffe0',
  lime: '00ff00',
  limegreen: '32cd32',
  linen: 'faf0e6',
  magenta: 'ff00ff',
  maroon: '800000',
  mediumaquamarine: '66cdaa',
  mediumblue: '0000cd',
  mediumorchid: 'ba55d3',
  mediumpurple: '9370db',
  mediumseagreen: '3cb371',
  mediumslateblue: '7b68ee',
  mediumspringgreen: '00fa9a',
  mediumturquoise: '48d1cc',
  mediumvioletred: 'c71585',
  midnightblue: '191970',
  mintcream: 'f5fffa',
  mistyrose: 'ffe4e1',
  moccasin: 'ffe4b5',
  navajowhite: 'ffdead',
  navy: '000080',
  oldlace: 'fdf5e6',
  olive: '808000',
  olivedrab: '6b8e23',
  orange: 'ffa500',
  orangered: 'ff4500',
  orchid: 'da70d6',
  palegoldenrod: 'eee8aa',
  palegreen: '98fb98',
  paleturquoise: 'afeeee',
  palevioletred: 'db7093',
  papayawhip: 'ffefd5',
  peachpuff: 'ffdab9',
  peru: 'cd853f',
  pink: 'ffc0cb',
  plum: 'dda0dd',
  powderblue: 'b0e0e6',
  purple: '800080',
  rebeccapurple: '663399',
  red: 'ff0000',
  rosybrown: 'bc8f8f',
  royalblue: '4169e1',
  saddlebrown: '8b4513',
  salmon: 'fa8072',
  sandybrown: 'f4a460',
  seagreen: '2e8b57',
  seashell: 'fff5ee',
  sienna: 'a0522d',
  silver: 'c0c0c0',
  skyblue: '87ceeb',
  slateblue: '6a5acd',
  slategray: '708090',
  slategrey: '708090',
  snow: 'fffafa',
  springgreen: '00ff7f',
  steelblue: '4682b4',
  tan: 'd2b48c',
  teal: '008080',
  thistle: 'd8bfd8',
  tomato: 'ff6347',
  turquoise: '40e0d0',
  violet: 'ee82ee',
  wheat: 'f5deb3',
  white: 'ffffff',
  whitesmoke: 'f5f5f5',
  yellow: 'ffff00',
  yellowgreen: '9acd32',
};

const MAX_VAR_DEPTH = 4;

/**
 * Parses a CSS color string into RGBA. Returns null for `none`, unparseable input and non-strings.
 * Supports hex (3/4/6/8 digits), rgb()/rgba(), hsl()/hsla(), named colors, `transparent`,
 * `var(--name[, fallback])` and `light-dark(light, dark)` (the light value is used).
 */
export function parseColor(input: unknown, resolveVariable?: CssVariableResolver): Color | null {
  if (typeof input !== 'string') return null;
  const parsed = parseInternal(input, resolveVariable, 0);
  if (!parsed) return null;
  return { ...parsed, source: input };
}

interface Rgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

function parseInternal(raw: string, resolve: CssVariableResolver | undefined, depth: number): Rgba | null {
  if (depth > MAX_VAR_DEPTH) return null;
  const s = raw.trim().toLowerCase();
  if (s === '' || s === 'none') return null;
  if (s === 'transparent') return { r: 0, g: 0, b: 0, a: 0 };
  if (s.startsWith('#')) return parseHex(s.slice(1));
  const named = NAMED_COLORS[s];
  if (named !== undefined) return parseHex(named);

  const fn = /^([a-z-]+)\((.*)\)$/s.exec(s);
  if (!fn) return null;
  const name = fn[1] ?? '';
  // Use the original-case body for var() names (custom properties are case-sensitive).
  const trimmedRaw = raw.trim();
  const body = trimmedRaw.slice(trimmedRaw.indexOf('(') + 1, -1);
  switch (name) {
    case 'rgb':
    case 'rgba':
      return parseRgbBody(body);
    case 'hsl':
    case 'hsla':
      return parseHslBody(body);
    case 'var':
      return parseVarBody(body, resolve, depth);
    case 'light-dark': {
      const [light] = splitTopLevel(body, ',');
      return light === undefined ? null : parseInternal(light, resolve, depth + 1);
    }
    default:
      return null;
  }
}

function parseHex(hex: string): Rgba | null {
  if (!/^[0-9a-f]+$/i.test(hex)) return null;
  const n = hex.length;
  if (n === 3 || n === 4) {
    const ch = (i: number): number => parseInt((hex[i] ?? '0') + (hex[i] ?? '0'), 16);
    return { r: ch(0), g: ch(1), b: ch(2), a: n === 4 ? roundAlpha(ch(3) / 255) : 1 };
  }
  if (n === 6 || n === 8) {
    const ch = (i: number): number => parseInt(hex.slice(i * 2, i * 2 + 2), 16);
    return { r: ch(0), g: ch(1), b: ch(2), a: n === 8 ? roundAlpha(ch(3) / 255) : 1 };
  }
  return null;
}

/** Splits `a, b / c` style function arguments into channel tokens and an optional alpha token. */
function splitArgs(body: string): { channels: string[]; alpha: string | undefined } | null {
  const trimmed = body.trim();
  if (trimmed.includes(',')) {
    const parts = trimmed.split(',').map((p) => p.trim());
    if (parts.some((p) => p === '')) return null;
    if (parts.length === 3) return { channels: parts, alpha: undefined };
    if (parts.length === 4) return { channels: parts.slice(0, 3), alpha: parts[3] };
    return null;
  }
  const [main = '', alpha, extra] = trimmed.split('/').map((p) => p.trim());
  if (extra !== undefined) return null;
  const channels = main.split(/\s+/).filter(Boolean);
  if (channels.length !== 3) return null;
  if (alpha !== undefined && alpha === '') return null;
  return { channels, alpha };
}

const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;

function parseNumber(token: string): number | null {
  return NUMBER_RE.test(token) ? Number(token) : null;
}

function parseAlpha(token: string | undefined): number | null {
  if (token === undefined) return 1;
  if (token.endsWith('%')) {
    const n = parseNumber(token.slice(0, -1));
    return n === null ? null : roundAlpha(clamp01(n / 100));
  }
  const n = parseNumber(token);
  return n === null ? null : roundAlpha(clamp01(n));
}

function parseRgbBody(body: string): Rgba | null {
  const args = splitArgs(body);
  if (!args) return null;
  const channels: number[] = [];
  for (const t of args.channels) {
    let v: number | null;
    if (t.endsWith('%')) {
      const n = parseNumber(t.slice(0, -1));
      v = n === null ? null : (n / 100) * 255;
    } else if (t === 'none') {
      v = 0;
    } else {
      v = parseNumber(t);
    }
    if (v === null) return null;
    channels.push(Math.round(Math.min(255, Math.max(0, v))));
  }
  const a = parseAlpha(args.alpha);
  if (a === null) return null;
  return { r: channels[0] ?? 0, g: channels[1] ?? 0, b: channels[2] ?? 0, a };
}

function parseHue(token: string): number | null {
  const m = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)(deg|rad|grad|turn)?$/i.exec(token);
  if (!m) return null;
  const n = Number(m[1]);
  switch ((m[2] ?? 'deg').toLowerCase()) {
    case 'rad':
      return (n * 180) / Math.PI;
    case 'grad':
      return n * 0.9;
    case 'turn':
      return n * 360;
    default:
      return n;
  }
}

function parsePercentish(token: string): number | null {
  const n = parseNumber(token.endsWith('%') ? token.slice(0, -1) : token);
  return n === null ? null : clamp01(n / 100);
}

function parseHslBody(body: string): Rgba | null {
  const args = splitArgs(body);
  if (!args) return null;
  const [ht = '', st = '', lt = ''] = args.channels;
  const h = parseHue(ht);
  const s = parsePercentish(st);
  const l = parsePercentish(lt);
  const a = parseAlpha(args.alpha);
  if (h === null || s === null || l === null || a === null) return null;
  const hue = (((h % 360) + 360) % 360) / 360;
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const conv = (t0: number): number => {
    let t = t0;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return {
    r: Math.round(conv(hue + 1 / 3) * 255),
    g: Math.round(conv(hue) * 255),
    b: Math.round(conv(hue - 1 / 3) * 255),
    a,
  };
}

function parseVarBody(body: string, resolve: CssVariableResolver | undefined, depth: number): Rgba | null {
  const commaIdx = findTopLevel(body, ',');
  const name = (commaIdx === -1 ? body : body.slice(0, commaIdx)).trim();
  const fallback = commaIdx === -1 ? undefined : body.slice(commaIdx + 1).trim();
  if (!name.startsWith('--')) return null;
  const resolved = resolve?.(name);
  if (typeof resolved === 'string' && resolved.trim() !== '') {
    const c = parseInternal(resolved, resolve, depth + 1);
    if (c) return c;
  }
  const builtin = HIGHCHARTS_CSS_VARIABLE_DEFAULTS[name];
  if (builtin !== undefined) {
    const c = parseInternal(builtin, resolve, depth + 1);
    if (c) return c;
  }
  if (fallback !== undefined && fallback !== '') return parseInternal(fallback, resolve, depth + 1);
  return null;
}

function findTopLevel(s: string, ch: string): number {
  let level = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === '(') level++;
    else if (c === ')') level--;
    else if (c === ch && level === 0) return i;
  }
  return -1;
}

function splitTopLevel(s: string, ch: string): string[] {
  const out: string[] = [];
  let rest = s;
  for (;;) {
    const i = findTopLevel(rest, ch);
    if (i === -1) {
      out.push(rest.trim());
      return out;
    }
    out.push(rest.slice(0, i).trim());
    rest = rest.slice(i + 1);
  }
}

function clamp01(n: number): number {
  return Math.min(1, Math.max(0, n));
}

function roundAlpha(a: number): number {
  return Math.round(a * 10000) / 10000;
}

function hex2(n: number): string {
  return Math.round(Math.min(255, Math.max(0, n)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase();
}

/** `RRGGBB`, uppercase, no leading `#`, alpha dropped. */
export function colorToHex(c: Color): string {
  return hex2(c.r) + hex2(c.g) + hex2(c.b);
}

/** `rgba(r, g, b, a)`. */
export function colorToCss(c: Color): string {
  return `rgba(${Math.round(c.r)}, ${Math.round(c.g)}, ${Math.round(c.b)}, ${roundAlpha(c.a)})`;
}

/** Composites `c` over `background` (Porter-Duff "over"). The result is opaque when the background is. */
export function flattenAlpha(c: Color, background: Color): Color {
  const a = clamp01(c.a);
  const ba = clamp01(background.a);
  const outA = a + ba * (1 - a);
  if (outA === 0) return { r: 0, g: 0, b: 0, a: 0 };
  const mix = (fg: number, bg: number): number => Math.round((fg * a + bg * ba * (1 - a)) / outA);
  const out: Color = {
    r: mix(c.r, background.r),
    g: mix(c.g, background.g),
    b: mix(c.b, background.b),
    a: roundAlpha(outA),
  };
  if (c.source !== undefined) out.source = c.source;
  return out;
}

/** Returns a copy of `c` with alpha replaced (clamped to 0-1). */
export function withAlpha(c: Color, a: number): Color {
  return { ...c, a: roundAlpha(clamp01(Number.isFinite(a) ? a : 1)) };
}

/** Channel-wise equality; `source` is ignored and alpha compared with a small tolerance. */
export function isColorEqual(a: Color, b: Color): boolean {
  return (
    Math.round(a.r) === Math.round(b.r) &&
    Math.round(a.g) === Math.round(b.g) &&
    Math.round(a.b) === Math.round(b.b) &&
    Math.abs(a.a - b.a) < 1e-3
  );
}
