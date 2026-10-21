/**
 * Infers an Excel number format code from text Highcharts actually rendered for known values:
 * axis tick labels, data labels, or tooltip text produced by a formatter. This is how formatter
 * callbacks (which have no declarative form) get an Excel equivalent: the samples are matched to
 * their values, and the shared pattern (currency prefix, unit suffix, k/M/B scaling, percent,
 * decimals, thousands grouping) becomes the format code.
 *
 * Pure and Highcharts-free: the caller collects `{ value, text }` samples.
 */

import { excelFormatCodeProblem } from '../utils/format-code';
import { escapeAffix, stripHtml } from './number-format-translator';

export interface FormatSample {
  /** The numeric value the text was rendered for. */
  value: number;
  /** The rendered text (HTML allowed; tags are stripped). */
  text: string;
}

export interface InferOptions {
  /** Decimal separator in the rendered text. Default '.'. */
  decimalPoint?: string;
  /** Preferred thousands separator; ',', ' ', '.', "'" are also accepted (except the decimal point). */
  thousandsSep?: string;
  /** `lang.numericSymbols` (index i stands for 1000^(i+1)). Common k/K/M/MM/B/bn/G/T are always known. */
  numericSymbols?: ReadonlyArray<string | null> | null;
  /**
   * When false, a pattern without prefix, suffix, scaling or percent yields null even when it shows
   * fixed decimals or grouping (used for Highcharts' default axis labels, whose plain numbers Excel's
   * General already approximates). Default true.
   */
  plainFormats?: boolean;
}

export interface InferredFormat {
  code: string;
  /** Samples whose text contained their value. */
  matched: number;
  total: number;
  /** One rendered text that follows the inferred pattern. */
  example: string;
}

type PercentKind = 'none' | 'fraction';

interface SampleMatch {
  prefix: string;
  /** Literal suffix; for percent fractions the text after the '%'. */
  suffix: string;
  /** Text between the number and '%' (' ' or ''), percent fractions only. */
  percentGap: string;
  scaleExp: number;
  percent: PercentKind;
  decimals: number;
  grouped: boolean;
  intDigits: number;
  /** More decimals shown than the value needs: evidence of a fixed decimals count. */
  padded: boolean;
  zero: boolean;
  text: string;
}

const COMMON_SYMBOLS: ReadonlyArray<[string, number]> = [
  ['k', 1],
  ['K', 1],
  ['M', 2],
  ['MM', 2],
  ['Mn', 2],
  ['mn', 2],
  ['B', 3],
  ['Bn', 3],
  ['bn', 3],
  ['G', 3],
  ['T', 4],
  ['Tn', 4],
];

const CURRENCY_TOKEN = /(?:^|[\s:(])([A-Z]{3}|[$€£¥₹₩₽¢])$/;
const SUFFIX_TOKEN = /^ ([A-Z]{3}\b|[$€£¥₹₩₽¢%])/;
const GLUED_STOP = '\\s\\d:;,|=()\\[\\]{}<>';

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function naturalDecimals(x: number): number {
  const s = String(Math.abs(Math.round(x * 1e10) / 1e10));
  if (s.includes('e')) return 0;
  const dot = s.indexOf('.');
  return dot === -1 ? 0 : s.length - dot - 1;
}

function close(shown: number, expected: number, decimals: number): boolean {
  const tol = 0.5 * 10 ** -decimals + 1e-9 * Math.max(1, Math.abs(expected));
  return Math.abs(shown - expected) <= tol;
}

function symbolTable(opts: InferOptions): Array<[string, number]> {
  const table = new Map<string, number>(COMMON_SYMBOLS);
  (opts.numericSymbols ?? []).forEach((sym, i) => {
    if (typeof sym === 'string' && sym.trim() !== '') table.set(sym.trim(), i + 1);
  });
  // Longest first so "MM" wins over "M".
  return [...table.entries()].sort((a, b) => b[0].length - a[0].length);
}

function matchSample(sample: FormatSample, opts: InferOptions, symbols: Array<[string, number]>): SampleMatch | null {
  const { value } = sample;
  if (!Number.isFinite(value)) return null;
  const text = stripHtml(sample.text)
    .replace(/[    ]/g, ' ')
    .replace(/−/g, '-')
    .trim();
  const dec = opts.decimalPoint ?? '.';
  const seps = [opts.thousandsSep, ',', ' ', "'", '.'].filter(
    (s): s is string => typeof s === 'string' && s.length === 1 && s !== dec,
  );
  const sepClass = `[${[...new Set(seps)].map(escapeRe).join('')}]`;
  const re = new RegExp(`(\\d{1,3}(?:(${sepClass})\\d{3})(?:\\2\\d{3})*|\\d+)(?:${escapeRe(dec)}(\\d+))?(?!\\d)`, 'g');

  let best: SampleMatch | null = null;
  let bestDigits = -1;
  for (const m of text.matchAll(re)) {
    const idx = m.index ?? 0;
    if (idx > 0 && /\d/.test(text[idx - 1] ?? '')) continue;
    const intRaw = m[1] ?? '';
    const sep = m[2];
    const decRaw = m[3] ?? '';
    const intStr = sep ? intRaw.split(sep).join('') : intRaw;
    const abs = Number(`${intStr}${decRaw ? `.${decRaw}` : ''}`);
    if (!Number.isFinite(abs)) continue;
    const decimals = decRaw.length;

    // Prefix: the text glued to the number ("$", "US$", "-€"), or a currency token one space before.
    const before = text.slice(0, idx);
    const gluedRaw = new RegExp(`[^${GLUED_STOP}]*$`).exec(before)?.[0] ?? '';
    const negative = /-/.test(gluedRaw) || /\(\s*$/.test(before.slice(0, before.length - gluedRaw.length));
    let prefix = gluedRaw.replace(/[-+]/g, '');
    if (prefix === '' && gluedRaw === '' && before.endsWith(' ')) {
      const cur = CURRENCY_TOKEN.exec(before.slice(0, -1));
      if (cur?.[1]) prefix = `${cur[1]} `;
    }

    // Suffix: the glued text ("K", "%", "M€"), or " %" / " USD" / " €" one space after.
    const after = text.slice(idx + m[0].length);
    let suffix = (new RegExp(`^[^${GLUED_STOP}]*`).exec(after)?.[0] ?? '').replace(/[.!?]+$/, '');
    if (suffix === '') {
      const tok = SUFFIX_TOKEN.exec(after);
      if (tok?.[1]) suffix = ` ${tok[1]}`;
      else {
        const sym = symbols.find(([s]) => after.startsWith(` ${s}`) && !/[A-Za-z]/.test(after[s.length + 1] ?? ''));
        if (sym) suffix = ` ${sym[0]}`;
      }
    }
    const shown = negative ? -abs : abs;

    const base = {
      prefix,
      decimals,
      grouped: sep !== undefined,
      intDigits: intStr.replace(/^0+(?=\d)/, '').length,
      zero: value === 0,
      text,
    };
    let found: SampleMatch | null = null;
    const trimmed = suffix.trimStart();
    // Scaled ("5K" for 5000) before literal, so a value of exactly 0 does not hide the scale.
    const sym = symbols.find(([s]) => trimmed.startsWith(s) && !/[A-Za-z]/.test(trimmed[s.length] ?? ''));
    if (sym && value !== 0) {
      const expected = value / 1000 ** sym[1];
      if (close(shown, expected, decimals) || close(abs, Math.abs(expected), decimals)) {
        found = {
          ...base,
          suffix,
          percentGap: '',
          scaleExp: sym[1],
          percent: 'none',
          padded: decimals > naturalDecimals(expected),
        };
      }
    }
    if (!found && trimmed.startsWith('%') && value !== 0) {
      const expected = value * 100;
      if (close(shown, expected, decimals) && !close(shown, value, decimals)) {
        found = {
          ...base,
          suffix: trimmed.slice(1),
          percentGap: suffix.startsWith(' ') ? ' ' : '',
          scaleExp: 0,
          percent: 'fraction',
          padded: decimals > naturalDecimals(expected),
        };
      }
    }
    if (!found && (close(shown, value, decimals) || close(abs, Math.abs(value), decimals))) {
      found = {
        ...base,
        suffix,
        percentGap: '',
        scaleExp: 0,
        percent: 'none',
        padded: decimals > naturalDecimals(value),
      };
    }
    if (found) {
      const digits = intStr.length + decimals;
      if (digits > bestDigits) {
        best = found;
        bestDigits = digits;
      }
    }
  }
  return best;
}

function patternKey(m: SampleMatch): string {
  return JSON.stringify([m.prefix, m.suffix, m.percentGap, m.scaleExp, m.percent]);
}

function core(decimals: number, grouped: boolean): string {
  return (grouped ? '#,##0' : '0') + (decimals > 0 ? `.${'0'.repeat(decimals)}` : '');
}

function buildCode(p: SampleMatch, decimals: number, grouped: boolean): string {
  const num = core(decimals, grouped) + ','.repeat(p.scaleExp);
  const pct = p.percent === 'fraction' ? `${p.percentGap}%` : '';
  return escapeAffix(p.prefix) + num + pct + escapeAffix(p.suffix);
}

/**
 * Infers the Excel number format that reproduces how the samples were rendered, or null when the
 * samples disagree, do not contain their values, or show nothing Excel's General would not.
 */
export function inferNumberFormat(samples: readonly FormatSample[], opts: InferOptions = {}): InferredFormat | null {
  if (samples.length === 0) return null;
  const symbols = symbolTable(opts);
  const matched = samples.map((s) => matchSample(s, opts, symbols)).filter((m): m is SampleMatch => m !== null);
  if (matched.length === 0 || matched.length < samples.length * 0.5) return null;
  const nonZero = matched.filter((m) => !m.zero);
  if (nonZero.length === 0) return null;

  const groups = new Map<string, SampleMatch[]>();
  for (const m of nonZero) {
    const k = patternKey(m);
    groups.set(k, [...(groups.get(k) ?? []), m]);
  }
  const winner = [...groups.values()].sort((a, b) => b.length - a.length)[0] ?? [];
  if (winner.length < nonZero.length * 0.6) return null;
  const p = winner[0];
  if (!p) return null;

  const decimalsSeen = winner.map((m) => m.decimals);
  const maxDecimals = Math.max(...decimalsSeen);
  const fixed = decimalsSeen.every((d) => d === decimalsSeen[0]) || winner.some((m) => m.padded);
  let grouped: boolean;
  if (winner.some((m) => m.grouped)) grouped = true;
  else if (winner.some((m) => m.intDigits >= 5)) grouped = false;
  else grouped = p.prefix !== '' || p.suffix !== '' || p.scaleExp > 0;

  const plain = p.prefix === '' && p.suffix === '' && p.scaleExp === 0 && p.percent === 'none';
  if (plain) {
    if (opts.plainFormats === false || !fixed) return null;
    if (!grouped && maxDecimals === 0) return null;
  }
  const decimals = Math.min(maxDecimals, fixed ? 30 : 6);
  let code = buildCode(p, decimals, grouped);

  // A zero rendered without the affixes ("0" next to "$5K") gets its own section.
  const zero = matched.find((m) => m.zero);
  const affixSuffix = p.percent === 'fraction' ? `${p.percentGap}%${p.suffix}` : p.suffix;
  if (zero && !plain && (zero.prefix !== p.prefix || zero.suffix !== affixSuffix)) {
    const zeroCode = escapeAffix(zero.prefix) + core(zero.decimals, false) + escapeAffix(zero.suffix);
    code = `${code};-${code};${zeroCode}`;
  }
  if (excelFormatCodeProblem(code) !== null) return null;
  return { code, matched: matched.length, total: samples.length, example: p.text };
}
