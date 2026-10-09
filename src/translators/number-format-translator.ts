/**
 * Translates Highcharts format strings (`{value:.2f}`, `{point.y:,.0f} %`, `{value:%Y-%m-%d}`)
 * and value options (valueDecimals/valuePrefix/valueSuffix) into Excel number format codes.
 * JavaScript formatter callbacks are never executed; they are reported as unsupported.
 */

import type { NumberFormat } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import { highchartsDateFormatToExcel } from '../utils/dates';
import { excelFormatCodeProblem } from '../utils/format-code';

export interface FormatContext {
  kind: 'axisLabel' | 'dataLabel' | 'tooltip' | 'value';
  valueDecimals?: number | null;
  valuePrefix?: string;
  valueSuffix?: string;
  thousandsSep?: string;
  decimalPoint?: string;
  axisType?: 'linear' | 'category' | 'datetime' | 'logarithmic';
  /** Option path of the object holding the format, e.g. "series[0].dataLabels" or "yAxis[0].labels". */
  property: string;
}

export interface TranslatedFormat {
  format: NumberFormat | null;
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  showPercentage: boolean;
  diagnostics: Diagnostic[];
}

/** Characters that may appear unquoted around a number in an Excel format code. */
const BARE_AFFIX_CHARS = new Set(['$', '€', '£', ' ', '-', '(', ')']);

/** Wraps text in double quotes for an Excel format code; embedded quotes become `\"`. */
export function escapeExcelLiteral(text: string): string {
  if (text === '') return '';
  return `"${text.replace(/"/g, '"\\""')}"`;
}

/** Escapes affix text: bare-safe characters at the edges stay bare, the rest is quoted. */
function escapeAffix(text: string): string {
  if (!text) return '';
  const chars = [...text];
  let first = -1;
  let last = -1;
  chars.forEach((c, i) => {
    if (!BARE_AFFIX_CHARS.has(c)) {
      if (first === -1) first = i;
      last = i;
    }
  });
  if (first === -1) return text;
  return (
    chars.slice(0, first).join('') +
    escapeExcelLiteral(chars.slice(first, last + 1).join('')) +
    chars.slice(last + 1).join('')
  );
}

function numberCore(decimals: number | null, thousands: boolean): string {
  if (decimals === null || !Number.isFinite(decimals) || decimals < 0) return thousands ? '#,##0' : 'General';
  const d = Math.min(30, Math.round(decimals));
  return (thousands ? '#,##0' : '0') + (d > 0 ? `.${'0'.repeat(d)}` : '');
}

/**
 * Excel number code for a decimals count. `null` decimals → `General` (or `#,##0` with grouping).
 * `prefix`/`suffix` are raw display text and are escaped. Examples: (2, true) → `#,##0.00`, (0, false) → `0`.
 */
export function decimalsToExcelCode(decimals: number | null, thousands: boolean, prefix = '', suffix = ''): string {
  return escapeAffix(prefix) + numberCore(decimals, thousands) + escapeAffix(suffix);
}

type Segment = { type: 'lit'; text: string } | { type: 'tok'; expr: string; spec: string | undefined; raw: string };
type TokenClass = 'value' | 'percentage' | 'category' | 'series' | 'unknown';

const VALUE_EXPRS = new Set(['value', 'y', 'point.y', 'point.value', 'text']);
const PERCENT_EXPRS = new Set(['point.percentage', 'percentage']);
const CATEGORY_EXPRS = new Set(['point.name', 'key', 'x', 'point.key', 'point.category', 'category', 'point.x']);
const SERIES_EXPRS = new Set(['series.name']);
const SEPARATOR_ONLY = /^[\s:;,|\-/()[\]]*$/;
const FLOAT_SPEC = /^(,)?(?:\.(\d+))?f$/;

function stripHtml(s: string): string {
  return s
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n: string) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, '&');
}

function tokenize(s: string): Segment[] {
  const out: Segment[] = [];
  let lit = '';
  let i = 0;
  while (i < s.length) {
    const open = s.indexOf('{', i);
    if (open === -1) {
      lit += s.slice(i);
      break;
    }
    const close = s.indexOf('}', open + 1);
    if (close === -1) {
      lit += s.slice(i);
      break;
    }
    lit += s.slice(i, open);
    if (lit) out.push({ type: 'lit', text: lit });
    lit = '';
    const inner = s.slice(open + 1, close);
    const colon = inner.indexOf(':');
    out.push({
      type: 'tok',
      expr: (colon === -1 ? inner : inner.slice(0, colon)).trim(),
      spec: colon === -1 ? undefined : inner.slice(colon + 1),
      raw: s.slice(open, close + 1),
    });
    i = close + 1;
  }
  if (lit) out.push({ type: 'lit', text: lit });
  return out;
}

function classify(expr: string): TokenClass {
  if (VALUE_EXPRS.has(expr)) return 'value';
  if (PERCENT_EXPRS.has(expr)) return 'percentage';
  if (CATEGORY_EXPRS.has(expr)) return 'category';
  if (SERIES_EXPRS.has(expr)) return 'series';
  return 'unknown';
}

function isFiniteNumber(v: unknown): v is number {
  return typeof v === 'number' && Number.isFinite(v);
}

/**
 * Translates a Highcharts `format` string (or `formatter` callback) into an Excel number format
 * plus data-label content switches.
 */
export function translateFormatString(format: unknown, formatter: unknown, ctx: FormatContext): TranslatedFormat {
  const r = translateFormatStringUnchecked(format, formatter, ctx);
  if (r.format?.kind !== 'excel') return r;
  const problem = excelFormatCodeProblem(r.format.code);
  if (problem === null) return r;
  // Excel refuses codes that break its structural rules (e.g. longer than 255 characters).
  const property = `${ctx.property}.${ctx.kind === 'tooltip' ? 'pointFormat' : 'format'}`;
  const fallback: NumberFormat =
    r.format.source !== undefined
      ? { kind: 'excel', code: 'General', source: r.format.source }
      : { kind: 'excel', code: 'General' };
  return {
    ...r,
    format: fallback,
    diagnostics: [
      ...r.diagnostics,
      createDiagnostic(
        'UNSUPPORTED_NUMBER_FORMAT',
        'unsupported',
        property,
        `The Excel number format for this label is invalid (${problem}); General is used.`,
        {
          details: { code: r.format.code.slice(0, 300) },
        },
      ),
    ],
  };
}

function translateFormatStringUnchecked(format: unknown, formatter: unknown, ctx: FormatContext): TranslatedFormat {
  const isDataLabel = ctx.kind === 'dataLabel';
  const formatProperty = `${ctx.property}.${ctx.kind === 'tooltip' ? 'pointFormat' : 'format'}`;
  const source = typeof format === 'string' ? format : undefined;

  const unsupported = (reason: string, diagnostic: Diagnostic): TranslatedFormat => ({
    format: source !== undefined ? { kind: 'unsupported', reason, source } : { kind: 'unsupported', reason },
    showValue: isDataLabel,
    showCategoryName: false,
    showSeriesName: false,
    showPercentage: false,
    diagnostics: [diagnostic],
  });

  if (typeof formatter === 'function') {
    return unsupported(
      'JavaScript formatter callbacks cannot be represented in Excel',
      createDiagnostic(
        'UNSUPPORTED_FORMATTER',
        'unsupported',
        `${ctx.property}.formatter`,
        isDataLabel
          ? 'Formatter functions are not executed; Excel shows the raw value instead.'
          : 'Formatter functions are not executed; Excel uses its default formatting.',
      ),
    );
  }

  const diagnostics: Diagnostic[] = [];
  const separatorDiagnostics = (thousands: boolean, decimals: number): void => {
    if (thousands && ctx.thousandsSep !== undefined && ctx.thousandsSep !== ',') {
      diagnostics.push(
        createDiagnostic(
          'APPROXIMATED_NUMBER_FORMAT',
          'approximated',
          'lang.thousandsSep',
          "Excel uses the viewer's locale thousands separator.",
          {
            details: { thousandsSep: ctx.thousandsSep },
          },
        ),
      );
    }
    if (decimals > 0 && ctx.decimalPoint !== undefined && ctx.decimalPoint !== '.') {
      diagnostics.push(
        createDiagnostic(
          'APPROXIMATED_NUMBER_FORMAT',
          'approximated',
          'lang.decimalPoint',
          "Excel uses the viewer's locale decimal separator.",
          {
            details: { decimalPoint: ctx.decimalPoint },
          },
        ),
      );
    }
  };
  const ctxDecimals =
    isFiniteNumber(ctx.valueDecimals) && ctx.valueDecimals >= 0 ? Math.round(ctx.valueDecimals) : null;
  const ctxPrefix = ctx.valuePrefix ?? '';
  const ctxSuffix = ctx.valueSuffix ?? '';

  // No format string: build from valueDecimals / valuePrefix / valueSuffix when present.
  if (source === undefined || source.trim() === '') {
    let fmt: NumberFormat | null = null;
    if (ctxDecimals !== null || ctxPrefix || ctxSuffix) {
      // Highcharts inserts `:,.Nf` for valueDecimals, i.e. with thousands grouping.
      const thousands = ctxDecimals !== null;
      fmt = { kind: 'excel', code: decimalsToExcelCode(ctxDecimals, thousands, ctxPrefix, ctxSuffix) };
      separatorDiagnostics(thousands, ctxDecimals ?? 0);
    }
    return {
      format: fmt,
      showValue: true,
      showCategoryName: false,
      showSeriesName: false,
      showPercentage: false,
      diagnostics,
    };
  }

  const segments = tokenize(stripHtml(source));
  const tokens = segments.flatMap((s, i) => (s.type === 'tok' ? [{ seg: s, index: i, cls: classify(s.expr) }] : []));

  // Which token classes are meaningful for this context.
  const allowed = (cls: TokenClass): boolean => {
    if (cls === 'value') return true;
    if (cls === 'unknown') return false;
    return isDataLabel;
  };
  if (ctx.kind !== 'tooltip') {
    const bad = tokens.find((t) => !allowed(t.cls));
    if (bad) {
      return unsupported(
        `Format token ${bad.seg.raw} has no Excel equivalent`,
        createDiagnostic(
          'UNSUPPORTED_NUMBER_FORMAT',
          'unsupported',
          formatProperty,
          `Format token ${bad.seg.raw} cannot be represented in Excel.`,
          {
            details: { format: source, token: bad.seg.raw },
          },
        ),
      );
    }
  }
  // Tooltips are not representable; only the value formatting is harvested, other tokens are ignored.
  const content = tokens.filter((t) => allowed(t.cls));

  const flags = {
    showValue: isDataLabel ? content.some((t) => t.cls === 'value') : true,
    showCategoryName: isDataLabel && content.some((t) => t.cls === 'category'),
    showSeriesName: isDataLabel && content.some((t) => t.cls === 'series'),
    showPercentage: isDataLabel && content.some((t) => t.cls === 'percentage'),
  };

  const numeric = content.find((t) => t.cls === 'value') ?? content.find((t) => t.cls === 'percentage');

  if (!numeric) {
    if (content.length === 0 && ctx.kind !== 'tooltip') {
      return unsupported(
        'Format has no value placeholder',
        createDiagnostic(
          'UNSUPPORTED_NUMBER_FORMAT',
          'unsupported',
          formatProperty,
          'Constant label text cannot be represented as an Excel number format.',
          {
            details: { format: source },
          },
        ),
      );
    }
    if (isDataLabel) reportDroppedLiterals(segments, diagnostics, formatProperty, source);
    return { format: null, ...flags, diagnostics };
  }

  // Affixes: all literal text when the numeric token is alone on that side; otherwise only the
  // text glued to it (e.g. "$" in "{series.name}: ${y}").
  const hasTokenBefore = tokens.some((t) => t.index < numeric.index);
  const hasTokenAfter = tokens.some((t) => t.index > numeric.index);
  let prefix = '';
  let suffix = '';
  const before = segments[numeric.index - 1];
  if (before?.type === 'lit') {
    prefix = hasTokenBefore ? (/[^\s:;,|]*$/.exec(before.text)?.[0] ?? '') : before.text;
  }
  const after = segments[numeric.index + 1];
  if (after?.type === 'lit') {
    suffix = hasTokenAfter ? (/^[^\s:;,|]*/.exec(after.text)?.[0] ?? '') : after.text;
  }
  const literalPrefix = prefix;
  const literalSuffix = suffix;
  if (ctx.kind === 'tooltip') {
    // Leading/trailing whitespace around the tooltip value is layout, not formatting.
    prefix = prefix.trimStart();
    suffix = suffix.trimEnd();
  }

  const spec = numeric.seg.spec;
  let code: string;

  if (numeric.cls === 'percentage') {
    let decimals = 0;
    if (spec !== undefined) {
      const m = FLOAT_SPEC.exec(spec.trim());
      if (!m) {
        return unsupported(
          `Percentage format "${spec}" has no Excel equivalent`,
          createDiagnostic(
            'UNSUPPORTED_NUMBER_FORMAT',
            'unsupported',
            formatProperty,
            `Format specifier "${spec}" cannot be represented in Excel.`,
            {
              details: { format: source },
            },
          ),
        );
      }
      decimals = m[2] !== undefined ? Number(m[2]) : 0;
    }
    // The literal "%" after the percentage is absorbed into Excel's percent format.
    let pct = '%';
    const pm = /^(\s*)%/.exec(suffix);
    if (pm) {
      pct = `${pm[1] ?? ''}%`;
      suffix = suffix.slice(pm[0].length);
    }
    code = escapeAffix(prefix) + numberCore(decimals, false) + pct + escapeAffix(suffix);
    separatorDiagnostics(false, decimals);
  } else {
    prefix += ctxPrefix;
    suffix = ctxSuffix + suffix;
    if (spec?.includes('%')) {
      const date = highchartsDateFormatToExcel(spec);
      if (date.kind === 'unsupported') {
        return unsupported(
          date.reason,
          createDiagnostic(
            'UNSUPPORTED_NUMBER_FORMAT',
            'unsupported',
            formatProperty,
            `Date format "${spec}" cannot be represented in Excel: ${date.reason}.`,
            {
              details: { format: source },
            },
          ),
        );
      }
      code = escapeAffix(prefix) + date.code + escapeAffix(suffix);
    } else if (spec !== undefined) {
      const m = FLOAT_SPEC.exec(spec.trim());
      if (!m) {
        return unsupported(
          `Format specifier "${spec}" has no Excel equivalent`,
          createDiagnostic(
            'UNSUPPORTED_NUMBER_FORMAT',
            'unsupported',
            formatProperty,
            `Format specifier "${spec}" cannot be represented in Excel.`,
            {
              details: { format: source },
            },
          ),
        );
      }
      const thousands = m[1] === ',';
      const decimals = m[2] !== undefined ? Number(m[2]) : null;
      code = escapeAffix(prefix) + numberCore(decimals, thousands) + escapeAffix(suffix);
      separatorDiagnostics(thousands, decimals ?? 0);
    } else if (ctx.kind === 'axisLabel' && (ctx.axisType === 'category' || ctx.axisType === 'datetime')) {
      // Category labels are text and datetime labels use the axis date format; a number format
      // cannot carry the literal text.
      if (prefix.trim() || suffix.trim()) {
        diagnostics.push(
          createDiagnostic(
            'APPROXIMATED_NUMBER_FORMAT',
            'approximated',
            formatProperty,
            `Literal text around ${ctx.axisType} axis labels is not reproduced.`,
            {
              details: { format: source },
            },
          ),
        );
      }
      return { format: null, ...flags, diagnostics };
    } else if (ctxDecimals !== null) {
      code = escapeAffix(prefix) + numberCore(ctxDecimals, true) + escapeAffix(suffix);
      separatorDiagnostics(true, ctxDecimals);
    } else if (prefix || suffix) {
      // e.g. "{value}%" → General"%" (the value is not a fraction, so "%" stays a literal; General keeps decimals).
      code = `${escapeAffix(prefix)}General${escapeAffix(suffix)}`;
    } else {
      code = 'General';
    }
  }

  if (isDataLabel)
    reportDroppedLiterals(segments, diagnostics, formatProperty, source, numeric.index, literalPrefix, literalSuffix);
  return { format: { kind: 'excel', code, source }, ...flags, diagnostics };
}

/** Reports literal label text that Excel's composed data labels cannot show. */
function reportDroppedLiterals(
  segments: Segment[],
  diagnostics: Diagnostic[],
  property: string,
  source: string,
  numericIndex = -1,
  prefix = '',
  suffix = '',
): void {
  let dropped = '';
  segments.forEach((s, i) => {
    if (s.type !== 'lit') return;
    let text = s.text;
    if (i === numericIndex - 1 && prefix && text.endsWith(prefix)) text = text.slice(0, -prefix.length);
    if (i === numericIndex + 1 && suffix && text.startsWith(suffix)) text = text.slice(suffix.length);
    dropped += text;
  });
  if (!SEPARATOR_ONLY.test(dropped)) {
    diagnostics.push(
      createDiagnostic(
        'APPROXIMATED_DATA_LABELS',
        'approximated',
        property,
        "Literal text between label parts is replaced by Excel's label separator.",
        {
          details: { format: source, droppedText: dropped.trim() },
        },
      ),
    );
  }
}
