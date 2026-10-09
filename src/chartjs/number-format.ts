/**
 * Chart.js tick formats → Excel number format codes.
 *
 * Chart.js formats numeric ticks with `Intl.NumberFormat(locale, ticks.format)`; time ticks with
 * the date adapter's tokens (date-fns style by default). Both are translated best-effort; callbacks
 * are never executed.
 */

import type { NumberFormat } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import { decimalsToExcelCode, escapeExcelLiteral } from '../translators/number-format-translator';
import { excelFormatCodeProblem } from '../utils/format-code';
import { num, rec, str } from './guards';

export interface ChartJsFormatResult {
  format: NumberFormat | null;
  diagnostic?: Diagnostic;
}

function fractionPart(min: number | undefined, max: number | undefined): string {
  const lo = Math.max(0, Math.min(30, Math.round(min ?? 0)));
  const hi = Math.max(lo, Math.min(30, Math.round(max ?? lo)));
  if (hi === 0) return '';
  return `.${'0'.repeat(lo)}${'#'.repeat(hi - lo)}`;
}

/** The currency symbol `Intl` prints for an ISO code (falls back to the code). */
function currencySymbol(code: string, display: unknown): string {
  try {
    const parts = new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: code,
      currencyDisplay:
        typeof display === 'string' ? (display as Intl.NumberFormatOptions['currencyDisplay']) : 'symbol',
    }).formatToParts(1);
    return parts.find((p) => p.type === 'currency')?.value ?? code;
  } catch {
    return code;
  }
}

/** Intl's default fraction digits for a currency (2 for USD, 0 for JPY). */
function currencyDigits(code: string): number {
  try {
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: code }).resolvedOptions()
      .maximumFractionDigits as number;
  } catch {
    return 2;
  }
}

/**
 * Translates `ticks.format` (an `Intl.NumberFormatOptions` object) into an Excel number format.
 * `property` is the option path, e.g. `options.scales.y.ticks.format`.
 */
export function intlFormatToExcel(format: unknown, property: string): ChartJsFormatResult {
  const f = rec(format);
  if (!f || Object.keys(f).length === 0) return { format: null };
  const source = JSON.stringify(f);
  const style = str(f.style) ?? 'decimal';
  const grouping = f.useGrouping !== false && f.useGrouping !== 'false';
  const minFd = num(f.minimumFractionDigits);
  const maxFd = num(f.maximumFractionDigits);
  const hasFraction = minFd !== undefined || maxFd !== undefined;
  const intPart = grouping ? '#,##0' : '0';
  let code: string;
  let approximated: string | null = null;

  switch (style) {
    case 'percent':
      code = `0${fractionPart(minFd, maxFd ?? minFd ?? 0)}%`;
      break;
    case 'currency': {
      const currency = str(f.currency)?.toUpperCase() ?? 'USD';
      const digits = currencyDigits(currency);
      const symbol = currencySymbol(currency, f.currencyDisplay);
      code = `[$${symbol.replace(/[[\]]/g, '')}]${intPart}${fractionPart(minFd ?? digits, maxFd ?? minFd ?? digits)}`;
      break;
    }
    case 'unit': {
      const unit = str(f.unit) ?? '';
      code = `${intPart}${hasFraction ? fractionPart(minFd, maxFd ?? minFd) : ''}${unit ? escapeExcelLiteral(` ${unit}`) : ''}`;
      approximated = `the unit "${unit}" is written as literal text after the number`;
      break;
    }
    default:
      code = hasFraction
        ? `${intPart}${fractionPart(minFd, maxFd ?? Math.max(minFd ?? 0, 3))}`
        : grouping
          ? decimalsToExcelCode(null, true)
          : 'General';
  }

  const notation = str(f.notation);
  if (notation === 'scientific' || notation === 'engineering') {
    code = `0${fractionPart(minFd ?? 2, maxFd ?? minFd ?? 2)}E+00`;
    if (notation === 'engineering') approximated = 'engineering notation is shown as scientific notation';
  } else if (notation === 'compact') {
    approximated = 'compact notation (1K, 1M) has no general Excel format; full numbers are shown';
  }
  if (f.signDisplay !== undefined && f.signDisplay !== 'auto') {
    approximated ??= `signDisplay "${String(f.signDisplay)}" is not reproduced`;
  }

  if (excelFormatCodeProblem(code) !== null) {
    return {
      format: { kind: 'excel', code: 'General', source },
      diagnostic: createDiagnostic(
        'APPROXIMATED_NUMBER_FORMAT',
        'approximated',
        property,
        'The tick format could not be expressed as an Excel number format; General is used.',
        { details: { format: f } },
      ),
    };
  }
  const result: ChartJsFormatResult = { format: { kind: 'excel', code, source } };
  if (approximated !== null) {
    result.diagnostic = createDiagnostic(
      'APPROXIMATED_NUMBER_FORMAT',
      'approximated',
      property,
      `Tick format approximated: ${approximated}.`,
      { details: { format: f, code } },
    );
  }
  return result;
}

/** date-fns / Chart.js time display tokens → Excel date tokens (longest first). */
const DATE_TOKENS: ReadonlyArray<[string, string]> = [
  ['yyyy', 'yyyy'],
  ['YYYY', 'yyyy'],
  ['yy', 'yy'],
  ['YY', 'yy'],
  ['MMMM', 'mmmm'],
  ['MMM', 'mmm'],
  ['MM', 'mm'],
  ['M', 'm'],
  ['dddd', 'dddd'],
  ['EEEE', 'dddd'],
  ['EEE', 'ddd'],
  ['ddd', 'ddd'],
  ['dd', 'dd'],
  ['DD', 'dd'],
  ['d', 'd'],
  ['D', 'd'],
  ['HH', 'hh'],
  ['H', 'h'],
  ['hh', 'hh'],
  ['h', 'h'],
  ['mm', 'mm'],
  ['m', 'm'],
  ['ss', 'ss'],
  ['s', 's'],
  ['a', 'AM/PM'],
  ['A', 'AM/PM'],
];

/**
 * Translates a date-fns style display format (`'MMM d, yyyy'`, `'HH:mm'`) into an Excel date code.
 * Returns null for tokens Excel has no equivalent for (week numbers, quarters, ordinals).
 */
export function dateDisplayFormatToExcel(format: unknown): string | null {
  if (typeof format !== 'string' || format.trim() === '') return null;
  let out = '';
  let i = 0;
  let hasHour12 = false;
  while (i < format.length) {
    const ch = format[i]!;
    if (ch === "'") {
      const end = format.indexOf("'", i + 1);
      const literal = end === -1 ? format.slice(i + 1) : format.slice(i + 1, end);
      out += escapeExcelLiteral(literal);
      i = end === -1 ? format.length : end + 1;
      continue;
    }
    if (/[A-Za-z]/.test(ch)) {
      const token = DATE_TOKENS.find(([t]) => format.startsWith(t, i));
      if (!token) return null;
      if (token[0] === 'h' || token[0] === 'hh') hasHour12 = true;
      out += token[1];
      i += token[0].length;
      continue;
    }
    out += /[\s,./:-]/.test(ch) ? ch : escapeExcelLiteral(ch);
    i++;
  }
  // Excel shows 24-hour clocks unless AM/PM is present; a 12-hour token without it gets one.
  if (hasHour12 && !out.includes('AM/PM')) out += ' AM/PM';
  return excelFormatCodeProblem(out) === null ? out : null;
}
