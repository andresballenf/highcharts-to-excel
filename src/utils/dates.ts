/**
 * Date helpers: Excel serial dates and Highcharts (strftime-style) date format conversion.
 */

import type { NumberFormat } from '../types/chart-model';

/** Days between the Excel 1900 epoch (with the 1900 leap-year bug) and 1970-01-01. */
export const EXCEL_EPOCH_OFFSET_DAYS = 25569;

const MS_PER_DAY = 86_400_000;

/** Converts a UTC timestamp in ms to an Excel serial date (no timezone shift). */
export function msToExcelSerial(ms: number): number {
  const serial = ms / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
  return Math.round(serial * 1e9) / 1e9;
}

/** Converts an Excel serial date back to a UTC timestamp in ms (rounded to the millisecond). */
export function excelSerialToMs(serial: number): number {
  return Math.round((serial - EXCEL_EPOCH_OFFSET_DAYS) * MS_PER_DAY);
}

/** Characters Excel shows literally in a date format without quoting. */
const BARE_DATE_LITERALS = new Set([' ', '-', '/', ':', ',', '.']);

/**
 * Converts a Highcharts/strftime date format (e.g. `%Y-%m-%d`) into an Excel number format code.
 * Tokens without an Excel equivalent (`%j`, `%w`, `%U`, ...) make the whole format unsupported.
 */
export function highchartsDateFormatToExcel(format: string): NumberFormat {
  if (typeof format !== 'string') {
    return { kind: 'unsupported', reason: 'Date format is not a string' };
  }
  if (format === '') {
    return { kind: 'unsupported', reason: 'Empty date format', source: format };
  }
  const hasAmPm = /%[pP]/.test(format);
  let code = '';
  let literal = '';
  let needsAmPm = false;

  const flushLiteral = (): void => {
    if (!literal) return;
    code += quoteDateLiteral(literal);
    literal = '';
  };
  const emit = (token: string): void => {
    flushLiteral();
    code += token;
  };

  for (let i = 0; i < format.length; i++) {
    const ch = format[i] ?? '';
    if (ch !== '%') {
      literal += ch;
      continue;
    }
    const t = format[i + 1];
    i++;
    if (t === undefined) {
      literal += '%';
      break;
    }
    switch (t) {
      case 'Y':
        emit('yyyy');
        break;
      case 'y':
        emit('yy');
        break;
      case 'm':
        emit('mm');
        break;
      case 'd':
        emit('dd');
        break;
      case 'e':
        emit('d');
        break;
      case 'b':
        emit('mmm');
        break;
      case 'B':
        emit('mmmm');
        break;
      case 'a':
        emit('ddd');
        break;
      case 'A':
        emit('dddd');
        break;
      case 'H':
        emit('hh');
        break;
      case 'k':
        emit('h');
        break;
      case 'I':
      case 'l':
        emit(t === 'I' ? 'hh' : 'h');
        if (!hasAmPm) needsAmPm = true;
        break;
      case 'M':
        // Excel reads `mm` as minutes when it follows an hour token or precedes seconds.
        emit('mm');
        break;
      case 'S':
        emit('ss');
        break;
      case 'L':
        // Avoid a double dot for the common `%S.%L` pattern.
        if (literal.endsWith('.')) literal = literal.slice(0, -1);
        emit('.000');
        break;
      case 'p':
      case 'P':
        emit('AM/PM');
        break;
      case '%':
        literal += '%';
        break;
      default:
        return {
          kind: 'unsupported',
          reason: `Date format token %${t} has no Excel equivalent`,
          source: format,
        };
    }
  }
  flushLiteral();
  if (needsAmPm) code += ' AM/PM';
  return { kind: 'excel', code, source: format };
}

/** Quotes literal text for a date format; bare-safe characters stay unquoted. */
function quoteDateLiteral(text: string): string {
  let out = '';
  let run = '';
  const flush = (): void => {
    if (run) out += `"${run.replace(/"/g, '"\\""')}"`;
    run = '';
  };
  for (const ch of text) {
    if (BARE_DATE_LITERALS.has(ch)) {
      flush();
      out += ch;
    } else {
      run += ch;
    }
  }
  flush();
  return out;
}

/** Picks a readable Excel date format for a datetime axis spanning `spanMs`. */
export function guessExcelDateFormatForRange(spanMs: number): string {
  const day = MS_PER_DAY;
  const span = Math.abs(spanMs);
  if (span <= day) return 'hh:mm';
  if (span <= 7 * day) return 'ddd d mmm';
  if (span <= 180 * day) return 'd mmm';
  if (span <= 2 * 365 * day) return 'mmm yyyy';
  return 'yyyy';
}
