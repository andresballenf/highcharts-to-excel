/**
 * Date helpers: Excel serial dates and Highcharts (strftime-style) date format conversion.
 */

import type { NumberFormat } from '../types/chart-model';
import { excelFormatCodeProblem } from './format-code';

/** Days between the Excel 1900 epoch (with the 1900 leap-year bug) and 1970-01-01. */
export const EXCEL_EPOCH_OFFSET_DAYS = 25569;

const MS_PER_DAY = 86_400_000;

/** Serial (offset-based, before the leap-year correction) of 1900-03-01, the first date after Excel's fake 1900-02-29. */
const FIRST_SERIAL_AFTER_FAKE_LEAP_DAY = 61;

/**
 * Converts a UTC timestamp in ms to an Excel serial date (no timezone shift).
 * Excel treats 1900 as a leap year (serial 60 = the non-existent 1900-02-29), so dates before
 * 1900-03-01 are one lower than the plain offset gives: 1900-01-01 → 1, 1900-02-28 → 59,
 * 1899-12-31 → 0. Dates before 1899-12-31 have no Excel date: the result is negative and callers
 * must not write it as a date.
 */
export function msToExcelSerial(ms: number): number {
  let serial = ms / MS_PER_DAY + EXCEL_EPOCH_OFFSET_DAYS;
  if (serial < FIRST_SERIAL_AFTER_FAKE_LEAP_DAY) serial -= 1;
  return Math.round(serial * 1e9) / 1e9;
}

/** Converts an Excel serial date back to a UTC timestamp in ms (rounded to the millisecond). */
export function excelSerialToMs(serial: number): number {
  const s = serial < FIRST_SERIAL_AFTER_FAKE_LEAP_DAY - 1 ? serial + 1 : serial;
  return Math.round((s - EXCEL_EPOCH_OFFSET_DAYS) * MS_PER_DAY);
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
  let seenHour = false;

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
        seenHour = true;
        emit('hh');
        break;
      case 'k':
        seenHour = true;
        emit('h');
        break;
      case 'I':
      case 'l':
        seenHour = true;
        emit(t === 'I' ? 'hh' : 'h');
        if (!hasAmPm) needsAmPm = true;
        break;
      case 'M':
        // Excel reads `mm` as minutes only when it follows an hour token or precedes seconds;
        // anywhere else it would show the month.
        if (!seenHour && !format.slice(i + 1).includes('%S')) {
          return {
            kind: 'unsupported',
            reason: 'Minutes (%M) without an hour or seconds token read as months in Excel',
            source: format,
          };
        }
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
  const problem = excelFormatCodeProblem(code);
  if (problem !== null)
    return { kind: 'unsupported', reason: `Excel date format would be invalid (${problem})`, source: format };
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
