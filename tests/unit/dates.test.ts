import { describe, expect, it } from 'vitest';
import {
  EXCEL_EPOCH_OFFSET_DAYS,
  excelSerialToMs,
  guessExcelDateFormatForRange,
  highchartsDateFormatToExcel,
  msToExcelSerial,
} from '../../src/utils/dates';

const code = (f: string) => {
  const r = highchartsDateFormatToExcel(f);
  return r.kind === 'excel' ? r.code : `unsupported:${r.reason}`;
};

describe('Excel serial dates', () => {
  it('converts UTC ms to serials', () => {
    expect(EXCEL_EPOCH_OFFSET_DAYS).toBe(25569);
    expect(msToExcelSerial(0)).toBe(25569);
    expect(msToExcelSerial(Date.UTC(2024, 0, 1))).toBe(45292);
    expect(msToExcelSerial(Date.UTC(2024, 0, 1, 12))).toBe(45292.5);
    expect(msToExcelSerial(Date.UTC(2024, 0, 1, 8))).toBe(45292.333333333);
  });
  it('round-trips', () => {
    const ms = Date.UTC(2023, 6, 15, 13, 45, 30, 250);
    expect(excelSerialToMs(msToExcelSerial(ms))).toBe(ms);
    expect(excelSerialToMs(45292)).toBe(Date.UTC(2024, 0, 1));
  });
});

describe('highchartsDateFormatToExcel', () => {
  it('converts common tokens', () => {
    expect(highchartsDateFormatToExcel('%Y-%m-%d')).toEqual({ kind: 'excel', code: 'yyyy-mm-dd', source: '%Y-%m-%d' });
    expect(code('%e. %b')).toBe('d. mmm');
    expect(code('%A, %B %e, %Y')).toBe('dddd, mmmm d, yyyy');
    expect(code('%a %y')).toBe('ddd yy');
    expect(code('%H:%M')).toBe('hh:mm');
    expect(code('%k:%M:%S')).toBe('h:mm:ss');
    expect(code('%H:%M:%S.%L')).toBe('hh:mm:ss.000');
    expect(code('%S%L')).toBe('ss.000');
  });
  it('handles 12-hour clocks and AM/PM', () => {
    expect(code('%I:%M')).toBe('hh:mm AM/PM');
    expect(code('%l:%M %p')).toBe('h:mm AM/PM');
    expect(code('%l:%M %P')).toBe('h:mm AM/PM');
  });
  it('quotes literal text and handles %%', () => {
    expect(code('Week of %b %e')).toBe('"Week" "of" mmm d');
    expect(code('%Y年')).toBe('yyyy"年"');
    expect(code('%d%%')).toBe('dd"%"');
  });
  it('rejects unsupported tokens', () => {
    for (const f of ['%j', '%w', '%u', '%U', 'Week %W']) {
      expect(highchartsDateFormatToExcel(f)).toMatchObject({ kind: 'unsupported', source: f });
    }
    expect(highchartsDateFormatToExcel(42 as unknown as string).kind).toBe('unsupported');
  });
});

describe('guessExcelDateFormatForRange', () => {
  const day = 86_400_000;
  it('picks by span', () => {
    expect(guessExcelDateFormatForRange(6 * 3_600_000)).toBe('hh:mm');
    expect(guessExcelDateFormatForRange(day)).toBe('hh:mm');
    expect(guessExcelDateFormatForRange(5 * day)).toBe('ddd d mmm');
    expect(guessExcelDateFormatForRange(90 * day)).toBe('d mmm');
    expect(guessExcelDateFormatForRange(400 * day)).toBe('mmm yyyy');
    expect(guessExcelDateFormatForRange(10 * 365 * day)).toBe('yyyy');
  });
});
