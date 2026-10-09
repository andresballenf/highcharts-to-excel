/**
 * Chart.js tick formats (Intl.NumberFormat options, date-fns display formats) → Excel codes.
 */

import { describe, expect, it } from 'vitest';
import { dateDisplayFormatToExcel, intlFormatToExcel } from '../../src/chartjs';
import { isValidExcelFormatCode } from '../../src/utils/format-code';

const P = 'options.scales.y.ticks.format';

describe('intlFormatToExcel', () => {
  it.each([
    [{ style: 'percent' }, '0%'],
    [{ style: 'percent', minimumFractionDigits: 1 }, '0.0%'],
    [{ style: 'currency', currency: 'USD' }, '[$$]#,##0.00'],
    [{ style: 'currency', currency: 'EUR' }, '[$€]#,##0.00'],
    [{ style: 'currency', currency: 'JPY' }, '[$¥]#,##0'],
    [{ minimumFractionDigits: 2 }, '#,##0.00#'],
    [{ minimumFractionDigits: 2, maximumFractionDigits: 2 }, '#,##0.00'],
    [{ maximumFractionDigits: 0, useGrouping: false }, '0'],
    [{ useGrouping: true }, '#,##0'],
    [{ useGrouping: false }, 'General'],
    [{ notation: 'scientific' }, '0.00E+00'],
  ])('%j → %s', (format, code) => {
    const r = intlFormatToExcel(format, P);
    expect(r.format).toEqual({ kind: 'excel', code, source: JSON.stringify(format) });
    expect(r.diagnostic).toBeUndefined();
    expect(isValidExcelFormatCode(code)).toBe(true);
  });

  it('returns no format for missing or empty options', () => {
    expect(intlFormatToExcel(undefined, P).format).toBeNull();
    expect(intlFormatToExcel({}, P).format).toBeNull();
  });

  it('approximates compact notation and units with a diagnostic at the option path', () => {
    const compact = intlFormatToExcel({ notation: 'compact' }, P);
    expect(compact.format).toMatchObject({ kind: 'excel', code: '#,##0' });
    expect(compact.diagnostic).toMatchObject({ code: 'APPROXIMATED_NUMBER_FORMAT', property: P });
    const unit = intlFormatToExcel({ style: 'unit', unit: 'kilometer', maximumFractionDigits: 1 }, P);
    expect(unit.format).toMatchObject({ code: '#,##0.#" kilometer"' });
    expect(unit.diagnostic?.code).toBe('APPROXIMATED_NUMBER_FORMAT');
  });
});

describe('dateDisplayFormatToExcel', () => {
  it.each([
    ['MMM d', 'mmm d'],
    ['yyyy-MM-dd', 'yyyy-mm-dd'],
    ['MMM yyyy', 'mmm yyyy'],
    ['HH:mm', 'hh:mm'],
    ['h:mm a', 'h:mm AM/PM'],
    ['ha', 'hAM/PM'],
    ['h:mm', 'h:mm AM/PM'],
    ["d 'de' MMMM", 'd "de" mmmm'],
  ])('%s → %s', (input, code) => {
    expect(dateDisplayFormatToExcel(input)).toBe(code);
  });

  it('returns null for tokens Excel cannot show (quarters, week numbers) and non-strings', () => {
    expect(dateDisplayFormatToExcel('qqq - yyyy')).toBeNull();
    expect(dateDisplayFormatToExcel('PP')).toBeNull();
    expect(dateDisplayFormatToExcel(undefined)).toBeNull();
    expect(dateDisplayFormatToExcel('')).toBeNull();
  });
});
