import { describe, expect, it } from 'vitest';
import { inferNumberFormat } from '../../src/translators/format-inference';

const s = (value: number, text: string) => ({ value, text });

describe('inferNumberFormat', () => {
  it('reads a currency tooltip with fixed decimals and grouping', () => {
    const r = inferNumberFormat([
      s(12500.5, 'Performance: <b>$12,500.50</b>'),
      s(23000, 'Performance: <b>$23,000.00</b>'),
      s(950, 'Performance: <b>$950.00</b>'),
    ]);
    expect(r?.code).toBe('$#,##0.00');
    expect(r?.matched).toBe(3);
  });

  it('reads k/M scaling and gives an unscaled zero its own section', () => {
    expect(inferNumberFormat([s(0, '$0'), s(5000, '$5K'), s(10000, '$10K'), s(15000, '$15K')])?.code).toBe(
      '$#,##0,"K";-$#,##0,"K";$0',
    );
    expect(inferNumberFormat([s(0, '0'), s(500000, '0.5M'), s(1e6, '1M'), s(1.5e6, '1.5M')])?.code).toBe(
      '#,##0.0,,"M";-#,##0.0,,"M";0',
    );
    expect(inferNumberFormat([s(2.1e9, '€2.1bn'), s(3e9, '€3.0bn')])?.code).toBe('€#,##0.0,,,"bn"');
  });

  it('tells percent fractions from literal percents', () => {
    expect(inferNumberFormat([s(0.453, '45.3%'), s(0.1, '10.0%')])?.code).toBe('0.0%');
    expect(inferNumberFormat([s(45.3, '45.3 %'), s(10, '10.0 %')])?.code).toBe('#,##0.0 "%"');
  });

  it('honours locale separators and trailing currency', () => {
    const r = inferNumberFormat([s(1234.5, '1.234,50 €'), s(99, '99,00 €')], { decimalPoint: ',', thousandsSep: '.' });
    expect(r?.code).toBe('#,##0.00 €');
  });

  it('picks the number that matches the value among several', () => {
    const r = inferNumberFormat([
      s(1234.5, 'Jan 2024<br>Mainstream: $1,234.50<br>Previous: $999.00'),
      s(999, 'Feb 2024<br>Mainstream: $2,000.00<br>Previous: $999.00'),
    ]);
    expect(r?.code).toBe('$#,##0.00');
  });

  it('returns null for plain natural numbers, unmatched or inconsistent samples', () => {
    expect(inferNumberFormat([s(12500.5, '12 500.5'), s(23000, '23 000')])).toBeNull();
    expect(inferNumberFormat([s(5, 'five'), s(6, 'six')])).toBeNull();
    expect(inferNumberFormat([s(5, '$5'), s(6, '6 USD'), s(7, '7%')])).toBeNull();
    expect(inferNumberFormat([])).toBeNull();
  });

  it('keeps plain fixed decimals unless plain formats are disabled', () => {
    const samples = [s(1, '1.00'), s(2.5, '2.50')];
    expect(inferNumberFormat(samples)?.code).toBe('0.00');
    expect(inferNumberFormat(samples, { plainFormats: false })).toBeNull();
  });

  it('handles negatives and currency codes', () => {
    expect(inferNumberFormat([s(-1500, '-$1,500'), s(2000, '$2,000')])?.code).toBe('$#,##0');
    expect(inferNumberFormat([s(1500, 'USD 1,500'), s(20, 'USD 20')])?.code).toBe('"USD" #,##0');
  });
});

describe('inferNumberFormat zero labels', () => {
  it('adds no zero section when zero follows the pattern', () => {
    expect(inferNumberFormat([s(0, '$0K'), s(5000, '$5K')])?.code).toBe('$#,##0,"K"');
    expect(inferNumberFormat([s(0, '0.0%'), s(0.25, '25.0%')])?.code).toBe('0.0%');
  });
});
