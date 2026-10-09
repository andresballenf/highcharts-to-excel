import { describe, expect, it } from 'vitest';
import {
  decimalsToExcelCode,
  escapeExcelLiteral,
  translateFormatString,
  type FormatContext,
} from '../../src/translators/number-format-translator';

const axis: FormatContext = { kind: 'axisLabel', property: 'yAxis[0].labels', axisType: 'linear' };
const dl: FormatContext = { kind: 'dataLabel', property: 'series[0].dataLabels' };
const codeOf = (format: string, ctx: FormatContext = axis) => {
  const r = translateFormatString(format, undefined, ctx);
  return r.format?.kind === 'excel' ? r.format.code : r.format;
};

describe('decimalsToExcelCode / escapeExcelLiteral', () => {
  it('builds codes', () => {
    expect(decimalsToExcelCode(2, true)).toBe('#,##0.00');
    expect(decimalsToExcelCode(0, false)).toBe('0');
    expect(decimalsToExcelCode(null, false)).toBe('General');
    expect(decimalsToExcelCode(1, false, '$', ' USD')).toBe('$0.0 "USD"');
  });
  it('escapes literals', () => {
    expect(escapeExcelLiteral('km')).toBe('"km"');
    expect(escapeExcelLiteral('a"b')).toBe('"a"\\""b"');
    expect(escapeExcelLiteral('')).toBe('');
  });
});

describe('translateFormatString', () => {
  it('maps a bare value to General', () => {
    const r = translateFormatString('{value}', undefined, axis);
    expect(r.format).toEqual({ kind: 'excel', code: 'General', source: '{value}' });
    expect(r.showValue).toBe(true);
    expect(r.diagnostics).toEqual([]);
  });

  it('keeps a literal percent sign quoted', () => {
    expect(codeOf('{value}%')).toBe('General"%"');
    expect(codeOf('{value:.1f}%')).toBe('0.0"%"');
  });

  it('handles decimals, grouping and currency prefixes', () => {
    // biome-ignore lint/suspicious/noTemplateCurlyInString: a Highcharts format string, not a template
    expect(codeOf('${point.y:,.2f}', dl)).toBe('$#,##0.00');
    expect(codeOf('{value:.2f}')).toBe('0.00');
    expect(codeOf('{y:.1f}', dl)).toBe('0.0');
    expect(codeOf('{point.y:,.0f}', dl)).toBe('#,##0');
    expect(codeOf('{value} km')).toBe('General "km"');
    expect(codeOf('€{value:,.2f}')).toBe('€#,##0.00');
    expect(codeOf('<b>{value:.0f}</b><br/>units')).toBe('0 "units"');
  });

  it('translates date tokens', () => {
    expect(codeOf('{value:%Y-%m-%d}', { ...axis, axisType: 'datetime' })).toBe('yyyy-mm-dd');
    const r = translateFormatString('{value:%j}', undefined, axis);
    expect(r.format?.kind).toBe('unsupported');
    expect(r.diagnostics[0]).toMatchObject({ code: 'UNSUPPORTED_NUMBER_FORMAT', property: 'yAxis[0].labels.format' });
  });

  it('reports formatter functions as unsupported', () => {
    const r = translateFormatString('{value}', () => 'x', dl);
    expect(r.format?.kind).toBe('unsupported');
    expect(r).toMatchObject({ showValue: true, showCategoryName: false, showSeriesName: false, showPercentage: false });
    expect(r.diagnostics).toHaveLength(1);
    expect(r.diagnostics[0]).toMatchObject({
      code: 'UNSUPPORTED_FORMATTER',
      outcome: 'unsupported',
      property: 'series[0].dataLabels.formatter',
    });
    expect(translateFormatString(undefined, () => 'x', axis).showValue).toBe(false);
  });

  it('builds codes from valueDecimals/prefix/suffix when no format is given', () => {
    const tooltip: FormatContext = { kind: 'tooltip', property: 'tooltip', valueDecimals: 2, valueSuffix: ' °C' };
    expect(translateFormatString(undefined, undefined, tooltip).format).toEqual({
      kind: 'excel',
      code: '#,##0.00 "°C"',
    });
    const none = translateFormatString('', undefined, axis);
    expect(none.format).toBeNull();
    expect(none.showValue).toBe(true);
  });

  it('harvests the value format from tooltip point formats', () => {
    const tooltip: FormatContext = { kind: 'tooltip', property: 'tooltip', valueDecimals: 1, valuePrefix: '$' };
    const r = translateFormatString(
      '<span style="color:{point.color}">●</span> {series.name}: <b>{point.y}</b><br/>',
      undefined,
      tooltip,
    );
    expect(r.format).toMatchObject({ kind: 'excel', code: '$#,##0.0' });
    expect(r.diagnostics).toEqual([]);
  });

  it('sets pie data label switches', () => {
    const pie = translateFormatString('<b>{point.name}</b>: {point.percentage:.1f} %', undefined, dl);
    expect(pie).toMatchObject({
      showValue: false,
      showCategoryName: true,
      showSeriesName: false,
      showPercentage: true,
    });
    expect(pie.format).toMatchObject({ kind: 'excel', code: '0.0 %' });
    expect(pie.diagnostics).toEqual([]);

    expect(codeOf('{point.percentage:.0f}%', dl)).toBe('0%');
    const named = translateFormatString('{point.name}', undefined, dl);
    expect(named).toMatchObject({ format: null, showValue: false, showCategoryName: true });
    const series = translateFormatString('{series.name}: {y}', undefined, dl);
    expect(series).toMatchObject({ showValue: true, showSeriesName: true });
    expect(series.format).toMatchObject({ code: 'General' });
  });

  it('reports dropped label text', () => {
    const r = translateFormatString('{point.name} has {y} items', undefined, dl);
    expect(r.diagnostics.map((d) => d.code)).toEqual(['APPROXIMATED_DATA_LABELS']);
    expect(r.format).toMatchObject({ code: 'General "items"' });
  });

  it('rejects unknown tokens', () => {
    for (const f of ['{point.custom.x}', '{point.z}', '{#if point.y}{point.y}{/if}']) {
      const r = translateFormatString(f, undefined, dl);
      expect(r.format?.kind).toBe('unsupported');
      expect(r.diagnostics[0]).toMatchObject({ code: 'UNSUPPORTED_NUMBER_FORMAT', outcome: 'unsupported' });
    }
    // percentage is not meaningful for axis labels
    expect(translateFormatString('{point.percentage:.1f}', undefined, axis).format?.kind).toBe('unsupported');
  });

  it('flags non-default separators', () => {
    const r = translateFormatString('{value:,.2f}', undefined, { ...axis, thousandsSep: ' ', decimalPoint: ',' });
    expect(r.format).toMatchObject({ code: '#,##0.00' });
    expect(r.diagnostics.map((d) => [d.code, d.property])).toEqual([
      ['APPROXIMATED_NUMBER_FORMAT', 'lang.thousandsSep'],
      ['APPROXIMATED_NUMBER_FORMAT', 'lang.decimalPoint'],
    ]);
    expect(translateFormatString('{value}', undefined, { ...axis, thousandsSep: ' ' }).diagnostics).toEqual([]);
  });

  it('leaves category and datetime axis labels unformatted', () => {
    expect(translateFormatString('{value}', undefined, { ...axis, axisType: 'category' }).format).toBeNull();
    const r = translateFormatString('Q {value}', undefined, { ...axis, axisType: 'category' });
    expect(r.format).toBeNull();
    expect(r.diagnostics[0]).toMatchObject({ code: 'APPROXIMATED_NUMBER_FORMAT', outcome: 'approximated' });
  });
});

describe('W5: invalid Excel number format codes', () => {
  it('falls back to General with UNSUPPORTED_NUMBER_FORMAT when the code exceeds 255 characters', () => {
    const long = 'x'.repeat(300);
    const r = translateFormatString(`${long}{value}`, undefined, axis);
    expect(r.format).toMatchObject({ kind: 'excel', code: 'General' });
    expect(r.diagnostics.find((d) => d.code === 'UNSUPPORTED_NUMBER_FORMAT')).toMatchObject({
      outcome: 'unsupported',
      property: 'yAxis[0].labels.format',
    });
    const viaSuffix = translateFormatString(undefined, undefined, { ...axis, kind: 'tooltip', valueSuffix: long });
    expect(viaSuffix.format).toMatchObject({ kind: 'excel', code: 'General' });
    expect(viaSuffix.diagnostics.some((d) => d.code === 'UNSUPPORTED_NUMBER_FORMAT')).toBe(true);
  });
});
