import { describe, expect, it } from 'vitest';
import {
  DEFAULT_HIGHCHARTS_FONT,
  fontToOoxml,
  primaryFontFamily,
  toFont,
} from '../../src/translators/typography-translator';
import { colorToHex } from '../../src/utils/colors';

describe('toFont', () => {
  it('parses sizes in px, pt, em and %', () => {
    expect(toFont({ fontSize: 14 }, null).size).toBe(14);
    expect(toFont({ fontSize: '12px' }, null).size).toBe(12);
    expect(toFont({ fontSize: '9pt' }, null).size).toBe(12);
    expect(toFont({ fontSize: '0.8em' }, DEFAULT_HIGHCHARTS_FONT).size).toBe(12.8);
    expect(toFont({ fontSize: '80%' }, null).size).toBe(12.8);
    expect(toFont({ fontSize: '1.2em' }, null).size).toBe(19.2);
    expect(toFont({ fontSize: '1rem' }, DEFAULT_HIGHCHARTS_FONT).size).toBe(16);
    expect(toFont({ fontSize: 'garbage' }, DEFAULT_HIGHCHARTS_FONT).size).toBe(16);
  });

  it('parses weight, style and color', () => {
    expect(toFont({ fontWeight: 'bold' }, null).bold).toBe(true);
    expect(toFont({ fontWeight: 'bolder' }, null).bold).toBe(true);
    expect(toFont({ fontWeight: 600 }, null).bold).toBe(true);
    expect(toFont({ fontWeight: '700' }, null).bold).toBe(true);
    expect(toFont({ fontWeight: 400 }, { ...DEFAULT_HIGHCHARTS_FONT, bold: true }).bold).toBe(false);
    expect(toFont({ fontStyle: 'italic' }, null).italic).toBe(true);
    expect(toFont({ fontStyle: 'oblique 10deg' }, null).italic).toBe(true);
    expect(colorToHex(toFont({ color: 'var(--highcharts-neutral-color-60)' }, null).color!)).toBe('666666');
  });

  it('inherits missing fields from base', () => {
    const f = toFont({ fontWeight: 'bold' }, DEFAULT_HIGHCHARTS_FONT);
    expect(f).toMatchObject({ family: DEFAULT_HIGHCHARTS_FONT.family, size: 16, bold: true, italic: false });
    expect(colorToHex(f.color!)).toBe('333333');
    expect(toFont(null, null)).toEqual({ family: null, size: null, bold: false, italic: false, color: null });
  });

  it('parses the font shorthand', () => {
    const f = toFont({ font: 'italic bold 12px/1.4 "Open Sans", sans-serif' }, null);
    expect(f).toMatchObject({ italic: true, bold: true, size: 12, family: '"Open Sans", sans-serif' });
    const g = toFont({ font: '10pt Arial', fontWeight: 'bold' }, null);
    expect(g).toMatchObject({ size: 13.333, family: 'Arial', bold: true });
  });
});

describe('primaryFontFamily', () => {
  it('takes the first family and maps generics', () => {
    expect(primaryFontFamily('"Open Sans", Arial')).toEqual({ name: 'Open Sans', generic: false });
    expect(primaryFontFamily("'Segoe UI'")).toEqual({ name: 'Segoe UI', generic: false });
    expect(primaryFontFamily('sans-serif')).toEqual({ name: 'Arial', generic: true });
    expect(primaryFontFamily('serif')).toEqual({ name: 'Times New Roman', generic: true });
    expect(primaryFontFamily('monospace')).toEqual({ name: 'Consolas', generic: true });
    expect(primaryFontFamily('system-ui')).toEqual({ name: 'Segoe UI', generic: true });
    expect(primaryFontFamily(null)).toEqual({ name: null, generic: false });
    expect(primaryFontFamily('')).toEqual({ name: null, generic: false });
  });
  it('skips leading system font aliases', () => {
    expect(
      primaryFontFamily("-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif"),
    ).toEqual({ name: 'Segoe UI', generic: false });
    expect(primaryFontFamily('-apple-system, sans-serif')).toEqual({ name: 'Arial', generic: true });
    expect(primaryFontFamily('-apple-system, BlinkMacSystemFont')).toEqual({ name: 'Segoe UI', generic: true });
  });
});

describe('fontToOoxml', () => {
  it('converts the default font', () => {
    expect(fontToOoxml(DEFAULT_HIGHCHARTS_FONT)).toEqual({
      typeface: 'Segoe UI',
      sizeHundredthsPt: 1200,
      bold: false,
      italic: false,
      colorHex: '333333',
    });
  });
  it('clamps sizes and reports generic families', () => {
    const r = fontToOoxml(
      { family: 'monospace', size: 1, bold: true, italic: true, color: null },
      'title.style.fontFamily',
    );
    expect(r).toMatchObject({ typeface: 'Consolas', sizeHundredthsPt: 100, bold: true, italic: true, colorHex: null });
    expect(r.diagnostic).toMatchObject({
      code: 'APPROXIMATED_FONT',
      outcome: 'approximated',
      property: 'title.style.fontFamily',
    });
    expect(fontToOoxml({ family: null, size: 10000, bold: false, italic: false, color: null }).sizeHundredthsPt).toBe(
      40000,
    );
    expect(fontToOoxml(null)).toEqual({
      typeface: null,
      sizeHundredthsPt: null,
      bold: false,
      italic: false,
      colorHex: null,
    });
  });
});
