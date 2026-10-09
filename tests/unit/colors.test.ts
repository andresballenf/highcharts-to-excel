import { describe, expect, it } from 'vitest';
import {
  HIGHCHARTS_CSS_VARIABLE_DEFAULTS,
  HIGHCHARTS_DEFAULT_PALETTE,
  colorToCss,
  colorToHex,
  flattenAlpha,
  isColorEqual,
  parseColor,
  withAlpha,
} from '../../src/utils/colors';
import { fillToSolidColor, resolveSeriesColor, toFill } from '../../src/translators/color-translator';
import { isFormulaLike, stripControlChars, truncate } from '../../src/utils/text';
import { clamp, emuToPx, ptToPx, pxToEmu, pxToHundredthsPt, pxToLineWidthEmu, pxToPt, roundTo } from '../../src/utils/units';

const rgba = (input: string) => {
  const c = parseColor(input);
  return c ? { r: c.r, g: c.g, b: c.b, a: c.a } : null;
};

describe('parseColor', () => {
  it('parses hex forms', () => {
    expect(rgba('#f00')).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(rgba('#f008')).toEqual({ r: 255, g: 0, b: 0, a: 0.5333 });
    expect(rgba('#2caffe')).toEqual({ r: 44, g: 175, b: 254, a: 1 });
    expect(rgba('#2CAFFE80')).toEqual({ r: 44, g: 175, b: 254, a: 0.502 });
    expect(rgba('#12')).toBeNull();
    expect(rgba('#gggggg')).toBeNull();
  });

  it('parses rgb()/rgba() in comma and space syntax', () => {
    expect(rgba('rgba(255, 0, 0, .5)')).toEqual({ r: 255, g: 0, b: 0, a: 0.5 });
    expect(rgba('rgb(10,20,30)')).toEqual({ r: 10, g: 20, b: 30, a: 1 });
    expect(rgba('rgb(10 20 30 / 25%)')).toEqual({ r: 10, g: 20, b: 30, a: 0.25 });
    expect(rgba('rgb(100%, 50%, 0%)')).toEqual({ r: 255, g: 128, b: 0, a: 1 });
    expect(rgba('RGBA(300, -5, 0, 2)')).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(rgba('rgb(1, 2)')).toBeNull();
  });

  it('parses hsl()/hsla()', () => {
    expect(rgba('hsl(0, 100%, 50%)')).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(rgba('hsl(120deg 100% 25%)')).toEqual({ r: 0, g: 128, b: 0, a: 1 });
    expect(rgba('hsla(240, 100%, 50%, 0.3)')).toEqual({ r: 0, g: 0, b: 255, a: 0.3 });
    expect(rgba('hsl(0.5turn, 100%, 50%)')).toEqual({ r: 0, g: 255, b: 255, a: 1 });
  });

  it('parses named colors case-insensitively, transparent and none', () => {
    expect(rgba('RebeccaPurple')).toEqual({ r: 102, g: 51, b: 153, a: 1 });
    expect(rgba('lightgoldenrodyellow')).toEqual({ r: 250, g: 250, b: 210, a: 1 });
    expect(rgba('transparent')).toEqual({ r: 0, g: 0, b: 0, a: 0 });
    expect(parseColor('none')).toBeNull();
    expect(parseColor('notacolor')).toBeNull();
    expect(parseColor(42)).toBeNull();
    expect(parseColor(undefined)).toBeNull();
  });

  it('keeps the original string as source', () => {
    expect(parseColor(' Red ')?.source).toBe(' Red ');
  });

  it('resolves var() via resolver, Highcharts defaults, then fallback', () => {
    const resolver = (name: string) => (name === '--brand' ? '#123456' : name === '--alias' ? 'var(--brand)' : null);
    expect(rgba('var(--highcharts-color-0)')).toEqual({ r: 44, g: 175, b: 254, a: 1 });
    expect(rgba('var(--highcharts-neutral-color-80)')).toEqual({ r: 51, g: 51, b: 51, a: 1 });
    const c = parseColor('var(--alias)', resolver);
    expect(c && colorToHex(c)).toBe('123456');
    expect(c?.source).toBe('var(--alias)');
    expect(rgba('var(--unknown, rgb(1, 2, 3))')).toEqual({ r: 1, g: 2, b: 3, a: 1 });
    expect(rgba('var(--unknown, var(--highcharts-color-1))')).toEqual({ r: 84, g: 79, b: 197, a: 1 });
    expect(parseColor('var(--unknown)')).toBeNull();
    // resolver overrides built-in defaults
    expect(colorToHex(parseColor('var(--highcharts-color-0)', () => 'black')!)).toBe('000000');
    // light-dark() as emitted by Highcharts 13 CSS uses the light value
    expect(colorToHex(parseColor('var(--highcharts-background-color)', () => 'light-dark(#ffffff, #141414)')!)).toBe('FFFFFF');
  });

  it('stops recursion at depth 4', () => {
    const loop = (name: string) => `var(${name})`;
    expect(parseColor('var(--a)', loop)).toBeNull();
  });
});

describe('palette constants', () => {
  it('exposes the Highcharts palette and CSS variables', () => {
    expect(HIGHCHARTS_DEFAULT_PALETTE).toHaveLength(10);
    expect(HIGHCHARTS_DEFAULT_PALETTE[0]).toBe('#2caffe');
    expect(HIGHCHARTS_CSS_VARIABLE_DEFAULTS['--highcharts-color-9']).toBe('#91e8e1');
    expect(HIGHCHARTS_CSS_VARIABLE_DEFAULTS['--highcharts-neutral-color-10']).toBe('#e6e6e6');
    expect(HIGHCHARTS_CSS_VARIABLE_DEFAULTS['--highcharts-highlight-color-100']).toBe('#0022ff');
  });
});

describe('color helpers', () => {
  const red = { r: 255, g: 0, b: 0, a: 0.5 };
  const white = { r: 255, g: 255, b: 255, a: 1 };
  it('formats hex and css', () => {
    expect(colorToHex({ r: 44, g: 175, b: 254, a: 0.2 })).toBe('2CAFFE');
    expect(colorToCss(red)).toBe('rgba(255, 0, 0, 0.5)');
  });
  it('flattens alpha over a background', () => {
    expect(flattenAlpha(red, white)).toEqual({ r: 255, g: 128, b: 128, a: 1 });
    expect(flattenAlpha({ r: 0, g: 0, b: 0, a: 0 }, white)).toEqual({ r: 255, g: 255, b: 255, a: 1 });
  });
  it('withAlpha and isColorEqual', () => {
    expect(withAlpha(white, 0.25)).toEqual({ r: 255, g: 255, b: 255, a: 0.25 });
    expect(withAlpha(white, 3).a).toBe(1);
    expect(isColorEqual(parseColor('#ff0000')!, parseColor('red')!)).toBe(true);
    expect(isColorEqual(parseColor('#ff0000')!, parseColor('rgba(255,0,0,0.5)')!)).toBe(false);
  });
});

describe('toFill', () => {
  it('handles solid colors, none and unresolvable strings', () => {
    expect(toFill('#ff0000').fill).toMatchObject({ type: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } });
    expect(toFill(undefined)).toEqual({ fill: null });
    expect(toFill(null)).toEqual({ fill: null });
    expect(toFill('none')).toEqual({ fill: null });
    const bad = toFill('bogus', undefined, 'series[0].color');
    expect(bad.fill).toBeNull();
    expect(bad.diagnostic).toMatchObject({ code: 'UNRESOLVED_COLOR', outcome: 'approximated', property: 'series[0].color' });
  });

  it('converts linear gradients with an angle', () => {
    const r = toFill({ linearGradient: { x1: 0, y1: 0, x2: 0, y2: 1 }, stops: [[1, '#000000'], [0, '#ffffff']] });
    expect(r.diagnostic).toBeUndefined();
    expect(r.fill).toMatchObject({
      type: 'gradient',
      angle: 90,
      stops: [
        { offset: 0, color: { r: 255, g: 255, b: 255 } },
        { offset: 1, color: { r: 0, g: 0, b: 0 } },
      ],
    });
    expect(toFill({ linearGradient: [0, 0, 1, 0], stops: [[0, 'red'], [1, 'blue']] }).fill).toMatchObject({ angle: 0 });
    expect(toFill({ linearGradient: [0, 0, 1, 1], stops: [[0, 'red'], [1, 'blue']] }).fill).toMatchObject({ angle: 45 });
    expect(toFill({ linearGradient: [1, 0, 0, 0], stops: [[0, 'red']] }).fill).toMatchObject({ angle: 180 });
  });

  it('approximates radial gradients', () => {
    const r = toFill({ radialGradient: { cx: 0.5, cy: 0.5, r: 0.5 }, stops: [[0, 'red'], [1, 'var(--highcharts-color-0)']] }, undefined, 'series[1].color');
    expect(r.fill).toMatchObject({ type: 'gradient', angle: 90 });
    expect(r.diagnostic).toMatchObject({ code: 'UNSUPPORTED_GRADIENT', outcome: 'approximated', property: 'series[1].color.radialGradient' });
  });

  it('fillToSolidColor', () => {
    expect(fillToSolidColor(null)).toBeNull();
    expect(fillToSolidColor({ type: 'none' })).toBeNull();
    expect(colorToHex(fillToSolidColor(toFill({ linearGradient: [0, 0, 0, 1], stops: [[0, '#abcdef'], [1, '#000']] }).fill)!)).toBe('ABCDEF');
  });

  it('resolveSeriesColor prefers explicit color, then palette', () => {
    const palette = [parseColor('#111111')!, parseColor('#222222')!];
    expect(colorToHex(resolveSeriesColor('#ff0000', 0, palette)!)).toBe('FF0000');
    expect(colorToHex(resolveSeriesColor(undefined, 3, palette)!)).toBe('222222');
    expect(colorToHex(resolveSeriesColor('garbage', 0, palette)!)).toBe('111111');
    expect(colorToHex(resolveSeriesColor(undefined, 11, [])!)).toBe('544FC5');
  });
});

describe('text utils', () => {
  it('isFormulaLike', () => {
    for (const s of ['=SUM(A1)', '+1', '-5', '@x', '\tfoo', '\rfoo', '  =1']) expect(isFormulaLike(s)).toBe(true);
    for (const s of ['', 'abc', 'a=b', '5', ' x']) expect(isFormulaLike(s)).toBe(false);
  });
  it('stripControlChars keeps tab/LF/CR and removes invalid XML chars', () => {
    expect(stripControlChars('a\u0000b\u0007c\td\ne\rf￾g\uD800h😀')).toBe('abc\td\ne\rfgh😀');
  });
  it('truncate', () => {
    expect(truncate('abcdef', 3)).toBe('abc');
    expect(truncate('abc', 10)).toBe('abc');
    expect(truncate('ab😀', 3)).toBe('ab');
    expect(truncate('abc', 0)).toBe('');
  });
});

describe('units', () => {
  it('converts', () => {
    expect(pxToEmu(1)).toBe(9525);
    expect(emuToPx(19050)).toBe(2);
    expect(pxToPt(16)).toBe(12);
    expect(ptToPx(12)).toBe(16);
    expect(pxToHundredthsPt(12)).toBe(900);
    expect(pxToLineWidthEmu(1.5)).toBe(14288);
    expect(clamp(5, 0, 3)).toBe(3);
    expect(clamp(-1, 0, 3)).toBe(0);
    expect(roundTo(1.23456, 2)).toBe(1.23);
  });
});
