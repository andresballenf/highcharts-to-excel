import { describe, expect, it } from 'vitest';
import {
  dashStyleFromHighcharts,
  dashStyleToOoxml,
  markerRadiusToOoxmlSize,
  markerSymbolFromHighcharts,
  markerSymbolToOoxml,
  strokeFromOptions,
} from '../../src/translators/style-translator';
import { colorToHex, parseColor } from '../../src/utils/colors';

describe('dash styles', () => {
  it('maps Highcharts names case-insensitively', () => {
    expect(dashStyleFromHighcharts('ShortDashDotDot')).toBe('shortdashdotdot');
    expect(dashStyleFromHighcharts('longdash')).toBe('longdash');
    expect(dashStyleFromHighcharts('Dot')).toBe('dot');
    expect(dashStyleFromHighcharts('Wavy')).toBe('solid');
    expect(dashStyleFromHighcharts(undefined)).toBe('solid');
  });
  it('maps to OOXML prstDash', () => {
    expect(dashStyleToOoxml('solid')).toBe('solid');
    expect(dashStyleToOoxml('dashdot')).toBe('dashDot');
    expect(dashStyleToOoxml('longdash')).toBe('lgDash');
    expect(dashStyleToOoxml('longdashdotdot')).toBe('lgDashDotDot');
    expect(dashStyleToOoxml('shortdash')).toBe('sysDash');
    expect(dashStyleToOoxml('shortdot')).toBe('sysDot');
    expect(dashStyleToOoxml('shortdashdotdot')).toBe('sysDashDotDot');
  });
});

describe('markers', () => {
  it('maps Highcharts symbols', () => {
    expect(markerSymbolFromHighcharts(undefined)).toBe('circle');
    expect(markerSymbolFromHighcharts('Diamond')).toBe('diamond');
    expect(markerSymbolFromHighcharts('triangle-down')).toBe('triangle-down');
    expect(markerSymbolFromHighcharts('url(https://example.com/sun.png)')).toBe('other');
  });
  it('maps to OOXML with approximations', () => {
    expect(markerSymbolToOoxml('square')).toEqual({ symbol: 'square' });
    expect(markerSymbolToOoxml('none')).toEqual({ symbol: 'none' });
    const down = markerSymbolToOoxml('triangle-down', 'series[2].marker.symbol');
    expect(down.symbol).toBe('triangle');
    expect(down.diagnostic).toMatchObject({
      code: 'APPROXIMATED_MARKER',
      outcome: 'approximated',
      property: 'series[2].marker.symbol',
    });
    expect(markerSymbolToOoxml('other')).toMatchObject({
      symbol: 'circle',
      diagnostic: { code: 'APPROXIMATED_MARKER' },
    });
  });
  it('converts radius to c:size', () => {
    expect(markerRadiusToOoxmlSize(4)).toBe(6);
    expect(markerRadiusToOoxmlSize(0)).toBe(2);
    expect(markerRadiusToOoxmlSize(100)).toBe(72);
  });
});

describe('strokeFromOptions', () => {
  it('reads options and falls back', () => {
    const s = strokeFromOptions({ color: '#ff0000', width: 2, dashStyle: 'Dash' }, {});
    expect(colorToHex(s.color!)).toBe('FF0000');
    expect(s.width).toBe(2);
    expect(s.dash).toBe('dash');

    const fallbackColor = parseColor('#00ff00')!;
    const f = strokeFromOptions({ width: 'bad' }, { color: fallbackColor, width: 3, dash: 'dot' });
    expect(f).toEqual({ color: fallbackColor, width: 3, dash: 'dot' });
    expect(strokeFromOptions({}, {})).toEqual({ color: null, width: 1, dash: 'solid' });
    expect(strokeFromOptions({ width: 0 }, { width: 5 }).width).toBe(0);
    expect(colorToHex(strokeFromOptions({ color: 'var(--highcharts-neutral-color-10)' }, {}).color!)).toBe('E6E6E6');
  });
});
