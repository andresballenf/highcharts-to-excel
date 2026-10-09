import { describe, expect, it } from 'vitest';
import { applyThemeOverrides } from '../../src/core/theme-overrides';
import * as F from '../fixtures/chart-models';

const rgbOf = (c: { r: number; g: number; b: number } | null | undefined) => (c ? [c.r, c.g, c.b] : null);

describe('applyThemeOverrides', () => {
  it('returns the same model without overrides', () => {
    const m = F.lineModel();
    expect(applyThemeOverrides(m, undefined)).toBe(m);
  });

  it('applies the palette to series (by index) and pie points', () => {
    const m = F.comboModel();
    const out = applyThemeOverrides(m, { colors: ['#ff0000', '#00ff00'] });
    expect(out.colors.map(rgbOf)).toEqual([
      [255, 0, 0],
      [0, 255, 0],
    ]);
    expect(rgbOf(out.series[0]!.color)).toEqual([255, 0, 0]);
    expect(out.series[0]!.fill).toEqual({ type: 'solid', color: expect.objectContaining({ r: 255, g: 0, b: 0 }) });
    expect(rgbOf(out.series[1]!.color)).toEqual([0, 255, 0]);
    expect(rgbOf(out.series[1]!.line!.color)).toEqual([0, 255, 0]);
    const pie = applyThemeOverrides(F.pieModel(), { colors: ['#111111', '#222222'] });
    expect(pie.series[0]!.points.map((p) => rgbOf(p.color))).toEqual([
      [17, 17, 17],
      [34, 34, 34],
      [17, 17, 17],
      [34, 34, 34],
    ]);
  });

  it('sets chart and plot backgrounds', () => {
    const out = applyThemeOverrides(F.lineModel(), { chartBackground: '#000000', plotBackground: 'white' });
    expect(out.background).toEqual({ type: 'solid', color: expect.objectContaining({ r: 0, g: 0, b: 0, a: 1 }) });
    expect(out.plotArea.background).toEqual({ type: 'solid', color: expect.objectContaining({ r: 255, g: 255, b: 255 }) });
  });

  it('applies fontFamily everywhere and merges component font overrides', () => {
    const m = F.styledModel();
    m.series[0] = { ...m.series[0]!, dataLabels: F.labels({ font: null }) };
    const out = applyThemeOverrides(m, {
      fontFamily: 'Arial',
      title: { size: 30, color: '#123456', bold: false },
      axisLabels: { italic: true },
      legend: { family: 'Calibri' },
      dataLabels: { size: 9 },
    });
    expect(out.title!.font).toMatchObject({ family: 'Arial', size: 30, bold: false, color: expect.objectContaining({ r: 0x12, g: 0x34, b: 0x56 }) });
    expect(out.subtitle!.font.family).toBe('Arial');
    expect(out.yAxes[0]!.labels.font).toMatchObject({ family: 'Arial', italic: true, size: 11 });
    // Null fonts are created from the Highcharts default font.
    expect(out.xAxes[0]!.labels.font).toMatchObject({ family: 'Arial', italic: true });
    expect(out.yAxes[0]!.title!.font.family).toBe('Arial');
    expect(out.legend.font!.family).toBe('Calibri');
    expect(out.series[0]!.dataLabels!.font).toMatchObject({ family: 'Arial', size: 9 });
  });

  it('updates gridlines and creates them on y axes when a width is given', () => {
    const m = F.lineModel();
    m.yAxes.push(F.axis({ index: 1, kind: 'linear' }));
    const out = applyThemeOverrides(m, { gridLineColor: '#ff00ff', gridLineWidth: 3 });
    expect(out.yAxes[0]!.gridLines).toEqual({ color: expect.objectContaining({ r: 255, g: 0, b: 255 }), width: 3, dash: 'solid' });
    expect(out.yAxes[1]!.gridLines).toMatchObject({ width: 3, color: expect.objectContaining({ r: 255, b: 255 }) });
    expect(out.xAxes[0]!.gridLines).toBeNull();
    const hidden = applyThemeOverrides(m, { gridLineWidth: 0 });
    expect(hidden.yAxes[0]!.gridLines!.width).toBe(0);
    expect(hidden.yAxes[1]!.gridLines).toBeNull();
  });

  it('applies per-series overrides', () => {
    const out = applyThemeOverrides(F.areaPercentModel(), { series: { 1: { color: '#abcdef', lineWidth: 4, fillOpacity: 0.3 } } });
    const s = out.series[1]!;
    expect(rgbOf(s.color)).toEqual([0xab, 0xcd, 0xef]);
    expect(s.line).toMatchObject({ width: 4 });
    expect(s.fillOpacity).toBe(0.3);
    expect(out.series[0]!.fillOpacity).toBe(0.75);
  });

  it('reports unparseable colors and never mutates the input', () => {
    const m = F.styledModel();
    const before = JSON.stringify(m);
    const out = applyThemeOverrides(m, {
      colors: ['nope'],
      chartBackground: 'also-bad',
      fontFamily: 'Arial',
      gridLineColor: '#000',
      gridLineWidth: 1,
      series: { 0: { color: '#00f', lineWidth: 1, fillOpacity: 0.1 } },
    });
    expect(JSON.stringify(m)).toBe(before);
    expect(m.warnings).toEqual([]);
    expect(out.warnings.map((w) => [w.code, w.property])).toEqual([
      ['UNRESOLVED_COLOR', 'themeOverrides.colors[0]'],
      ['UNRESOLVED_COLOR', 'themeOverrides.chartBackground'],
    ]);
    expect(out.background).toEqual(m.background);
  });
});
