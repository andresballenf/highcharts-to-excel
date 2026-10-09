/**
 * Chart.js configuration → ChartModel IR (no rendering).
 */

import { describe, expect, it } from 'vitest';
import { extractChartJsModel } from '../../src/chartjs';
import type { ChartModel } from '../../src/types/chart-model';
import { DiagnosticCollector, type Diagnostic } from '../../src/types/diagnostics';
import * as F from '../fixtures/chartjs-configs';

function extract(config: unknown, options: Parameters<typeof extractChartJsModel>[1] = {}): ChartModel {
  return extractChartJsModel(config, options);
}

function codes(model: ChartModel): string[] {
  return model.warnings.map((d) => `${d.code}@${d.property}`);
}

function find(model: ChartModel, code: Diagnostic['code'], property?: string): Diagnostic | undefined {
  return model.warnings.find((d) => d.code === code && (property === undefined || d.property === property));
}

const hex = (c: { r: number; g: number; b: number } | null | undefined): string | null =>
  c ? `#${[c.r, c.g, c.b].map((v) => v.toString(16).padStart(2, '0')).join('')}` : null;

describe('extractChartJsModel: input and meta', () => {
  it('rejects inputs that are neither a chart nor a configuration', () => {
    expect(() => extract({ foo: 1 })).toThrowError(/Chart\.js chart instance or a Chart\.js configuration/);
    expect(() => extract(null)).toThrow();
  });

  it('marks the model as coming from Chart.js', () => {
    const m = extract(F.lineChart());
    expect(m.meta.sourceLibrary).toBe('chartjs');
    expect(m.meta.sourceChartType).toBe('line');
    expect(m.meta.extraction).toBe('headless');
    expect(m.meta.sourceVersion).toBeNull();
  });

  it('never mutates the configuration', () => {
    for (const make of [F.lineChart, F.stackedBarChart, F.pieChart, F.autoColorChart, F.timeScaleChart]) {
      const config = make();
      const before = JSON.stringify(config);
      extract(config);
      expect(JSON.stringify(config)).toBe(before);
    }
  });

  it('sizes the chart from options.aspectRatio with a 600px default width', () => {
    expect(extract(F.lineChart())).toMatchObject({ width: 600, height: 300 });
    expect(extract(F.pieChart())).toMatchObject({ width: 600, height: 600 });
    const c = F.lineChart();
    c.options = { ...c.options, aspectRatio: 3 };
    expect(extract(c, { chartWidth: 900 })).toMatchObject({ width: 900, height: 300 });
  });
});

describe('extractChartJsModel: line charts and styles', () => {
  const m = extract(F.lineChart());

  it('maps labels to a category axis and data to points', () => {
    expect(m.xAxes).toHaveLength(1);
    expect(m.xAxes[0]!.kind).toBe('category');
    expect(m.xAxes[0]!.categories).toEqual(F.MONTHS);
    const s = m.series[0]!;
    expect(s.kind).toBe('line');
    expect(s.name).toBe('Revenue');
    expect(s.points.map((p) => p.y)).toEqual([12, 19, 3, 5, 2, 3]);
    expect(s.points.map((p) => p.x)).toEqual([0, 1, 2, 3, 4, 5]);
    expect(s.points[2]!.name).toBe('Mar');
  });

  it('translates colors, widths, dashes and markers', () => {
    const [rev, cost] = m.series;
    expect(hex(rev!.color)).toBe('#ff6384');
    expect(rev!.line).toMatchObject({ width: 2, dash: 'solid' });
    expect(rev!.marker).toMatchObject({ enabled: true, symbol: 'diamond', radius: 5 });
    expect(rev!.marker!.fill).toMatchObject({ r: 255, g: 99, b: 132, a: 0.5 });
    expect(cost!.line!.width).toBe(3);
    // [6, 4] on a 3px line: 2x / 1.3x the width → Excel's short dash preset.
    expect(cost!.line!.dash).toBe('shortdash');
    expect(find(m, 'UNSUPPORTED_STYLE', 'data.datasets[1].borderDash')).toBeDefined();
  });

  it('reads title, subtitle, legend and fonts (Chart.js defaults: Helvetica Neue 12px #666)', () => {
    expect(m.title).toMatchObject({ text: 'Monthly revenue', align: 'center' });
    expect(m.title!.font).toMatchObject({ size: 18, bold: true });
    expect(hex(m.title!.font.color)).toBe('#111111');
    expect(m.subtitle!.text).toBe('FY 2024');
    expect(m.subtitle!.font).toMatchObject({ size: 12, bold: false });
    expect(m.subtitle!.font.family).toContain('Helvetica Neue');
    expect(hex(m.subtitle!.font.color)).toBe('#666666');
    expect(m.legend).toMatchObject({ enabled: true, position: 'bottom', layout: 'horizontal' });
    expect(m.legend.font).toMatchObject({ size: 14 });
  });

  it('reads axis titles, bounds, grid and the Intl tick format', () => {
    const [x] = m.xAxes;
    const [y] = m.yAxes;
    expect(x!.title!.text).toBe('Month');
    expect(x!.gridLines).toBeNull();
    expect(y!.title!.text).toBe('USD');
    expect(hex(y!.title!.font.color)).toBe('#ff0000');
    expect(y).toMatchObject({ min: 0, max: 25, kind: 'linear' });
    expect(y!.gridLines).toMatchObject({ width: 2 });
    expect(hex(y!.gridLines!.color)).toBe('#dddddd');
    expect(y!.labels.format).toMatchObject({ kind: 'excel', code: '[$$]#,##0' });
    expect(m.series[0]!.yFormat).toMatchObject({ code: '[$$]#,##0' });
  });

  it('defaults the background to white and leaves the title off unless displayed', () => {
    expect(m.background).toMatchObject({ type: 'solid', color: { r: 255, g: 255, b: 255 } });
    expect(extract(F.hiddenDatasetChart()).title).toBeNull();
  });

  it('maps tension to smooth lines and fill to areas', () => {
    const s = extract(F.smoothAreaChart()).series[0]!;
    expect(s.kind).toBe('areaspline');
    expect(s.smooth).toBe(true);
    expect(s.fill).toMatchObject({ type: 'solid' });
    expect(s.fillOpacity).toBe(1);
  });
});

describe('extractChartJsModel: colors plugin and defaults', () => {
  it('applies the Chart.js colors plugin palette to datasets without colors', () => {
    const m = extract(F.autoColorChart());
    expect(m.series[0]!.fill).toMatchObject({ color: { r: 54, g: 162, b: 235, a: 0.5 } });
    expect(m.series[1]!.fill).toMatchObject({ color: { r: 255, g: 99, b: 132, a: 0.5 } });
    expect(m.series[0]!.border!.color).toMatchObject({ r: 54, g: 162, b: 235, a: 1 });
    expect(m.colors).toHaveLength(7);
  });

  it('falls back to rgba(0,0,0,0.1) when the colors plugin is disabled', () => {
    const c = F.autoColorChart();
    c.options = { plugins: { colors: { enabled: false } } };
    const m = extract(c);
    expect(m.series[0]!.fill).toMatchObject({ color: { r: 0, g: 0, b: 0, a: 0.1 } });
  });

  it('does not recolor when any dataset defines a color (unless forceOverride)', () => {
    const c = F.autoColorChart();
    (c.data.datasets[0] as { backgroundColor?: string }).backgroundColor = '#123456';
    expect(hex(extract(c).series[1]!.color)).toBe('#000000');
    c.options = { plugins: { colors: { forceOverride: true } } };
    expect(hex(extract(c).series[0]!.color)).toBe('#36a2eb');
  });
});

describe('extractChartJsModel: bar charts', () => {
  it('maps stacked bars with stack groups', () => {
    const m = extract(F.stackedBarChart());
    expect(m.series.map((s) => s.kind)).toEqual(['column', 'column', 'column']);
    expect(m.series.every((s) => s.stacking === 'normal' && s.stackGroup === 'a')).toBe(true);
    expect(m.series[0]!.bars).toMatchObject({ pointPadding: 0.05, groupPadding: 0.1, borderRadius: 0 });
    expect(m.yAxes[0]!.min).toBe(0);
  });

  it('maps indexAxis "y" to horizontal bars and an inverted chart', () => {
    const m = extract(F.horizontalBarChart());
    expect(m.inverted).toBe(true);
    expect(m.series[0]!.kind).toBe('bar');
    expect(m.xAxes[0]!.kind).toBe('category');
    expect(m.xAxes[0]!.id).toBe('y');
    expect(m.yAxes[0]!.id).toBe('x');
  });

  it('keeps per-point colors', () => {
    const s = extract(F.perPointColorChart()).series[0]!;
    expect(hex(s.color)).toBe('#ff0000');
    expect(s.points[0]!.color).toBeNull();
    expect(hex(s.points[1]!.color)).toBe('#008000');
    expect(hex(s.points[2]!.color)).toBe('#0000ff');
    expect(s.border).toMatchObject({ width: 1 });
  });

  it('honours parsing keys (including dotted paths)', () => {
    const m = extract(F.parsingKeysChart());
    const s = m.series[0]!;
    expect(s.points.map((p) => [p.name, p.y])).toEqual([
      ['Jan', 5],
      ['Feb', 7],
    ]);
  });

  it('maps mixed bar + line (dataset type overrides chart type)', () => {
    const m = extract(F.mixedBarLineChart());
    expect(m.series.map((s) => [s.kind, s.sourceType])).toEqual([
      ['column', 'bar'],
      ['line', 'line'],
    ]);
    expect(m.meta.sourceChartType).toBe('bar');
  });

  it('binds yAxisID to a secondary, opposite value axis', () => {
    const m = extract(F.secondaryAxisChart());
    expect(m.yAxes.map((a) => [a.id, a.opposite])).toEqual([
      ['y', false],
      ['y1', true],
    ]);
    expect(m.series.map((s) => s.yAxisIndex)).toEqual([0, 1]);
    expect(m.yAxes[1]!.title!.text).toBe('°C');
  });
});

describe('extractChartJsModel: pie family, scatter, bubble', () => {
  it('maps pie slices with names, colors and offsets', () => {
    const s = extract(F.pieChart()).series[0]!;
    expect(s.kind).toBe('pie');
    expect(s.points.map((p) => [p.name, p.y])).toEqual([
      ['Chrome', 62],
      ['Firefox', 20],
      ['Safari', 18],
    ]);
    expect(s.points.map((p) => hex(p.color))).toEqual(['#36a2eb', '#ff6384', '#ffcd56']);
    expect(s.points[1]!.sliced).toBe(12);
    expect(s.border).toMatchObject({ width: 2, color: { r: 255, g: 255, b: 255 } });
  });

  it('maps a doughnut cutout (percent or px) to innerSize and rotation/circumference to angles', () => {
    const s = extract(F.doughnutChart('60%')).series[0]!;
    expect(s.kind).toBe('doughnut');
    expect(s.pie).toEqual({ innerSize: 0.6, startAngle: 90, endAngle: 270 });
    const px = extract(F.doughnutChart(150));
    expect(px.series[0]!.pie!.innerSize).toBeCloseTo(0.5);
    expect(find(px, 'APPROXIMATED_LAYOUT', 'data.datasets[0].cutout')).toBeDefined();
    // Colors plugin colors doughnut slices individually.
    expect(s.points.map((p) => hex(p.color))).toEqual(['#36a2eb', '#ff6384', '#ff9f40']);
  });

  it('maps scatter {x, y} points on linear axes', () => {
    const m = extract(F.scatterChart());
    const s = m.series[0]!;
    expect(s.kind).toBe('scatter');
    expect(s.points.map((p) => [p.x, p.y])).toEqual([
      [-10, 0],
      [0, 10],
      [10, 5],
      [0.5, 5.5],
    ]);
    expect(m.xAxes[0]!.kind).toBe('linear');
    expect(s.line!.width).toBe(0);
    expect(s.marker).toMatchObject({ enabled: true, symbol: 'triangle', radius: 3 });
  });

  it('maps bubble [x, y, r] array data (one parser per dataset, as Chart.js)', () => {
    const m = extract({
      type: 'bubble',
      data: { datasets: [{ label: 'B', data: [[1, 2, 3], [4, 5], { x: 9, y: 9, r: 9 }] }] },
    });
    expect(m.series[0]!.points.map((p) => [p.x, p.y, p.z])).toEqual([
      [1, 2, 3],
      [4, 5, 3],
      [null, null, 3],
    ]);
  });

  it('maps bubble {x, y, r} points', () => {
    const m = extract(F.bubbleChart());
    expect(m.series[0]!.kind).toBe('bubble');
    expect(m.series[0]!.points.map((p) => [p.x, p.y, p.z])).toEqual([
      [20, 30, 15],
      [40, 10, 10],
      [30, 20, 5],
    ]);
    expect(find(m, 'APPROXIMATED_LAYOUT', 'data.datasets[0].data')).toBeDefined();
  });

  it('maps radar and polarArea to unknown series on a polar chart', () => {
    for (const make of [F.radarChart, F.polarAreaChart]) {
      const m = extract(make());
      expect(m.polar).toBe(true);
      expect(m.series[0]!.kind).toBe('unknown');
      expect(m.xAxes).toHaveLength(0);
      expect(find(m, 'UNSUPPORTED_SERIES_TYPE', 'data.datasets[0].type')).toBeDefined();
    }
  });
});

describe('extractChartJsModel: data semantics', () => {
  it('excludes hidden datasets by default and includes them with seriesVisibility "all"', () => {
    const m = extract(F.hiddenDatasetChart());
    expect(m.series.map((s) => s.name)).toEqual(['Shown']);
    expect(find(m, 'HIDDEN_SERIES_EXCLUDED', 'data.datasets[1].hidden')).toBeDefined();
    const all = extract(F.hiddenDatasetChart(), { seriesVisibility: 'all' });
    expect(all.series.map((s) => [s.name, s.visible])).toEqual([
      ['Shown', true],
      ['Hidden', false],
    ]);
    expect(find(all, 'HIDDEN_SERIES_INCLUDED')).toBeDefined();
  });

  it('keeps nulls as empty points and reports spanGaps and non-numeric values', () => {
    const m = extract(F.nullsChart());
    const s = m.series[0]!;
    expect(s.points.map((p) => [p.y, p.isNull])).toEqual([
      [1, false],
      [null, true],
      [3, false],
      [null, true],
    ]);
    expect(find(m, 'UNSUPPORTED_STYLE', 'data.datasets[0].spanGaps')).toBeDefined();
    expect(find(m, 'NON_NUMERIC_VALUE', 'data.datasets[0].data')).toMatchObject({ details: { count: 1 } });
  });

  it('maps a time scale to a datetime axis with wall-clock ms', () => {
    const m = extract(F.timeScaleChart());
    expect(m.xAxes[0]!.kind).toBe('datetime');
    expect(m.series[0]!.points.map((p) => p.x)).toEqual([
      Date.UTC(2024, 0, 1),
      Date.UTC(2024, 0, 2),
      Date.UTC(2024, 0, 5),
    ]);
    expect(m.xAxes[0]!.labels.format).toMatchObject({ kind: 'excel', code: 'mmm d' });
    expect(m.meta.datetimeOffsetMinutes).toBe(0);
  });

  it('parses time labels with primitive data and millisecond x values', () => {
    const ms = Date.UTC(2024, 5, 1, 12);
    const m = extract({
      type: 'line',
      data: { labels: ['2024-06-01T00:00', ms], datasets: [{ label: 'T', data: [1, 2] }] },
      options: { scales: { x: { type: 'time' } } },
    });
    const offset = -new Date(ms).getTimezoneOffset() * 60_000;
    expect(m.series[0]!.points.map((p) => p.x)).toEqual([Date.UTC(2024, 5, 1), ms + offset]);
  });
});

describe('extractChartJsModel: unsupported features are reported with precise paths', () => {
  const diagnostics = new DiagnosticCollector();
  const m = extract(F.callbackChart(), { diagnostics });

  it('reports tick callbacks, scriptable colors, tooltips and datalabels', () => {
    const list = codes(m);
    expect(list).toContain('UNSUPPORTED_FORMATTER@options.scales.y.ticks.callback');
    expect(list).toContain('UNSUPPORTED_STYLE@data.datasets[0].backgroundColor');
    expect(list).toContain('UNSUPPORTED_STYLE@data.datasets[0].borderColor');
    expect(list).toContain('UNSUPPORTED_TOOLTIP@options.plugins.tooltip.callbacks.label');
    expect(list).toContain('UNSUPPORTED_STYLE@options.plugins.datalabels');
    expect(diagnostics.items.length).toBe(m.warnings.length);
  });

  it('keeps logarithmic and reversed scales', () => {
    expect(m.yAxes[0]).toMatchObject({ kind: 'logarithmic', logBase: 10, reversed: true });
    expect(m.yAxes[0]!.labels.format).toMatchObject({ kind: 'unsupported' });
  });

  it('reports an empty chart as blocking', () => {
    const empty = extract({ type: 'bar', data: { labels: [], datasets: [] } });
    expect(find(empty, 'EMPTY_CHART')).toMatchObject({ outcome: 'blocking' });
  });
});
