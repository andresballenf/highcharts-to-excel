/**
 * Live Chart.js 4 instances (rendered into jsdom with a stub 2D context) → ChartModel IR.
 *
 * jsdom has no canvas: `installCanvasStub` gives Chart.js a no-op context so layout, parsing and
 * option resolution run for real. Time scales need a date adapter, which is not installed, so
 * time scales are covered by configuration tests only.
 */

import { Chart } from 'chart.js/auto';
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'vitest';
import { analyzeChartJsCompatibility, extractChartJsModel, isChartJsChart } from '../../src/chartjs';
import * as F from '../fixtures/chartjs-configs';

let restore: () => void;
const charts: Chart[] = [];

beforeAll(() => {
  restore = F.installCanvasStub();
});
afterAll(() => restore());
afterEach(() => {
  while (charts.length > 0) charts.pop()!.destroy();
  document.body.innerHTML = '';
});

// biome-ignore lint/suspicious/noExplicitAny: fixtures are loosely typed configurations
function render(config: any, width = 640, height = 320, id?: string): Chart {
  const chart = new Chart(F.makeCanvas(width, height, id), F.staticOptions(config));
  charts.push(chart);
  return chart;
}

const hex = (c: { r: number; g: number; b: number } | null | undefined): string | null =>
  c ? `#${[c.r, c.g, c.b].map((v) => v.toString(16).padStart(2, '0')).join('')}` : null;

describe('live Chart.js charts', () => {
  it('recognizes a Chart instance and reads its size, version and canvas id', () => {
    const chart = render(F.lineChart(), 640, 320, 'sales-canvas');
    expect(isChartJsChart(chart)).toBe(true);
    const m = extractChartJsModel(chart);
    expect(m).toMatchObject({ width: 640, height: 320 });
    expect(m.meta).toMatchObject({ sourceLibrary: 'chartjs', sourceVersion: Chart.version, chartId: 'sales-canvas' });
    expect(m.plotArea.box).not.toBeNull();
  });

  it('extracts the same data and styles as the configuration path', () => {
    const fromConfig = extractChartJsModel(F.lineChart());
    const live = extractChartJsModel(render(F.lineChart()));
    expect(live.series.map((s) => s.points.map((p) => [p.x, p.y]))).toEqual(
      fromConfig.series.map((s) => s.points.map((p) => [p.x, p.y])),
    );
    expect(live.series.map((s) => hex(s.color))).toEqual(fromConfig.series.map((s) => hex(s.color)));
    expect(live.series[0]!.marker).toEqual(fromConfig.series[0]!.marker);
    expect(live.xAxes[0]!.categories).toEqual(F.MONTHS);
    expect(live.title!.text).toBe('Monthly revenue');
    expect(live.yAxes[0]!.labels.format).toEqual(fromConfig.yAxes[0]!.labels.format);
  });

  it('does not treat Chart.js default tick callbacks as user formatters', () => {
    const m = extractChartJsModel(render(F.stackedBarChart()));
    expect(m.warnings.filter((d) => d.code === 'UNSUPPORTED_FORMATTER')).toEqual([]);
    expect(m.series.every((s) => s.stacking === 'normal')).toBe(true);
  });

  it('reads colors assigned by the colors plugin', () => {
    const m = extractChartJsModel(render(F.autoColorChart()));
    expect(hex(m.series[0]!.color)).toBe('#36a2eb');
    expect(m.series[0]!.fill).toMatchObject({ color: { a: 0.5 } });
    expect(hex(m.series[1]!.color)).toBe('#ff6384');
  });

  it('evaluates scriptable options through the rendered elements', () => {
    const config = F.perPointColorChart();
    config.data.datasets[0].backgroundColor = (ctx: { dataIndex: number }) =>
      ctx.dataIndex === 1 ? '#00ff00' : '#ff0000';
    const m = extractChartJsModel(render(config));
    expect(m.series[0]!.points.map((p) => hex(p.color))).toEqual([null, '#00ff00', null]);
    expect(m.warnings.some((d) => d.code === 'UNSUPPORTED_STYLE' && d.property.endsWith('backgroundColor'))).toBe(
      false,
    );
  });

  it('follows legend toggles: hidden datasets and hidden pie slices', () => {
    const chart = render(F.hiddenDatasetChart());
    chart.show(1);
    chart.hide(0);
    const m = extractChartJsModel(chart);
    expect(m.series.map((s) => s.name)).toEqual(['Hidden']);

    const pie = render(F.pieChart());
    pie.toggleDataVisibility(2);
    pie.update();
    const slices = extractChartJsModel(pie).series[0]!.points;
    expect(slices.map((p) => p.visible)).toEqual([true, true, false]);
  });

  it('maps mixed charts, horizontal bars and secondary axes through the live scales', () => {
    const mixed = extractChartJsModel(render(F.mixedBarLineChart()));
    expect(mixed.series.map((s) => s.kind)).toEqual(['column', 'line']);
    const horizontal = extractChartJsModel(render(F.horizontalBarChart()));
    expect(horizontal.inverted).toBe(true);
    expect(horizontal.series[0]!.kind).toBe('bar');
    expect(horizontal.series[0]!.points.map((p) => p.y)).toEqual([3, 7, 5]);
    const secondary = extractChartJsModel(render(F.secondaryAxisChart()));
    expect(secondary.yAxes.map((a) => [a.id, a.opposite])).toEqual([
      ['y', false],
      ['y1', true],
    ]);
    expect(secondary.series.map((s) => s.yAxisIndex)).toEqual([0, 1]);
  });

  it('reads the rendered doughnut cutout and scatter/bubble data', () => {
    expect(extractChartJsModel(render(F.doughnutChart('50%'))).series[0]!.pie!.innerSize).toBe(0.5);
    const scatter = extractChartJsModel(render(F.scatterChart()));
    expect(scatter.series[0]!.points.map((p) => [p.x, p.y])).toEqual([
      [-10, 0],
      [0, 10],
      [10, 5],
      [0.5, 5.5],
    ]);
    const bubble = extractChartJsModel(render(F.bubbleChart()));
    expect(bubble.series[0]!.points.map((p) => p.z)).toEqual([15, 10, 5]);
  });

  it('never mutates the chart configuration and reports radar charts as not editable', () => {
    const chart = render(F.lineChart());
    const before = JSON.stringify(chart.config.data);
    extractChartJsModel(chart);
    expect(JSON.stringify(chart.config.data)).toBe(before);

    const report = analyzeChartJsCompatibility(render(F.radarChart()));
    expect(report.editable).toBe(false);
    expect(report.blocking).toContain('chart.polar');
  });
});
