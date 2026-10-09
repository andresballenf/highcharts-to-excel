/**
 * Chart.js option edge cases → IR values and diagnostics (configuration input).
 */

import { describe, expect, it } from 'vitest';
import { extractChartJsModel } from '../../src/chartjs';
import type { ChartModel } from '../../src/types/chart-model';

// biome-ignore lint/suspicious/noExplicitAny: loose configurations for edge cases
function extract(config: any, options: Parameters<typeof extractChartJsModel>[1] = {}): ChartModel {
  return extractChartJsModel(config, options);
}

function has(m: ChartModel, code: string, property: string): boolean {
  return m.warnings.some((d) => d.code === code && d.property === property);
}

const line = (dataset: Record<string, unknown>, options: Record<string, unknown> = {}) => ({
  type: 'line',
  data: { labels: ['a', 'b', 'c'], datasets: [{ label: 'S', data: [1, 2, 3], ...dataset }] },
  options,
});

describe('data shapes', () => {
  it('reads object data ({ label: value }) and floating bars', () => {
    const m = extract({ type: 'bar', data: { datasets: [{ label: 'O', data: { Jan: 4, Feb: 6 } }] } });
    expect(m.series[0]!.points.map((p) => [p.name, p.y])).toEqual([
      ['Jan', 4],
      ['Feb', 6],
    ]);
    const floating = extract({
      type: 'bar',
      data: {
        labels: ['a', 'b'],
        datasets: [
          {
            label: 'F',
            data: [
              [1, 3],
              [2, 5],
            ],
          },
        ],
      },
    });
    expect(floating.series[0]!.points.every((p) => p.y === null && p.isNull)).toBe(true);
    expect(has(floating, 'NON_NUMERIC_VALUE', 'data.datasets[0].data')).toBe(true);
  });

  it('exports decimated or original data by dataMode', () => {
    const ds = {
      label: 'D',
      data: [{ x: 0, y: 1 }],
      _decimated: true,
      _data: [
        { x: 0, y: 1 },
        { x: 1, y: 2 },
      ],
    };
    const config = { type: 'line', data: { datasets: [ds] }, options: { scales: { x: { type: 'linear' } } } };
    const rendered = extract(config);
    expect(rendered.series[0]!.points).toHaveLength(1);
    expect(rendered.series[0]!.dataSemantics).toMatchObject({ grouped: true, sourcePointCount: 2 });
    expect(has(rendered, 'DATA_GROUPED', 'data.datasets[0].data')).toBe(true);
    const raw = extract(config, { dataMode: 'raw' });
    expect(raw.series[0]!.points.map((p) => [p.x, p.y])).toEqual([
      [0, 1],
      [1, 2],
    ]);
  });

  it('parses Date objects, numeric strings and free-form date strings on time scales', () => {
    const d = new Date(Date.UTC(2024, 2, 1));
    const m = extract({
      type: 'line',
      data: {
        datasets: [
          {
            label: 'T',
            data: [
              { x: d, y: 1 },
              { x: String(d.getTime()), y: 2 },
              { x: 'March 3, 2024', y: 3 },
              { x: 'not a date', y: 4 },
            ],
          },
        ],
      },
      options: {
        scales: {
          x: { type: 'timeseries', time: { parser: 'yyyy', unit: 'quarter' }, adapters: { date: { zone: 'UTC' } } },
        },
      },
    });
    const off = (ms: number) => -new Date(ms).getTimezoneOffset() * 60_000;
    const xs = m.series[0]!.points.map((p) => p.x);
    expect(xs.slice(0, 2)).toEqual([d.getTime() + off(d.getTime()), d.getTime() + off(d.getTime())]);
    expect(xs[2]).toBe(Date.UTC(2024, 2, 3));
    expect(xs[3]).toBeNull();
    expect(has(m, 'NON_NUMERIC_VALUE', 'data.datasets[0].data')).toBe(true);
    expect(has(m, 'APPROXIMATED_AXIS_SCALE', 'options.scales.x.type')).toBe(true);
    expect(has(m, 'APPROXIMATED_DATETIME', 'options.scales.x.time.parser')).toBe(true);
    expect(has(m, 'APPROXIMATED_DATETIME', 'options.scales.x.adapters.date.zone')).toBe(true);
    expect(has(m, 'APPROXIMATED_NUMBER_FORMAT', 'options.scales.x.time.displayFormats.quarter')).toBe(false);
  });
});

describe('dataset styles', () => {
  it.each([
    [[2, 2], 2, 'shortdot'],
    [[1, 3], 1, 'dot'],
    [[1, 1], 1, 'shortdot'],
    [[12, 4], 2, 'longdash'],
    [[10, 4, 2, 4], 2, 'longdashdot'],
    [[4, 2, 1, 2], 2, 'dashdot'],
    [[2, 2, 2, 2], 2, 'shortdashdot'],
    [[2, 2, 1, 2, 1, 2], 2, 'shortdashdotdot'],
    [[5], 1, 'longdash'],
  ])('borderDash %j at width %d → %s', (dash, width, style) => {
    expect(extract(line({ borderDash: dash, borderWidth: width })).series[0]!.line!.dash).toBe(style);
  });

  it('reports stepped lines, fill targets, scatter fill, bar thickness and unresolved colors', () => {
    const m = extract(
      line({ stepped: true, fill: 'end', borderColor: 'nonsense-color', backgroundColor: { pattern: 1 } }),
    );
    expect(has(m, 'UNSUPPORTED_STYLE', 'data.datasets[0].stepped')).toBe(true);
    expect(has(m, 'APPROXIMATED_CHART_TYPE', 'data.datasets[0].fill')).toBe(true);
    expect(has(m, 'UNRESOLVED_COLOR', 'data.datasets[0].borderColor')).toBe(true);
    expect(has(m, 'UNSUPPORTED_GRADIENT', 'data.datasets[0].backgroundColor')).toBe(true);
    expect(m.series[0]!.kind).toBe('area');

    const stackedFill = extract(line({ fill: '-1' }, { scales: { y: { stacked: true } } }));
    expect(has(stackedFill, 'APPROXIMATED_CHART_TYPE', 'data.datasets[0].fill')).toBe(false);
    expect(stackedFill.series[0]!.stacking).toBe('normal');

    const scatter = extract({
      type: 'scatter',
      data: { datasets: [{ label: 'S', data: [{ x: 1, y: 1 }], fill: true, showLine: true, borderWidth: 2 }] },
    });
    expect(has(scatter, 'UNSUPPORTED_STYLE', 'data.datasets[0].fill')).toBe(true);
    expect(scatter.series[0]!.line!.width).toBe(2);

    const bar = extract({
      type: 'bar',
      data: {
        labels: ['a'],
        datasets: [{ label: 'B', data: [1], barThickness: 10, borderRadius: { topLeft: 4, topRight: 6 } }],
      },
    });
    expect(has(bar, 'APPROXIMATED_LAYOUT', 'data.datasets[0].barThickness')).toBe(true);
    expect(bar.series[0]!.bars!.borderRadius).toBe(6);
  });

  it('resolves options.datasets / options.elements and per-point marker overrides', () => {
    const m = extract(
      line(
        { pointRadius: [3, 6, 0], pointStyle: ['circle', 'rect', 'star'] },
        {
          datasets: { line: { borderWidth: 5 } },
          elements: { point: { backgroundColor: '#010203' } },
        },
      ),
    );
    const s = m.series[0]!;
    expect(s.line!.width).toBe(5);
    expect(s.marker!.fill).toMatchObject({ r: 1, g: 2, b: 3 });
    expect(s.points[1]!.marker).toMatchObject({ radius: 6, symbol: 'square' });
    expect(s.points[2]!.marker).toMatchObject({ radius: 0, enabled: false, symbol: 'other' });
    const off = extract(line({ pointStyle: false, showLine: false }));
    expect(off.series[0]!.marker!.enabled).toBe(false);
    expect(off.series[0]!.line!.width).toBe(0);
  });
});

describe('chart chrome', () => {
  it('reads title/subtitle alignment, multi-line text and positions', () => {
    const m = extract(
      line(
        {},
        {
          plugins: {
            title: { display: true, text: ['Line one', 'Line two'], align: 'start', position: 'bottom' },
            subtitle: { display: true, text: () => 'x' },
            legend: {
              position: 'chartArea',
              align: 'start',
              reverse: true,
              title: { display: true, text: 'L' },
              labels: { filter: () => true },
            },
          },
        },
      ),
    );
    expect(m.title).toMatchObject({ text: 'Line one\nLine two', align: 'left' });
    expect(has(m, 'APPROXIMATED_LAYOUT', 'options.plugins.title.position')).toBe(true);
    expect(m.subtitle).toBeNull();
    expect(has(m, 'UNSUPPORTED_FORMATTER', 'options.plugins.subtitle.text')).toBe(true);
    expect(m.legend).toMatchObject({ position: 'topRight', overlay: true, reversed: true });
    expect(has(m, 'APPROXIMATED_LEGEND_POSITION', 'options.plugins.legend.align')).toBe(true);
    expect(has(m, 'UNSUPPORTED_STYLE', 'options.plugins.legend.title')).toBe(true);
    expect(has(m, 'UNSUPPORTED_FORMATTER', 'options.plugins.legend.labels.filter')).toBe(true);
    const left = extract(line({}, { plugins: { legend: { position: 'left', display: false } } }));
    expect(left.legend).toMatchObject({ position: 'left', layout: 'vertical', enabled: false });
  });

  it('uses the customCanvasBackgroundColor plugin option and chart-level font defaults', () => {
    const m = extract(
      line(
        {},
        {
          plugins: {
            customCanvasBackgroundColor: { color: '#fafafa' },
            title: { display: true, text: 'T', font: () => ({}) },
          },
          font: { family: 'Inter', size: 14 },
          color: '#222',
        },
      ),
    );
    expect(m.background).toMatchObject({ color: { r: 0xfa, g: 0xfa, b: 0xfa } });
    expect(m.title!.font).toMatchObject({ family: 'Inter', size: 14, bold: true });
    expect(has(m, 'UNSUPPORTED_STYLE', 'options.plugins.title.font')).toBe(true);
  });

  it('reports annotations and external tooltips', () => {
    const m = extract(
      line({}, { plugins: { annotation: { annotations: { a: {}, b: {} } }, tooltip: { external: () => {} } } }),
    );
    expect(has(m, 'UNSUPPORTED_ANNOTATION', 'options.plugins.annotation')).toBe(true);
    expect(has(m, 'UNSUPPORTED_TOOLTIP', 'options.plugins.tooltip.external')).toBe(true);
  });
});

describe('scales', () => {
  it('reads tick rotation, crossings, grid/border styles and bound hints', () => {
    const m = extract(
      line(
        { data: [5, 6, 7] },
        {
          scales: {
            x: {
              ticks: { minRotation: 45, maxRotation: 45 },
              position: { y: 2 },
              grid: { color: ['#ff0000', '#00ff00'] },
            },
            y: {
              position: 'center',
              suggestedMin: 1,
              suggestedMax: 10,
              grace: '10%',
              grid: { color: () => 'red', tickColor: '#0000ff', tickWidth: 3 },
              border: { dash: [4, 4], color: '#123456', width: 2 },
              ticks: { stepSize: 2, display: false },
              display: true,
            },
            y2: { axis: 'y', beginAtZero: true, display: false },
          },
        },
      ),
    );
    const [x] = m.xAxes;
    const [y, y2] = m.yAxes;
    expect(x!.labels.rotation).toBe(-45);
    expect(x!.crossing).toBe(2);
    expect(x!.gridLines!.color).toMatchObject({ r: 255, g: 0, b: 0 });
    expect(has(m, 'UNSUPPORTED_STYLE', 'options.scales.x.grid.color')).toBe(true);
    expect(has(m, 'UNSUPPORTED_STYLE', 'options.scales.y.grid.color')).toBe(true);
    expect(has(m, 'UNSUPPORTED_AXIS_FEATURE', 'options.scales.y.position')).toBe(true);
    expect(has(m, 'APPROXIMATED_AXIS_SCALE', 'options.scales.y.grace')).toBe(true);
    expect(y).toMatchObject({ min: 1, max: 10, tickInterval: 2 });
    expect(y!.labels.enabled).toBe(false);
    expect(y!.tickMarks).toMatchObject({ width: 3, color: { r: 0, g: 0, b: 255 } });
    expect(y!.axisLine).toMatchObject({ width: 2, dash: 'dash' });
    expect(y2).toMatchObject({ visible: false, id: 'y2' });
  });

  it('determines scale axes from position and dataset bindings; default ids otherwise', () => {
    const m = extract({
      type: 'line',
      data: {
        labels: ['a', 'b'],
        datasets: [
          { label: 'A', data: [1, 2], yAxisID: 'left' },
          { label: 'B', data: [3, 4], yAxisID: 'bound' },
          { label: 'C', data: [5, 6], yAxisID: 'right' },
        ],
      },
      options: { scales: { left: { position: 'left' }, bound: {}, right: { position: 'right' } } },
    });
    expect(m.yAxes.map((a) => a.id)).toEqual(['left', 'bound', 'right']);
    expect(m.series.map((s) => s.yAxisIndex)).toEqual([0, 1, 2]);
    expect(m.xAxes.map((a) => a.id)).toEqual(['x']);
  });

  it('reverses doughnut rings so the first dataset stays outside', () => {
    const m = extract({
      type: 'doughnut',
      data: {
        labels: ['a', 'b'],
        datasets: [
          { label: 'Outer', data: [1, 2] },
          { label: 'Inner', data: [3, 4] },
        ],
      },
    });
    expect(m.series.map((s) => s.name)).toEqual(['Inner', 'Outer']);
  });
});
