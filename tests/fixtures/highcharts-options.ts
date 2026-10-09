/**
 * Shared Highcharts option fixtures. Deterministic (no Math.random); never mutate them — render
 * through `tests/helpers/render-chart.ts`, which deep-clones.
 *
 * Fixtures marked [more] need `highcharts-more` loaded (bubble, polar, columnrange).
 */

import type { Options } from 'highcharts';

const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];

export const simpleLine: Options = {
  title: { text: 'Simple line' },
  xAxis: { categories: months },
  series: [{ type: 'line', name: 'Sales', data: [1, 3, 2, 4, 6, 5] }],
};

export const multiLine: Options = {
  title: { text: 'Multi line' },
  xAxis: { categories: months },
  series: [
    { type: 'line', name: 'North', data: [1, 2, 3, 4, 5, 6] },
    { type: 'line', name: 'South', data: [6, 5, 4, 3, 2, 1] },
    { type: 'line', name: 'East', data: [2, 4, 3, 5, 4, 6] },
  ],
};

export const splineChart: Options = {
  title: { text: 'Spline' },
  xAxis: { categories: months },
  series: [{ type: 'spline', name: 'Temperature', data: [7.0, 6.9, 9.5, 14.5, 18.2, 21.5] }],
};

export const columnChart: Options = {
  chart: { type: 'column' },
  title: { text: 'Column' },
  xAxis: { categories: ['Apples', 'Pears', 'Plums', 'Kiwis'] },
  series: [
    { type: 'column', name: '2023', data: [5, 3, 4, 7] },
    { type: 'column', name: '2024', data: [2, 2, 3, 2] },
  ],
};

export const stackedColumn: Options = {
  chart: { type: 'column' },
  title: { text: 'Stacked column' },
  xAxis: { categories: ['Q1', 'Q2', 'Q3', 'Q4'] },
  plotOptions: { column: { stacking: 'normal' } },
  series: [
    { type: 'column', name: 'A', data: [5, 3, 4, 7], stack: 'left' },
    { type: 'column', name: 'B', data: [2, 2, 3, 2], stack: 'left' },
    { type: 'column', name: 'C', data: [3, 4, 4, 2], stack: 'right' },
  ],
};

export const percentStackedColumn: Options = {
  chart: { type: 'column' },
  title: { text: 'Percent stacked column' },
  xAxis: { categories: ['Q1', 'Q2', 'Q3', 'Q4'] },
  plotOptions: { column: { stacking: 'percent' } },
  series: [
    { type: 'column', name: 'A', data: [5, 3, 4, 7] },
    { type: 'column', name: 'B', data: [2, 2, 3, 2] },
  ],
};

export const barChart: Options = {
  chart: { type: 'bar' },
  title: { text: 'Bar' },
  xAxis: { categories: ['Africa', 'America', 'Asia', 'Europe'] },
  series: [{ type: 'bar', name: 'Population', data: [1216, 1001, 4436, 738] }],
};

export const areaChart: Options = {
  chart: { type: 'area' },
  title: { text: 'Area' },
  xAxis: { categories: months },
  series: [{ type: 'area', name: 'Usage', data: [3, 4, 3, 5, 4, 10] }],
};

export const stackedArea: Options = {
  chart: { type: 'area' },
  title: { text: 'Stacked area' },
  xAxis: { categories: months },
  plotOptions: { area: { stacking: 'normal' } },
  series: [
    { type: 'area', name: 'A', data: [1, 2, 3, 4, 5, 6] },
    { type: 'area', name: 'B', data: [2, 2, 2, 2, 2, 2] },
  ],
};

export const pieChart: Options = {
  title: { text: 'Pie' },
  series: [
    {
      type: 'pie',
      name: 'Share',
      data: [
        { name: 'Chrome', y: 61.4, sliced: true },
        { name: 'Edge', y: 11.8, color: '#ff0000' },
        { name: 'Firefox', y: 10.9 },
        ['Safari', 4.6],
      ],
    },
  ],
};

export const doughnutChart: Options = {
  title: { text: 'Doughnut' },
  series: [
    {
      type: 'pie',
      name: 'Share',
      innerSize: '55%',
      data: [
        ['A', 45],
        ['B', 30],
        ['C', 25],
      ],
    },
  ],
};

export const scatterChart: Options = {
  chart: { type: 'scatter' },
  title: { text: 'Scatter' },
  series: [
    {
      type: 'scatter',
      name: 'Female',
      data: [
        [161.2, 51.6],
        [167.5, 59.0],
        [159.5, 49.2],
        [157.0, 63.0],
      ],
    },
    {
      type: 'scatter',
      name: 'Male',
      data: [
        [174.0, 65.6],
        [175.3, 71.8],
        [193.5, 80.7],
      ],
    },
  ],
};

export const datetimeChart: Options = {
  title: { text: 'Datetime' },
  xAxis: { type: 'datetime', labels: { format: '{value:%b %e}' } },
  series: [
    {
      type: 'line',
      name: 'Visits',
      data: [
        [Date.UTC(2024, 0, 1), 10],
        [Date.UTC(2024, 0, 2), 12],
        [Date.UTC(2024, 0, 3), 9],
        [Date.UTC(2024, 0, 4), 14],
      ],
    },
  ],
};

export const percentChart: Options = {
  chart: { type: 'column' },
  title: { text: 'Percent' },
  xAxis: { categories: ['A', 'B', 'C'] },
  yAxis: { labels: { format: '{value}%' } },
  plotOptions: { series: { dataLabels: { enabled: true, format: '{point.y:.1f}%' } } },
  series: [{ type: 'column', name: 'Rate', data: [12.5, 40.25, 70] }],
};

export const customColors: Options = {
  chart: {
    backgroundColor: {
      linearGradient: { x1: 0, y1: 0, x2: 0, y2: 1 },
      stops: [
        [0, '#ffffff'],
        [1, '#dddddd'],
      ],
    },
  },
  title: { text: 'Custom colors' },
  xAxis: { categories: ['A', 'B', 'C'] },
  series: [
    { type: 'column', name: 'Red-ish', color: '#aa3333', data: [1, { y: 2, color: '#ff0000' }, 3] },
    { type: 'column', name: 'Green-ish', color: 'rgb(0, 128, 0)', data: [3, 2, 1] },
  ],
};

export const styledMode: Options = {
  chart: { styledMode: true },
  title: { text: 'Styled mode' },
  xAxis: { categories: ['A', 'B', 'C'] },
  series: [
    { type: 'column', name: 'First', data: [1, 2, 3] },
    { type: 'line', name: 'Second', data: [3, 2, 1] },
  ],
};

export const customAxes: Options = {
  title: { text: 'Custom axes' },
  xAxis: {
    categories: ['A', 'B', 'C', 'D'],
    title: { text: 'Category', style: { color: '#123456', fontSize: '14px' } },
    lineColor: '#0000ff',
    lineWidth: 2,
  },
  yAxis: {
    min: 0,
    max: 100,
    tickInterval: 25,
    reversed: true,
    opposite: true,
    gridLineWidth: 2,
    gridLineColor: '#cccccc',
    gridLineDashStyle: 'Dash',
    title: { text: 'Value', style: { color: '#654321', fontWeight: 'bold' } },
  },
  series: [{ type: 'line', name: 'S', data: [10, 50, 30, 90] }],
};

export const nullNegative: Options = {
  title: { text: 'Nulls and negatives' },
  xAxis: { categories: ['A', 'B', 'C', 'D', 'E'] },
  series: [
    { type: 'line', name: 'With nulls', data: [1, null, 3, -4, 5] },
    { type: 'column', name: 'Negatives', data: [-1, -2, 0, 2, null] },
  ],
};

export const hiddenSeries: Options = {
  title: { text: 'Hidden series' },
  xAxis: { categories: ['A', 'B', 'C'] },
  series: [
    { type: 'line', name: 'Shown', data: [1, 2, 3] },
    { type: 'line', name: 'Hidden', visible: false, data: [3, 2, 1] },
  ],
};

export const comboChart: Options = {
  title: { text: 'Combo' },
  xAxis: { categories: ['A', 'B', 'C', 'D'] },
  series: [
    { type: 'column', name: 'Bars', data: [3, 2, 1, 3] },
    { type: 'line', name: 'Line', data: [2, 3, 5, 7] },
    { type: 'spline', name: 'Average', data: [3, 2.67, 3, 6.33] },
  ],
};

export const secondaryAxis: Options = {
  title: { text: 'Secondary axis' },
  xAxis: { categories: months },
  yAxis: [{ title: { text: 'Rainfall' } }, { title: { text: 'Temperature' }, opposite: true }],
  series: [
    { type: 'column', name: 'Rainfall', data: [49.9, 71.5, 106.4, 129.2, 144.0, 176.0] },
    { type: 'spline', name: 'Temperature', yAxis: 1, data: [7.0, 6.9, 9.5, 14.5, 18.2, 21.5] },
  ],
};

export const customStyling: Options = {
  chart: { width: 900, height: 500, plotBackgroundColor: '#fafafa' },
  title: { text: 'Custom styling', style: { fontFamily: 'Georgia, serif', fontSize: '24px', color: '#112233' } },
  legend: { layout: 'vertical', align: 'right', verticalAlign: 'middle' },
  xAxis: { categories: ['A', 'B', 'C'] },
  plotOptions: { series: { dataLabels: { enabled: true } } },
  series: [
    { type: 'column', name: 'One', data: [1, 2, 3] },
    { type: 'column', name: 'Two', data: [2, 3, 1] },
  ],
};

/** [more] */
export const bubbleChart: Options = {
  chart: { type: 'bubble' },
  title: { text: 'Bubble' },
  series: [
    {
      type: 'bubble',
      name: 'B1',
      data: [
        [9, 81, 63],
        [98, 5, 89],
        [51, 50, 73],
      ],
    },
    {
      type: 'bubble',
      name: 'B2',
      data: [
        [42, 38, 20],
        [6, 18, 1],
        [1, 93, 55],
      ],
    },
  ],
};

/** [more] A type the core + more bundle renders but the exporter does not support. */
export const unsupportedType: Options = {
  chart: { type: 'columnrange' },
  title: { text: 'Column range' },
  xAxis: { categories: ['Jan', 'Feb', 'Mar'] },
  series: [
    {
      type: 'columnrange',
      name: 'Temperatures',
      data: [
        [-9.5, 8.0],
        [-7.8, 8.3],
        [-13.1, 9.2],
      ],
    },
  ],
};

/** [more] */
export const polarChart: Options = {
  chart: { polar: true },
  title: { text: 'Polar' },
  xAxis: { categories: ['N', 'E', 'S', 'W'] },
  series: [{ type: 'line', name: 'Wind', data: [3, 5, 2, 4] }],
};

export const emptyChart: Options = {
  title: { text: 'Empty' },
  series: [],
};

/** Deterministic pseudo-data: 3 series × 1,000 points. */
function wave(seriesIndex: number, n: number): number[] {
  const out: number[] = [];
  for (let i = 0; i < n; i++) {
    out.push(Math.round((Math.sin((i + seriesIndex * 17) / 25) * 50 + 100 + seriesIndex * 10) * 100) / 100);
  }
  return out;
}

export const largeChart: Options = {
  title: { text: 'Large' },
  series: [0, 1, 2].map((i) => ({ type: 'line' as const, name: `Wave ${i + 1}`, data: wave(i, 1000) })),
};

/** Every fixture by name (core-only fixtures plus the [more] ones). */
export const allFixtures: Readonly<Record<string, Options>> = {
  simpleLine,
  multiLine,
  splineChart,
  columnChart,
  stackedColumn,
  percentStackedColumn,
  barChart,
  areaChart,
  stackedArea,
  pieChart,
  doughnutChart,
  scatterChart,
  datetimeChart,
  percentChart,
  customColors,
  styledMode,
  customAxes,
  nullNegative,
  hiddenSeries,
  comboChart,
  secondaryAxis,
  customStyling,
  bubbleChart,
  unsupportedType,
  polarChart,
  emptyChart,
  largeChart,
};
