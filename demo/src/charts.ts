/**
 * The demo's chart catalogue: one Highcharts options object per card.
 *
 * Copied (and lightly adapted) from tests/fixtures/highcharts-options.ts so this folder is
 * self-contained. All data is deterministic — no Math.random — so exports are reproducible.
 */

import type { Chart, Options } from 'highcharts';
import type { PerChartExportConfig } from 'highcharts-editable-excel';

// Teach the Highcharts typings about the per-chart `exporting.editableExcel` block read by the library.
declare module 'highcharts' {
  interface ExportingOptions {
    editableExcel?: PerChartExportConfig;
  }
}

export interface DemoChart {
  /** Stable name: the card's chart div is `#chart-<name>`, the buttons `export-<name>` etc. */
  name: string;
  title: string;
  /** One sentence shown on the card explaining what it demonstrates. */
  note: string;
  options: Options;
}

const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];

/** Values used by the dynamic card; `nextDynamicValue` continues the sequence deterministically. */
export const dynamicInitialData = [12, 15, 11, 18, 16, 21];

/** Deterministic pseudo-random value for step `n` (a small linear congruential sequence). */
export function deterministicValue(n: number): number {
  return ((n * 37 + 11) % 29) + 5;
}

/** The styled-mode series colors, also defined as CSS variables in index.html. */
export const STYLED_MODE_COLORS = ['#8e44ad', '#16a085', '#d35400'] as const;

export const demoCharts: DemoChart[] = [
  {
    name: 'line',
    title: 'Line',
    note: 'A single line series over categories.',
    options: {
      title: { text: 'Monthly sales' },
      xAxis: { categories: months },
      yAxis: { title: { text: 'Units' } },
      series: [{ type: 'line', name: 'Sales', data: [1, 3, 2, 4, 6, 5] }],
    },
  },
  {
    name: 'multi-line',
    title: 'Multi-series line',
    note: 'Three line series sharing one category column.',
    options: {
      title: { text: 'Sales by region' },
      xAxis: { categories: months },
      yAxis: { title: { text: 'Units' } },
      series: [
        { type: 'line', name: 'North', data: [1, 2, 3, 4, 5, 6] },
        { type: 'line', name: 'South', data: [6, 5, 4, 3, 2, 1] },
        { type: 'line', name: 'East', data: [2, 4, 3, 5, 4, 6] },
      ],
    },
  },
  {
    name: 'column',
    title: 'Column',
    note: 'Clustered columns, two series.',
    options: {
      chart: { type: 'column' },
      title: { text: 'Fruit harvest' },
      xAxis: { categories: ['Apples', 'Pears', 'Plums', 'Kiwis'] },
      yAxis: { title: { text: 'Tonnes' } },
      series: [
        { type: 'column', name: '2023', data: [5, 3, 4, 7] },
        { type: 'column', name: '2024', data: [2, 2, 3, 2] },
      ],
    },
  },
  {
    name: 'bar',
    title: 'Bar (custom menu text)',
    note: 'Per-chart override: exporting.editableExcel.menuText = "Save as Excel chart".',
    options: {
      chart: { type: 'bar' },
      title: { text: 'Population (millions)' },
      xAxis: { categories: ['Africa', 'America', 'Asia', 'Europe'] },
      yAxis: { title: { text: 'Millions' } },
      series: [{ type: 'bar', name: 'Population', data: [1216, 1001, 4436, 738] }],
      exporting: { editableExcel: { menuText: 'Save as Excel chart' } },
    },
  },
  {
    name: 'area',
    title: 'Area',
    note: 'A filled area series.',
    options: {
      chart: { type: 'area' },
      title: { text: 'Bandwidth usage' },
      xAxis: { categories: months },
      yAxis: { title: { text: 'TB' } },
      series: [{ type: 'area', name: 'Usage', data: [3, 4, 3, 5, 4, 10] }],
    },
  },
  {
    name: 'pie',
    title: 'Pie (custom menu items)',
    note: 'Defines its own menuItems plus an app-specific "myCustomItem"; the Excel item is added alongside.',
    options: {
      title: { text: 'Browser share' },
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
      exporting: {
        menuItemDefinitions: {
          myCustomItem: {
            text: 'Show total (custom item)',
            onclick(this: Chart) {
              const total = this.series[0]?.points.reduce((sum, p) => sum + (p.y ?? 0), 0) ?? 0;
              this.setTitle(undefined, { text: `Total: ${total.toFixed(1)}%` });
            },
          },
        },
        buttons: {
          contextButton: { menuItems: ['viewFullscreen', 'printChart', 'separator', 'downloadPNG', 'downloadSVG', 'separator', 'myCustomItem'] },
        },
      },
    },
  },
  {
    name: 'doughnut',
    title: 'Doughnut',
    note: 'A pie with innerSize becomes an Excel doughnut chart.',
    options: {
      title: { text: 'Budget split' },
      series: [
        {
          type: 'pie',
          name: 'Share',
          innerSize: '55%',
          data: [
            ['Payroll', 45],
            ['Rent', 30],
            ['Other', 25],
          ],
        },
      ],
    },
  },
  {
    name: 'scatter',
    title: 'Scatter',
    note: 'Two x/y series; each gets its own X and Y columns.',
    options: {
      chart: { type: 'scatter' },
      title: { text: 'Height vs weight' },
      xAxis: { title: { text: 'Height (cm)' } },
      yAxis: { title: { text: 'Weight (kg)' } },
      series: [
        { type: 'scatter', name: 'Female', data: [[161.2, 51.6], [167.5, 59.0], [159.5, 49.2], [157.0, 63.0]] },
        { type: 'scatter', name: 'Male', data: [[174.0, 65.6], [175.3, 71.8], [193.5, 80.7]] },
      ],
    },
  },
  {
    name: 'custom',
    title: 'Custom styling',
    note: 'Georgia font, custom colors, dashed colored gridlines, data labels, vertical legend on the right, 800×450, plot background.',
    options: {
      chart: {
        width: 800,
        height: 450,
        backgroundColor: '#fffdf7',
        plotBackgroundColor: '#f3f6fb',
        style: { fontFamily: 'Georgia, serif' },
      },
      colors: ['#c0392b', '#2c3e50', '#27ae60'],
      title: { text: 'Custom styling', style: { fontFamily: 'Georgia, serif', fontSize: '22px', color: '#112233' } },
      subtitle: { text: 'Serif fonts, dashed grid, labels on every point' },
      legend: { layout: 'vertical', align: 'right', verticalAlign: 'middle' },
      xAxis: { categories: ['North', 'South', 'East', 'West'] },
      yAxis: {
        title: { text: 'Revenue' },
        gridLineColor: '#e67e22',
        gridLineWidth: 1,
        gridLineDashStyle: 'Dash',
      },
      plotOptions: { series: { dataLabels: { enabled: true } } },
      series: [
        { type: 'column', name: 'Plan', data: [10, 12, 9, 14] },
        { type: 'column', name: 'Actual', data: [11, 10, 13, 15] },
      ],
    },
  },
  {
    name: 'datetime',
    title: 'Datetime axis',
    note: 'Timestamps on the x axis become Excel dates with a date number format.',
    options: {
      title: { text: 'Daily visits' },
      xAxis: { type: 'datetime', labels: { format: '{value:%b %e}' } },
      yAxis: { title: { text: 'Visits' } },
      series: [
        {
          type: 'line',
          name: 'Visits',
          data: [
            [Date.UTC(2024, 0, 1), 10],
            [Date.UTC(2024, 0, 2), 12],
            [Date.UTC(2024, 0, 3), 9],
            [Date.UTC(2024, 0, 4), 14],
            [Date.UTC(2024, 0, 5), 13],
          ],
        },
      ],
    },
  },
  {
    name: 'dynamic',
    title: 'Dynamic data',
    note: 'Add points or replace the data, then export: the workbook contains what the chart shows now.',
    options: {
      title: { text: 'Live readings' },
      yAxis: { title: { text: 'Reading' } },
      series: [{ type: 'line', name: 'Sensor', data: [...dynamicInitialData] }],
    },
  },
  {
    name: 'styled',
    title: 'Styled mode',
    note: 'chart.styledMode: colors come from CSS variables (--highcharts-color-N) resolved by the browser.',
    options: {
      chart: { styledMode: true },
      title: { text: 'Styled mode (CSS)' },
      xAxis: { categories: ['A', 'B', 'C', 'D'] },
      series: [
        { type: 'column', name: 'First', data: [1, 2, 3, 2] },
        { type: 'column', name: 'Second', data: [2, 1, 2, 3] },
        { type: 'line', name: 'Third', data: [3, 2, 1, 2] },
      ],
    },
  },
  {
    name: 'combo',
    title: 'Combo',
    note: 'Column + line + spline in one plot area.',
    options: {
      title: { text: 'Combo' },
      xAxis: { categories: ['A', 'B', 'C', 'D'] },
      series: [
        { type: 'column', name: 'Bars', data: [3, 2, 1, 3] },
        { type: 'line', name: 'Line', data: [2, 3, 5, 7] },
        { type: 'spline', name: 'Average', data: [3, 2.67, 3, 6.33] },
      ],
    },
  },
  {
    name: 'secondary',
    title: 'Secondary axis',
    note: 'Rainfall on the left axis, temperature on an opposite (secondary) axis.',
    options: {
      title: { text: 'Rainfall and temperature' },
      xAxis: { categories: months },
      yAxis: [{ title: { text: 'Rainfall (mm)' } }, { title: { text: 'Temperature (°C)' }, opposite: true }],
      series: [
        { type: 'column', name: 'Rainfall', data: [49.9, 71.5, 106.4, 129.2, 144.0, 176.0] },
        { type: 'spline', name: 'Temperature', yAxis: 1, data: [7.0, 6.9, 9.5, 14.5, 18.2, 21.5] },
      ],
    },
  },
  {
    name: 'stacked',
    title: 'Stacked column',
    note: 'plotOptions.column.stacking = "normal".',
    options: {
      chart: { type: 'column' },
      title: { text: 'Quarterly revenue by product' },
      xAxis: { categories: ['Q1', 'Q2', 'Q3', 'Q4'] },
      yAxis: { title: { text: 'Revenue' } },
      plotOptions: { column: { stacking: 'normal' } },
      series: [
        { type: 'column', name: 'A', data: [5, 3, 4, 7] },
        { type: 'column', name: 'B', data: [2, 2, 3, 2] },
        { type: 'column', name: 'C', data: [3, 4, 4, 2] },
      ],
    },
  },
];
