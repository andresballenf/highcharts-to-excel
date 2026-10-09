/**
 * Chart.js 4 configurations used by the Chart.js adapter tests, plus a minimal 2D canvas context
 * stub so `new Chart(canvas, config)` runs under jsdom (which has no canvas implementation).
 *
 * Every factory returns a fresh object: Chart.js mutates the configuration it is given
 * (`options.scales` is replaced by the merged scale config; the colors plugin writes dataset colors).
 */

import type { ChartConfiguration } from 'chart.js';

// biome-ignore lint/suspicious/noExplicitAny: Chart.js configuration unions are too narrow for mixed fixtures
type Config = ChartConfiguration<any, any, any>;

export const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];

export function lineChart(): Config {
  return {
    type: 'line',
    data: {
      labels: MONTHS,
      datasets: [
        {
          label: 'Revenue',
          data: [12, 19, 3, 5, 2, 3],
          borderColor: '#ff6384',
          backgroundColor: 'rgba(255, 99, 132, 0.5)',
          borderWidth: 2,
          pointStyle: 'rectRot',
          pointRadius: 5,
        },
        { label: 'Costs', data: [7, 11, 5, 8, 3, 7], borderColor: 'rgb(54, 162, 235)', borderDash: [6, 4] },
      ],
    },
    options: {
      plugins: {
        title: { display: true, text: 'Monthly revenue', color: '#111111', font: { size: 18 } },
        subtitle: { display: true, text: 'FY 2024' },
        legend: { position: 'bottom', labels: { font: { size: 14 }, color: '#333333' } },
      },
      scales: {
        x: { title: { display: true, text: 'Month' }, grid: { display: false } },
        y: {
          title: { display: true, text: 'USD', color: 'rgb(255, 0, 0)' },
          min: 0,
          max: 25,
          grid: { color: '#dddddd', lineWidth: 2 },
          ticks: { format: { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 } },
        },
      },
    },
  };
}

export function smoothAreaChart(): Config {
  return {
    type: 'line',
    data: {
      labels: MONTHS,
      datasets: [{ label: 'Visitors', data: [3, 4, 6, 5, 8, 9], fill: true, tension: 0.4 }],
    },
    options: { plugins: { title: { display: true, text: 'Smooth area' } } },
  };
}

export function stackedBarChart(): Config {
  return {
    type: 'bar',
    data: {
      labels: ['Q1', 'Q2', 'Q3'],
      datasets: [
        { label: 'North', data: [10, 20, 30], backgroundColor: '#4bc0c0', stack: 'a' },
        { label: 'South', data: [5, 15, 25], backgroundColor: '#9966ff', stack: 'a' },
        { label: 'West', data: [8, 6, 4], backgroundColor: '#ff9f40', stack: 'a' },
      ],
    },
    options: {
      plugins: { title: { display: true, text: 'Stacked sales' } },
      scales: { x: { stacked: true }, y: { stacked: true } },
    },
  };
}

export function horizontalBarChart(): Config {
  return {
    type: 'bar',
    data: { labels: ['A', 'B', 'C'], datasets: [{ label: 'Score', data: [3, 7, 5] }] },
    options: { indexAxis: 'y' },
  };
}

export function pieChart(): Config {
  return {
    type: 'pie',
    data: {
      labels: ['Chrome', 'Firefox', 'Safari'],
      datasets: [
        {
          label: 'Share',
          data: [62, 20, 18],
          backgroundColor: ['#36a2eb', '#ff6384', '#ffcd56'],
          offset: [0, 12, 0],
        },
      ],
    },
    options: { plugins: { title: { display: true, text: 'Browser share' } } },
  };
}

export function doughnutChart(cutout: string | number = '60%'): Config {
  return {
    type: 'doughnut',
    data: { labels: ['Red', 'Blue', 'Yellow'], datasets: [{ label: 'Votes', data: [300, 50, 100] }] },
    options: { cutout, rotation: 90, circumference: 180 },
  };
}

export function scatterChart(): Config {
  return {
    type: 'scatter',
    data: {
      datasets: [
        {
          label: 'Samples',
          data: [
            { x: -10, y: 0 },
            { x: 0, y: 10 },
            { x: 10, y: 5 },
            { x: 0.5, y: 5.5 },
          ],
          backgroundColor: 'rgb(255, 99, 132)',
          pointStyle: 'triangle',
        },
      ],
    },
    options: { scales: { x: { type: 'linear', position: 'bottom' } } },
  };
}

export function bubbleChart(): Config {
  return {
    type: 'bubble',
    data: {
      datasets: [
        {
          label: 'Bubbles',
          data: [
            { x: 20, y: 30, r: 15 },
            { x: 40, y: 10, r: 10 },
            { x: 30, y: 20, r: 5 },
          ],
          backgroundColor: 'rgba(255, 99, 132, 0.6)',
        },
      ],
    },
  };
}

export function radarChart(): Config {
  return {
    type: 'radar',
    data: { labels: ['Eating', 'Drinking', 'Sleeping'], datasets: [{ label: 'Me', data: [65, 59, 90] }] },
  };
}

export function polarAreaChart(): Config {
  return {
    type: 'polarArea',
    data: { labels: ['Red', 'Green'], datasets: [{ label: 'Area', data: [11, 16] }] },
  };
}

export function mixedBarLineChart(): Config {
  return {
    type: 'bar',
    data: {
      labels: MONTHS.slice(0, 4),
      datasets: [
        { label: 'Bars', data: [10, 20, 30, 40] },
        { label: 'Average', type: 'line', data: [15, 15, 25, 35] },
      ],
    },
    options: { plugins: { title: { display: true, text: 'Combo' } } },
  };
}

export function timeScaleChart(): Config {
  return {
    type: 'line',
    data: {
      datasets: [
        {
          label: 'Visits',
          data: [
            { x: '2024-01-01', y: 5 },
            { x: '2024-01-02', y: 8 },
            { x: '2024-01-05', y: 3 },
          ],
        },
      ],
    },
    options: {
      plugins: { title: { display: true, text: 'Daily visits' } },
      scales: { x: { type: 'time', time: { unit: 'day', displayFormats: { day: 'MMM d' } } } },
    },
  };
}

export function perPointColorChart(): Config {
  return {
    type: 'bar',
    data: {
      labels: ['A', 'B', 'C'],
      datasets: [
        {
          label: 'Colored',
          data: [1, 2, 3],
          backgroundColor: ['rgb(255, 0, 0)', 'rgb(0, 128, 0)', 'rgb(0, 0, 255)'],
          borderColor: '#000000',
          borderWidth: 1,
        },
      ],
    },
  };
}

export function hiddenDatasetChart(): Config {
  return {
    type: 'line',
    data: {
      labels: ['a', 'b', 'c'],
      datasets: [
        { label: 'Shown', data: [1, 2, 3] },
        { label: 'Hidden', data: [3, 2, 1], hidden: true },
      ],
    },
  };
}

export function nullsChart(): Config {
  return {
    type: 'line',
    data: {
      labels: ['a', 'b', 'c', 'd'],
      datasets: [{ label: 'Gappy', data: [1, null, 3, 'oops'], spanGaps: true }],
    },
  };
}

export function secondaryAxisChart(): Config {
  return {
    type: 'line',
    data: {
      labels: MONTHS.slice(0, 3),
      datasets: [
        { label: 'Rainfall', type: 'bar', data: [49.9, 71.5, 106.4], yAxisID: 'y' },
        { label: 'Temperature', data: [7, 6.9, 9.5], yAxisID: 'y1' },
      ],
    },
    options: {
      plugins: { title: { display: true, text: 'Secondary axis' } },
      scales: {
        y: { type: 'linear', position: 'left', title: { display: true, text: 'mm' } },
        y1: {
          type: 'linear',
          position: 'right',
          title: { display: true, text: '°C' },
          grid: { drawOnChartArea: false },
        },
      },
    },
  };
}

export function parsingKeysChart(): Config {
  return {
    type: 'bar',
    data: {
      datasets: [
        {
          label: 'Sales',
          data: [
            { month: 'Jan', sales: { total: 5 } },
            { month: 'Feb', sales: { total: 7 } },
          ],
          parsing: { xAxisKey: 'month', yAxisKey: 'sales.total' },
        },
      ],
    },
  };
}

export function callbackChart(): Config {
  return {
    type: 'line',
    data: {
      labels: ['a', 'b'],
      datasets: [{ label: 'S', data: [1, 2], backgroundColor: () => 'red', borderColor: () => 'blue' }],
    },
    options: {
      scales: {
        y: { type: 'logarithmic', reverse: true, ticks: { callback: (v: unknown) => `${String(v)} u` } },
      },
      plugins: {
        tooltip: { callbacks: { label: () => 'x' } },
        datalabels: { color: '#fff' },
      },
    },
  };
}

/** A chart without any color: the colors plugin assigns its palette. */
export function autoColorChart(): Config {
  return {
    type: 'bar',
    data: {
      labels: ['a', 'b'],
      datasets: [
        { label: 'One', data: [1, 2] },
        { label: 'Two', data: [3, 4] },
      ],
    },
  };
}

// ---------------------------------------------------------------------------
// Canvas stub for live Chart.js instances under jsdom
// ---------------------------------------------------------------------------

function stubContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const target: Record<string, unknown> = {
    canvas,
    measureText: (t: string) => ({
      width: String(t).length * 6,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
    getImageData: () => ({ data: new Uint8ClampedArray(4) }),
    createLinearGradient: () => ({ addColorStop() {} }),
    createRadialGradient: () => ({ addColorStop() {} }),
    createPattern: () => ({}),
    getLineDash: () => [],
    isPointInPath: () => false,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
  };
  return new Proxy(target, {
    get: (t, k) => (k in t ? t[k as string] : () => undefined),
    set: (t, k, v) => {
      t[k as string] = v;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

/** Makes `canvas.getContext('2d')` return a no-op 2D context. Returns a restore function. */
export function installCanvasStub(): () => void {
  const proto = HTMLCanvasElement.prototype as unknown as { getContext: unknown };
  const original = proto.getContext;
  proto.getContext = function (this: HTMLCanvasElement) {
    return stubContext(this);
  };
  return () => {
    proto.getContext = original;
  };
}

/** A canvas of the given CSS size attached to the document (use with `responsive: false`). */
export function makeCanvas(width = 640, height = 320, id?: string): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  if (id) canvas.id = id;
  document.body.appendChild(canvas);
  return canvas;
}

/** `config` with animation and responsiveness off, for deterministic live charts. */
export function staticOptions(config: Config): Config {
  return { ...config, options: { ...(config.options ?? {}), animation: false, responsive: false } };
}
