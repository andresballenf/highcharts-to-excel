/**
 * Export performance benchmark. Skipped unless BENCH=1 (run it with `pnpm bench`, which spawns
 * vitest on this file and prints the table from tests/output/bench.json).
 *
 * Measures, per size, the median of 3 runs of `ExportResult.timings` (extract / translate / write /
 * total), the output size and the heap delta around the export call. Chart rendering in jsdom is
 * NOT included in any timing. Excel's per-series guidance is 32,000 points (see
 * EXCEL_MAX_POINTS_PER_SERIES); the 100,000-point cases exceed it on purpose, and the table
 * records whether a ROW_LIMIT_EXCEEDED warning was emitted.
 *
 * Responsiveness: a 10 ms `setInterval` runs during each export and records the longest gap
 * between two ticks (start → first tick and last tick → end included). In Node this measures
 * event-loop starvation the same way a browser's main thread would starve: a long synchronous
 * stretch inside the export shows up as one long gap.
 *
 * jsdom has no `Worker`, so the default writer zips with `zipSync` on the main thread here. The
 * "worker zip" columns rerun each case with `OoxmlExcelWriter({ zip: 'async' })` and a placeholder
 * global `Worker`: under vitest fflate resolves to its Node build, whose async zip runs on
 * worker_threads, which mirrors the browser path (fflate's Blob-URL Web Workers).
 */

import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Options } from 'highcharts';
import { unzipSync } from 'fflate';
import { afterAll, afterEach, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx, type ExportResult } from '../../src/index';
import { OoxmlExcelWriter } from '../../src/excel/ooxml-writer';
import { EXCEL_MAX_POINTS_PER_SERIES } from '../../src/excel/writer-interface';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

const H = Highcharts as unknown as HighchartsLike;
const RUNS = 3;

interface BenchCase {
  label: string;
  series: number;
  pointsPerSeries: number;
}

const CASES: BenchCase[] = [
  { label: '100 (1 × 100)', series: 1, pointsPerSeries: 100 },
  { label: '1,000 (1 × 1,000)', series: 1, pointsPerSeries: 1_000 },
  { label: '10,000 (5 × 2,000)', series: 5, pointsPerSeries: 2_000 },
  { label: '10,000 (1 × 10,000)', series: 1, pointsPerSeries: 10_000 },
  { label: '100,000 (1 × 100,000)', series: 1, pointsPerSeries: 100_000 },
  { label: '500,000 (5 × 100,000)', series: 5, pointsPerSeries: 100_000 },
];

export interface BenchRow {
  label: string;
  totalPoints: number;
  series: number;
  pointsPerSeries: number;
  extractMs: number;
  translateMs: number;
  writeMs: number;
  totalMs: number;
  bytes: number;
  heapDeltaMB: number;
  rowLimitWarning: boolean;
  /** Longest gap (ms) between two ticks of a 10 ms interval running during the export (median of runs). */
  maxTickGapMs: number;
  /** Interval ticks observed during the export (median of runs). */
  ticks: number;
  /** `ExportTimings.zipMs` (median of runs). */
  zipMs: number;
  /** Same case with the worker zip (see the file comment): total ms and longest tick gap. */
  workerZipTotalMs: number;
  workerZipMaxTickGapMs: number;
  runs: number;
}

function sameParts(a: Uint8Array, b: Uint8Array): boolean {
  const pa = unzipSync(a);
  const pb = unzipSync(b);
  const names = Object.keys(pa);
  if (names.join() !== Object.keys(pb).join()) return false;
  return names.every((n) => n === 'docProps/core.xml' || Buffer.from(pa[n]!).equals(Buffer.from(pb[n]!)));
}

/** Runs `fn` while a 10 ms interval ticks; returns its result and the longest gap between ticks. */
async function withTickGap<T>(fn: () => Promise<T>): Promise<{ value: T; maxGapMs: number; ticks: number }> {
  let last = performance.now();
  let maxGapMs = 0;
  let ticks = 0;
  const timer = setInterval(() => {
    const t = performance.now();
    maxGapMs = Math.max(maxGapMs, t - last);
    last = t;
    ticks++;
  }, 10);
  try {
    const value = await fn();
    maxGapMs = Math.max(maxGapMs, performance.now() - last);
    return { value, maxGapMs, ticks };
  } finally {
    clearInterval(timer);
  }
}

function options(c: BenchCase): Options {
  return {
    title: { text: `Bench ${c.label}` },
    series: Array.from({ length: c.series }, (_, s) => ({
      type: 'line' as const,
      name: `S${s + 1}`,
      data: Array.from(
        { length: c.pointsPerSeries },
        (_, i) => Math.round((Math.sin((i + s * 13) / 40) * 50 + 100) * 100) / 100,
      ),
    })),
  };
}

const median = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)]!;
};
const round = (n: number, d = 1): number => Math.round(n * 10 ** d) / 10 ** d;
const gc = (globalThis as { gc?: () => void }).gc;

const rows: BenchRow[] = [];
const BENCH_ENABLED = process.env.BENCH === '1';
if (!BENCH_ENABLED) console.info('[bench] skipped: set BENCH=1 (or run `pnpm bench`) to run the export benchmark');

describe.skipIf(!BENCH_ENABLED)('export benchmark (BENCH=1)', () => {
  afterEach(() => destroyAll());

  afterAll(() => {
    const out = resolve(__dirname, '../output');
    mkdirSync(out, { recursive: true });
    writeFileSync(
      resolve(out, 'bench.json'),
      `${JSON.stringify(
        {
          date: new Date().toISOString(),
          node: process.version,
          highchartsVersion: H.version ?? null,
          environment: 'vitest + jsdom (chart rendering excluded from timings)',
          gcExposed: typeof gc === 'function',
          excelMaxPointsPerSeries: EXCEL_MAX_POINTS_PER_SERIES,
          rows,
        },
        null,
        2,
      )}\n`,
    );
  });

  it.each(CASES.map((c) => [c.label, c] as const))(
    '%s',
    async (_l, c) => {
      // Warm-up (JIT) on the same chart shape, not recorded.
      {
        const chart = renderChart(H, options(c));
        await exportHighchartsToXlsx(chart);
        destroyAll();
      }
      const samples: Array<{
        extractMs: number;
        translateMs: number;
        writeMs: number;
        totalMs: number;
        bytes: number;
        heap: number;
        warn: boolean;
        zipMs: number;
        gap: number;
        ticks: number;
        workerTotal: number;
        workerGap: number;
      }> = [];
      for (let r = 0; r < RUNS; r++) {
        const chart = renderChart(H, options(c));
        gc?.();
        const heap0 = process.memoryUsage().heapUsed;
        const { value: result, maxGapMs, ticks } = await withTickGap(() => exportHighchartsToXlsx(chart));
        const heap1 = process.memoryUsage().heapUsed;
        // Browser-like path: worker zip (placeholder Worker global; fflate's Node build uses worker_threads).
        const g = globalThis as { Worker?: unknown };
        const hadWorker = 'Worker' in g;
        g.Worker ??= class {};
        let worker: { value: ExportResult; maxGapMs: number };
        try {
          worker = await withTickGap(() =>
            exportHighchartsToXlsx(chart, { writer: new OoxmlExcelWriter({ zip: 'async' }) }),
          );
        } finally {
          if (!hadWorker) delete g.Worker;
        }
        destroyAll();
        expect(result.report.editable).toBe(true);
        // Same package; only docProps/core.xml differs when the runs straddle a second (created = now).
        expect(sameParts(worker.value.bytes, result.bytes)).toBe(true);
        samples.push({
          ...result.timings,
          bytes: result.bytes.byteLength,
          heap: heap1 - heap0,
          warn: result.warnings.some((w) => w.code === 'ROW_LIMIT_EXCEEDED'),
          gap: maxGapMs,
          ticks,
          zipMs: result.timings.zipMs,
          workerTotal: worker.value.timings.totalMs,
          workerGap: worker.maxGapMs,
        });
      }
      rows.push({
        label: c.label,
        totalPoints: c.series * c.pointsPerSeries,
        series: c.series,
        pointsPerSeries: c.pointsPerSeries,
        extractMs: round(median(samples.map((s) => s.extractMs))),
        translateMs: round(median(samples.map((s) => s.translateMs))),
        writeMs: round(median(samples.map((s) => s.writeMs))),
        totalMs: round(median(samples.map((s) => s.totalMs))),
        bytes: median(samples.map((s) => s.bytes)),
        heapDeltaMB: round(median(samples.map((s) => s.heap)) / (1024 * 1024), 2),
        rowLimitWarning: samples.some((s) => s.warn),
        maxTickGapMs: round(median(samples.map((s) => s.gap))),
        ticks: median(samples.map((s) => s.ticks)),
        zipMs: round(median(samples.map((s) => s.zipMs))),
        workerZipTotalMs: round(median(samples.map((s) => s.workerTotal))),
        workerZipMaxTickGapMs: round(median(samples.map((s) => s.workerGap))),
        runs: RUNS,
      });
      expect(samples.some((s) => s.warn)).toBe(c.pointsPerSeries > EXCEL_MAX_POINTS_PER_SERIES);
    },
    600_000,
  );
});
