/**
 * Export performance benchmark. Skipped unless BENCH=1 (run it with `pnpm bench`, which spawns
 * vitest on this file and prints the table from tests/output/bench.json).
 *
 * Measures, per size, the median of 3 runs of `ExportResult.timings` (extract / translate / write /
 * total), the output size and the heap delta around the export call. Chart rendering in jsdom is
 * NOT included in any timing. Excel's per-series guidance is 32,000 points (see
 * EXCEL_MAX_POINTS_PER_SERIES); every size here stays below it, and the table records whether a
 * ROW_LIMIT_EXCEEDED warning was emitted.
 */

import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Options } from 'highcharts';
import { afterAll, afterEach, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx } from '../../src/index';
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
  runs: number;
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
      }> = [];
      for (let r = 0; r < RUNS; r++) {
        const chart = renderChart(H, options(c));
        gc?.();
        const heap0 = process.memoryUsage().heapUsed;
        const result = await exportHighchartsToXlsx(chart);
        const heap1 = process.memoryUsage().heapUsed;
        destroyAll();
        expect(result.report.editable).toBe(true);
        samples.push({
          ...result.timings,
          bytes: result.bytes.byteLength,
          heap: heap1 - heap0,
          warn: result.warnings.some((w) => w.code === 'ROW_LIMIT_EXCEEDED'),
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
        runs: RUNS,
      });
      expect(samples.some((s) => s.warn)).toBe(c.pointsPerSeries > EXCEL_MAX_POINTS_PER_SERIES);
    },
    120_000,
  );
});
