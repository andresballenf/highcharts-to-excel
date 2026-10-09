/**
 * `pnpm bench`: runs tests/integration/bench.test.ts under vitest (jsdom, BENCH=1) and prints the
 * results from tests/output/bench.json as a markdown table.
 */

import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { BenchRow } from '../tests/integration/bench.test';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const vitest = resolve(root, 'node_modules/vitest/vitest.mjs');

const benchJson = resolve(root, 'tests/output/bench.json');
const startedAt = Date.now();

// The benchmark suite only runs when BENCH === '1'; this script always sets it for the child.
if (process.env.BENCH !== undefined && process.env.BENCH !== '1') {
  console.info(`bench: ignoring BENCH=${process.env.BENCH}; the child vitest run gets BENCH=1`);
}

const nodeOptions = [process.env.NODE_OPTIONS ?? '', '--expose-gc'].join(' ').trim();
const run = spawnSync(process.execPath, [vitest, 'run', 'tests/integration/bench.test.ts'], {
  cwd: root,
  stdio: 'inherit',
  env: { ...process.env, BENCH: '1', NODE_OPTIONS: nodeOptions },
});
if (run.status !== 0) {
  console.error(`bench: vitest exited with status ${String(run.status)}`);
  process.exit(run.status ?? 1);
}

if (!existsSync(benchJson) || statSync(benchJson).mtimeMs < startedAt) {
  console.error('bench: tests/output/bench.json was not written by this run (was the suite skipped? it needs BENCH=1)');
  process.exit(1);
}

const report = JSON.parse(readFileSync(benchJson, 'utf8')) as {
  node: string;
  highchartsVersion: string | null;
  gcExposed: boolean;
  excelMaxPointsPerSeries: number;
  rows: BenchRow[];
};

const fmt = (n: number): string => n.toLocaleString('en-US');
const lines = [
  `Highcharts ${report.highchartsVersion ?? '?'} · Node ${report.node} · jsdom · median of ${report.rows[0]?.runs ?? 3} runs (rendering excluded) · gc exposed: ${String(report.gcExposed)}`,
  '',
  '| Points | extract ms | translate ms | write ms (zip) | total ms | max tick gap ms | ticks | worker zip: total ms | worker zip: max tick gap ms | XLSX bytes | heap Δ MB | ROW_LIMIT warning |',
  '|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|',
  ...report.rows.map(
    (r) =>
      `| ${r.label} | ${r.extractMs} | ${r.translateMs} | ${r.writeMs} (${r.zipMs}) | ${r.totalMs} | ${r.maxTickGapMs} | ${r.ticks} | ${r.workerZipTotalMs} | ${r.workerZipMaxTickGapMs} | ${fmt(r.bytes)} | ${r.heapDeltaMB} | ${r.rowLimitWarning ? 'yes' : 'no'} |`,
  ),
  '',
  `Excel guidance: at most ${fmt(report.excelMaxPointsPerSeries)} points per chart series; only the 100,000-points-per-series cases exceed it, so only they carry a ROW_LIMIT_EXCEEDED warning.`,
  'max tick gap: longest gap between two ticks of a 10 ms interval running during the export (event-loop starvation).',
  "Default columns: zipSync on the main thread (jsdom has no Worker). Worker zip: OoxmlExcelWriter({ zip: 'async' }) on fflate's worker_threads build, the browser-like path; same bytes.",
];
console.log(`\n${lines.join('\n')}\n`);
console.log('Wrote tests/output/bench.json');
