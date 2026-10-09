/**
 * `pnpm bench`: runs tests/integration/bench.test.ts under vitest (jsdom, BENCH=1) and prints the
 * results from tests/output/bench.json as a markdown table.
 */

import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { BenchRow } from '../tests/integration/bench.test';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const vitest = resolve(root, 'node_modules/vitest/vitest.mjs');

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

const report = JSON.parse(readFileSync(resolve(root, 'tests/output/bench.json'), 'utf8')) as {
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
  '| Points | extract ms | translate ms | write ms | total ms | XLSX bytes | heap Δ MB | ROW_LIMIT warning |',
  '|---|---:|---:|---:|---:|---:|---:|---|',
  ...report.rows.map(
    (r) =>
      `| ${r.label} | ${r.extractMs} | ${r.translateMs} | ${r.writeMs} | ${r.totalMs} | ${fmt(r.bytes)} | ${r.heapDeltaMB} | ${r.rowLimitWarning ? 'yes' : 'no'} |`,
  ),
  '',
  `Excel guidance: at most ${fmt(report.excelMaxPointsPerSeries)} points per chart series; none of these sizes exceeds it, so no ROW_LIMIT_EXCEEDED warning is expected.`,
];
console.log(`\n${lines.join('\n')}\n`);
console.log('Wrote tests/output/bench.json');
