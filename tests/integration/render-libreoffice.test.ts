/**
 * LibreOffice render smoke check.
 *
 * This is a SMOKE CHECK that a third-party spreadsheet application opens the exported workbooks and
 * draws the native charts (PDF → PNG, with the chart's own text present in the PDF). It is NOT a
 * claim about Microsoft Excel's rendering fidelity: Microsoft Excel was not available here, and
 * LibreOffice's chart renderer differs from Excel's in layout, fonts and defaults.
 *
 * The workbooks are generated inside this test (into tests/output/render/) rather than read from
 * tests/output/export/v13/, so the result does not depend on test-file order and two files never
 * write the same path concurrently. They are exported with `includeSourceData: false`: the hidden
 * Data sheet is not printed, so the PDF holds only the Chart sheet (cell A1 = title, plus the
 * chart). Legend/axis strings asserted below therefore can only come from the rendered chart.
 */

import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { Options } from 'highcharts';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx } from '../../src/index';
import * as F from '../fixtures/highcharts-options';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

const H = Highcharts as unknown as HighchartsLike;
const OUT = resolve(__dirname, '../output/render');

function which(bin: string): string | null {
  const r = spawnSync('which', [bin], { encoding: 'utf8' });
  return r.status === 0 && r.stdout.trim() !== '' ? r.stdout.trim() : null;
}

const SOFFICE = which('soffice') ?? which('libreoffice');
const PDFTOPPM = which('pdftoppm');
const PDFTOTEXT = which('pdftotext');
const SKIP_REASON = !SOFFICE
  ? 'soffice/libreoffice not on PATH'
  : !PDFTOPPM || !PDFTOTEXT
    ? 'pdftoppm/pdftotext (poppler-utils) not on PATH'
    : null;

interface RenderCase {
  name: string;
  fixture: Options;
  title: string;
  /** Strings that only the chart itself draws (legend entries / axis titles). */
  chartText: string[];
}

const CASES: RenderCase[] = [
  { name: 'line', fixture: F.simpleLine, title: 'Simple line', chartText: ['Sales'] },
  { name: 'column', fixture: F.columnChart, title: 'Column', chartText: ['2023', '2024', 'Apples'] },
  { name: 'pie', fixture: F.pieChart, title: 'Pie', chartText: ['Chrome', 'Firefox'] },
  { name: 'combo', fixture: F.comboChart, title: 'Combo', chartText: ['Bars', 'Average'] },
  { name: 'secondary', fixture: F.secondaryAxis, title: 'Secondary axis', chartText: ['Rainfall', 'Temperature'] },
  { name: 'datetime', fixture: F.datetimeChart, title: 'Datetime', chartText: ['Visits'] },
];

const pdfText = new Map<string, string>();

describe.skipIf(SKIP_REASON !== null)('LibreOffice render smoke check (not an Excel fidelity claim)', () => {
  beforeAll(async () => {
    mkdirSync(OUT, { recursive: true });
    const files: string[] = [];
    for (const c of CASES) {
      const chart = renderChart(H, c.fixture);
      const result = await exportHighchartsToXlsx(chart, { includeSourceData: false });
      const p = join(OUT, `${c.name}.xlsx`);
      writeFileSync(p, result.bytes);
      files.push(p);
    }
    destroyAll();
    for (const c of CASES) rmSync(join(OUT, `${c.name}.pdf`), { force: true });
    const profile = mkdtempSync(join(tmpdir(), 'hc2xl-lo-'));
    try {
      execFileSync(
        SOFFICE!,
        [
          `-env:UserInstallation=${pathToFileURL(profile).href}`,
          '--headless',
          '--convert-to',
          'pdf',
          '--outdir',
          OUT,
          ...files,
        ],
        {
          stdio: 'pipe',
          timeout: 180_000,
        },
      );
    } finally {
      rmSync(profile, { recursive: true, force: true });
    }
    for (const c of CASES) {
      const pdf = join(OUT, `${c.name}.pdf`);
      if (!existsSync(pdf)) continue;
      execFileSync(PDFTOPPM!, ['-r', '60', '-png', '-f', '1', '-singlefile', pdf, join(OUT, c.name)], {
        stdio: 'pipe',
        timeout: 60_000,
      });
      pdfText.set(c.name, execFileSync(PDFTOTEXT!, [pdf, '-'], { encoding: 'utf8', timeout: 60_000 }));
    }
  }, 240_000);

  afterAll(() => {
    const found = CASES.map((c) => `${c.name}: ${pdfText.get(c.name)?.includes(c.title) ? `"${c.title}"` : 'MISSING'}`);
    console.info(`[render-libreoffice] titles found in PDFs → ${found.join('; ')} (output: ${OUT})`);
  });

  it.each(CASES.map((c) => [c.name, c] as const))('%s renders to PDF and PNG with the chart text', (_n, c) => {
    const pdf = join(OUT, `${c.name}.pdf`);
    const png = join(OUT, `${c.name}.png`);
    expect(existsSync(pdf), `${pdf} not produced`).toBe(true);
    expect(statSync(pdf).size).toBeGreaterThan(5 * 1024);
    expect(existsSync(png), `${png} not produced`).toBe(true);
    // A blank 60-dpi A4 page compresses to well under 3 KB; a drawn chart does not.
    expect(statSync(png).size).toBeGreaterThan(3 * 1024);
    const text = pdfText.get(c.name) ?? '';
    expect(text).toContain(c.title);
    for (const s of c.chartText) expect(text, `chart text "${s}" in ${c.name}.pdf`).toContain(s);
  });
});

if (SKIP_REASON !== null) {
  if (process.env.REQUIRE_RENDER === '1') {
    // CI sets REQUIRE_RENDER=1 after installing LibreOffice + poppler: a missing tool is a failure there.
    describe('LibreOffice render smoke check', () => {
      it('requires the render toolchain (REQUIRE_RENDER=1)', () => {
        throw new Error(
          `[render-libreoffice] REQUIRE_RENDER=1 but ${SKIP_REASON}. Install libreoffice-calc and poppler-utils.`,
        );
      });
    });
  } else {
    console.warn(`[render-libreoffice] skipped: ${SKIP_REASON} (set REQUIRE_RENDER=1 to make this a failure)`);
    describe('LibreOffice render smoke check', () => {
      it.skip(`skipped: ${SKIP_REASON}`, () => {});
    });
  }
}
