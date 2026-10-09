/**
 * Visual regression of LibreOffice renders (NOT an Excel fidelity claim).
 *
 * Each fixture is rendered with Highcharts 13 in jsdom, exported with `includeSourceData: false`
 * (the PDF then holds only the Chart sheet: A1 title + the native chart), converted to PDF by
 * `soffice --headless`, rasterized at 60 dpi (first page) by `pdftoppm`, cropped to the bounding
 * box of the non-white pixels (the drawn chart; this also makes the comparison independent of the
 * locale's default page size), and compared with `tests/baselines/render/<name>.png` by pixelmatch
 * (threshold 0.1, at most 1.5% of the pixels may differ). On a mismatch the actual image and a diff
 * are written to `tests/output/visual/<name>.actual.png` / `<name>.diff.png`.
 *
 * Baselines: `UPDATE_BASELINES=1 pnpm test:visual` writes missing baselines and rewrites the ones
 * that no longer match; review the PNGs before committing. Baselines are meant to be produced on
 * CI's image (ubuntu-latest + apt `libreoffice-calc` + `poppler-utils`, default fonts; the
 * committed set was produced on Ubuntu 24.04 with LibreOffice 24.2 and poppler 24.02). Fonts
 * differ between machines: the pixelmatch tolerance absorbs anti-aliasing noise, but not font
 * substitution or a different LibreOffice major version, which move text and legend boxes. That is
 * why the crop stays on the chart area and the data sheet is left out. If the renderer or fonts
 * change on purpose, regenerate the baselines on that image.
 *
 * Skips with a printed reason when soffice/pdftoppm are missing, unless REQUIRE_RENDER=1 (CI),
 * in which case a missing tool fails.
 */

import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { Options } from 'highcharts';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx } from '../../src/index';
import * as F from '../fixtures/highcharts-options';
import { destroyAll, type HighchartsLike, renderChart } from '../helpers/render-chart';

const H = Highcharts as unknown as HighchartsLike;
const OUT = resolve(__dirname, '../output/visual');
const BASELINES = resolve(__dirname, '../baselines/render');
const UPDATE = process.env.UPDATE_BASELINES === '1';
/** pixelmatch per-pixel color threshold (0..1, smaller is stricter). */
const THRESHOLD = 0.1;
/** Largest share of differing pixels that still passes. */
const MAX_RATIO = 0.005; // 0.5%: a vanished series on these crops is ~0.7%, local noise is 0.000%
/** Channel value below which a pixel counts as drawn when finding the chart's bounding box. */
const INK = 245;
const MARGIN = 4;

function which(bin: string): string | null {
  const r = spawnSync('which', [bin], { encoding: 'utf8' });
  return r.status === 0 && r.stdout.trim() !== '' ? r.stdout.trim() : null;
}

const SOFFICE = which('soffice') ?? which('libreoffice');
const PDFTOPPM = which('pdftoppm');
const SKIP_REASON = !SOFFICE
  ? 'soffice/libreoffice not on PATH'
  : !PDFTOPPM
    ? 'pdftoppm (poppler-utils) not on PATH'
    : null;

const CASES: ReadonlyArray<readonly [name: string, fixture: Options]> = [
  ['line', F.simpleLine],
  ['column', F.columnChart],
  ['bar', F.barChart],
  ['area', F.areaChart],
  ['pie', F.pieChart],
  ['doughnut', F.doughnutChart],
  ['combo', F.comboChart],
  ['secondary', F.secondaryAxis],
];

/** Crops a page to the bounding box of its non-white pixels plus a small white margin. */
function cropToInk(page: PNG): PNG {
  let x0 = page.width;
  let y0 = page.height;
  let x1 = -1;
  let y1 = -1;
  for (let y = 0; y < page.height; y++) {
    for (let x = 0; x < page.width; x++) {
      const i = (y * page.width + x) * 4;
      const d = page.data;
      if (d[i]! < INK || d[i + 1]! < INK || d[i + 2]! < INK) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  }
  if (x1 < 0) throw new Error('rendered page is blank');
  x0 = Math.max(0, x0 - MARGIN);
  y0 = Math.max(0, y0 - MARGIN);
  x1 = Math.min(page.width - 1, x1 + MARGIN);
  y1 = Math.min(page.height - 1, y1 + MARGIN);
  const out = new PNG({ width: x1 - x0 + 1, height: y1 - y0 + 1 });
  PNG.bitblt(page, out, x0, y0, out.width, out.height, 0, 0);
  return out;
}

/** Copies an image onto a white canvas of the given size (top-left aligned). */
function padTo(img: PNG, width: number, height: number): PNG {
  if (img.width === width && img.height === height) return img;
  const out = new PNG({ width, height });
  out.data.fill(255);
  PNG.bitblt(img, out, 0, 0, img.width, img.height, 0, 0);
  return out;
}

function writePng(path: string, img: PNG): void {
  writeFileSync(path, PNG.sync.write(img, { colorType: 2, deflateLevel: 9 }));
}

interface Comparison {
  ratio: number;
  diffPixels: number;
  sizeNote: string;
  diff: PNG;
}

function compare(actual: PNG, baseline: PNG): Comparison {
  const width = Math.max(actual.width, baseline.width);
  const height = Math.max(actual.height, baseline.height);
  const a = padTo(actual, width, height);
  const b = padTo(baseline, width, height);
  const diff = new PNG({ width, height });
  const diffPixels = pixelmatch(b.data, a.data, diff.data, width, height, { threshold: THRESHOLD });
  const sizeNote =
    actual.width === baseline.width && actual.height === baseline.height
      ? `${width}x${height}`
      : `actual ${actual.width}x${actual.height} vs baseline ${baseline.width}x${baseline.height}`;
  return { ratio: diffPixels / (width * height), diffPixels, sizeNote, diff };
}

const actuals = new Map<string, PNG>();
const report: string[] = [];

describe.skipIf(SKIP_REASON !== null)('LibreOffice visual regression (not an Excel fidelity claim)', () => {
  beforeAll(async () => {
    mkdirSync(OUT, { recursive: true });
    const files: string[] = [];
    for (const [name, fixture] of CASES) {
      const chart = renderChart(H, fixture);
      const result = await exportHighchartsToXlsx(chart, { includeSourceData: false });
      const p = join(OUT, `${name}.xlsx`);
      writeFileSync(p, result.bytes);
      files.push(p);
      for (const ext of ['pdf', 'actual.png', 'diff.png', 'png']) rmSync(join(OUT, `${name}.${ext}`), { force: true });
    }
    destroyAll();
    const profile = mkdtempSync(join(tmpdir(), 'hc2xl-visual-'));
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
        { stdio: 'pipe', timeout: 180_000 },
      );
    } finally {
      rmSync(profile, { recursive: true, force: true });
    }
    for (const [name] of CASES) {
      const pdf = join(OUT, `${name}.pdf`);
      if (!existsSync(pdf)) continue;
      execFileSync(PDFTOPPM!, ['-png', '-r', '60', '-f', '1', '-l', '1', '-singlefile', pdf, join(OUT, name)], {
        stdio: 'pipe',
        timeout: 60_000,
      });
      actuals.set(name, cropToInk(PNG.sync.read(readFileSync(join(OUT, `${name}.png`)))));
    }
  }, 240_000);

  afterAll(() => {
    console.info(`[visual-regression] ${report.join('; ')} (limit ${(MAX_RATIO * 100).toFixed(1)}%, output: ${OUT})`);
  });

  it.each(CASES.map(([name]) => [name] as const))('%s matches its baseline render', (name) => {
    const actual = actuals.get(name);
    expect(actual, `${name}: no PDF/PNG produced by soffice/pdftoppm`).toBeDefined();
    const baselinePath = join(BASELINES, `${name}.png`);
    if (!existsSync(baselinePath)) {
      if (UPDATE) {
        mkdirSync(BASELINES, { recursive: true });
        writePng(baselinePath, actual!);
        report.push(`${name}: baseline written`);
        return;
      }
      throw new Error(`${name}: no baseline at ${baselinePath}; run UPDATE_BASELINES=1 pnpm test:visual`);
    }
    const result = compare(actual!, PNG.sync.read(readFileSync(baselinePath)));
    const pct = `${(result.ratio * 100).toFixed(3)}%`;
    if (result.ratio <= MAX_RATIO) {
      report.push(`${name}: ${pct}`);
      return;
    }
    if (UPDATE) {
      writePng(baselinePath, actual!);
      report.push(`${name}: ${pct} → baseline rewritten`);
      return;
    }
    writePng(join(OUT, `${name}.actual.png`), actual!);
    writePng(join(OUT, `${name}.diff.png`), result.diff);
    report.push(`${name}: ${pct} FAIL`);
    throw new Error(
      `${name}: ${result.diffPixels} differing pixels = ${pct} > ${(MAX_RATIO * 100).toFixed(1)}% (${result.sizeNote}). ` +
        `See ${join(OUT, `${name}.diff.png`)} and ${name}.actual.png; if the change is intended, ` +
        'run UPDATE_BASELINES=1 pnpm test:visual and review the new baseline.',
    );
  });
});

if (SKIP_REASON !== null) {
  if (process.env.REQUIRE_RENDER === '1') {
    describe('LibreOffice visual regression', () => {
      it('requires the render toolchain (REQUIRE_RENDER=1)', () => {
        throw new Error(
          `[visual-regression] REQUIRE_RENDER=1 but ${SKIP_REASON}. Install libreoffice-calc and poppler-utils.`,
        );
      });
    });
  } else {
    console.warn(`[visual-regression] skipped: ${SKIP_REASON} (set REQUIRE_RENDER=1 to make this a failure)`);
    describe('LibreOffice visual regression', () => {
      it.skip(`skipped: ${SKIP_REASON}`, () => {});
    });
  }
}
