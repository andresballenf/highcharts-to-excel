// @vitest-environment node
import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromArrayBuffer, loadWorkbook } from '@office-kit/xlsx/io';
import { getCellAtAddress, getSheet, getSheetState, sheetNames } from '@office-kit/xlsx/workbook';
import { listChartsOnSheet, listImagesOnSheet } from '@office-kit/xlsx/drawing';
import { OoxmlExcelWriter } from '../../src/excel/ooxml-writer';
import { inspectXlsx } from '../helpers/inspect-xlsx';
import { chartSpecs, fullWorkbook, type ChartSpecName } from '../fixtures/writer-specs';

const OUT_DIR = resolve(__dirname, '../output/writer');

function findSoffice(): string | null {
  for (const bin of ['soffice', 'libreoffice']) {
    const r = spawnSync('which', [bin], { encoding: 'utf8' });
    if (r.status === 0 && r.stdout.trim() !== '') return r.stdout.trim();
  }
  return null;
}

const SOFFICE = findSoffice();

/** Why the XSD check cannot run here, or null when OOXML_SCHEMA_DIR and python3 + lxml are available. */
function xsdSkipReason(): string | null {
  const dir = process.env.OOXML_SCHEMA_DIR;
  if (!dir) return 'OOXML_SCHEMA_DIR is not set';
  if (!existsSync(join(dir, 'ISO-IEC29500-4_2016', 'sml.xsd'))) return `no ISO-IEC29500-4_2016/sml.xsd under ${dir}`;
  const py = spawnSync('python3', ['-c', 'import lxml'], { encoding: 'utf8' });
  if (py.status !== 0) return 'python3 with lxml is not available';
  return null;
}
const XSD_SKIP = xsdSkipReason();
const allChartNames = Object.keys(chartSpecs) as ChartSpecName[];

async function writeAll(): Promise<{ allPath: string; bytes: Uint8Array }> {
  mkdirSync(OUT_DIR, { recursive: true });
  const writer = new OoxmlExcelWriter();
  for (const name of allChartNames) {
    const bytes = await writer.write(fullWorkbook({ charts: [chartSpecs[name]()], image: false }));
    writeFileSync(join(OUT_DIR, `${name}.xlsx`), bytes);
  }
  const bytes = await writer.write(fullWorkbook({ charts: allChartNames.map((n) => chartSpecs[n]()) }));
  const allPath = join(OUT_DIR, 'all-charts.xlsx');
  writeFileSync(allPath, bytes);
  return { allPath, bytes };
}

describe('writer integration', () => {
  it('round-trips through @office-kit/xlsx (strict chart parser)', async () => {
    const { bytes } = await writeAll();
    const x = await inspectXlsx(bytes);
    x.assertWellFormed();
    expect(x.chartPaths()).toHaveLength(allChartNames.length);

    const wb = await loadWorkbook(fromArrayBuffer(bytes));
    expect(sheetNames(wb)).toEqual(['Data', 'Chart', 'Meta']);
    expect(getSheetState(wb, 'Meta')).toBe('hidden');
    expect(getCellAtAddress(wb, 'Data!A1')?.value).toBe('Quarter');
    expect(getCellAtAddress(wb, 'Data!C2')?.value).toBe(4);
    expect(getCellAtAddress(wb, 'Data!A8')?.value).toBe('=1+1');
    expect(getCellAtAddress(wb, 'Data!E8')?.value).toBe('Ünïcødé & <tags> "q" end');

    const ws = getSheet(wb, 'Chart');
    expect(ws).toBeDefined();
    const charts = listChartsOnSheet(ws!);
    expect(charts).toHaveLength(allChartNames.length);
    expect(listImagesOnSheet(ws!)).toHaveLength(1);
    const kinds = charts.map((c) => {
      const content = c.content as { kind: string; chart?: { space?: { plotArea?: { chart?: { kind?: string } } } } };
      expect(content.kind).toBe('chart');
      return content.chart?.space?.plotArea?.chart?.kind;
    });
    expect(kinds).toEqual([
      'bar',
      'bar',
      'bar',
      'line',
      'area',
      'scatter',
      'bubble',
      'pie',
      'doughnut',
      'bar',
      'bar',
      'line',
      'line',
      'bar',
    ]);
  });

  it.skipIf(!SOFFICE)(
    'renders in LibreOffice to a PDF',
    async () => {
      const { allPath } = await writeAll();
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
            OUT_DIR,
            allPath,
          ],
          { stdio: 'pipe', timeout: 180_000 },
        );
      } finally {
        rmSync(profile, { recursive: true, force: true });
      }
      const pdf = join(OUT_DIR, 'all-charts.pdf');
      expect(existsSync(pdf)).toBe(true);
      expect(statSync(pdf).size).toBeGreaterThan(5 * 1024);
      console.info(`[writer-render] LibreOffice PDF: ${pdf} (${statSync(pdf).size} bytes)`);
    },
    200_000,
  );

  // Optional: never fails because the schemas or lxml are absent (not even with REQUIRE_RENDER=1).
  it.skipIf(XSD_SKIP !== null)(
    'validates every part against the ECMA-376 schemas (scripts/validate-ooxml.py)',
    async () => {
      await writeAll();
      const files = [...allChartNames.map((n) => join(OUT_DIR, `${n}.xlsx`)), join(OUT_DIR, 'all-charts.xlsx')];
      const r = spawnSync('python3', [resolve(__dirname, '../../scripts/validate-ooxml.py'), ...files], {
        encoding: 'utf8',
        timeout: 120_000,
      });
      console.info(`[writer-render] ${r.stdout.trim().split('\n').pop()}`);
      expect(r.status, `${r.stdout}\n${r.stderr}`).toBe(0);
    },
    150_000,
  );
  if (XSD_SKIP !== null) console.warn(`[writer-render] skipping XSD validation: ${XSD_SKIP}`);

  if (!SOFFICE) {
    if (process.env.REQUIRE_RENDER === '1') {
      // CI sets REQUIRE_RENDER=1 after installing LibreOffice: a missing soffice is a failure there.
      it('requires soffice/libreoffice (REQUIRE_RENDER=1)', () => {
        throw new Error(
          '[writer-render] REQUIRE_RENDER=1 but soffice/libreoffice is not on PATH. Install libreoffice-calc.',
        );
      });
    } else {
      it.skip('renders in LibreOffice to a PDF (soffice/libreoffice not on PATH)', () => {});
      console.warn(
        '[writer-render] skipping LibreOffice render: soffice/libreoffice not found on PATH (set REQUIRE_RENDER=1 to make this a failure)',
      );
    }
  }
});
