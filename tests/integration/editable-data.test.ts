/**
 * Editable-data validation (spec 12.3).
 *
 * WHAT THIS PROVES — AND WHAT IT DOES NOT:
 * This test validates STRUCTURE: after a data cell on the Data sheet is edited and the workbook is
 * saved, the chart part still exists, its series formulas (`<c:f>`) are unchanged, and they still
 * point at the range that contains the edited cell, with the drawing/chart relationships intact.
 * That is the precondition for Excel to redraw the chart from the edited value. It does NOT
 * exercise Microsoft Excel's recalculation/redraw itself: Microsoft Excel was not available in this
 * environment.
 *
 * Two independent edit paths are used:
 *  1. @office-kit/xlsx: load → setCellAtAddress → workbookToBytes (a third-party round-trip).
 *     It runs in a plain Node child process: under vitest's jsdom environment the global
 *     `Uint8Array` belongs to another realm and office-kit's zip writer rejects its own entries
 *     ("ReadableStream entries are not yet supported"). That is a test-environment quirk, not a
 *     defect of either library.
 *  2. Library-independent: unzip with fflate, rewrite the `<v>` of the cell in the sheet XML, rezip.
 *     This does not depend on any library preserving the chart on save, so it isolates the claim
 *     that the chart's references are real cell references.
 */

import Highcharts from 'highcharts';
import 'highcharts/modules/exporting';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx } from '../../src/index';
import * as F from '../fixtures/highcharts-options';
import { inspectXlsx, type XlsxInspection } from '../helpers/inspect-xlsx';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';
import { parseRange, valuesAtRange } from './export-fixtures.shared';
import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

const H = Highcharts as unknown as HighchartsLike;
const REPO = resolve(__dirname, '../..');

/** Edits one cell with @office-kit/xlsx in a plain Node process (see header). */
const OFFICE_KIT_EDIT = `
import { readFileSync, writeFileSync } from 'node:fs';
import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import { getCellAtAddress, setCellAtAddress } from '@office-kit/xlsx/workbook';
const [input, output, address, value] = process.argv.slice(1);
const wb = await loadWorkbook(fromArrayBuffer(new Uint8Array(readFileSync(input))));
const before = getCellAtAddress(wb, address)?.value;
setCellAtAddress(wb, address, Number(value));
const bytes = await workbookToBytes(wb);
writeFileSync(output, bytes);
const reloaded = getCellAtAddress(await loadWorkbook(fromArrayBuffer(bytes)), address)?.value;
process.stdout.write(JSON.stringify({ before, reloaded }));
`;
const OUT = resolve(__dirname, '../output/editable');
const EDIT_REF = 'B2';
const NEW_VALUE = 999;

let original: Uint8Array;
let originalFormulas: ReturnType<XlsxInspection['seriesFormulas']>;
let originalChartXml: string;

async function assertEditedInvariants(bytes: Uint8Array, label: string): Promise<XlsxInspection> {
  const x = await inspectXlsx(bytes);
  x.assertWellFormed();
  const dataPath = x.sheetPath('Data');
  // The edited cell holds the new value; its neighbours are untouched.
  expect(x.cellValue(dataPath, EDIT_REF), `${label}: edited cell`).toBe(NEW_VALUE);
  expect(x.cellValue(dataPath, 'B3')).toBe(2);
  expect(x.cellValue(dataPath, 'C2')).toBe(6);
  // The chart part still exists and is still wired to the drawing on the Chart sheet.
  const charts = x.chartPaths();
  expect(charts, `${label}: chart parts`).toHaveLength(1);
  const drawingRels = x.parts.filter((p) => /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/.test(p));
  expect(drawingRels).toHaveLength(1);
  const chartFile = charts[0]!.split('/').pop()!;
  expect(x.text(drawingRels[0]!)).toContain(chartFile);
  // Formulas unchanged, and the first series' value range still covers the edited cell.
  const formulas = x.seriesFormulas(charts[0]!);
  expect(formulas, `${label}: formulas`).toEqual(originalFormulas);
  const range = parseRange(formulas[0]!.val!);
  expect(range.sheet).toBe('Data');
  expect(range.cells).toContain(EDIT_REF);
  expect(valuesAtRange(x, formulas[0]!.val!)).toEqual([NEW_VALUE, 2, 3, 4, 5, 6]);
  return x;
}

describe('editable data round-trip (structure, not Excel recalculation)', () => {
  beforeAll(async () => {
    const chart = renderChart(H, F.multiLine);
    const result = await exportHighchartsToXlsx(chart);
    original = result.bytes;
    destroyAll();
    const x = await inspectXlsx(original);
    originalFormulas = x.seriesFormulas(x.chartPaths()[0]!);
    originalChartXml = x.chartXml(0);
    mkdirSync(OUT, { recursive: true });
    writeFileSync(join(OUT, 'multi-line-original.xlsx'), original);
  });
  afterEach(() => destroyAll());

  it('the exported chart references Data!B2:B7 and B2 holds the source value', async () => {
    const x = await inspectXlsx(original);
    expect(originalFormulas.map((f) => f.val)).toEqual(['Data!$B$2:$B$7', 'Data!$C$2:$C$7', 'Data!$D$2:$D$7']);
    expect(originalFormulas.map((f) => f.cat)).toEqual(['Data!$A$2:$A$7', 'Data!$A$2:$A$7', 'Data!$A$2:$A$7']);
    expect(originalFormulas.map((f) => f.name)).toEqual(['Data!$B$1', 'Data!$C$1', 'Data!$D$1']);
    expect(x.cellValue(x.sheetPath('Data'), EDIT_REF)).toBe(1);
    console.info(`[editable-data] formulas: ${JSON.stringify(originalFormulas)}; ${EDIT_REF} before = 1`);
  });

  it('@office-kit/xlsx: edit Data!B2 → 999, save, chart formulas still point at the edited range', async () => {
    const input = join(OUT, 'multi-line-original.xlsx');
    const output = join(OUT, 'multi-line-edited-officekit.xlsx');
    const stdout = execFileSync(
      process.execPath,
      ['--input-type=module', '-e', OFFICE_KIT_EDIT, input, output, `Data!${EDIT_REF}`, String(NEW_VALUE)],
      {
        cwd: REPO,
        encoding: 'utf8',
        timeout: 60_000,
      },
    );
    const { before, reloaded } = JSON.parse(stdout) as { before: unknown; reloaded: unknown };
    expect(before).toBe(1);
    expect(reloaded).toBe(NEW_VALUE);
    const saved = new Uint8Array(readFileSync(output));
    const x = await assertEditedInvariants(saved, 'office-kit');
    console.info(
      `[editable-data] office-kit: ${EDIT_REF} after = ${String(x.cellValue(x.sheetPath('Data'), EDIT_REF))}`,
    );
  });

  it('fflate (library-independent): rewrite <v> of B2 in the sheet XML, rezip, same invariants', async () => {
    const files = unzipSync(original);
    const x0 = await inspectXlsx(original);
    const dataPath = x0.sheetPath('Data');
    const sheetXml = strFromU8(files[dataPath]!);
    const cellRe = new RegExp(`(<c r="${EDIT_REF}"[^>]*>)<v>[^<]*</v>(</c>)`);
    expect(sheetXml).toMatch(cellRe);
    files[dataPath] = strToU8(sheetXml.replace(cellRe, `$1<v>${NEW_VALUE}</v>$2`));
    const rezipped = zipSync(files);
    writeFileSync(join(OUT, 'multi-line-edited-fflate.xlsx'), rezipped);
    const x = await assertEditedInvariants(rezipped, 'fflate');
    // Only the one sheet part changed: the chart part is byte-identical.
    expect(x.chartXml(0)).toBe(originalChartXml);
    console.info(`[editable-data] fflate: ${EDIT_REF} after = ${String(x.cellValue(dataPath, EDIT_REF))}`);
  });
});
