/**
 * Integration suite: the 20 required export cases, run against a real Highcharts build (one
 * version per test file — see export-fixtures.v11/v12/v13.test.ts). Every case renders a live chart in
 * jsdom, exports it with `exportHighchartsToXlsx`, saves the workbook under
 * `tests/output/export/<version>/<case>.xlsx` and inspects the package structurally:
 * package parts, content types, chart XML groups/axes/formatting, and that every series formula
 * points at Data-sheet cells holding the expected numbers. Nothing in the pipeline is mocked.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import type { Chart, Options } from 'highcharts';
import { afterEach, describe, expect, it } from 'vitest';
import { exportHighchartsToXlsx, type ExportOptions, type ExportResult } from '../../src/index';
import { colorToHex, parseColor } from '../../src/utils/colors';
import * as F from '../fixtures/highcharts-options';
import { decodeXmlEntities, inspectXlsx, type XlsxInspection } from '../helpers/inspect-xlsx';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

// biome-ignore lint/suspicious/noExplicitAny: tests read live Highcharts internals that the public types omit
type AnyChart = any;

// ---------------------------------------------------------------------------
// Helpers (exported for the other integration files)
// ---------------------------------------------------------------------------

/** Column letters → 1-based index. */
function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function colLetters(index: number): string {
  let s = '';
  let n = index;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export interface ParsedRange {
  sheet: string;
  cells: string[];
}

/** Parses `Data!$B$2:$B$7` / `'My Data'!$A$1` into the sheet name and the list of cell refs. */
export function parseRange(formula: string): ParsedRange {
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(formula.trim());
  if (!m) throw new Error(`Unparseable range formula: ${formula}`);
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]!;
  const c1 = colIndex(m[3]!);
  const r1 = Number(m[4]);
  const c2 = m[5] ? colIndex(m[5]) : c1;
  const r2 = m[6] ? Number(m[6]) : r1;
  const cells: string[] = [];
  for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) cells.push(`${colLetters(c)}${r}`);
  }
  return { sheet, cells };
}

/** Values of every cell a chart formula references (null for empty / missing cells). */
export function valuesAtRange(insp: XlsxInspection, formula: string): Array<string | number | null> {
  const { sheet, cells } = parseRange(formula);
  const path = insp.sheetPath(sheet);
  return cells.map((ref) => insp.cellValue(path, ref));
}

/** Raw `<c:ser>` blocks of a chart part, in document order. */
export function serBlocks(chartXml: string): string[] {
  return [...chartXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map((m) => m[1]!);
}

/**
 * First `srgbClr` of the series-level `<c:spPr>` (fill for bars/areas, line for lines); for
 * marker-only series (scatter: line is noFill) the marker fill.
 */
export function seriesColor(ser: string): string | null {
  const noPts = ser.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, '');
  const first = (block: string | undefined): string | null =>
    block === undefined ? null : (/<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(block)?.[1]?.toUpperCase() ?? null);
  const sp = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(noPts.replace(/<c:marker>[\s\S]*?<\/c:marker>/g, ''));
  return first(sp?.[1]) ?? first(/<c:marker>([\s\S]*?)<\/c:marker>/.exec(noPts)?.[1]);
}

/** Data-point overrides: point index → first srgbClr. */
export function dPtColors(ser: string): Map<number, string> {
  const out = new Map<number, string>();
  for (const m of ser.matchAll(/<c:dPt>([\s\S]*?)<\/c:dPt>/g)) {
    const idx = Number(/<c:idx val="(\d+)"\/>/.exec(m[1]!)?.[1]);
    const color = /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[1]!)?.[1];
    if (color) out.set(idx, color.toUpperCase());
  }
  return out;
}

export function hexOf(css: unknown): string | null {
  const c = parseColor(css);
  return c ? colorToHex(c).toUpperCase() : null;
}

/** Axis element names (catAx/valAx/dateAx/serAx) directly in the plot area, in order. */
export function axisElements(chartXml: string): Array<{ kind: string; xml: string }> {
  const plot = /<c:plotArea>([\s\S]*)<\/c:plotArea>/.exec(chartXml)?.[1] ?? '';
  return [...plot.matchAll(/<c:(catAx|valAx|dateAx|serAx)>([\s\S]*?)<\/c:\1>/g)].map((m) => ({
    kind: m[1]!,
    xml: m[2]!,
  }));
}

/** Expected numeric y values per exported series, derived from the fixture options. */
function yValues(data: unknown[]): Array<number | null> {
  return data.map((p) => {
    if (p === null || p === undefined) return null;
    if (typeof p === 'number') return p;
    if (Array.isArray(p)) return (p[p.length - 1] as number | null) ?? null;
    return ((p as { y?: number | null }).y ?? null) as number | null;
  });
}

function seriesData(o: Options, i: number): unknown[] {
  return ((o.series?.[i] as { data?: unknown[] } | undefined)?.data ?? []) as unknown[];
}

/** Assertions shared by every successful export. */
export async function assertCommonPackage(result: ExportResult, expectedExcelType: string): Promise<XlsxInspection> {
  expect(result.bytes[0]).toBe(0x50);
  expect(result.bytes[1]).toBe(0x4b);
  const x = await inspectXlsx(result.bytes);
  x.assertWellFormed();
  for (const part of [
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/sheet2.xml',
    'xl/worksheets/_rels/sheet1.xml.rels',
    'xl/charts/chart1.xml',
    'xl/drawings/drawing1.xml',
    'xl/drawings/_rels/drawing1.xml.rels',
    '[Content_Types].xml',
  ]) {
    expect(x.parts, `missing part ${part}`).toContain(part);
  }
  const ct = x.text('[Content_Types].xml');
  expect(ct).toMatch(
    /<Override[^>]*PartName="\/xl\/charts\/chart1\.xml"[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/,
  );
  expect(x.contentTypes()).toContain('application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
  // Relationship chain: chart sheet → drawing → chart.
  expect(x.text('xl/worksheets/_rels/sheet1.xml.rels')).toMatch(/Target="[^"]*drawings\/drawing1\.xml"/);
  expect(x.text('xl/drawings/_rels/drawing1.xml.rels')).toMatch(/Target="[^"]*charts\/chart1\.xml"/);
  expect(x.text('xl/drawings/drawing1.xml')).toContain('<c:chart ');
  // Native chart, never a picture.
  expect(x.hasImages()).toBe(false);
  expect(x.parts.filter((p) => p.startsWith('xl/media/'))).toEqual([]);
  expect(x.text('xl/drawings/drawing1.xml')).not.toContain('<xdr:pic>');
  // Report.
  expect(result.report.editable).toBe(true);
  expect(result.report.excelChartType).toBe(expectedExcelType);
  expect(result.report.blocking).toEqual([]);
  expect(result.warnings.filter((w) => w.outcome === 'blocking')).toEqual([]);
  // Every series references the Data sheet with numbers.
  const chartXml = x.chartXml(0);
  const sers = serBlocks(chartXml);
  expect(sers.length).toBeGreaterThan(0);
  for (const ser of sers) {
    const valBlock = /<c:(val|yVal)>([\s\S]*?)<\/c:\1>/.exec(ser);
    expect(valBlock, 'series without c:val / c:yVal').not.toBeNull();
    const f = /<c:numRef><c:f>([^<]+)<\/c:f>/.exec(valBlock![2]!);
    expect(f, 'value block is not a numRef').not.toBeNull();
    expect(parseRange(decodeXmlEntities(f![1]!)).sheet).toBe('Data');
  }
  return x;
}

function save(dir: string, name: string, bytes: Uint8Array): string {
  mkdirSync(dir, { recursive: true });
  const p = join(dir, `${name}.xlsx`);
  writeFileSync(p, bytes);
  return p;
}

// ---------------------------------------------------------------------------
// Case table
// ---------------------------------------------------------------------------

interface CaseContext {
  x: XlsxInspection;
  chartXml: string;
  result: ExportResult;
  chart: AnyChart;
  options: Options;
}

interface ExportCase {
  name: string;
  fixture: Options;
  excelType: string;
  /** Expected plot groups in plot-area order. */
  groups: string[];
  exportOptions?: ExportOptions;
  /** Mutate the live chart before export. */
  before?: (chart: AnyChart) => void;
  /** Check series values / categories against the fixture (default true; scatter handled separately). */
  categories?: boolean;
  check?: (ctx: CaseContext) => void;
}

const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];

export const EXPORT_CASES: ExportCase[] = [
  {
    name: 'line',
    fixture: F.simpleLine,
    excelType: 'line',
    groups: ['lineChart'],
    check: ({ chartXml }) => expect(chartXml).toContain('<c:smooth val="0"/>'),
  },
  { name: 'multi-line', fixture: F.multiLine, excelType: 'line', groups: ['lineChart'] },
  {
    name: 'spline',
    fixture: F.splineChart,
    excelType: 'line',
    groups: ['lineChart'],
    check: ({ chartXml, result }) => {
      expect(chartXml).toContain('<c:smooth val="1"/>');
      expect(result.report.sourceChartType).toBe('spline');
    },
  },
  {
    name: 'column',
    fixture: F.columnChart,
    excelType: 'column',
    groups: ['barChart'],
    check: ({ chartXml }) => {
      expect(chartXml).toContain('<c:barDir val="col"/>');
      expect(chartXml).toContain('<c:grouping val="clustered"/>');
    },
  },
  {
    name: 'stacked-column',
    fixture: F.stackedColumn,
    excelType: 'stackedColumn',
    groups: ['barChart'],
    check: ({ chartXml }) => {
      expect(chartXml).toContain('<c:barDir val="col"/>');
      expect(chartXml).toContain('<c:grouping val="stacked"/>');
      expect(chartXml).toContain('<c:overlap val="100"/>');
    },
  },
  {
    name: 'bar',
    fixture: F.barChart,
    excelType: 'bar',
    groups: ['barChart'],
    check: ({ chartXml }) => {
      expect(chartXml).toContain('<c:barDir val="bar"/>');
      expect(chartXml).not.toContain('<c:barDir val="col"/>');
      const axes = axisElements(chartXml);
      expect(axes.find((a) => a.kind === 'catAx')?.xml).toContain('<c:axPos val="l"/>');
    },
  },
  { name: 'area', fixture: F.areaChart, excelType: 'area', groups: ['areaChart'] },
  {
    name: 'pie',
    fixture: F.pieChart,
    excelType: 'pie',
    groups: ['pieChart'],
    check: ({ chartXml, chart }) => {
      expect(chartXml).toContain('<c:varyColors val="1"/>');
      const pts = dPtColors(serBlocks(chartXml)[0]!);
      const live = (chart.series[0].points as Array<{ color: unknown }>).map((p) => hexOf(p.color));
      live.forEach((hex, i) => {
        expect(pts.get(i), `pie dPt ${i}`).toBe(hex);
      });
      expect(pts.get(1)).toBe('FF0000'); // explicit point color in the fixture
      expect(chartXml).toMatch(/<c:dPt><c:idx val="0"\/>(?:(?!<\/c:dPt>)[\s\S])*<c:explosion val="\d+"\/>/); // sliced
    },
  },
  {
    name: 'doughnut',
    fixture: F.doughnutChart,
    excelType: 'doughnut',
    groups: ['doughnutChart'],
    check: ({ chartXml }) => expect(chartXml).toContain('<c:holeSize val="55"/>'),
  },
  {
    name: 'scatter',
    fixture: F.scatterChart,
    excelType: 'scatter',
    groups: ['scatterChart'],
    categories: false,
    check: ({ x }) => {
      const fs = x.seriesFormulas(x.chartPaths()[0]!);
      expect(fs).toHaveLength(2);
      const female = seriesData(F.scatterChart, 0) as number[][];
      const male = seriesData(F.scatterChart, 1) as number[][];
      expect(valuesAtRange(x, fs[0]!.xVal!)).toEqual(female.map((p) => p[0]));
      expect(valuesAtRange(x, fs[0]!.yVal!)).toEqual(female.map((p) => p[1]));
      expect(valuesAtRange(x, fs[1]!.xVal!)).toEqual(male.map((p) => p[0]));
      expect(valuesAtRange(x, fs[1]!.yVal!)).toEqual(male.map((p) => p[1]));
      expect(fs.every((f) => f.val === undefined && f.cat === undefined)).toBe(true);
      // Scatter blocks have "<name> X/Y" headers, so the series name is a literal.
      expect(serBlocks(x.chartXml(0)).map((s) => /<c:tx><c:v>([^<]*)<\/c:v><\/c:tx>/.exec(s)?.[1])).toEqual([
        'Female',
        'Male',
      ]);
    },
  },
  {
    name: 'datetime',
    fixture: F.datetimeChart,
    excelType: 'line',
    groups: ['lineChart'],
    categories: false,
    check: ({ x, chartXml }) => {
      const axes = axisElements(chartXml);
      const dateAx = axes.find((a) => a.kind === 'dateAx');
      expect(dateAx, 'no c:dateAx').toBeDefined();
      expect(dateAx!.xml).toMatch(/<c:numFmt formatCode="[^"]+" sourceLinked="0"\/>/);
      const f = x.seriesFormulas(x.chartPaths()[0]!)[0]!;
      // Date.UTC(2024,0,1) → Excel serial 45292.
      expect(valuesAtRange(x, f.cat!)).toEqual([45292, 45293, 45294, 45295]);
      expect(chartXml).toMatch(/<c:cat><c:numRef><c:f>Data!\$A\$2:\$A\$5<\/c:f>/);
      // The category cells carry a date number format (built-in 14-22 or a custom code with d/m/y).
      const a2 = x.cellXml(x.sheetPath('Data'), 'A2')!;
      const styleIdx = Number(/\bs="(\d+)"/.exec(a2)?.[1]);
      const styles = x.text('xl/styles.xml');
      const cellXfs = [.../<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)![1]!.matchAll(/<xf\s[^>]*?\/?>/g)].map(
        (m) => m[0],
      );
      const numFmtId = Number(/numFmtId="(\d+)"/.exec(cellXfs[styleIdx]!)?.[1]);
      const code =
        numFmtId >= 164
          ? decodeXmlEntities(new RegExp(`<numFmt numFmtId="${numFmtId}" formatCode="([^"]*)"`).exec(styles)?.[1] ?? '')
          : String(numFmtId);
      expect(code, 'date cell format').toMatch(/^(?:1[4-9]|2[0-2]|.*[dmy].*)$/);
    },
  },
  {
    name: 'percent',
    fixture: F.percentChart,
    excelType: 'column',
    groups: ['barChart'],
    check: ({ chartXml }) => {
      const fmts = [...chartXml.matchAll(/<c:numFmt formatCode="([^"]*)"/g)].map((m) => decodeXmlEntities(m[1]!));
      expect(
        fmts.some((f) => f.includes('%')),
        `numFmts: ${fmts.join(' | ')}`,
      ).toBe(true);
      const valAx = axisElements(chartXml).find((a) => a.kind === 'valAx')!;
      expect(valAx.xml).toMatch(/<c:numFmt formatCode="[^"]*%[^"]*" sourceLinked="0"\/>/);
    },
  },
  {
    name: 'custom-colors',
    fixture: F.customColors,
    excelType: 'column',
    groups: ['barChart'],
    check: ({ chartXml }) => {
      const sers = serBlocks(chartXml);
      expect(seriesColor(sers[0]!)).toBe('AA3333');
      expect(seriesColor(sers[1]!)).toBe('008000');
      expect(dPtColors(sers[0]!).get(1)).toBe('FF0000');
      expect(dPtColors(sers[1]!).size).toBe(0);
    },
  },
  {
    name: 'styled-mode',
    fixture: F.styledMode,
    excelType: 'combo:column+line',
    groups: ['barChart', 'lineChart'],
    check: ({ chartXml, result }) => {
      // jsdom has no Highcharts CSS: colors fall back to the default palette by colorIndex
      // (the values of --highcharts-color-0/1) and the fallback is reported.
      expect(result.warnings.map((w) => w.code)).toContain('STYLED_MODE_FALLBACK');
      const sers = serBlocks(chartXml);
      expect(seriesColor(sers[0]!)).toBe('2CAFFE');
      expect(seriesColor(sers[1]!)).toBe('544FC5');
    },
  },
  {
    name: 'custom-axes',
    fixture: F.customAxes,
    excelType: 'line',
    groups: ['lineChart'],
    check: ({ chartXml }) => {
      const axes = axisElements(chartXml);
      const cat = axes.find((a) => a.kind === 'catAx')!;
      const val = axes.find((a) => a.kind === 'valAx')!;
      expect(cat.xml).toContain('<a:t>Category</a:t>');
      expect(cat.xml).toContain('<a:srgbClr val="123456"/>');
      expect(cat.xml).toMatch(/<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"\/>/); // lineWidth 2, lineColor
      expect(val.xml).toContain('<a:t>Value</a:t>');
      expect(val.xml).toContain('<c:min val="0"/>');
      expect(val.xml).toContain('<c:max val="100"/>');
      expect(val.xml).toContain('<c:orientation val="maxMin"/>'); // reversed
      expect(val.xml).toContain('<c:axPos val="r"/>'); // opposite
      expect(val.xml).toContain('<c:majorUnit val="25"/>');
      const grid = /<c:majorGridlines>([\s\S]*?)<\/c:majorGridlines>/.exec(val.xml)?.[1];
      expect(grid, 'no majorGridlines').toBeDefined();
      expect(grid).toContain('<a:srgbClr val="CCCCCC"/>');
      expect(grid).toContain('<a:prstDash val="dash"/>');
    },
  },
  {
    name: 'runtime-updated',
    fixture: F.simpleLine,
    excelType: 'line',
    groups: ['lineChart'],
    categories: false,
    before: (chart) => {
      chart.series[0].setData([10, 20, 30, 40, 50, 60], false);
      chart.addSeries({ type: 'line', name: 'Added', data: [7, 8, 9, 10, 11, 12] }, false);
      chart.update({ title: { text: 'Updated title' } }, false);
      chart.redraw(false);
    },
    check: ({ x, chartXml, result }) => {
      const fs = x.seriesFormulas(x.chartPaths()[0]!);
      expect(fs).toHaveLength(2);
      expect(valuesAtRange(x, fs[0]!.val!)).toEqual([10, 20, 30, 40, 50, 60]);
      expect(valuesAtRange(x, fs[1]!.val!)).toEqual([7, 8, 9, 10, 11, 12]);
      expect(valuesAtRange(x, fs[1]!.name!)).toEqual(['Added']);
      expect(valuesAtRange(x, fs[0]!.cat!)).toEqual(months);
      expect(chartXml).toContain('<a:t>Updated title</a:t>');
      expect(chartXml).not.toContain('<a:t>Simple line</a:t>');
      expect(x.cellValue(x.sheetPath('Chart'), 'A1')).toBe('Updated title');
      expect(result.filename).toBe('Updated title.xlsx');
    },
  },
  {
    name: 'null-negative',
    fixture: F.nullNegative,
    excelType: 'combo:line+column',
    groups: ['lineChart', 'barChart'],
    check: ({ x, chartXml, result }) => {
      const data = x.sheetPath('Data');
      // Null → no cell at all (or a cell without <v>), so Excel treats it as blank.
      const b3 = x.cellXml(data, 'B3');
      expect(b3 === null || !/<v>/.test(b3)).toBe(true);
      const c6 = x.cellXml(data, 'C6');
      expect(c6 === null || !/<v>/.test(c6)).toBe(true);
      expect(x.cellValue(data, 'B5')).toBe(-4);
      expect(x.cellValue(data, 'C2')).toBe(-1);
      expect(x.cellValue(data, 'C3')).toBe(-2);
      expect(x.cellValue(data, 'C4')).toBe(0);
      expect(chartXml).toContain('<c:dispBlanksAs val="gap"/>');
      expect(result.warnings.filter((w) => w.code === 'NULL_VALUES').map((w) => w.seriesIndex)).toEqual([0, 1]);
    },
  },
  {
    name: 'hidden-series',
    fixture: F.hiddenSeries,
    excelType: 'line',
    groups: ['lineChart'],
    categories: false,
    check: ({ x, result }) => {
      const fs = x.seriesFormulas(x.chartPaths()[0]!);
      expect(fs).toHaveLength(1);
      expect(valuesAtRange(x, fs[0]!.name!)).toEqual(['Shown']);
      expect(valuesAtRange(x, fs[0]!.val!)).toEqual([1, 2, 3]);
      expect(x.cellValue(x.sheetPath('Data'), 'C1')).toBeNull();
      expect(result.warnings.map((w) => w.code)).toContain('HIDDEN_SERIES_EXCLUDED');
    },
  },
  {
    name: 'combo',
    fixture: F.comboChart,
    excelType: 'combo:column+line',
    groups: ['barChart', 'lineChart'],
    check: ({ chartXml }) => {
      const line = /<c:lineChart>([\s\S]*?)<\/c:lineChart>/.exec(chartXml)![1]!;
      expect(serBlocks(line)).toHaveLength(2);
      expect(line).toContain('<c:smooth val="1"/>'); // the spline series
      const bar = /<c:barChart>([\s\S]*?)<\/c:barChart>/.exec(chartXml)![1]!;
      expect(serBlocks(bar)).toHaveLength(1);
      // Both groups share one axis pair.
      expect(axisElements(chartXml)).toHaveLength(2);
    },
  },
  {
    name: 'secondary',
    fixture: F.secondaryAxis,
    excelType: 'combo:column+line',
    groups: ['barChart', 'lineChart'],
    check: ({ chartXml, result }) => {
      const axes = axisElements(chartXml);
      expect(axes).toHaveLength(4);
      const right = axes.filter((a) => a.kind === 'valAx' && a.xml.includes('<c:axPos val="r"/>'));
      expect(right).toHaveLength(1);
      expect(right[0]!.xml).toContain('<c:crosses val="max"/>');
      expect(right[0]!.xml).toContain('<a:t>Temperature</a:t>');
      expect(axes.filter((a) => a.kind === 'valAx' && a.xml.includes('<a:t>Rainfall</a:t>'))).toHaveLength(1);
      // Each group references its own axis pair.
      const barIds = [
        .../<c:barChart>[\s\S]*?<\/c:barChart>/.exec(chartXml)![0].matchAll(/<c:axId val="(\d+)"\/>/g),
      ].map((m) => m[1]);
      const lineIds = [
        .../<c:lineChart>[\s\S]*?<\/c:lineChart>/.exec(chartXml)![0].matchAll(/<c:axId val="(\d+)"\/>/g),
      ].map((m) => m[1]);
      expect(barIds).toHaveLength(2);
      expect(lineIds).toHaveLength(2);
      expect(new Set([...barIds, ...lineIds]).size).toBe(4);
      expect(result.warnings.map((w) => w.code)).toContain('SECONDARY_AXIS');
    },
  },
];

// ---------------------------------------------------------------------------
// Suite
// ---------------------------------------------------------------------------

export function runExportFixtureSuite(Highcharts: HighchartsLike, version: 'v11' | 'v12' | 'v13'): void {
  const outDir = resolve(__dirname, '../output/export', version);
  const render = (o: Options): Chart => renderChart(Highcharts, o);

  describe(`export fixtures (Highcharts ${Highcharts.version ?? version})`, () => {
    afterEach(() => destroyAll());

    it.each(EXPORT_CASES.map((c, i) => [`${String(i + 1).padStart(2, '0')} ${c.name}`, c] as const))(
      '%s',
      async (_label, c) => {
        const chart = render(c.fixture) as AnyChart;
        c.before?.(chart);
        const result = await exportHighchartsToXlsx(chart, c.exportOptions);
        save(outDir, c.name, result.bytes);
        const x = await assertCommonPackage(result, c.excelType);
        const chartXml = x.chartXml(0);
        expect(x.plotGroupKinds(x.chartPaths()[0]!)).toEqual(c.groups);
        expect(x.sheetNames()).toEqual([
          { name: 'Chart', hidden: false },
          { name: 'Data', hidden: false },
        ]);

        if (c.categories !== false) {
          // Values + categories straight from the fixture: every visible series in order.
          const fs = x.seriesFormulas(x.chartPaths()[0]!);
          const series = (c.fixture.series ?? []) as Array<{ data?: unknown[]; name?: string; visible?: boolean }>;
          expect(fs).toHaveLength(series.length);
          const fixtureCats = (Array.isArray(c.fixture.xAxis) ? c.fixture.xAxis[0] : c.fixture.xAxis)?.categories;
          series.forEach((s, i) => {
            const f = fs[i]!;
            expect(f.val, `series ${i} c:val`).toBeDefined();
            expect(valuesAtRange(x, f.val!)).toEqual(yValues(s.data ?? []));
            expect(f.cat, `series ${i} c:cat`).toBeDefined();
            const cats = valuesAtRange(x, f.cat!);
            if (fixtureCats) expect(cats).toEqual(fixtureCats.slice(0, cats.length));
            else expect(cats).toEqual((chart.series[i].points as Array<{ name: string }>).map((p) => p.name));
            if (f.name) expect(valuesAtRange(x, f.name)).toEqual([s.name]);
          });
        }

        // Series colors follow the live chart (pie/doughnut use per-point colors instead).
        if (!['pie', 'doughnut', 'styled-mode'].includes(c.name)) {
          const sers = serBlocks(chartXml);
          const visible = (chart.series as AnyChart[]).filter((s) => s.visible !== false);
          expect(sers).toHaveLength(visible.length);
          visible.forEach((s, i) => {
            expect(seriesColor(sers[i]!), `series ${i} color`).toBe(hexOf(s.color));
          });
        }

        c.check?.({ x, chartXml, result, chart, options: c.fixture });
      },
    );

    it('18b hidden-series with seriesVisibility:"all" exports both series', async () => {
      const chart = render(F.hiddenSeries);
      const result = await exportHighchartsToXlsx(chart, { seriesVisibility: 'all' });
      save(outDir, 'hidden-series-all', result.bytes);
      const x = await assertCommonPackage(result, 'line');
      const fs = x.seriesFormulas(x.chartPaths()[0]!);
      expect(fs).toHaveLength(2);
      expect(valuesAtRange(x, fs[1]!.name!)).toEqual(['Hidden']);
      expect(valuesAtRange(x, fs[1]!.val!)).toEqual([3, 2, 1]);
      expect(result.warnings.map((w) => w.code)).toContain('HIDDEN_SERIES_INCLUDED');
    });

    it('includeReferenceImage:true in jsdom → still no image, WRITER_LIMITATION reported', async () => {
      const chart = render(F.simpleLine);
      const result = await exportHighchartsToXlsx(chart, { includeReferenceImage: true });
      save(outDir, 'reference-image-jsdom', result.bytes);
      const x = await assertCommonPackage(result, 'line');
      expect(x.hasImages()).toBe(false);
      const lim = result.warnings.filter((w) => w.code === 'WRITER_LIMITATION');
      expect(lim).toHaveLength(1);
      expect(lim[0]!.property).toBe('includeReferenceImage');
      expect(lim[0]!.outcome).not.toBe('blocking');
      expect(lim[0]!.message.toLowerCase()).toMatch(/browser|canvas|render/);
    });

    it('includeSourceData:false hides the Data sheet but the chart still references it', async () => {
      const chart = render(F.multiLine);
      const result = await exportHighchartsToXlsx(chart, { includeSourceData: false });
      save(outDir, 'hidden-data-sheet', result.bytes);
      const x = await assertCommonPackage(result, 'line');
      expect(x.sheetNames()).toEqual([
        { name: 'Chart', hidden: false },
        { name: 'Data', hidden: true },
      ]);
      const fs = x.seriesFormulas(x.chartPaths()[0]!);
      expect(fs.map((f) => f.val)).toEqual(['Data!$B$2:$B$7', 'Data!$C$2:$C$7', 'Data!$D$2:$D$7']);
      expect(valuesAtRange(x, fs[0]!.val!)).toEqual([1, 2, 3, 4, 5, 6]);
    });
  });
}
