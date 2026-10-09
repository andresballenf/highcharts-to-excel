import { createHash } from 'node:crypto';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  createDefaultExcelWriter,
  OoxmlExcelWriter,
  resetWorkerProbe,
  resolveZipMode,
  ZIP_DETAIL,
} from '../../src/excel/ooxml-writer';
import type { WorkbookSpec, WriteProgress } from '../../src/excel/writer-interface';
import { elementOrder, inspectXlsx, type XlsxInspection } from '../helpers/inspect-xlsx';
import {
  CAT_F,
  NAME1_F,
  NAME2_F,
  PNG_1X1,
  VAL1_F,
  VAL2_F,
  axis,
  chart,
  chartSpecs,
  fullWorkbook,
  series,
} from '../fixtures/writer-specs';

async function build(wb: WorkbookSpec = fullWorkbook()): Promise<{ bytes: Uint8Array; x: XlsxInspection }> {
  const bytes = await new OoxmlExcelWriter().write(wb);
  return { bytes, x: await inspectXlsx(bytes) };
}

const DATA = 'xl/worksheets/sheet1.xml';

describe('OoxmlExcelWriter – package', () => {
  it('produces the expected parts, [Content_Types].xml first, all well-formed', async () => {
    const { bytes, x } = await build();
    expect(x.parts[0]).toBe('[Content_Types].xml');
    expect([...x.parts].sort()).toEqual(
      [
        '[Content_Types].xml',
        '_rels/.rels',
        'docProps/app.xml',
        'docProps/core.xml',
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/styles.xml',
        'xl/theme/theme1.xml',
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/sheet2.xml',
        'xl/worksheets/sheet3.xml',
        'xl/worksheets/_rels/sheet2.xml.rels',
        'xl/drawings/drawing1.xml',
        'xl/drawings/_rels/drawing1.xml.rels',
        'xl/charts/chart1.xml',
        'xl/media/image1.png',
      ].sort(),
    );
    x.assertWellFormed();
    expect(x.bytes('xl/media/image1.png')).toEqual(PNG_1X1);
    // zip local header signature
    expect([...bytes.slice(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it('is deterministic for identical input', async () => {
    const a = await new OoxmlExcelWriter().write(fullWorkbook());
    const b = await createDefaultExcelWriter().write(fullWorkbook());
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
    expect(createDefaultExcelWriter().name).toBe('ooxml');
  });

  it('declares content types for every part', async () => {
    const { x } = await build();
    const ct = x.text('[Content_Types].xml');
    expect(ct).toContain(
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    );
    expect(ct).toContain('<Default Extension="xml" ContentType="application/xml"/>');
    expect(ct).toContain('<Default Extension="png" ContentType="image/png"/>');
    const overrides = [...ct.matchAll(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g)].map((m) => [
      m[1],
      m[2],
    ]);
    expect(Object.fromEntries(overrides)).toEqual({
      '/xl/workbook.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      '/xl/worksheets/sheet1.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      '/xl/worksheets/sheet2.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      '/xl/worksheets/sheet3.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      '/xl/drawings/drawing1.xml': 'application/vnd.openxmlformats-officedocument.drawing+xml',
      '/xl/charts/chart1.xml': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      '/xl/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml',
      '/xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
      '/docProps/core.xml': 'application/vnd.openxmlformats-package.core-properties+xml',
      '/docProps/app.xml': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    });
    expect(x.contentTypes()).toHaveLength(13);
    // every non-rels/png part is declared via Override or the xml default
    for (const p of x.parts) {
      if (p.endsWith('.rels') || p.endsWith('.png') || p === '[Content_Types].xml') continue;
      expect(
        overrides.some(([name]) => name === `/${p}`),
        p,
      ).toBe(true);
    }
  });

  it('writes doc props', async () => {
    const { x } = await build();
    const core = x.text('docProps/core.xml');
    expect(core).toContain('<dc:title>Writer test &amp; co</dc:title>');
    expect(core).toContain('<dc:creator>vitest</dc:creator>');
    expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2024-05-06T07:08:09Z</dcterms:created>');
    expect(x.text('docProps/app.xml')).toContain('<Application>');
    const rels = x.text('_rels/.rels');
    expect(rels).toContain('Target="xl/workbook.xml"');
    expect(rels).toContain('Target="docProps/core.xml"');
  });

  it('writes the workbook with sheet ids, rel ids and hidden state', async () => {
    const { x } = await build();
    expect(x.sheetNames()).toEqual([
      { name: 'Data', hidden: false },
      { name: 'Chart', hidden: false },
      { name: 'Meta', hidden: true },
    ]);
    const wb = x.text('xl/workbook.xml');
    expect(elementOrder(wb, 'workbook')).toEqual(['workbookPr', 'bookViews', 'sheets', 'calcPr']);
    expect(wb).toContain('<sheet name="Meta" sheetId="3" state="hidden" r:id="rId3"/>');
    expect(x.sheetPath('Chart')).toBe('xl/worksheets/sheet2.xml');
    const wbRels = x.text('xl/_rels/workbook.xml.rels');
    expect(wbRels).toContain('Target="styles.xml"');
    expect(wbRels).toContain('Target="theme/theme1.xml"');
  });

  it('first visible sheet is active when leading sheets are hidden', async () => {
    const wb = fullWorkbook();
    wb.sheets[0]!.hidden = true;
    const { x } = await build(wb);
    expect(x.text('xl/workbook.xml')).toContain('activeTab="1"');
    expect(x.text('xl/worksheets/sheet2.xml')).toContain('tabSelected="1"');
    expect(x.text('xl/worksheets/sheet1.xml')).not.toContain('tabSelected');
  });

  it('embeds a complete theme', async () => {
    const { x } = await build();
    const theme = x.text('xl/theme/theme1.xml');
    expect(elementOrder(theme, 'a:themeElements')).toEqual(['a:clrScheme', 'a:fontScheme', 'a:fmtScheme']);
    expect(elementOrder(theme, 'a:clrScheme')).toEqual([
      'a:dk1',
      'a:lt1',
      'a:dk2',
      'a:lt2',
      'a:accent1',
      'a:accent2',
      'a:accent3',
      'a:accent4',
      'a:accent5',
      'a:accent6',
      'a:hlink',
      'a:folHlink',
    ]);
    expect(theme).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');
    expect(theme).toContain('typeface="Calibri Light"');
    expect(elementOrder(theme, 'a:fmtScheme')).toEqual([
      'a:fillStyleLst',
      'a:lnStyleLst',
      'a:effectStyleLst',
      'a:bgFillStyleLst',
    ]);
    for (const lst of ['a:fillStyleLst', 'a:lnStyleLst', 'a:effectStyleLst', 'a:bgFillStyleLst']) {
      expect(elementOrder(theme, lst), lst).toHaveLength(3);
    }
  });
});

describe('OoxmlExcelWriter – worksheet', () => {
  it('orders worksheet children and writes freeze pane, widths and dimension', async () => {
    const { x } = await build();
    const s = x.text(DATA);
    expect(elementOrder(s, 'worksheet')).toEqual([
      'dimension',
      'sheetViews',
      'sheetFormatPr',
      'cols',
      'sheetData',
      'pageMargins',
    ]);
    expect(s).toContain('<dimension ref="A1:G8"/>');
    expect(s).toContain('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>');
    expect(s).toContain(
      '<col min="1" max="1" width="14" customWidth="1"/><col min="5" max="5" width="30.5" customWidth="1"/>',
    );
    expect(s).toContain('<sheetFormatPr defaultRowHeight="15"/>');
    const rowNums = [...s.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]));
    expect(rowNums).toEqual([1, 2, 3, 4, 5, 7, 8]);
    // chart sheet references its drawing
    const chartSheet = x.text('xl/worksheets/sheet2.xml');
    expect(elementOrder(chartSheet, 'worksheet')).toEqual([
      'dimension',
      'sheetViews',
      'sheetFormatPr',
      'sheetData',
      'pageMargins',
      'drawing',
    ]);
    expect(chartSheet).toContain('<drawing r:id="rId1"/>');
    expect(x.text('xl/worksheets/_rels/sheet2.xml.rels')).toContain('Target="../drawings/drawing1.xml"');
  });

  it('writes typed cell values', async () => {
    const { x } = await build();
    expect(x.cellValue(DATA, 'A1')).toBe('Quarter');
    expect(x.cellValue(DATA, 'B2')).toBe(1);
    expect(x.cellValue(DATA, 'G2')).toBe(0.125);
    expect(x.cellValue(DATA, 'G5')).toBe(0); // -0 written as 0
    expect(x.cellXml(DATA, 'G5')).toContain('<v>0</v>');
    expect(x.cellValue(DATA, 'F8')).toBe(1);
    expect(x.cellXml(DATA, 'F8')).toContain('t="b"');
    // #N/A error cell (alignment gaps: line/area charts connect across it)
    expect(x.cellXml(DATA, 'G8')).toBe('<c r="G8" t="e"><v>#N/A</v></c>');
    expect(x.cellValue(DATA, 'G8')).toBe('#N/A');
    // blank without style → omitted; blank with style → styled empty cell
    expect(x.cellXml(DATA, 'B3')).toBeNull();
    expect(x.cellXml(DATA, 'A4')).toMatch(/^<c r="A4" s="\d+"\/>$/);
    // NaN → blank (no style → omitted)
    expect(x.cellXml(DATA, 'B5')).toBeNull();
  });

  it('keeps formula-like strings as inline text (formula-injection protection)', async () => {
    const { x } = await build();
    const s = x.text(DATA);
    expect(s).not.toContain('<f>');
    expect(s).not.toContain('<f ');
    for (const [ref, text] of [
      ['A8', '=1+1'],
      ['B8', '+foo'],
      ['C8', '-5'],
      ['D8', '@x'],
    ] as const) {
      expect(x.cellXml(DATA, ref)).toContain('t="inlineStr"');
      expect(x.cellValue(DATA, ref)).toBe(text);
    }
    expect(x.cellXml(DATA, 'A8')).toBe('<c r="A8" t="inlineStr"><is><t xml:space="preserve">=1+1</t></is></c>');
  });

  it('escapes XML specials and strips illegal control characters', async () => {
    const { x } = await build();
    const raw = x.cellXml(DATA, 'E8')!;
    expect(raw).toContain('Ünïcødé &amp; &lt;tags&gt; "q" end');
    expect(raw).not.toContain('\u0001');
    expect(x.cellValue(DATA, 'E8')).toBe('Ünïcødé & <tags> "q" end');
  });
});

describe('OoxmlExcelWriter – strictness (W4, W5, W8, W9, W10)', () => {
  const oneSheet = (cells: WorkbookSpec['sheets'][number]['rows'][number]['cells']): WorkbookSpec => {
    const wb = fullWorkbook();
    wb.sheets[0]!.rows = [{ row0: 0, cells }];
    return wb;
  };

  it('W4: escapes _xHHHH_ sequences in cell text', async () => {
    const { x } = await build(oneSheet([{ col0: 0, row0: 0, value: { type: 'string', value: '_x0041_BC' } }]));
    expect(x.cellXml(DATA, 'A1')).toBe(
      '<c r="A1" t="inlineStr"><is><t xml:space="preserve">_x005F_x0041_BC</t></is></c>',
    );
  });

  it('W8: encodes carriage returns in text content', async () => {
    const { x } = await build(oneSheet([{ col0: 0, row0: 0, value: { type: 'string', value: 'a\r\nb' } }]));
    expect(x.cellXml(DATA, 'A1')).toContain('>a&#13;\nb</t>');
  });

  it('W5: rejects invalid number format codes in cell styles', async () => {
    for (const numberFormat of ['0"x', '[Red0', '0;0;0;0;0', '0'.repeat(256)]) {
      const wb = oneSheet([{ col0: 0, row0: 0, value: { type: 'number', value: 1 }, style: { numberFormat } }]);
      await expect(new OoxmlExcelWriter().write(wb), numberFormat).rejects.toThrow(/number format/);
    }
  });

  it('W9: only left/center/right alignments are written', async () => {
    const wb = oneSheet([
      {
        col0: 0,
        row0: 0,
        value: { type: 'number', value: 1 },
        style: { align: 'center" foo="x' as 'center', bold: true },
      },
      { col0: 1, row0: 0, value: { type: 'number', value: 2 }, style: { align: 'right' } },
    ]);
    const { x } = await build(wb);
    const styles = x.text('xl/styles.xml');
    expect(styles).not.toContain('foo=');
    expect(styles.match(/<alignment /g)).toHaveLength(1);
    expect(styles).toContain('<alignment horizontal="right"/>');
  });

  it('W10: resolves every registered style object (same index for equal styles)', async () => {
    const shared = { bold: true };
    const wb = oneSheet([
      { col0: 0, row0: 0, value: { type: 'number', value: 1 }, style: shared },
      { col0: 1, row0: 0, value: { type: 'number', value: 2 }, style: shared },
      { col0: 2, row0: 0, value: { type: 'number', value: 3 }, style: { bold: true } },
    ]);
    const { x } = await build(wb);
    const s = (ref: string) => /\bs="(\d+)"/.exec(x.cellXml(DATA, ref)!)?.[1];
    expect(s('A1')).toBeDefined();
    expect(s('B1')).toBe(s('A1'));
    expect(s('C1')).toBe(s('A1'));
  });
});

describe('OoxmlExcelWriter – styles', () => {
  it('writes mandatory fills, built-in and custom numFmts, fonts and cellXfs', async () => {
    const { x } = await build();
    const st = x.text('xl/styles.xml');
    expect(elementOrder(st, 'styleSheet')).toEqual([
      'numFmts',
      'fonts',
      'fills',
      'borders',
      'cellStyleXfs',
      'cellXfs',
      'cellStyles',
      'dxfs',
      'tableStyles',
    ]);
    const fills = elementOrder(st, 'fills');
    expect(fills.length).toBeGreaterThanOrEqual(3);
    expect(st).toMatch(
      /<fills count="\d+"><fill><patternFill patternType="none"\/><\/fill><fill><patternFill patternType="gray125"\/><\/fill>/,
    );
    expect(st).toContain('<fgColor rgb="FFDDEBF7"/>');
    // custom formats start at 164; "0.00" maps to built-in id 2
    expect(st).toContain('<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>');
    expect(st).toContain('<numFmt numFmtId="165" formatCode="0.0%"/>');
    expect(st).not.toContain('formatCode="0.00"');
    expect(st).toMatch(
      /<xf numFmtId="2" fontId="\d+" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"\/>/,
    );
    expect(st).toContain('<cellXfs count="6">'); // default + header, date, percent, italic blank, notes
    expect(st).toContain('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
    const parsed = x.xml('xl/styles.xml');
    expect(parsed.styleSheet.cellXfs.xf[0]['@_numFmtId']).toBe('0');
    // header style: bold + fill + center alignment
    const s = x.cellXml(DATA, 'A1')!;
    const idx = Number(/s="(\d+)"/.exec(s)![1]);
    const xf = parsed.styleSheet.cellXfs.xf[idx];
    expect(xf['@_applyFont']).toBe('1');
    expect(xf['@_applyFill']).toBe('1');
    expect(xf['@_applyAlignment']).toBe('1');
    expect(xf.alignment['@_horizontal']).toBe('center');
    expect(parsed.styleSheet.fonts.font[Number(xf['@_fontId'])].b).toBeDefined();
    // the date cells reference the yyyy-mm-dd xf
    const dateIdx = Number(/s="(\d+)"/.exec(x.cellXml(DATA, 'F2')!)![1]);
    expect(parsed.styleSheet.cellXfs.xf[dateIdx]['@_numFmtId']).toBe('164');
  });
});

describe('OoxmlExcelWriter – drawings', () => {
  it('writes chart + image anchors and relationships', async () => {
    const { x } = await build();
    const d = x.text('xl/drawings/drawing1.xml');
    expect(elementOrder(d, 'xdr:wsDr')).toEqual(['xdr:oneCellAnchor', 'xdr:oneCellAnchor']);
    expect(elementOrder(d, 'xdr:oneCellAnchor')).toEqual(['xdr:from', 'xdr:ext', 'xdr:graphicFrame', 'xdr:clientData']);
    expect(elementOrder(d, 'xdr:from')).toEqual(['xdr:col', 'xdr:colOff', 'xdr:row', 'xdr:rowOff']);
    expect(d).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>38100</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>',
    );
    expect(d).toContain('<xdr:ext cx="5715000" cy="3810000"/>');
    expect(d).toContain('<xdr:cNvPr id="2" name="Chart 1"/>');
    expect(d).toContain('r:id="rId1"/></a:graphicData>');
    expect(d).toContain('<a:blip r:embed="rId2"/>');
    expect(elementOrder(d, 'xdr:graphicFrame')).toEqual(['xdr:nvGraphicFramePr', 'xdr:xfrm', 'a:graphic']);
    expect(elementOrder(d, 'xdr:pic')).toEqual(['xdr:nvPicPr', 'xdr:blipFill', 'xdr:spPr']);
    const rels = x.text('xl/drawings/_rels/drawing1.xml.rels');
    expect(rels).toContain(
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
    );
    expect(rels).toContain(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>',
    );
    expect(x.hasImages()).toBe(true);
  });

  it('chart part references worksheet cells', async () => {
    const { x } = await build();
    expect(x.chartPaths()).toEqual(['xl/charts/chart1.xml']);
    expect(x.plotGroupKinds('xl/charts/chart1.xml')).toEqual(['barChart', 'lineChart']);
    expect(x.seriesFormulas('xl/charts/chart1.xml')).toEqual([
      { name: NAME1_F, cat: CAT_F, val: VAL1_F },
      { name: NAME2_F, cat: CAT_F, val: VAL2_F },
    ]);
    expect(x.chartXml(0)).toContain('<c:chartSpace');
  });

  it('numbers charts globally and omits media without images', async () => {
    const wb = fullWorkbook({ charts: [chartSpecs.pie(), chartSpecs.scatter()], image: false });
    wb.sheets[0]!.drawings = [
      {
        kind: 'chart',
        chart: chartSpecs.bubble(),
        name: 'On data',
        anchor: { col0: 9, row0: 0, colOffsetPx: 0, rowOffsetPx: 0, widthPx: 300, heightPx: 200 },
      },
    ];
    const { x } = await build(wb);
    x.assertWellFormed();
    expect(x.hasImages()).toBe(false);
    expect(x.parts.some((p) => p.startsWith('xl/media/'))).toBe(false);
    expect(x.chartPaths()).toEqual(['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml']);
    expect(x.plotGroupKinds('xl/charts/chart1.xml')).toEqual(['bubbleChart']);
    expect(x.plotGroupKinds('xl/charts/chart2.xml')).toEqual(['pieChart']);
    expect(x.text('xl/drawings/_rels/drawing2.xml.rels')).toContain('Target="../charts/chart3.xml"');
    expect(x.seriesFormulas('xl/charts/chart3.xml')).toEqual([
      { name: NAME1_F, xVal: "'Data'!$D$2:$D$5", yVal: VAL1_F },
    ]);
    expect(x.text('xl/worksheets/sheet1.xml')).toMatch(/<pageMargins [^>]*\/><drawing r:id="rId1"\/><\/worksheet>/);
  });
});

describe('OoxmlExcelWriter – invalid input', () => {
  const w = new OoxmlExcelWriter();
  it('rejects workbooks without visible sheets', async () => {
    const wb = fullWorkbook();
    for (const s of wb.sheets) s.hidden = true;
    await expect(w.write(wb)).rejects.toThrow(/visible/);
    await expect(w.write({ properties: {}, sheets: [] })).rejects.toThrow(/at least one sheet/);
  });
  it('rejects bad / duplicate sheet names', async () => {
    const wb = fullWorkbook();
    wb.sheets[1]!.name = 'data';
    await expect(w.write(wb)).rejects.toThrow(/Duplicate/);
    const wb2 = fullWorkbook();
    wb2.sheets[1]!.name = 'a/b';
    await expect(w.write(wb2)).rejects.toThrow(/Invalid sheet name/);
  });
  it('rejects duplicate cells, out-of-range rows and non-PNG images', async () => {
    const wb = fullWorkbook();
    wb.sheets[0]!.rows.push({ row0: 0, cells: [{ col0: 0, row0: 0, value: { type: 'number', value: 1 } }] });
    await expect(w.write(wb)).rejects.toThrow(/duplicate cell A1/);
    const wb2 = fullWorkbook();
    wb2.sheets[0]!.rows.push({
      row0: 2_000_000,
      cells: [{ col0: 0, row0: 2_000_000, value: { type: 'number', value: 1 } }],
    });
    await expect(w.write(wb2)).rejects.toThrow(/outside/);
    const wb3 = fullWorkbook();
    const img = wb3.sheets[1]!.drawings.find((d) => d.kind === 'image')!;
    if (img.kind === 'image') img.png = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]);
    await expect(w.write(wb3)).rejects.toThrow(/not a PNG/);
  });
});

// ---------------------------------------------------------------------------
// Chunked writing: golden bytes, zip mode, abort, progress, yielding
// ---------------------------------------------------------------------------

/**
 * A Data sheet of `n` rows (strings, numbers, NaN, -0, blanks, booleans, a shared date style) and a
 * Chart sheet with a line chart (string categories) and a scatter chart (numeric X), whose caches
 * hold `n` points each. Large enough to cross many 2,000-item chunk boundaries and to make the
 * worksheet and chart parts exceed fflate's 160,000-byte worker threshold.
 */
function largeWorkbook(n = 25_000): WorkbookSpec {
  const date = { numberFormat: 'yyyy-mm-dd' };
  const header = { bold: true, fillHex: 'DDEBF7' };
  const rows: WorkbookSpec['sheets'][number]['rows'] = [
    {
      row0: 0,
      cells: ['Label', 'Sales', 'Costs', 'Date', 'Flag'].map((t, c) => ({
        col0: c,
        row0: 0,
        value: { type: 'string' as const, value: t },
        style: header,
      })),
    },
  ];
  const cats: Array<string | null> = [];
  const xs: number[] = [];
  const v1: Array<number | null> = [];
  const v2: Array<number | null> = [];
  for (let i = 0; i < n; i++) {
    const r = i + 1;
    const label = i % 97 === 0 ? null : `P${i} & <${i % 7}>`;
    const a = i % 50 === 0 ? Number.NaN : Math.round(Math.sin(i / 40) * 5_000) / 100;
    const b = i % 333 === 0 ? -0 : (i * 7) % 1_000;
    cats.push(label);
    xs.push(i / 4);
    v1.push(Number.isNaN(a) ? null : a);
    v2.push(b);
    rows.push({
      row0: r,
      cells: [
        label === null
          ? { col0: 0, row0: r, value: { type: 'blank' } }
          : { col0: 0, row0: r, value: { type: 'string', value: label } },
        { col0: 1, row0: r, value: { type: 'number', value: a } },
        { col0: 2, row0: r, value: { type: 'number', value: b } },
        { col0: 3, row0: r, value: { type: 'number', value: 45_000 + i }, style: date },
        ...(i % 3 === 0 ? [{ col0: 4, row0: r, value: { type: 'boolean' as const, value: i % 2 === 0 } }] : []),
      ],
    });
  }
  const last = n + 1;
  const ref = (col: string) => `'Data'!$${col}$2:$${col}$${last}`;
  const line = chart(
    [
      {
        kind: 'line',
        grouping: 'standard',
        varyColors: false,
        showMarkers: false,
        axisIds: [100, 200],
        dataLabels: null,
        series: [
          series(0, {
            categories: { kind: 'str', formula: ref('A'), cache: cats },
            values: { formula: ref('B'), cache: v1 },
          }),
          series(1, {
            categories: { kind: 'str', formula: ref('A'), cache: cats },
            values: { formula: ref('C'), cache: v2, formatCode: '0.0' },
          }),
        ],
      },
    ],
    [axis(100, 'cat', 200), axis(200, 'val', 100)],
  );
  const scatter = chart(
    [
      {
        kind: 'scatter',
        scatterStyle: 'lineMarker',
        varyColors: false,
        axisIds: [300, 400],
        dataLabels: null,
        series: [
          series(0, {
            categories: { kind: 'num', formula: ref('D'), cache: xs },
            values: { formula: ref('C'), cache: v2 },
          }),
        ],
      },
    ],
    [axis(300, 'val', 400, { position: 'b' }), axis(400, 'val', 300)],
  );
  return {
    properties: { title: 'Large', creator: 'vitest', created: new Date('2024-05-06T07:08:09.123Z') },
    sheets: [
      {
        name: 'Data',
        hidden: false,
        columns: [{ col0: 0, widthChars: 20 }],
        rows,
        freezeHeaderRow: true,
        drawings: [],
      },
      {
        name: 'Chart',
        hidden: false,
        columns: [],
        rows: [],
        freezeHeaderRow: false,
        drawings: [line, scatter].map((c, i) => ({
          kind: 'chart' as const,
          chart: c,
          name: `Chart ${i + 1}`,
          anchor: { col0: 1, row0: 1 + i * 22, colOffsetPx: 0, rowOffsetPx: 0, widthPx: 600, heightPx: 400 },
        })),
      },
    ],
  };
}

const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

/** Fixture specs → sha256 of the package, captured with the pre-chunking writer (zipSync, one pass). */
const GOLDEN: Record<string, { spec: () => WorkbookSpec; sha256: string }> = {
  fullWorkbook: {
    spec: () => fullWorkbook(),
    sha256: '87865b893d6f0a50670189df6c3d72200c960560bbc7b4adaf0efc13031a5063',
  },
  allCharts: {
    spec: () => fullWorkbook({ charts: Object.values(chartSpecs).map((make) => make()), image: false }),
    sha256: '637a7c76bc69d42ee1521c46556555e657cb6c326027a2918baf152da9d9f358',
  },
  large25k: { spec: () => largeWorkbook(), sha256: '34bbe3102f7b9bdf445f5a5dbb5bc6ec33630f93a33f318593ecc6429cb5c22c' },
};

describe('OoxmlExcelWriter – golden bytes', () => {
  it.each(Object.keys(GOLDEN))('%s matches the pre-chunking bytes', async (name) => {
    const { spec, sha256: expected } = GOLDEN[name]!;
    const bytes = await new OoxmlExcelWriter().write(spec());
    expect(sha256(bytes)).toBe(expected);
  });
});

/** A stand-in for the browser `Worker` that answers the writer's start-up probe (or fails it). */
function fakeWorkerClass(behaviour: 'message' | 'error' | 'throw') {
  return class FakeWorker {
    onmessage: ((e: unknown) => void) | null = null;
    onerror: ((e: unknown) => void) | null = null;
    constructor() {
      if (behaviour === 'throw') throw new Error('blocked by CSP');
      setTimeout(() => (behaviour === 'message' ? this.onmessage?.({ data: 1 }) : this.onerror?.({})), 1);
    }
    postMessage(): void {}
    terminate(): void {}
  };
}

describe('OoxmlExcelWriter – zip mode', () => {
  const hadCreateObjectURL = typeof URL.createObjectURL === 'function';
  afterEach(() => {
    vi.unstubAllGlobals();
    resetWorkerProbe();
    if (!hadCreateObjectURL) delete (URL as { createObjectURL?: unknown }).createObjectURL;
  });
  const stubBlobUrls = () => {
    if (!hadCreateObjectURL) (URL as { createObjectURL?: unknown }).createObjectURL = () => 'blob:probe';
  };
  const zipDetail = async (writer: OoxmlExcelWriter, spec: WorkbookSpec): Promise<string | undefined> => {
    const events: WriteProgress[] = [];
    await writer.write(spec, { onProgress: (p) => events.push(p) });
    return events.find((e) => e.phase === 'zip')?.detail;
  };

  it("'auto' (the default) zips synchronously under jsdom, which has no Worker", async () => {
    expect(typeof (globalThis as { Worker?: unknown }).Worker).toBe('undefined');
    expect(await resolveZipMode('auto', 50_000_000)).toBe('sync');
    expect(await zipDetail(new OoxmlExcelWriter(), largeWorkbook(3_000))).toBe(ZIP_DETAIL.sync);
  });

  it("'auto' picks the worker zip only for large parts and only when a Blob worker starts", async () => {
    stubBlobUrls();
    vi.stubGlobal('Worker', fakeWorkerClass('message'));
    expect(await resolveZipMode('auto', 1_000)).toBe('sync');
    expect(await resolveZipMode('auto', 5_000_000)).toBe('async');
    resetWorkerProbe();
    vi.stubGlobal('Worker', fakeWorkerClass('error'));
    expect(await resolveZipMode('auto', 5_000_000)).toBe('sync');
    resetWorkerProbe();
    vi.stubGlobal('Worker', fakeWorkerClass('throw'));
    expect(await resolveZipMode('auto', 5_000_000)).toBe('sync');
    expect(await resolveZipMode('sync', 5_000_000)).toBe('sync');
  });

  it("forcing 'async' without Web Workers rejects with a clear error", async () => {
    await expect(new OoxmlExcelWriter({ zip: 'async' }).write(fullWorkbook())).rejects.toThrow(
      /zip: 'async' needs Web Workers/,
    );
  });

  it('the worker zip produces the same bytes as zipSync', async () => {
    // Satisfies the guard only: under vitest fflate resolves to its Node build, whose async zip runs
    // on worker_threads. The browser build (Blob URL workers) is exercised by the Playwright downloads.
    vi.stubGlobal('Worker', class {});
    const writer = new OoxmlExcelWriter({ zip: 'async' });
    expect(await zipDetail(writer, largeWorkbook(3_000))).toBe(ZIP_DETAIL.async);
    expect(sha256(await writer.write(GOLDEN.large25k!.spec()))).toBe(GOLDEN.large25k!.sha256);
    expect(sha256(await writer.write(GOLDEN.fullWorkbook!.spec()))).toBe(GOLDEN.fullWorkbook!.sha256);
  });

  it('aborting during the worker zip terminates it and rejects with the reason', async () => {
    vi.stubGlobal('Worker', class {});
    const controller = new AbortController();
    const reason = new Error('stop zipping');
    const write = new OoxmlExcelWriter({ zip: 'async' }).write(largeWorkbook(), {
      signal: controller.signal,
      onProgress: (p) => {
        if (p.phase === 'zip') controller.abort(reason);
      },
    });
    await expect(write).rejects.toBe(reason);
  });

  it('rejects an unknown zip mode', () => {
    expect(() => new OoxmlExcelWriter({ zip: 'fast' as 'sync' })).toThrow(/zip/);
  });
});

describe('OoxmlExcelWriter – chunks, progress and abort', () => {
  it.each(Object.keys(GOLDEN))('%s keeps its golden bytes when written in small chunks with progress', async (name) => {
    const { spec, sha256: expected } = GOLDEN[name]!;
    const events: WriteProgress[] = [];
    const bytes = await new OoxmlExcelWriter().write(spec(), { yieldEvery: 700, onProgress: (p) => events.push(p) });
    expect(sha256(bytes)).toBe(expected);
    expect(events.length).toBeGreaterThan(1);
  });

  it('reports monotonic write progress per chunk, then zip, ending at 1', async () => {
    const events: WriteProgress[] = [];
    await new OoxmlExcelWriter().write(largeWorkbook(10_000), { yieldEvery: 2_000, onProgress: (p) => events.push(p) });
    const write = events.filter((e) => e.phase === 'write');
    const zip = events.filter((e) => e.phase === 'zip');
    // 10,001 rows × ≤5 cells (grouped and emitted) + 3 × 10,000 cache points, in 2,000-item chunks.
    expect(write.length).toBeGreaterThanOrEqual(30);
    const lastWrite = events.map((e) => e.phase).lastIndexOf('write');
    expect(events.findIndex((e) => e.phase === 'zip')).toBeGreaterThan(lastWrite);
    for (const list of [write, zip]) {
      for (let i = 1; i < list.length; i++) expect(list[i]!.fraction).toBeGreaterThanOrEqual(list[i - 1]!.fraction);
      for (const e of list) expect(e.fraction).toBeGreaterThanOrEqual(0);
      expect(list.at(-1)!.fraction).toBe(1);
    }
    expect(zip[0]!.fraction).toBe(0);
  });

  it('encodes large chart and sheet parts piece by piece, pausing without advancing the fraction', async () => {
    const events: WriteProgress[] = [];
    await new OoxmlExcelWriter().write(largeWorkbook(), { yieldEvery: 2_000, onProgress: (p) => events.push(p) });
    // Encoding pauses report 0 new items, so they repeat the fraction of the event before.
    const pauses = new Set(
      events
        .filter((e, i) => i > 0 && e.phase === 'write' && e.detail && e.fraction === events[i - 1]!.fraction)
        .map((e) => e.detail),
    );
    expect(pauses).toEqual(new Set(['sheet "Data"', 'chart 1', 'chart 2']));
  });

  it('rejects with signal.reason when the signal is already aborted', async () => {
    const reason = new Error('cancelled');
    const progress = vi.fn();
    await expect(
      new OoxmlExcelWriter().write(fullWorkbook(), { signal: AbortSignal.abort(reason), onProgress: progress }),
    ).rejects.toBe(reason);
    expect(progress).not.toHaveBeenCalled();
  });

  it('stops between chunks when aborted mid-write and never zips', async () => {
    const controller = new AbortController();
    const events: WriteProgress[] = [];
    const write = new OoxmlExcelWriter().write(largeWorkbook(10_000), {
      signal: controller.signal,
      yieldEvery: 1_000,
      onProgress: (p) => {
        events.push(p);
        if (p.phase === 'write' && p.fraction > 0.3) controller.abort();
      },
    });
    await expect(write).rejects.toMatchObject({ name: 'AbortError' });
    const at = events.findIndex((e) => e.fraction > 0.3);
    expect(events).toHaveLength(at + 1);
    expect(events.some((e) => e.phase === 'zip')).toBe(false);
  });

  it('yields to the event loop between chunks in a browser-like environment (jsdom has window)', async () => {
    let ticks = 0;
    const timer = setInterval(() => ticks++, 1);
    try {
      await new OoxmlExcelWriter().write(largeWorkbook(), { yieldEvery: 2_000 });
    } finally {
      clearInterval(timer);
    }
    expect(ticks).toBeGreaterThan(0);
  });
});
