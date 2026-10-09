import { describe, expect, it } from 'vitest';
import { createDefaultExcelWriter, OoxmlExcelWriter } from '../../src/excel/ooxml-writer';
import type { WorkbookSpec } from '../../src/excel/writer-interface';
import { elementOrder, inspectXlsx, type XlsxInspection } from '../helpers/inspect-xlsx';
import { CAT_F, NAME1_F, NAME2_F, PNG_1X1, VAL1_F, VAL2_F, chartSpecs, fullWorkbook } from '../fixtures/writer-specs';

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
    expect(ct).toContain('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    expect(ct).toContain('<Default Extension="xml" ContentType="application/xml"/>');
    expect(ct).toContain('<Default Extension="png" ContentType="image/png"/>');
    const overrides = [...ct.matchAll(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g)].map((m) => [m[1], m[2]]);
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
      expect(overrides.some(([name]) => name === `/${p}`), p).toBe(true);
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
      'a:dk1', 'a:lt1', 'a:dk2', 'a:lt2', 'a:accent1', 'a:accent2', 'a:accent3', 'a:accent4', 'a:accent5', 'a:accent6', 'a:hlink', 'a:folHlink',
    ]);
    expect(theme).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');
    expect(theme).toContain('typeface="Calibri Light"');
    expect(elementOrder(theme, 'a:fmtScheme')).toEqual(['a:fillStyleLst', 'a:lnStyleLst', 'a:effectStyleLst', 'a:bgFillStyleLst']);
    for (const lst of ['a:fillStyleLst', 'a:lnStyleLst', 'a:effectStyleLst', 'a:bgFillStyleLst']) {
      expect(elementOrder(theme, lst), lst).toHaveLength(3);
    }
  });
});

describe('OoxmlExcelWriter – worksheet', () => {
  it('orders worksheet children and writes freeze pane, widths and dimension', async () => {
    const { x } = await build();
    const s = x.text(DATA);
    expect(elementOrder(s, 'worksheet')).toEqual(['dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'pageMargins']);
    expect(s).toContain('<dimension ref="A1:G8"/>');
    expect(s).toContain('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>');
    expect(s).toContain('<col min="1" max="1" width="14" customWidth="1"/><col min="5" max="5" width="30.5" customWidth="1"/>');
    expect(s).toContain('<sheetFormatPr defaultRowHeight="15"/>');
    const rowNums = [...s.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]));
    expect(rowNums).toEqual([1, 2, 3, 4, 5, 7, 8]);
    // chart sheet references its drawing
    const chartSheet = x.text('xl/worksheets/sheet2.xml');
    expect(elementOrder(chartSheet, 'worksheet')).toEqual(['dimension', 'sheetViews', 'sheetFormatPr', 'sheetData', 'pageMargins', 'drawing']);
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

describe('OoxmlExcelWriter – styles', () => {
  it('writes mandatory fills, built-in and custom numFmts, fonts and cellXfs', async () => {
    const { x } = await build();
    const st = x.text('xl/styles.xml');
    expect(elementOrder(st, 'styleSheet')).toEqual(['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles']);
    const fills = elementOrder(st, 'fills');
    expect(fills.length).toBeGreaterThanOrEqual(3);
    expect(st).toMatch(/<fills count="\d+"><fill><patternFill patternType="none"\/><\/fill><fill><patternFill patternType="gray125"\/><\/fill>/);
    expect(st).toContain('<fgColor rgb="FFDDEBF7"/>');
    // custom formats start at 164; "0.00" maps to built-in id 2
    expect(st).toContain('<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>');
    expect(st).toContain('<numFmt numFmtId="165" formatCode="0.0%"/>');
    expect(st).not.toContain('formatCode="0.00"');
    expect(st).toMatch(/<xf numFmtId="2" fontId="\d+" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"\/>/);
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
    expect(d).toContain('<xdr:from><xdr:col>1</xdr:col><xdr:colOff>38100</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>');
    expect(d).toContain('<xdr:ext cx="5715000" cy="3810000"/>');
    expect(d).toContain('<xdr:cNvPr id="2" name="Chart 1"/>');
    expect(d).toContain('r:id="rId1"/></a:graphicData>');
    expect(d).toContain('<a:blip r:embed="rId2"/>');
    expect(elementOrder(d, 'xdr:graphicFrame')).toEqual(['xdr:nvGraphicFramePr', 'xdr:xfrm', 'a:graphic']);
    expect(elementOrder(d, 'xdr:pic')).toEqual(['xdr:nvPicPr', 'xdr:blipFill', 'xdr:spPr']);
    const rels = x.text('xl/drawings/_rels/drawing1.xml.rels');
    expect(rels).toContain('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>');
    expect(rels).toContain('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>');
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
      { kind: 'chart', chart: chartSpecs.bubble(), name: 'On data', anchor: { col0: 9, row0: 0, colOffsetPx: 0, rowOffsetPx: 0, widthPx: 300, heightPx: 200 } },
    ];
    const { x } = await build(wb);
    x.assertWellFormed();
    expect(x.hasImages()).toBe(false);
    expect(x.parts.some((p) => p.startsWith('xl/media/'))).toBe(false);
    expect(x.chartPaths()).toEqual(['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml']);
    expect(x.plotGroupKinds('xl/charts/chart1.xml')).toEqual(['bubbleChart']);
    expect(x.plotGroupKinds('xl/charts/chart2.xml')).toEqual(['pieChart']);
    expect(x.text('xl/drawings/_rels/drawing2.xml.rels')).toContain('Target="../charts/chart3.xml"');
    expect(x.seriesFormulas('xl/charts/chart3.xml')).toEqual([{ name: NAME1_F, xVal: "'Data'!$D$2:$D$5", yVal: VAL1_F }]);
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
    wb2.sheets[0]!.rows.push({ row0: 2_000_000, cells: [{ col0: 0, row0: 2_000_000, value: { type: 'number', value: 1 } }] });
    await expect(w.write(wb2)).rejects.toThrow(/outside/);
    const wb3 = fullWorkbook();
    const img = wb3.sheets[1]!.drawings.find((d) => d.kind === 'image')!;
    if (img.kind === 'image') img.png = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]);
    await expect(w.write(wb3)).rejects.toThrow(/not a PNG/);
  });
});
