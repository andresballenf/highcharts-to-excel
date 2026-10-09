/**
 * Hand-written OOXML (SpreadsheetML + DrawingML) writer producing native, editable Excel charts.
 * Browser-safe: depends only on fflate for zipping.
 */
import { zipSync, type ZipOptions } from 'fflate';
import type { CellStyleSpec, ExcelWriter, SheetSpec, WorkbookSpec } from './writer-interface';
import { buildChartXml } from './chart-xml';
import { buildDrawingXml, type DrawingTarget } from './drawing-xml';
import {
  buildAppXml,
  buildContentTypesXml,
  buildCoreXml,
  buildRootRelsXml,
  buildThemeXml,
  buildWorkbookRelsXml,
  buildWorkbookXml,
  CT,
  type ContentTypeOverride,
} from './package-parts';
import { buildStyles } from './styles-xml';
import { buildSheetRelsXml, buildWorksheetXml } from './worksheet-xml';

/** Fixed zip entry timestamp so identical input yields identical bytes. */
const ZIP_MTIME = new Date(2000, 0, 1, 0, 0, 0);
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const INVALID_SHEET_NAME_CHARS = /[\][:*?/\\]/;

function validateSheets(sheets: readonly SheetSpec[]): void {
  if (!Array.isArray(sheets) || sheets.length === 0) throw new Error('Workbook must contain at least one sheet');
  if (!sheets.some((s) => !s.hidden)) throw new Error('Workbook must contain at least one visible sheet');
  const seen = new Set<string>();
  for (const s of sheets) {
    const name = s.name;
    if (typeof name !== 'string' || name.length === 0 || name.length > 31) {
      throw new Error(`Invalid sheet name "${String(name)}": must be 1-31 characters`);
    }
    if (INVALID_SHEET_NAME_CHARS.test(name) || name.startsWith("'") || name.endsWith("'")) {
      throw new Error(`Invalid sheet name "${name}": contains characters Excel does not allow`);
    }
    const key = name.toLowerCase();
    if (seen.has(key)) throw new Error(`Duplicate sheet name "${name}"`);
    seen.add(key);
  }
}

function isPng(bytes: Uint8Array): boolean {
  return bytes instanceof Uint8Array && bytes.length > PNG_SIGNATURE.length && PNG_SIGNATURE.every((b, i) => bytes[i] === b);
}

/** Build every package part (path → bytes) in zip order, `[Content_Types].xml` first. */
export function buildPackageParts(wb: WorkbookSpec): Map<string, Uint8Array> {
  validateSheets(wb.sheets);
  const enc = new TextEncoder();

  const allStyles: Array<CellStyleSpec | undefined> = [];
  for (const sheet of wb.sheets) for (const row of sheet.rows) for (const c of row.cells) allStyles.push(c.style);
  const styles = buildStyles(allStyles);

  const overrides: ContentTypeOverride[] = [
    { partName: '/xl/workbook.xml', contentType: CT.workbook },
  ];
  const xlParts: Array<[string, Uint8Array]> = [];
  const activeTab = wb.sheets.findIndex((s) => !s.hidden);
  const { xml: workbookRels, sheetRelIds } = buildWorkbookRelsXml(wb.sheets.length);

  let drawingCount = 0;
  let chartCount = 0;
  let imageCount = 0;
  const chartParts: Array<[string, Uint8Array]> = [];
  const mediaParts: Array<[string, Uint8Array]> = [];

  wb.sheets.forEach((sheet, i) => {
    const sheetNo = i + 1;
    let drawingRelId: string | null = null;
    if (sheet.drawings.length > 0) {
      const drawingNo = ++drawingCount;
      drawingRelId = 'rId1';
      const targets: DrawingTarget[] = sheet.drawings.map((d, j) => {
        const relId = `rId${j + 1}`;
        if (d.kind === 'chart') {
          const n = ++chartCount;
          chartParts.push([`xl/charts/chart${n}.xml`, enc.encode(buildChartXml(d.chart))]);
          overrides.push({ partName: `/xl/charts/chart${n}.xml`, contentType: CT.chart });
          return { relId, target: `../charts/chart${n}.xml` };
        }
        if (!isPng(d.png)) throw new Error(`Image drawing "${d.name}" is not a PNG`);
        const n = ++imageCount;
        mediaParts.push([`xl/media/image${n}.png`, d.png]);
        return { relId, target: `../media/image${n}.png` };
      });
      const drawing = buildDrawingXml(sheet.drawings, targets);
      xlParts.push([`xl/drawings/drawing${drawingNo}.xml`, enc.encode(drawing.xml)]);
      xlParts.push([`xl/drawings/_rels/drawing${drawingNo}.xml.rels`, enc.encode(drawing.relsXml)]);
      overrides.push({ partName: `/xl/drawings/drawing${drawingNo}.xml`, contentType: CT.drawing });
      xlParts.push([
        `xl/worksheets/_rels/sheet${sheetNo}.xml.rels`,
        enc.encode(buildSheetRelsXml(`../drawings/drawing${drawingNo}.xml`, drawingRelId)),
      ]);
    }
    xlParts.push([
      `xl/worksheets/sheet${sheetNo}.xml`,
      enc.encode(buildWorksheetXml(sheet, styles, { tabSelected: i === activeTab, drawingRelId })),
    ]);
    overrides.push({ partName: `/xl/worksheets/sheet${sheetNo}.xml`, contentType: CT.worksheet });
  });

  overrides.push(
    { partName: '/xl/theme/theme1.xml', contentType: CT.theme },
    { partName: '/xl/styles.xml', contentType: CT.styles },
    { partName: '/docProps/core.xml', contentType: CT.core },
    { partName: '/docProps/app.xml', contentType: CT.app },
  );

  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', enc.encode(buildContentTypesXml(overrides)));
  parts.set('_rels/.rels', enc.encode(buildRootRelsXml()));
  parts.set('docProps/app.xml', enc.encode(buildAppXml()));
  parts.set('docProps/core.xml', enc.encode(buildCoreXml(wb.properties ?? {})));
  parts.set(
    'xl/workbook.xml',
    enc.encode(
      buildWorkbookXml(
        wb.sheets.map((s, i) => ({ name: s.name, hidden: s.hidden, relId: sheetRelIds[i]! })),
        activeTab,
      ),
    ),
  );
  parts.set('xl/_rels/workbook.xml.rels', enc.encode(workbookRels));
  parts.set('xl/styles.xml', enc.encode(styles.xml));
  parts.set('xl/theme/theme1.xml', enc.encode(buildThemeXml()));
  for (const [p, b] of xlParts.sort((a, b) => a[0].localeCompare(b[0]))) parts.set(p, b);
  for (const [p, b] of chartParts) parts.set(p, b);
  for (const [p, b] of mediaParts) parts.set(p, b);
  return parts;
}

export class OoxmlExcelWriter implements ExcelWriter {
  readonly name = 'ooxml';

  async write(wb: WorkbookSpec): Promise<Uint8Array> {
    const parts = buildPackageParts(wb);
    const files: Record<string, [Uint8Array, ZipOptions]> = {};
    for (const [path, bytes] of parts) {
      files[path] = [bytes, { level: path.endsWith('.png') ? 0 : 6, mtime: ZIP_MTIME }];
    }
    return zipSync(files, { level: 6, mtime: ZIP_MTIME });
  }
}

export function createDefaultExcelWriter(): ExcelWriter {
  return new OoxmlExcelWriter();
}
