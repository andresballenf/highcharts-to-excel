/**
 * xl/worksheets/sheetN.xml builder.
 *
 * Strings are always written as inline strings (`t="inlineStr"`), never as formulas: a value such
 * as "=1+1" stays literal text. This is the writer's formula-injection protection.
 */
import { EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, type CellSpec, type SheetSpec } from './writer-interface';
import type { StylesResult } from './styles-xml';
import { a1, el, escapeXstring, formatNumber, xmlDocument } from './xml';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Excel's per-cell text limit. */
export const EXCEL_MAX_CELL_CHARS = 32_767;

export interface WorksheetXmlOptions {
  /** Marks this sheet as the selected tab. */
  tabSelected: boolean;
  /** Relationship id of the sheet's drawing part, when it has drawings. */
  drawingRelId: string | null;
}

function assertIndex(n: number, max: number, what: string, sheet: string): void {
  if (!Number.isInteger(n) || n < 0 || n >= max) {
    throw new Error(`Sheet "${sheet}": ${what} ${String(n)} is outside 0..${max - 1}`);
  }
}

function truncateCellText(s: string): string {
  if (s.length <= EXCEL_MAX_CELL_CHARS) return s;
  let cut = s.slice(0, EXCEL_MAX_CELL_CHARS);
  // Don't leave half a surrogate pair.
  const last = cut.charCodeAt(cut.length - 1);
  if (last >= 0xd800 && last <= 0xdbff) cut = cut.slice(0, -1);
  return cut;
}

function cellXml(cell: CellSpec, styles: StylesResult): string | null {
  const ref = a1(cell.col0, cell.row0);
  const s = styles.indexOf(cell.style);
  const sAttr = s !== 0 ? ` s="${s}"` : '';
  const v = cell.value;
  switch (v.type) {
    case 'number':
      if (!Number.isFinite(v.value)) return s !== 0 ? `<c r="${ref}"${sAttr}/>` : null;
      return `<c r="${ref}"${sAttr}><v>${formatNumber(v.value)}</v></c>`;
    case 'string':
      return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXstring(truncateCellText(v.value))}</t></is></c>`;
    case 'boolean':
      return `<c r="${ref}"${sAttr} t="b"><v>${v.value ? 1 : 0}</v></c>`;
    case 'blank':
      return s !== 0 ? `<c r="${ref}"${sAttr}/>` : null;
    case 'error':
      // Only #N/A is modelled: line/area charts skip it and connect the neighbouring points.
      if ((v.value as string) !== '#N/A') throw new Error(`Unsupported cell error value ${JSON.stringify(v.value)}`);
      return `<c r="${ref}"${sAttr} t="e"><v>#N/A</v></c>`;
    default: {
      const never: never = v;
      throw new Error(`Unknown cell value type ${JSON.stringify(never)}`);
    }
  }
}

export function buildWorksheetXml(sheet: SheetSpec, styles: StylesResult, opts: WorksheetXmlOptions): string {
  // Group by the cell's own coordinates (RowSpec.row0 is advisory); reject duplicates.
  const rows = new Map<number, Map<number, CellSpec>>();
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      assertIndex(cell.row0, EXCEL_MAX_ROWS, 'row', sheet.name);
      assertIndex(cell.col0, EXCEL_MAX_COLUMNS, 'column', sheet.name);
      let r = rows.get(cell.row0);
      if (!r) {
        r = new Map();
        rows.set(cell.row0, r);
      }
      if (r.has(cell.col0)) throw new Error(`Sheet "${sheet.name}": duplicate cell ${a1(cell.col0, cell.row0)}`);
      r.set(cell.col0, cell);
    }
  }

  let minRow = Infinity;
  let maxRow = -Infinity;
  let minCol = Infinity;
  let maxCol = -Infinity;
  const rowXml: string[] = [];
  for (const r0 of [...rows.keys()].sort((a, b) => a - b)) {
    const cells = rows.get(r0)!;
    const out: string[] = [];
    for (const c0 of [...cells.keys()].sort((a, b) => a - b)) {
      const x = cellXml(cells.get(c0)!, styles);
      if (x === null) continue;
      out.push(x);
      minCol = Math.min(minCol, c0);
      maxCol = Math.max(maxCol, c0);
    }
    if (out.length === 0) continue;
    minRow = Math.min(minRow, r0);
    maxRow = Math.max(maxRow, r0);
    rowXml.push(`<row r="${r0 + 1}">${out.join('')}</row>`);
  }

  const dimension =
    rowXml.length === 0
      ? 'A1'
      : minRow === maxRow && minCol === maxCol
        ? a1(minCol, minRow)
        : `${a1(minCol, minRow)}:${a1(maxCol, maxRow)}`;

  const sheetView = sheet.freezeHeaderRow
    ? el('sheetView', { tabSelected: opts.tabSelected ? 1 : null, workbookViewId: 0 }, [
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>',
      ])
    : el('sheetView', { tabSelected: opts.tabSelected ? 1 : null, workbookViewId: 0 });

  const colWidths = new Map<number, number>();
  for (const c of sheet.columns) {
    assertIndex(c.col0, EXCEL_MAX_COLUMNS, 'column', sheet.name);
    if (!Number.isFinite(c.widthChars)) continue;
    colWidths.set(c.col0, Math.min(255, Math.max(0, c.widthChars)));
  }
  const cols = [...colWidths.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(
      ([c0, w]) =>
        `<col min="${c0 + 1}" max="${c0 + 1}" width="${formatNumber(Math.round(w * 100) / 100)}" customWidth="1"/>`,
    );

  const body = [
    `<dimension ref="${dimension}"/>`,
    `<sheetViews>${sheetView}</sheetViews>`,
    '<sheetFormatPr defaultRowHeight="15"/>',
    cols.length > 0 ? `<cols>${cols.join('')}</cols>` : '',
    rowXml.length > 0 ? `<sheetData>${rowXml.join('')}</sheetData>` : '<sheetData/>',
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    opts.drawingRelId ? `<drawing r:id="${opts.drawingRelId}"/>` : '',
  ].join('');

  return xmlDocument(`<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${body}</worksheet>`);
}

export function buildSheetRelsXml(drawingTarget: string, relId: string): string {
  return xmlDocument(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${drawingTarget}"/>` +
      '</Relationships>',
  );
}
