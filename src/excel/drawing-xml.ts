/**
 * xl/drawings/drawingN.xml (+ rels) builder: one oneCellAnchor per DrawingSpec.
 */
import type { AnchorSpec, DrawingSpec } from './writer-interface';
import { EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS } from './writer-interface';
import { escapeAttr, xmlDocument } from './xml';

const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

export const EMU_PER_PX = 9525;
/** Largest coordinate DrawingML accepts (ST_Coordinate). */
const MAX_COORD = 27_273_042_316_900;

export function pxToEmu(px: number): number {
  if (!Number.isFinite(px)) return 0;
  const v = Math.round(Math.max(0, px) * EMU_PER_PX);
  return Math.min(v, MAX_COORD);
}

export interface DrawingTarget {
  /** Relationship id inside the drawing rels, e.g. "rId1". */
  relId: string;
  /** Target relative to xl/drawings/, e.g. "../charts/chart1.xml". */
  target: string;
}

function anchorXml(anchor: AnchorSpec, content: string): string {
  if (!Number.isInteger(anchor.col0) || anchor.col0 < 0 || anchor.col0 >= EXCEL_MAX_COLUMNS) {
    throw new Error(`Drawing anchor column ${String(anchor.col0)} is invalid`);
  }
  if (!Number.isInteger(anchor.row0) || anchor.row0 < 0 || anchor.row0 >= EXCEL_MAX_ROWS) {
    throw new Error(`Drawing anchor row ${String(anchor.row0)} is invalid`);
  }
  return (
    '<xdr:oneCellAnchor>' +
    `<xdr:from><xdr:col>${anchor.col0}</xdr:col><xdr:colOff>${pxToEmu(anchor.colOffsetPx)}</xdr:colOff>` +
    `<xdr:row>${anchor.row0}</xdr:row><xdr:rowOff>${pxToEmu(anchor.rowOffsetPx)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${pxToEmu(anchor.widthPx)}" cy="${pxToEmu(anchor.heightPx)}"/>` +
    content +
    '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>'
  );
}

function chartFrameXml(id: number, name: string, relId: string): string {
  return (
    '<xdr:graphicFrame macro="">' +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${relId}"/></a:graphicData></a:graphic>` +
    '</xdr:graphicFrame>'
  );
}

function pictureXml(id: number, name: string, relId: string, anchor: AnchorSpec): string {
  return (
    '<xdr:pic>' +
    `<xdr:nvPicPr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(anchor.widthPx)}" cy="${pxToEmu(anchor.heightPx)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
    '</xdr:pic>'
  );
}

/**
 * Build a drawing part. `targets[i]` is the relationship for `drawings[i]` (chart part or image).
 */
export function buildDrawingXml(
  drawings: readonly DrawingSpec[],
  targets: readonly DrawingTarget[],
): { xml: string; relsXml: string } {
  if (targets.length !== drawings.length) throw new Error('buildDrawingXml: one target per drawing is required');
  const anchors: string[] = [];
  const rels: string[] = [];
  drawings.forEach((d, i) => {
    const t = targets[i]!;
    const id = i + 2;
    const name = d.name && d.name.trim() !== '' ? d.name : `${d.kind === 'chart' ? 'Chart' : 'Picture'} ${i + 1}`;
    if (d.kind === 'chart') {
      anchors.push(anchorXml(d.anchor, chartFrameXml(id, name, t.relId)));
      rels.push(`<Relationship Id="${t.relId}" Type="${REL_CHART}" Target="${escapeAttr(t.target)}"/>`);
    } else {
      anchors.push(anchorXml(d.anchor, pictureXml(id, name, t.relId, d.anchor)));
      rels.push(`<Relationship Id="${t.relId}" Type="${REL_IMAGE}" Target="${escapeAttr(t.target)}"/>`);
    }
  });
  const xml = xmlDocument(
    `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">${anchors.join('')}</xdr:wsDr>`,
  );
  const relsXml = xmlDocument(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
  );
  return { xml, relsXml };
}
