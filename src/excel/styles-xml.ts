/**
 * xl/styles.xml builder. Collects distinct CellStyleSpecs and assigns cellXfs indices.
 */
import type { CellStyleSpec } from './writer-interface';
import { assertExcelFormatCode } from '../utils/format-code';
import { el, escapeAttr, xmlDocument } from './xml';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Built-in number formats that must be referenced by id rather than redeclared. */
const BUILTIN_NUMFMTS: Readonly<Record<string, number>> = {
  General: 0,
  '0': 1,
  '0.00': 2,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
};

const FIRST_CUSTOM_NUMFMT_ID = 164;

export interface StylesResult {
  xml: string;
  /** cellXfs index for a style (0 = default when undefined / empty). */
  indexOf(style: CellStyleSpec | undefined): number;
}

interface NormalizedStyle {
  bold: boolean;
  italic: boolean;
  numberFormat: string | null;
  fillHex: string | null;
  fontColorHex: string | null;
  align: 'left' | 'center' | 'right' | null;
}

function normalizeHex(hex: string | undefined, what: string): string | null {
  if (hex === undefined || hex === null || hex === '') return null;
  const h = hex.replace(/^#/, '');
  if (!/^[0-9a-fA-F]{6}$/.test(h)) throw new Error(`Invalid ${what} color "${hex}" (expected RRGGBB)`);
  return h.toUpperCase();
}

const ALIGNMENTS: ReadonlySet<string> = new Set(['left', 'center', 'right']);

function normalize(style: CellStyleSpec): NormalizedStyle {
  const nf = style.numberFormat;
  const align = style.align;
  return {
    bold: style.bold === true,
    italic: style.italic === true,
    numberFormat: nf && nf !== 'General' ? assertExcelFormatCode(nf, 'a cell style') : null,
    fillHex: normalizeHex(style.fillHex, 'fill'),
    fontColorHex: normalizeHex(style.fontColorHex, 'font'),
    // Whitelisted: anything else is ignored rather than written into an attribute.
    align: typeof align === 'string' && ALIGNMENTS.has(align) ? (align as NormalizedStyle['align']) : null,
  };
}

function isDefault(s: NormalizedStyle): boolean {
  return (
    !s.bold && !s.italic && s.numberFormat === null && s.fillHex === null && s.fontColorHex === null && s.align === null
  );
}

function keyOf(s: NormalizedStyle): string {
  return JSON.stringify([s.bold, s.italic, s.numberFormat, s.fillHex, s.fontColorHex, s.align]);
}

function fontXml(bold: boolean, italic: boolean, colorHex: string | null): string {
  return el('font', undefined, [
    bold ? '<b/>' : '',
    italic ? '<i/>' : '',
    '<sz val="11"/>',
    colorHex ? `<color rgb="FF${colorHex}"/>` : '<color theme="1"/>',
    '<name val="Calibri"/>',
    '<family val="2"/>',
    '<scheme val="minor"/>',
  ]);
}

/**
 * Build styles.xml from every style used in the workbook. Pass all cell styles (duplicates and
 * undefined are fine); `indexOf` resolves any of them afterwards.
 */
export function buildStyles(styles: Iterable<CellStyleSpec | undefined>): StylesResult {
  const xfIndexByKey = new Map<string, number>();
  const xfs: Array<{ numFmtId: number; fontId: number; fillId: number; align: string | null }> = [];

  const numFmtIdByCode = new Map<string, number>();
  const customNumFmts: Array<{ id: number; code: string }> = [];
  const fontIdByKey = new Map<string, number>([[JSON.stringify([false, false, null]), 0]]);
  const fonts: string[] = [fontXml(false, false, null)];
  const fillIdByHex = new Map<string, number>();
  const fills: string[] = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ];

  const numFmtId = (code: string | null): number => {
    if (code === null) return 0;
    const builtin = BUILTIN_NUMFMTS[code];
    if (builtin !== undefined) return builtin;
    let id = numFmtIdByCode.get(code);
    if (id === undefined) {
      id = FIRST_CUSTOM_NUMFMT_ID + customNumFmts.length;
      numFmtIdByCode.set(code, id);
      customNumFmts.push({ id, code });
    }
    return id;
  };
  const fontId = (s: NormalizedStyle): number => {
    const k = JSON.stringify([s.bold, s.italic, s.fontColorHex]);
    let id = fontIdByKey.get(k);
    if (id === undefined) {
      id = fonts.length;
      fonts.push(fontXml(s.bold, s.italic, s.fontColorHex));
      fontIdByKey.set(k, id);
    }
    return id;
  };
  const fillId = (hex: string | null): number => {
    if (hex === null) return 0;
    let id = fillIdByHex.get(hex);
    if (id === undefined) {
      id = fills.length;
      fills.push(
        `<fill><patternFill patternType="solid"><fgColor rgb="FF${hex}"/><bgColor indexed="64"/></patternFill></fill>`,
      );
      fillIdByHex.set(hex, id);
    }
    return id;
  };

  // Index per style object, recorded at registration so indexOf does not normalise again.
  const xfIndexByObject = new WeakMap<CellStyleSpec, number>();
  const register = (style: CellStyleSpec | undefined): number => {
    if (!style) return 0;
    const known = xfIndexByObject.get(style);
    if (known !== undefined) return known;
    const idx = registerNormalized(normalize(style));
    xfIndexByObject.set(style, idx);
    return idx;
  };
  const registerNormalized = (n: NormalizedStyle): number => {
    if (isDefault(n)) return 0;
    const k = keyOf(n);
    const existing = xfIndexByKey.get(k);
    if (existing !== undefined) return existing;
    const idx = xfs.length + 1;
    xfs.push({ numFmtId: numFmtId(n.numberFormat), fontId: fontId(n), fillId: fillId(n.fillHex), align: n.align });
    xfIndexByKey.set(k, idx);
    return idx;
  };

  for (const s of styles) register(s);

  const cellXfs = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    ...xfs.map((x) =>
      el(
        'xf',
        [
          ['numFmtId', x.numFmtId],
          ['fontId', x.fontId],
          ['fillId', x.fillId],
          ['borderId', 0],
          ['xfId', 0],
          ['applyNumberFormat', x.numFmtId !== 0 ? 1 : null],
          ['applyFont', x.fontId !== 0 ? 1 : null],
          ['applyFill', x.fillId !== 0 ? 1 : null],
          ['applyAlignment', x.align ? 1 : null],
        ],
        x.align ? [`<alignment horizontal="${x.align}"/>`] : undefined,
      ),
    ),
  ];

  const parts: string[] = [];
  if (customNumFmts.length > 0) {
    parts.push(
      `<numFmts count="${customNumFmts.length}">` +
        customNumFmts.map((f) => `<numFmt numFmtId="${f.id}" formatCode="${escapeAttr(f.code)}"/>`).join('') +
        '</numFmts>',
    );
  }
  parts.push(`<fonts count="${fonts.length}">${fonts.join('')}</fonts>`);
  parts.push(`<fills count="${fills.length}">${fills.join('')}</fills>`);
  parts.push('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
  parts.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
  parts.push(`<cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>`);
  parts.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
  parts.push('<dxfs count="0"/>');
  parts.push('<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>');

  const xml = xmlDocument(`<styleSheet xmlns="${NS_MAIN}">${parts.join('')}</styleSheet>`);

  return {
    xml,
    indexOf(style) {
      if (!style) return 0;
      const known = xfIndexByObject.get(style);
      if (known !== undefined) return known;
      const n = normalize(style);
      if (isDefault(n)) return 0;
      const idx = xfIndexByKey.get(keyOf(n));
      if (idx === undefined) throw new Error('Cell style was not registered with buildStyles()');
      return idx;
    },
  };
}
