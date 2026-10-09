/**
 * DrawingML primitives shared by the chart writer: fills, lines, shape properties and text
 * properties. Split out of chart-xml.ts to keep that file focused on chart structure.
 */
import type { ExcelFillSpec, ExcelFontSpec, ExcelLineSpec, ExcelTextSpec, OoxmlDash } from './writer-interface';
import { clampInt, escapeAttr, escapeXml, finite } from './xml';

export const EMU_PER_PX = 9525;
/** ST_LineWidth upper bound. */
const MAX_LINE_WIDTH_EMU = 20_116_800;

const VALID_DASHES: ReadonlySet<OoxmlDash> = new Set<OoxmlDash>([
  'solid', 'dash', 'dot', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot',
  'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot',
]);

/** Validate and normalise an RRGGBB colour (a leading '#' is tolerated). */
export function hexColor(hex: string, context = 'color'): string {
  const h = String(hex).replace(/^#/, '');
  if (!/^[0-9a-fA-F]{6}$/.test(h)) throw new Error(`Invalid ${context} "${hex}" (expected RRGGBB)`);
  return h.toUpperCase();
}

/** Alpha 0..1 → a:alpha val in 1000ths of a percent; null when fully opaque. */
export function alphaVal(alpha: number | null | undefined): number | null {
  const a = finite(alpha, 1);
  if (a >= 1) return null;
  return clampInt(a * 100_000, 0, 100_000, 100_000);
}

export function pxToLineEmu(px: number): number {
  return clampInt(finite(px, 0) * EMU_PER_PX, 0, MAX_LINE_WIDTH_EMU, 0);
}

/** Degrees → 60000ths of a degree. */
export function degToAngle(deg: number): number {
  return Math.round(finite(deg, 0) * 60_000);
}

export function srgbClr(hex: string, alpha: number | null | undefined): string {
  const a = alphaVal(alpha);
  const h = hexColor(hex);
  return a === null ? `<a:srgbClr val="${h}"/>` : `<a:srgbClr val="${h}"><a:alpha val="${a}"/></a:srgbClr>`;
}

export function solidFill(hex: string, alpha: number | null | undefined): string {
  return `<a:solidFill>${srgbClr(hex, alpha)}</a:solidFill>`;
}

export function fillXml(fill: ExcelFillSpec | null | undefined): string {
  if (!fill) return '';
  switch (fill.type) {
    case 'none':
      return '<a:noFill/>';
    case 'solid':
      return solidFill(fill.hex, fill.alpha);
    case 'gradient': {
      const stops = fill.stops.filter((s) => s && typeof s.hex === 'string');
      if (stops.length === 0) return '';
      if (stops.length === 1) return solidFill(stops[0]!.hex, stops[0]!.alpha);
      const sorted = [...stops].sort((a, b) => finite(a.pos, 0) - finite(b.pos, 0));
      const gs = sorted
        .map((s) => `<a:gs pos="${clampInt(finite(s.pos, 0) * 100_000, 0, 100_000, 0)}">${srgbClr(s.hex, s.alpha)}</a:gs>`)
        .join('');
      // ST_PositiveFixedAngle: [0, 21600000).
      const ang = ((degToAngle(fill.angle) % 21_600_000) + 21_600_000) % 21_600_000;
      return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst><a:lin ang="${ang}" scaled="1"/></a:gradFill>`;
    }
    default:
      return '';
  }
}

export function lineXml(line: ExcelLineSpec | null | undefined): string {
  if (!line) return '';
  const w = pxToLineEmu(line.widthPx);
  const children: string[] = [];
  if (line.noFill) children.push('<a:noFill/>');
  else if (line.hex) children.push(solidFill(line.hex, line.alpha));
  if (!line.noFill && line.dash && line.dash !== 'solid' && VALID_DASHES.has(line.dash)) {
    children.push(`<a:prstDash val="${line.dash}"/>`);
  }
  children.push('<a:round/>');
  return `<a:ln w="${w}">${children.join('')}</a:ln>`;
}

/** `<c:spPr>` with fill then a:ln; empty string when neither is set. */
export function spPrXml(fill: ExcelFillSpec | null | undefined, line: ExcelLineSpec | null | undefined): string {
  const inner = fillXml(fill) + lineXml(line);
  return inner === '' ? '' : `<c:spPr>${inner}</c:spPr>`;
}

/** Attributes + children for a:defRPr / a:rPr from a font. */
function runPropsXml(tag: 'a:defRPr' | 'a:rPr', font: ExcelFontSpec | null | undefined, extraAttrs = ''): string {
  if (!font) return `<${tag}${extraAttrs}/>`;
  let attrs = extraAttrs;
  if (font.sizeHundredthsPt !== null && Number.isFinite(font.sizeHundredthsPt)) {
    attrs += ` sz="${clampInt(font.sizeHundredthsPt, 100, 400_000, 1000)}"`;
  }
  attrs += ` b="${font.bold ? 1 : 0}" i="${font.italic ? 1 : 0}"`;
  const children =
    (font.colorHex ? solidFill(font.colorHex, 1) : '') +
    (font.typeface ? `<a:latin typeface="${escapeAttr(font.typeface)}"/>` : '');
  return children === '' ? `<${tag}${attrs}/>` : `<${tag}${attrs}>${children}</${tag}>`;
}

export function bodyPrXml(rotationDeg: number | null | undefined): string {
  if (rotationDeg === null || rotationDeg === undefined || !Number.isFinite(rotationDeg)) return '<a:bodyPr/>';
  const rot = clampInt(degToAngle(rotationDeg), -5_400_000, 5_400_000, 0);
  return `<a:bodyPr rot="${rot}" vert="horz"/>`;
}

/** `<c:txPr>`: default text properties. Empty string when there is nothing to say. */
export function txPrXml(font: ExcelFontSpec | null | undefined, rotationDeg: number | null = null): string {
  const hasRot = rotationDeg !== null && Number.isFinite(rotationDeg);
  if (!font && !hasRot) return '';
  return (
    '<c:txPr>' +
    bodyPrXml(rotationDeg) +
    '<a:lstStyle/>' +
    `<a:p><a:pPr>${runPropsXml('a:defRPr', font)}</a:pPr><a:endParaRPr lang="en-US"/></a:p>` +
    '</c:txPr>'
  );
}

/** `<c:tx><c:rich>…` for titles: one a:p per line. */
export function richTextXml(text: ExcelTextSpec): string {
  const paras = text.lines.map(
    (line) =>
      '<a:p>' +
      `<a:pPr>${runPropsXml('a:defRPr', text.font)}</a:pPr>` +
      `<a:r>${runPropsXml('a:rPr', text.font, ' lang="en-US"')}<a:t>${escapeXml(line)}</a:t></a:r>` +
      '</a:p>',
  );
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/>${paras.join('')}</c:rich></c:tx>`;
}

/** `<c:title>`; empty string when the text has no lines. */
export function titleXml(text: ExcelTextSpec | null | undefined): string {
  if (!text || text.lines.length === 0) return '';
  return `<c:title>${richTextXml(text)}<c:layout/><c:overlay val="${text.overlay ? 1 : 0}"/></c:title>`;
}
