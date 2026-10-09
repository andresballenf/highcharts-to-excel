/**
 * Minimal XML building helpers for the OOXML writer. Browser-safe, no dependencies.
 */

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Characters that may not appear in an XML 1.0 document: C0 controls other than TAB/LF/CR,
 * U+FFFE/U+FFFF and unpaired surrogates.
 */
const ILLEGAL_XML_CHARS =
  // biome-ignore lint/suspicious/noControlCharactersInRegex: XML 1.0 forbids these code points; the regex must name them
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

export function stripIllegalXmlChars(text: string): string {
  return text.replace(ILLEGAL_XML_CHARS, '');
}

/**
 * Escape text content. Strips characters that are illegal in XML 1.0; TAB/LF are kept and CR is
 * encoded as `&#13;` so XML end-of-line normalisation does not turn it into LF.
 */
export function escapeXml(text: string): string {
  return stripIllegalXmlChars(String(text))
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\r/g, '&#13;');
}

/**
 * Escape text for an ST_Xstring context (cell `<t>`, chart `c:v`): Excel decodes `_xHHHH_` there as
 * the character U+HHHH, so a literal `_x` that starts such a sequence is written as `_x005F_x`.
 * Not for DrawingML `a:t` (plain xsd:string).
 */
export function escapeXstring(text: string): string {
  return escapeXml(String(text).replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1'));
}

/** Escape an attribute value (double-quoted). Whitespace control chars are encoded so they survive normalisation. */
export function escapeAttr(text: string): string {
  return stripIllegalXmlChars(String(text))
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(/\t/g, '&#9;')
    .replace(/\n/g, '&#10;')
    .replace(/\r/g, '&#13;');
}

export type AttrValue = string | number | boolean | null | undefined;
export type Attrs = Record<string, AttrValue> | Array<[string, AttrValue]>;
/** Children: pre-rendered XML strings; null/undefined/false/'' entries are skipped. */
export type Children = string | null | undefined | false | ReadonlyArray<string | null | undefined | false>;

function renderAttrValue(v: string | number | boolean): string {
  if (typeof v === 'boolean') return v ? '1' : '0';
  if (typeof v === 'number') return formatNumber(v);
  return escapeAttr(v);
}

export function renderAttrs(attrs: Attrs | undefined): string {
  if (!attrs) return '';
  const entries = Array.isArray(attrs) ? attrs : Object.entries(attrs);
  let out = '';
  for (const [k, v] of entries) {
    if (v === null || v === undefined) continue;
    out += ` ${k}="${renderAttrValue(v)}"`;
  }
  return out;
}

function renderChildren(children: Children): string {
  if (children === null || children === undefined || children === false) return '';
  if (typeof children === 'string') return children;
  let out = '';
  for (const c of children) if (c) out += c;
  return out;
}

/**
 * Build an element. Children are raw XML (already escaped); use `text()` / `escapeXml()` for text.
 * Elements with no children are self-closing.
 */
export function el(name: string, attrs?: Attrs, children?: Children): string {
  const inner = renderChildren(children);
  const a = renderAttrs(attrs);
  return inner === '' ? `<${name}${a}/>` : `<${name}${a}>${inner}</${name}>`;
}

/** Element whose content is escaped text; always rendered with open/close tags. */
export function textEl(name: string, text: string, attrs?: Attrs): string {
  return `<${name}${renderAttrs(attrs)}>${escapeXml(text)}</${name}>`;
}

/** `<name val="…"/>` — the ubiquitous OOXML single-value element. */
export function valEl(name: string, val: string | number | boolean): string {
  return `<${name} val="${renderAttrValue(val)}"/>`;
}

/**
 * Render a finite number for XML. `-0` becomes `0`; non-finite numbers throw (callers must
 * sanitise first).
 */
export function formatNumber(n: number): string {
  if (!Number.isFinite(n)) throw new Error(`Cannot write non-finite number ${String(n)} to XML`);
  if (Object.is(n, -0) || n === 0) return '0';
  return String(n);
}

/** Coerce to a finite number (fallback when not), with `-0` normalised. */
export function finite(n: number | null | undefined, fallback: number): number {
  if (typeof n !== 'number' || !Number.isFinite(n)) return fallback;
  return n === 0 ? 0 : n;
}

/** Clamp then round to an integer; non-finite → fallback. */
export function clampInt(n: number | null | undefined, min: number, max: number, fallback: number): number {
  const v = finite(n, fallback);
  const r = Math.round(Math.min(max, Math.max(min, v)));
  return r === 0 ? 0 : r;
}

export function clampNum(n: number | null | undefined, min: number, max: number, fallback: number): number {
  const v = finite(n, fallback);
  const r = Math.min(max, Math.max(min, v));
  return r === 0 ? 0 : r;
}

/** Convert a zero-based column index to letters (0 → A, 26 → AA). */
export function columnLetters(col0: number): string {
  let n = col0 + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function a1(col0: number, row0: number): string {
  return `${columnLetters(col0)}${row0 + 1}`;
}

/** Wrap a document body with the standard XML declaration. */
export function xmlDocument(body: string): string {
  return XML_DECLARATION + body;
}
