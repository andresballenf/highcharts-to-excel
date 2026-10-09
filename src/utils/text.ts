/**
 * Text safety helpers for cell values and XML.
 */

const FORMULA_TRIGGERS = new Set(['=', '+', '-', '@', '\t', '\r']);

/**
 * True when Excel (or a CSV consumer) could interpret the string as a formula: it starts with
 * `= + - @`, a tab or CR, possibly after leading whitespace. Numbers-as-text like "-5" count too.
 */
export function isFormulaLike(s: string): boolean {
  if (typeof s !== 'string' || s.length === 0) return false;
  if (FORMULA_TRIGGERS.has(s[0] ?? '')) return true;
  return /^\s+[=+\-@]/.test(s);
}

// Anything outside the XML 1.0 Char production: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF].
// With the `u` flag lone surrogates are matched as single code points and removed.
const XML_INVALID = /[^\t\n\r -퟿-�\u{10000}-\u{10FFFF}]/gu;

/** Removes characters not allowed in XML 1.0 (tab, LF and CR are kept). */
export function stripControlChars(s: string): string {
  return s.replace(XML_INVALID, '');
}

/** Cuts `s` to at most `max` UTF-16 units without splitting a surrogate pair. No ellipsis is added. */
export function truncate(s: string, max: number): string {
  if (max <= 0) return '';
  if (s.length <= max) return s;
  let end = max;
  const code = s.charCodeAt(end - 1);
  if (code >= 0xd800 && code <= 0xdbff) end--;
  return s.slice(0, end);
}
