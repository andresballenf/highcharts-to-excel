/**
 * File, sheet and cell reference helpers.
 */

const MAX_FILENAME_LENGTH = 120;
const MAX_SHEET_NAME_LENGTH = 31;
const XLSX_SUFFIX = '.xlsx';

// biome-ignore lint/suspicious/noControlCharactersInRegex: filenames must not contain control characters
const CONTROL_CHARS = /[\u0000-\u001f\u007f-\u009f]/g;

/**
 * Produces a safe download filename ending in `.xlsx`. Strips control characters and
 * `\ / : * ? " < > |`, collapses whitespace, trims dots/spaces, limits length to 120 chars.
 */
export function sanitizeFilename(name: string | undefined | null, fallback = 'chart'): string {
  const clean = (raw: string): { base: string; suffix: string } => {
    // Whitespace controls (tab, CR, LF) become spaces; other controls are removed.
    let s = raw
      .replace(/\s+/g, ' ')
      .replace(CONTROL_CHARS, '')
      .replace(/[\\/:*?"<>|]/g, '');
    s = s.replace(/\s+/g, ' ').trim();
    let suffix = XLSX_SUFFIX;
    if (s.toLowerCase().endsWith(XLSX_SUFFIX)) {
      suffix = s.slice(-XLSX_SUFFIX.length);
      s = s.slice(0, -XLSX_SUFFIX.length);
    }
    s = trimDotsAndSpaces(s);
    const maxBase = MAX_FILENAME_LENGTH - suffix.length;
    if (s.length > maxBase) s = trimDotsAndSpaces(sliceSafe(s, maxBase));
    return { base: s, suffix };
  };
  const primary = clean(typeof name === 'string' ? name : '');
  if (primary.base) return primary.base + primary.suffix;
  const fb = clean(fallback);
  return (fb.base || 'chart') + XLSX_SUFFIX;
}

function trimDotsAndSpaces(s: string): string {
  return s.replace(/^[.\s]+|[.\s]+$/g, '');
}

/** Slices to `max` UTF-16 units without splitting a surrogate pair. */
function sliceSafe(s: string, max: number): string {
  if (s.length <= max) return s;
  let end = Math.max(0, max);
  const code = s.charCodeAt(end - 1);
  if (end > 0 && code >= 0xd800 && code <= 0xdbff) end--;
  return s.slice(0, end);
}

/**
 * Applies Excel's worksheet naming rules: no `[ ] : * ? / \`, no leading/trailing apostrophe,
 * 1-31 chars, not `History`, unique (case-insensitive) against `taken` by suffixing ` (2)`, ` (3)`...
 */
export function sanitizeSheetName(
  name: string | undefined | null,
  taken: ReadonlySet<string> = new Set(),
  fallback = 'Sheet',
): { name: string; adjusted: boolean } {
  const provided = typeof name === 'string';
  const clean = (raw: string): string => {
    let s = raw
      .replace(CONTROL_CHARS, '')
      .replace(/[[\]:*?/\\]/g, '')
      .trim();
    s = stripApostrophes(s);
    if (s.length > MAX_SHEET_NAME_LENGTH) s = stripApostrophes(sliceSafe(s, MAX_SHEET_NAME_LENGTH));
    return s;
  };
  let base = provided ? clean(name) : '';
  if (!base) base = clean(fallback) || 'Sheet';

  const takenLower = new Set<string>();
  for (const t of taken) takenLower.add(t.toLowerCase());
  const isFree = (candidate: string): boolean =>
    candidate.toLowerCase() !== 'history' && !takenLower.has(candidate.toLowerCase());

  let result = base;
  if (!isFree(result)) {
    for (let n = 2; ; n++) {
      const suffix = ` (${n})`;
      const head = stripApostrophes(sliceSafe(base, MAX_SHEET_NAME_LENGTH - suffix.length)).trimEnd();
      const candidate = head + suffix;
      if (isFree(candidate)) {
        result = candidate;
        break;
      }
    }
  }
  const adjusted = provided ? result !== name : result !== fallback;
  return { name: result, adjusted };
}

function stripApostrophes(s: string): string {
  return s.replace(/^'+|'+$/g, '').trim();
}

/**
 * Quotes a sheet name for use in a formula/reference. Names made only of `[A-Za-z0-9_]`
 * that do not start with a digit (and do not look like a cell reference) stay bare.
 */
export function quoteSheetNameForFormula(name: string): string {
  const bare =
    /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) &&
    !/^[A-Za-z]{1,3}[0-9]+$/.test(name) &&
    !/^[Rr][0-9]*([Cc][0-9]*)?$/.test(name) &&
    !/^[Cc][0-9]*$/.test(name) &&
    !/^(true|false)$/i.test(name);
  return bare ? name : `'${name.replace(/'/g, "''")}'`;
}

/** 0 → A, 25 → Z, 26 → AA. */
export function columnIndexToLetters(index0: number): string {
  if (!Number.isInteger(index0) || index0 < 0) {
    throw new RangeError(`Invalid column index: ${index0}`);
  }
  let n = index0 + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Cell reference from zero-based indices: (1, 1) → `$B$2` (or `B2` when not absolute). */
export function cellRef(col0: number, row0: number, absolute = true): string {
  if (!Number.isInteger(row0) || row0 < 0) throw new RangeError(`Invalid row index: ${row0}`);
  const col = columnIndexToLetters(col0);
  const row = String(row0 + 1);
  return absolute ? `$${col}$${row}` : `${col}${row}`;
}

/** Single-column range reference: `'My Sheet'!$B$2:$B$5` (sheet quoted only when required). */
export function rangeRef(sheet: string, col0: number, row0From: number, row0To: number): string {
  return `${quoteSheetNameForFormula(sheet)}!${cellRef(col0, row0From)}:${cellRef(col0, row0To)}`;
}
