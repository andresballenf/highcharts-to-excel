/**
 * Structural checks for Excel number format codes (ECMA-376 Part 1, 18.8.31 numFmts).
 * Excel refuses to open (or "repairs") a workbook whose format codes break these rules.
 */

/** Excel's limit for a number format code. */
export const EXCEL_MAX_FORMAT_CODE_LENGTH = 255;

/**
 * Returns why `code` is not a usable Excel number format code, or null when it is structurally valid:
 * at most 255 characters, balanced `"…"` literals and `[…]` sections, no dangling `\`, and at most
 * four `;`-separated sections.
 */
export function excelFormatCodeProblem(code: string): string | null {
  if (typeof code !== 'string' || code.length === 0) return 'empty';
  if (code.length > EXCEL_MAX_FORMAT_CODE_LENGTH) return `longer than ${EXCEL_MAX_FORMAT_CODE_LENGTH} characters`;
  let inQuote = false;
  let inBracket = false;
  let sections = 1;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (inQuote) {
      if (ch === '"') inQuote = false;
      continue;
    }
    if (inBracket) {
      if (ch === ']') inBracket = false;
      else if (ch === '[') return 'nested "["';
      continue;
    }
    if (ch === '\\') {
      if (i === code.length - 1) return 'dangling "\\"';
      i++;
    } else if (ch === '"') {
      inQuote = true;
    } else if (ch === '[') {
      inBracket = true;
    } else if (ch === ']') {
      return 'unbalanced "]"';
    } else if (ch === ';') {
      sections++;
    }
  }
  if (inQuote) return 'unbalanced quote';
  if (inBracket) return 'unbalanced "["';
  if (sections > 4) return 'more than 4 sections';
  return null;
}

export function isValidExcelFormatCode(code: string): boolean {
  return excelFormatCodeProblem(code) === null;
}

/** Throws (writer side) when a format code is invalid. */
export function assertExcelFormatCode(code: string, where: string): string {
  const problem = excelFormatCodeProblem(code);
  if (problem !== null)
    throw new Error(
      `Invalid number format code ${JSON.stringify(code.length > 60 ? `${code.slice(0, 60)}…` : code)} in ${where}: ${problem}`,
    );
  return code;
}
