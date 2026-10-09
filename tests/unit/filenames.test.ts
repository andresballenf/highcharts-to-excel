import { describe, expect, it } from 'vitest';
import {
  cellRef,
  columnIndexToLetters,
  quoteSheetNameForFormula,
  rangeRef,
  sanitizeFilename,
  sanitizeSheetName,
} from '../../src/utils/filenames';

describe('sanitizeFilename', () => {
  it('cleans names and appends .xlsx', () => {
    expect(sanitizeFilename('Sales: Q1/Q2 <draft>?')).toBe('Sales Q1Q2 draft.xlsx');
    expect(sanitizeFilename('  report   2024 ')).toBe('report 2024.xlsx');
    expect(sanitizeFilename('report.XLSX')).toBe('report.XLSX');
    expect(sanitizeFilename('..hidden..')).toBe('hidden.xlsx');
    expect(sanitizeFilename('a\u0000b\nc')).toBe('ab c.xlsx');
  });
  it('falls back when empty', () => {
    expect(sanitizeFilename(undefined)).toBe('chart.xlsx');
    expect(sanitizeFilename(null)).toBe('chart.xlsx');
    expect(sanitizeFilename('???')).toBe('chart.xlsx');
    expect(sanitizeFilename('', 'my export')).toBe('my export.xlsx');
  });
  it('limits length to 120 chars', () => {
    const out = sanitizeFilename('x'.repeat(300));
    expect(out).toHaveLength(120);
    expect(out.endsWith('.xlsx')).toBe(true);
  });
});

describe('sanitizeSheetName', () => {
  it('removes invalid characters', () => {
    expect(sanitizeSheetName('My: Data [2024]')).toEqual({ name: 'My Data 2024', adjusted: true });
    expect(sanitizeSheetName('a/b\\c*d?e')).toEqual({ name: 'abcde', adjusted: true });
    expect(sanitizeSheetName('Data')).toEqual({ name: 'Data', adjusted: false });
  });
  it('strips apostrophes at the edges and limits to 31 chars', () => {
    expect(sanitizeSheetName("'Quoted'")).toEqual({ name: 'Quoted', adjusted: true });
    expect(sanitizeSheetName("It's")).toEqual({ name: "It's", adjusted: false });
    const long = sanitizeSheetName('A'.repeat(40));
    expect(long.name).toBe('A'.repeat(31));
    expect(long.adjusted).toBe(true);
  });
  it('uses fallback for empty names and avoids History', () => {
    expect(sanitizeSheetName('')).toEqual({ name: 'Sheet', adjusted: true });
    expect(sanitizeSheetName(undefined, new Set(), 'Chart')).toEqual({ name: 'Chart', adjusted: false });
    expect(sanitizeSheetName('history')).toEqual({ name: 'history (2)', adjusted: true });
  });
  it('makes names unique case-insensitively', () => {
    expect(sanitizeSheetName('data', new Set(['Data']))).toEqual({ name: 'data (2)', adjusted: true });
    expect(sanitizeSheetName('Data', new Set(['Data', 'DATA (2)']))).toEqual({ name: 'Data (3)', adjusted: true });
    const r = sanitizeSheetName('B'.repeat(31), new Set(['B'.repeat(31)]));
    expect(r.name).toBe(`${'B'.repeat(27)} (2)`);
    expect(r.name.length).toBe(31);
  });
});

describe('formula references', () => {
  it('quotes sheet names when needed', () => {
    expect(quoteSheetNameForFormula("My Data's")).toBe("'My Data''s'");
    expect(quoteSheetNameForFormula('Data')).toBe('Data');
    expect(quoteSheetNameForFormula('Data_2')).toBe('Data_2');
    expect(quoteSheetNameForFormula('2024')).toBe("'2024'");
    expect(quoteSheetNameForFormula('Sales-2024')).toBe("'Sales-2024'");
    expect(quoteSheetNameForFormula('A1')).toBe("'A1'");
    expect(quoteSheetNameForFormula('R1C1')).toBe("'R1C1'");
  });
  it('converts column indices', () => {
    expect(columnIndexToLetters(0)).toBe('A');
    expect(columnIndexToLetters(25)).toBe('Z');
    expect(columnIndexToLetters(26)).toBe('AA');
    expect(columnIndexToLetters(701)).toBe('ZZ');
    expect(columnIndexToLetters(702)).toBe('AAA');
    expect(columnIndexToLetters(16383)).toBe('XFD');
    expect(() => columnIndexToLetters(-1)).toThrow(RangeError);
  });
  it('builds cell and range refs', () => {
    expect(cellRef(1, 1)).toBe('$B$2');
    expect(cellRef(0, 0, false)).toBe('A1');
    expect(rangeRef('Data', 1, 1, 4)).toBe('Data!$B$2:$B$5');
    expect(rangeRef('My Data', 0, 1, 10)).toBe("'My Data'!$A$2:$A$11");
  });
});
