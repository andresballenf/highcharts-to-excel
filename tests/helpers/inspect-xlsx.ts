/**
 * Test helper: open an XLSX (zip) and inspect its parts without mocks.
 */
import { strFromU8, unzipSync } from 'fflate';
import { XMLParser, XMLValidator } from 'fast-xml-parser';

export interface SeriesFormulas {
  name?: string;
  cat?: string;
  val?: string;
  xVal?: string;
  yVal?: string;
  bubbleSize?: string;
}

export interface XlsxInspection {
  parts: string[];
  has(path: string): boolean;
  bytes(path: string): Uint8Array;
  text(path: string): string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  xml(path: string): any;
  chartPaths(): string[];
  chartXml(i: number): string;
  seriesFormulas(chartPath: string): SeriesFormulas[];
  plotGroupKinds(chartPath: string): string[];
  sheetNames(): Array<{ name: string; hidden: boolean }>;
  /** Resolve a sheet name to its part path (e.g. "xl/worksheets/sheet1.xml"). */
  sheetPath(name: string): string;
  cellValue(sheetPath: string, ref: string): string | number | null;
  /** Raw `<c …>…</c>` markup for a cell, or null. */
  cellXml(sheetPath: string, ref: string): string | null;
  hasImages(): boolean;
  contentTypes(): string[];
  assertWellFormed(): void;
}

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', preserveOrder: false });

export function decodeXmlEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Direct child tag names of the FIRST occurrence of `parentTag` (in document order). */
export function elementOrder(xmlText: string, parentTag: string): string[] {
  const open = new RegExp(`<${escapeRe(parentTag)}(?=[\\s/>])(?:\\s+[^\\s=/>]+\\s*=\\s*(?:"[^"]*"|'[^']*'))*\\s*(/?)>`, 'g');
  const m = open.exec(xmlText);
  if (!m) throw new Error(`elementOrder: <${parentTag}> not found`);
  if (m[1] === '/') return [];
  const token = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
  token.lastIndex = open.lastIndex;
  const out: string[] = [];
  let depth = 0;
  let t: RegExpExecArray | null;
  while ((t = token.exec(xmlText))) {
    const name = t[2];
    if (!name) continue; // comment / CDATA / PI
    const closing = t[1] === '/';
    const selfClosing = t[4] === '/';
    if (closing) {
      if (depth === 0) return out;
      depth--;
      continue;
    }
    if (depth === 0) out.push(name);
    if (!selfClosing) depth++;
  }
  throw new Error(`elementOrder: <${parentTag}> is not closed`);
}

function firstF(block: string | undefined): string | undefined {
  if (block === undefined) return undefined;
  const m = /<c:f>([\s\S]*?)<\/c:f>/.exec(block);
  return m ? decodeXmlEntities(m[1]!) : undefined;
}

function childBlock(serXml: string, tag: string): string | undefined {
  const m = new RegExp(`<${escapeRe(tag)}>([\\s\\S]*?)</${escapeRe(tag)}>`).exec(serXml);
  return m ? m[1] : undefined;
}

function relsTargets(text: string): Array<{ id: string; type: string; target: string }> {
  const out: Array<{ id: string; type: string; target: string }> = [];
  const re = /<Relationship\s([^>]*?)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const attrs = m[1]!;
    const get = (n: string): string => decodeXmlEntities(new RegExp(`\\b${n}="([^"]*)"`).exec(attrs)?.[1] ?? '');
    out.push({ id: get('Id'), type: get('Type'), target: get('Target') });
  }
  return out;
}

export async function inspectXlsx(bytes: Uint8Array): Promise<XlsxInspection> {
  const files = unzipSync(bytes);
  const parts = Object.keys(files);
  const textCache = new Map<string, string>();

  const get = (path: string): Uint8Array => {
    const f = files[path];
    if (!f) throw new Error(`Part not found: ${path} (have: ${parts.join(', ')})`);
    return f;
  };
  const text = (path: string): string => {
    let t = textCache.get(path);
    if (t === undefined) {
      t = strFromU8(get(path));
      textCache.set(path, t);
    }
    return t;
  };
  const chartPaths = (): string[] =>
    parts
      .filter((p) => /^xl\/charts\/chart\d+\.xml$/.test(p))
      .sort((a, b) => Number(/(\d+)\.xml$/.exec(a)![1]) - Number(/(\d+)\.xml$/.exec(b)![1]));

  const sheetNames = (): Array<{ name: string; hidden: boolean }> => {
    const wb = text('xl/workbook.xml');
    const out: Array<{ name: string; hidden: boolean }> = [];
    const re = /<sheet\s([^>]*?)\/?>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(wb))) {
      const attrs = m[1]!;
      const name = decodeXmlEntities(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '');
      const state = /\bstate="([^"]*)"/.exec(attrs)?.[1];
      out.push({ name, hidden: state === 'hidden' || state === 'veryHidden' });
    }
    return out;
  };

  const cellXml = (sheetPath: string, ref: string): string | null => {
    const t = text(sheetPath);
    const re = new RegExp(`<c\\s[^>]*\\br="${escapeRe(ref)}"[^>]*?(?:/>|>([\\s\\S]*?)</c>)`);
    const m = re.exec(t);
    return m ? m[0] : null;
  };

  const sharedStrings = (): string[] => {
    if (!files['xl/sharedStrings.xml']) return [];
    const out: string[] = [];
    const re = /<si>([\s\S]*?)<\/si>/g;
    let m: RegExpExecArray | null;
    const sst = text('xl/sharedStrings.xml');
    while ((m = re.exec(sst))) {
      const ts = [...m[1]!.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) => decodeXmlEntities(x[1]!));
      out.push(ts.join(''));
    }
    return out;
  };

  return {
    parts,
    has: (p) => p in files,
    bytes: get,
    text,
    xml: (path) => parser.parse(text(path)),
    chartPaths,
    chartXml: (i) => {
      const p = chartPaths()[i];
      if (!p) throw new Error(`No chart #${i}`);
      return text(p);
    },
    seriesFormulas(chartPath) {
      const t = text(chartPath);
      const out: SeriesFormulas[] = [];
      for (const m of t.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
        const ser = m[1]!;
        const entry: SeriesFormulas = {};
        const assign = (key: keyof SeriesFormulas, tag: string): void => {
          const f = firstF(childBlock(ser, tag));
          if (f !== undefined) entry[key] = f;
        };
        assign('name', 'c:tx');
        assign('cat', 'c:cat');
        assign('val', 'c:val');
        assign('xVal', 'c:xVal');
        assign('yVal', 'c:yVal');
        assign('bubbleSize', 'c:bubbleSize');
        out.push(entry);
      }
      return out;
    },
    plotGroupKinds(chartPath) {
      return elementOrder(text(chartPath), 'c:plotArea')
        .filter((n) => /Chart$/.test(n))
        .map((n) => n.replace(/^c:/, ''));
    },
    sheetNames,
    sheetPath(name) {
      const wb = text('xl/workbook.xml');
      const re = /<sheet\s([^>]*?)\/?>/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(wb))) {
        const attrs = m[1]!;
        if (decodeXmlEntities(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '') !== name) continue;
        const rid = /\br:id="([^"]*)"/.exec(attrs)?.[1];
        const rel = relsTargets(text('xl/_rels/workbook.xml.rels')).find((r) => r.id === rid);
        if (!rel) break;
        return rel.target.startsWith('/') ? rel.target.slice(1) : `xl/${rel.target}`;
      }
      throw new Error(`Sheet not found: ${name}`);
    },
    cellValue(sheetPath, ref) {
      const c = cellXml(sheetPath, ref);
      if (c === null) return null;
      const type = /^<c\s[^>]*\bt="([^"]*)"/.exec(c)?.[1];
      if (type === 'inlineStr') {
        const ts = [...c.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) => decodeXmlEntities(x[1]!));
        return ts.join('');
      }
      const v = /<v>([\s\S]*?)<\/v>/.exec(c)?.[1];
      if (v === undefined) return null;
      if (type === 's') return sharedStrings()[Number(v)] ?? null;
      if (type === 'str' || type === 'e') return decodeXmlEntities(v);
      return Number(v);
    },
    cellXml,
    hasImages: () => parts.some((p) => p.startsWith('xl/media/')),
    contentTypes() {
      const t = text('[Content_Types].xml');
      return [...t.matchAll(/<(?:Default|Override)\s[^>]*\bContentType="([^"]*)"/g)].map((m) => decodeXmlEntities(m[1]!));
    },
    assertWellFormed() {
      for (const p of parts) {
        if (!/\.(xml|rels)$/.test(p)) continue;
        const res = XMLValidator.validate(text(p));
        if (res !== true) {
          throw new Error(`Malformed XML in ${p}: ${res.err.msg} (line ${res.err.line}, col ${res.err.col})`);
        }
      }
    },
  };
}
