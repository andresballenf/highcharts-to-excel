/**
 * Hand-written OOXML (SpreadsheetML + DrawingML) writer producing native, editable Excel charts.
 * Browser-safe: depends only on fflate for zipping.
 *
 * Large workbooks are serialized in chunks (worksheet cells, chart cache points). Between chunks
 * `write` reports progress, checks the abort signal and, in a browser, yields to the event loop so
 * the page can repaint. In browsers the package is compressed by fflate's worker-based `zip`; in
 * Node and jsdom by `zipSync`. Every path produces the same bytes.
 */
import { zip, zipSync, type AsyncZippable, type ZipOptions, type Zippable } from 'fflate';
import type {
  CellStyleSpec,
  ExcelChartSpec,
  ExcelSeriesRef,
  ExcelWriter,
  SheetSpec,
  WorkbookSpec,
  WriteContext,
  WriteProgress,
} from './writer-interface';
import { buildChartXml, cachePointsXml, chartCacheRefs, type CacheKind } from './chart-xml';
import { buildDrawingXml, type DrawingTarget } from './drawing-xml';
import {
  buildAppXml,
  buildContentTypesXml,
  buildCoreXml,
  buildRootRelsXml,
  buildThemeXml,
  buildWorkbookRelsXml,
  buildWorkbookXml,
  CT,
  type ContentTypeOverride,
} from './package-parts';
import { buildStyles } from './styles-xml';
import { buildSheetRelsXml, runWork, worksheetXmlWork, type WorkStep } from './worksheet-xml';

/** Fixed zip entry timestamp so identical input yields identical bytes. */
const ZIP_MTIME = new Date(2000, 0, 1, 0, 0, 0);
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const INVALID_SHEET_NAME_CHARS = /[\][:*?/\\]/;
/** Default `WriteContext.yieldEvery`: cells or cache points per chunk (a few ms of work each). */
export const DEFAULT_YIELD_EVERY = 10_000;
/** fflate's `zip` deflates smaller entries on the calling thread anyway; below this, workers buy nothing. */
export const ASYNC_ZIP_MIN_BYTES = 160_000;
/** How long `'auto'` waits for a probe worker to start before falling back to `zipSync`. */
const WORKER_PROBE_TIMEOUT_MS = 3_000;

/**
 * How `OoxmlExcelWriter` compresses the package.
 * - `'auto'` (default): fflate's worker-based async `zip` when some part is large enough for it to
 *   matter, `Worker`, `Blob` and `URL.createObjectURL` exist, and a Blob-URL worker actually starts
 *   (a strict Content-Security-Policy can forbid it); otherwise `zipSync`.
 * - `'sync'`: always `zipSync` on the calling thread.
 * - `'async'`: always the worker-based `zip`; rejects where `Worker` is undefined (Node, jsdom).
 */
export type ZipMode = 'sync' | 'async' | 'auto';

export interface OoxmlWriterOptions {
  zip?: ZipMode;
}

/** `WriteProgress.detail` of the zip phase, naming the path taken. */
export const ZIP_DETAIL = { sync: 'zipSync (main thread)', async: 'zip (Web Workers)' } as const;

function validateSheets(sheets: readonly SheetSpec[]): void {
  if (!Array.isArray(sheets) || sheets.length === 0) throw new Error('Workbook must contain at least one sheet');
  if (!sheets.some((s) => !s.hidden)) throw new Error('Workbook must contain at least one visible sheet');
  const seen = new Set<string>();
  for (const s of sheets) {
    const name = s.name;
    if (typeof name !== 'string' || name.length === 0 || name.length > 31) {
      throw new Error(`Invalid sheet name "${String(name)}": must be 1-31 characters`);
    }
    if (INVALID_SHEET_NAME_CHARS.test(name) || name.startsWith("'") || name.endsWith("'")) {
      throw new Error(`Invalid sheet name "${name}": contains characters Excel does not allow`);
    }
    const key = name.toLowerCase();
    if (seen.has(key)) throw new Error(`Duplicate sheet name "${name}"`);
    seen.add(key);
  }
}

function isPng(bytes: Uint8Array): boolean {
  return (
    bytes instanceof Uint8Array && bytes.length > PNG_SIGNATURE.length && PNG_SIGNATURE.every((b, i) => bytes[i] === b)
  );
}

/** Cache points per reference, as chunk strings built before the chart XML that embeds them. */
type CachePointsMemo = Map<ExcelSeriesRef<unknown>, Partial<Record<CacheKind, string[]>>>;

/** Characters encoded between two pauses when a large part is encoded piece by piece. */
const ENCODE_CHUNK_CHARS = 1_000_000;

/**
 * UTF-8 encodes `pieces` (each a whole run of XML elements, so no surrogate pair is split) into one
 * array, pausing between pieces about every ENCODE_CHUNK_CHARS characters. Same bytes as encoding
 * the joined string, without building it.
 */
function* encodeWork(
  enc: TextEncoder,
  pieces: readonly string[],
  chunkItems: number,
  detail: string,
): Generator<WorkStep, Uint8Array, void> {
  if (pieces.length === 1) return enc.encode(pieces[0]);
  const encoded: Uint8Array[] = [];
  let total = 0;
  let chars = 0;
  for (const piece of pieces) {
    if (Number.isFinite(chunkItems) && chars >= ENCODE_CHUNK_CHARS) {
      yield { items: 0, detail };
      chars = 0;
    }
    const bytes = enc.encode(piece);
    encoded.push(bytes);
    total += bytes.byteLength;
    chars += piece.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const bytes of encoded) {
    out.set(bytes, offset);
    offset += bytes.byteLength;
  }
  return out;
}

/** Marks where a reference's cache points go in the chart XML skeleton (NUL never survives escaping). */
const CACHE_MARK = '\u0000';

/**
 * The chart XML as pieces: the skeleton from `buildChartXml`, split where the pre-built cache
 * chunks go. Falls back to one string when the skeleton does not split cleanly.
 */
function chartXmlPieces(spec: ExcelChartSpec, memo: CachePointsMemo): string[] {
  const slots: string[][] = [];
  const skeleton = buildChartXml(spec, {
    cachePoints: (ref, kind) => {
      const chunks = memo.get(ref)?.[kind];
      if (chunks === undefined) return undefined;
      slots.push(chunks);
      return `${CACHE_MARK}${slots.length - 1}${CACHE_MARK}`;
    },
  });
  const split = skeleton.split(CACHE_MARK);
  // Markers can appear out of lookup order (values are looked up before categories); each must
  // name a distinct slot, and every slot must appear once.
  const used = new Set<number>();
  for (let k = 1; k < split.length; k += 2) {
    const slot = /^\d+$/.test(split[k]!) ? Number(split[k]) : -1;
    if (slot < 0 || slot >= slots.length || used.has(slot)) break;
    used.add(slot);
  }
  if (split.length !== 2 * slots.length + 1 || used.size !== slots.length) {
    return [buildChartXml(spec, { cachePoints: (ref, kind) => memo.get(ref)?.[kind]?.join('') })];
  }
  const pieces: string[] = [];
  for (const [k, piece] of split.entries()) {
    if (k % 2 === 0) pieces.push(piece);
    else for (const chunk of slots[Number(piece)]!) pieces.push(chunk);
  }
  return pieces;
}

function* chartCacheWork(
  refs: ReadonlyArray<{ ref: ExcelSeriesRef<unknown>; kind: CacheKind }>,
  memo: CachePointsMemo,
  chunkItems: number,
  detail: string,
): Generator<WorkStep, void, void> {
  let pending = 0;
  for (const { ref, kind } of refs) {
    if (memo.get(ref)?.[kind] !== undefined) continue;
    const parts: string[] = [];
    const n = ref.cache.length;
    for (let from = 0; from < n; ) {
      // Chunks span references, so many short series do not each cost a pause.
      const to = Math.min(n, from + (chunkItems - pending));
      parts.push(cachePointsXml(ref, kind, from, to));
      pending += to - from;
      from = to;
      if (pending >= chunkItems) {
        yield { items: pending, detail };
        pending = 0;
      }
    }
    memo.set(ref, { ...memo.get(ref), [kind]: parts });
  }
  if (pending > 0) yield { items: pending, detail };
}

/**
 * Chunked package builder: yields a `WorkStep` after about `chunkItems` cells or cache points and
 * returns every package part (path → bytes) in zip order, `[Content_Types].xml` first.
 * `buildPackageParts` runs it without pausing; the bytes are the same for any `chunkItems`.
 */
function* packageWork(wb: WorkbookSpec, chunkItems: number): Generator<WorkStep, Map<string, Uint8Array>, void> {
  validateSheets(wb.sheets);
  const enc = new TextEncoder();

  // First-seen order of distinct style objects: what buildStyles would register from every cell.
  const uniqueStyles = new Set<CellStyleSpec | undefined>();
  let pending = 0;
  for (const sheet of wb.sheets) {
    for (const row of sheet.rows) {
      if (pending >= chunkItems) {
        yield { items: pending, detail: 'cell styles' };
        pending = 0;
      }
      for (const c of row.cells) uniqueStyles.add(c.style);
      pending += row.cells.length;
    }
  }
  if (pending > 0) yield { items: pending, detail: 'cell styles' };
  const styles = buildStyles(uniqueStyles);

  const overrides: ContentTypeOverride[] = [{ partName: '/xl/workbook.xml', contentType: CT.workbook }];
  const xlParts: Array<[string, Uint8Array]> = [];
  const activeTab = wb.sheets.findIndex((s) => !s.hidden);
  const { xml: workbookRels, sheetRelIds } = buildWorkbookRelsXml(wb.sheets.length);

  let drawingCount = 0;
  let chartCount = 0;
  let imageCount = 0;
  const chartParts: Array<[string, Uint8Array]> = [];
  const mediaParts: Array<[string, Uint8Array]> = [];

  for (const [i, sheet] of wb.sheets.entries()) {
    const sheetNo = i + 1;
    let drawingRelId: string | null = null;
    if (sheet.drawings.length > 0) {
      const drawingNo = ++drawingCount;
      drawingRelId = 'rId1';
      const targets: DrawingTarget[] = [];
      for (const [j, d] of sheet.drawings.entries()) {
        const relId = `rId${j + 1}`;
        if (d.kind === 'chart') {
          const n = ++chartCount;
          const memo: CachePointsMemo = new Map();
          const detail = `chart ${n}`;
          yield* chartCacheWork(chartCacheRefs(d.chart), memo, chunkItems, detail);
          const bytes = yield* encodeWork(enc, chartXmlPieces(d.chart, memo), chunkItems, detail);
          chartParts.push([`xl/charts/chart${n}.xml`, bytes]);
          overrides.push({ partName: `/xl/charts/chart${n}.xml`, contentType: CT.chart });
          targets.push({ relId, target: `../charts/chart${n}.xml` });
          continue;
        }
        if (!isPng(d.png)) throw new Error(`Image drawing "${d.name}" is not a PNG`);
        const n = ++imageCount;
        mediaParts.push([`xl/media/image${n}.png`, d.png]);
        targets.push({ relId, target: `../media/image${n}.png` });
      }
      const drawing = buildDrawingXml(sheet.drawings, targets);
      xlParts.push([`xl/drawings/drawing${drawingNo}.xml`, enc.encode(drawing.xml)]);
      xlParts.push([`xl/drawings/_rels/drawing${drawingNo}.xml.rels`, enc.encode(drawing.relsXml)]);
      overrides.push({ partName: `/xl/drawings/drawing${drawingNo}.xml`, contentType: CT.drawing });
      xlParts.push([
        `xl/worksheets/_rels/sheet${sheetNo}.xml.rels`,
        enc.encode(buildSheetRelsXml(`../drawings/drawing${drawingNo}.xml`, drawingRelId)),
      ]);
    }
    const sheetOpts = { tabSelected: i === activeTab, drawingRelId };
    const sheetPieces = yield* worksheetXmlWork(sheet, styles, sheetOpts, chunkItems);
    const sheetBytes = yield* encodeWork(enc, sheetPieces, chunkItems, `sheet "${sheet.name}"`);
    xlParts.push([`xl/worksheets/sheet${sheetNo}.xml`, sheetBytes]);
    overrides.push({ partName: `/xl/worksheets/sheet${sheetNo}.xml`, contentType: CT.worksheet });
  }

  overrides.push(
    { partName: '/xl/theme/theme1.xml', contentType: CT.theme },
    { partName: '/xl/styles.xml', contentType: CT.styles },
    { partName: '/docProps/core.xml', contentType: CT.core },
    { partName: '/docProps/app.xml', contentType: CT.app },
  );

  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', enc.encode(buildContentTypesXml(overrides)));
  parts.set('_rels/.rels', enc.encode(buildRootRelsXml()));
  parts.set('docProps/app.xml', enc.encode(buildAppXml()));
  parts.set('docProps/core.xml', enc.encode(buildCoreXml(wb.properties ?? {})));
  parts.set(
    'xl/workbook.xml',
    enc.encode(
      buildWorkbookXml(
        wb.sheets.map((s, i) => ({ name: s.name, hidden: s.hidden, relId: sheetRelIds[i]! })),
        activeTab,
      ),
    ),
  );
  parts.set('xl/_rels/workbook.xml.rels', enc.encode(workbookRels));
  parts.set('xl/styles.xml', enc.encode(styles.xml));
  parts.set('xl/theme/theme1.xml', enc.encode(buildThemeXml()));
  for (const [p, b] of xlParts.sort((a, b) => a[0].localeCompare(b[0]))) parts.set(p, b);
  for (const [p, b] of chartParts) parts.set(p, b);
  for (const [p, b] of mediaParts) parts.set(p, b);
  return parts;
}

/** Build every package part (path → bytes) in zip order, `[Content_Types].xml` first. */
export function buildPackageParts(wb: WorkbookSpec): Map<string, Uint8Array> {
  return runWork(packageWork(wb, Number.POSITIVE_INFINITY));
}

/**
 * Work items `packageWork` reports for a workbook: every cell three times (styles, grouping,
 * serializing) plus every chart cache point. Tolerates malformed specs (0 for what it cannot read).
 */
function countWork(wb: WorkbookSpec): number {
  let cells = 0;
  let points = 0;
  for (const sheet of Array.isArray(wb?.sheets) ? wb.sheets : []) {
    for (const row of Array.isArray(sheet?.rows) ? sheet.rows : []) {
      if (Array.isArray(row?.cells)) cells += row.cells.length;
    }
    for (const d of Array.isArray(sheet?.drawings) ? sheet.drawings : []) {
      if (d?.kind === 'chart') for (const { ref } of chartCacheRefs(d.chart)) points += ref.cache.length;
    }
  }
  return 3 * cells + points;
}

function abortReason(signal: AbortSignal): unknown {
  if (signal.reason !== undefined) return signal.reason;
  const error = new Error('The operation was aborted.');
  error.name = 'AbortError';
  return error;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) throw abortReason(signal);
}

/**
 * A macrotask boundary, so the browser can paint between chunks: `scheduler.yield()` where it
 * exists, else a MessageChannel message (not clamped to 4 ms like nested timeouts), else setTimeout.
 */
function yieldToEventLoop(): Promise<void> {
  const scheduler = (globalThis as { scheduler?: { yield?: () => Promise<void> } }).scheduler;
  if (typeof scheduler?.yield === 'function') return scheduler.yield();
  if (typeof MessageChannel !== 'undefined') {
    return new Promise((resolve) => {
      const channel = new MessageChannel();
      channel.port1.onmessage = () => {
        channel.port1.close();
        resolve();
      };
      channel.port2.postMessage(null);
    });
  }
  return new Promise((resolve) => setTimeout(resolve, 0));
}

let workerProbe: Promise<boolean> | null = null;

/** Clears the cached Blob-worker probe result (tests). */
export function resetWorkerProbe(): void {
  workerProbe = null;
}

/**
 * Whether a Blob-URL worker starts here (fflate's async zip needs one; a strict CSP forbids it and
 * then reports the failure as an event fflate never listens to). Probed once per page.
 */
function blobWorkersUsable(): Promise<boolean> {
  workerProbe ??= new Promise<boolean>((resolve) => {
    let url: string | null = null;
    let worker: Worker | null = null;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const done = (ok: boolean): void => {
      if (timer !== undefined) clearTimeout(timer);
      try {
        worker?.terminate();
        if (url !== null && typeof URL.revokeObjectURL === 'function') URL.revokeObjectURL(url);
      } catch {
        // Cleanup only.
      }
      resolve(ok);
    };
    try {
      url = URL.createObjectURL(new Blob(['postMessage(1)'], { type: 'text/javascript' }));
      worker = new Worker(url);
      worker.onmessage = () => done(true);
      worker.onerror = () => done(false);
      timer = setTimeout(() => done(false), WORKER_PROBE_TIMEOUT_MS);
    } catch {
      done(false);
    }
  });
  return workerProbe;
}

/** The zip path `write` takes for a package whose largest part has `largestPartBytes` bytes. */
export async function resolveZipMode(mode: ZipMode, largestPartBytes: number): Promise<'sync' | 'async'> {
  if (mode === 'sync') return 'sync';
  const hasWorker = typeof Worker !== 'undefined';
  if (mode === 'async') {
    if (!hasWorker) {
      throw new Error(
        "OoxmlExcelWriter zip: 'async' needs Web Workers (Worker is undefined here); use 'auto' or 'sync'",
      );
    }
    return 'async';
  }
  if (largestPartBytes < ASYNC_ZIP_MIN_BYTES || !hasWorker) return 'sync';
  if (typeof Blob === 'undefined' || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
    return 'sync';
  }
  return (await blobWorkersUsable()) ? 'async' : 'sync';
}

/** fflate's worker-based zip; an abort terminates its workers and rejects with the signal's reason. */
function zipAsync(files: AsyncZippable, signal: AbortSignal | undefined): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    let settled = false;
    let terminate: (() => void) | undefined;
    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      terminate?.();
      reject(abortReason(signal!));
    };
    try {
      terminate = zip(files, { level: 6, mtime: ZIP_MTIME }, (error, data) => {
        if (settled) return;
        settled = true;
        signal?.removeEventListener('abort', onAbort);
        if (error) reject(error);
        else resolve(data);
      });
    } catch (error) {
      settled = true;
      reject(error);
      return;
    }
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

export class OoxmlExcelWriter implements ExcelWriter {
  readonly name = 'ooxml';
  readonly zipMode: ZipMode;

  constructor(options: OoxmlWriterOptions = {}) {
    const mode = options.zip ?? 'auto';
    if (mode !== 'sync' && mode !== 'async' && mode !== 'auto') {
      throw new Error(`OoxmlExcelWriter: zip must be 'sync', 'async' or 'auto', got ${JSON.stringify(mode)}`);
    }
    this.zipMode = mode;
  }

  /**
   * Serializes and zips the workbook. With a context: reports progress after every chunk of
   * `yieldEvery` cells or cache points, rejects with `signal.reason` once the signal is aborted
   * (checked between chunks and around the zip), and in a browser yields to the event loop
   * between chunks. The bytes do not depend on the context or the zip path.
   */
  async write(wb: WorkbookSpec, context: WriteContext = {}): Promise<Uint8Array> {
    const { signal, onProgress } = context;
    const chunkItems =
      typeof context.yieldEvery === 'number' && context.yieldEvery >= 1
        ? Math.floor(context.yieldEvery)
        : DEFAULT_YIELD_EVERY;
    const report = (p: WriteProgress): void => onProgress?.(p);
    throwIfAborted(signal);

    const browser = typeof window !== 'undefined';
    const total = Math.max(1, countWork(wb));
    const work = packageWork(wb, chunkItems);
    let done = 0;
    let step = work.next();
    while (!step.done) {
      done += step.value.items;
      report({ phase: 'write', fraction: Math.min(1, done / total), detail: step.value.detail });
      throwIfAborted(signal);
      if (browser) {
        await yieldToEventLoop();
        throwIfAborted(signal);
      }
      step = work.next();
    }
    const parts = step.value;
    report({ phase: 'write', fraction: 1 });

    let largest = 0;
    const files: Record<string, [Uint8Array, ZipOptions]> = {};
    for (const [path, bytes] of parts) {
      files[path] = [bytes, { level: path.endsWith('.png') ? 0 : 6, mtime: ZIP_MTIME }];
      largest = Math.max(largest, bytes.byteLength);
    }
    let mode = await resolveZipMode(this.zipMode, largest);
    throwIfAborted(signal);
    report({ phase: 'zip', fraction: 0, detail: ZIP_DETAIL[mode] });
    let bytes: Uint8Array;
    if (mode === 'async') {
      try {
        bytes = await zipAsync(files, signal);
      } catch (error) {
        // 'auto' recovers from a worker failure on the main thread; an abort or a forced 'async' does not.
        if (signal?.aborted || this.zipMode === 'async') throw error;
        mode = 'sync';
        bytes = zipSync(files as Zippable, { level: 6, mtime: ZIP_MTIME });
      }
    } else {
      bytes = zipSync(files as Zippable, { level: 6, mtime: ZIP_MTIME });
    }
    throwIfAborted(signal);
    report({ phase: 'zip', fraction: 1, detail: ZIP_DETAIL[mode] });
    return bytes;
  }
}

export function createDefaultExcelWriter(options: OoxmlWriterOptions = {}): ExcelWriter {
  return new OoxmlExcelWriter(options);
}
