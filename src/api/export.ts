/**
 * Programmatic export API: Highcharts chart (or options) → XLSX bytes with a native, editable chart.
 *
 * Pipeline: resolve options → extract ChartModel → hooks.transformModel → theme overrides →
 * (optional reference image) → translate to sheet specs → write OOXML package.
 * Nothing here triggers a download (see src/browser/download.ts).
 */

import type { ChartModel } from '../types/chart-model';
import {
  DiagnosticCollector,
  buildCompatibilityReport,
  type CompatibilityReport,
  type Diagnostic,
} from '../types/diagnostics';
import {
  ExportError,
  XLSX_MIME_TYPE,
  resolveExportOptions,
  type ExportOptions,
  type ExportPhase,
  type ExportProgress,
  type ExportResult,
  type MultiChartExportEntry,
  type MultiChartExportResult,
} from '../types/public-api';
import type { ExcelWriter, SheetSpec, WorkbookSpec, WriteProgress } from '../excel/writer-interface';
import { createDefaultExcelWriter } from '../excel/ooxml-writer';
import { extractChartModel, extractChartModelFromOptions } from '../highcharts/extract-chart';
import { getHighchartsVersion } from '../highcharts/guards';
import { applyThemeOverrides } from '../core/theme-overrides';
import { translateChartModel, type TranslationResult } from '../core/translate-chart';
import { sanitizeFilename } from '../utils/filenames';
import {
  REFERENCE_IMAGE_PROPERTY,
  renderReferenceImage,
  reportReferenceImageUnavailable,
  type ReferenceImage,
} from '../browser/reference-image';

type ResolvedOptions = ReturnType<typeof resolveExportOptions>;

const CREATOR = 'highcharts-editable-excel';

type Source = { kind: 'chart'; chart: unknown } | { kind: 'options'; options: object };

interface Prepared {
  model: ChartModel;
  collector: DiagnosticCollector;
}

interface Translated extends Prepared {
  translation: TranslationResult;
  warnings: Diagnostic[];
  report: CompatibilityReport;
}

function now(): number {
  return typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now();
}

function elapsed(from: number): number {
  return Math.max(0, now() - from);
}

function diagnosticKey(d: Diagnostic): string {
  return `${d.code}|${d.property}|${d.seriesIndex ?? ''}`;
}

function mergeWarnings(...lists: ReadonlyArray<readonly Diagnostic[]>): Diagnostic[] {
  const seen = new Set<string>();
  const out: Diagnostic[] = [];
  for (const list of lists) {
    for (const d of list) {
      const key = diagnosticKey(d);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(d);
    }
  }
  return out;
}

/**
 * Lets the browser paint (e.g. a spinner) between the synchronous phases. A macrotask, only where
 * there is a `window`; Node and workers are not slowed down.
 */
function yieldToEventLoop(): Promise<void> {
  if (typeof window === 'undefined') return Promise.resolve();
  return new Promise((resolve) => setTimeout(resolve, 0));
}

// ---------------------------------------------------------------------------
// Cancellation and progress
// ---------------------------------------------------------------------------

/** ExportError ABORTED for an aborted signal; `cause` is `signal.reason`. */
export function abortedError(signal: AbortSignal): ExportError {
  return new ExportError('ABORTED', 'The export was aborted (ExportOptions.signal).', {}, { cause: signal.reason });
}

function isAbortSignal(value: unknown): value is AbortSignal {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as { aborted?: unknown; addEventListener?: unknown };
  return typeof v.aborted === 'boolean' && typeof v.addEventListener === 'function';
}

/**
 * Validates `signal` and `onProgress` (the other options are checked by `resolveExportOptions`).
 *
 * @throws ExportError INVALID_OPTIONS naming the option.
 */
function validateRuntimeOptions(o: { signal?: unknown; onProgress?: unknown }): void {
  const shown = (v: unknown) => (typeof v === 'string' ? JSON.stringify(v) : typeof v);
  if (o.signal !== undefined && !isAbortSignal(o.signal)) {
    throw new ExportError(
      'INVALID_OPTIONS',
      `Invalid export option "signal": expected an AbortSignal, got ${shown(o.signal)}.`,
      { property: 'signal' },
    );
  }
  if (o.onProgress !== undefined && typeof o.onProgress !== 'function') {
    throw new ExportError(
      'INVALID_OPTIONS',
      `Invalid export option "onProgress": expected a function, got ${shown(o.onProgress)}.`,
      { property: 'onProgress' },
    );
  }
}

/**
 * Overall fraction at which each phase starts. Extraction and translation are cheap next to
 * serializing and zipping (see the README benchmark), so they share the first 20 %.
 */
const WRITE_START = 0.2;
const ZIP_START = 0.6;

/**
 * Turns phase events into monotonic `ExportProgress` calls and owns the abort checks.
 * Remembers an error thrown by the user's callback so it can propagate unwrapped.
 */
class ExportRun {
  private last = 0;
  private zipStartedAt: number | null = null;
  private thrown: { error: unknown } | null = null;

  constructor(
    readonly signal: AbortSignal | undefined,
    private readonly callback: ((progress: ExportProgress) => void) | undefined,
  ) {}

  /** @throws ExportError ABORTED once the signal is aborted. */
  checkpoint(): void {
    if (this.signal?.aborted) throw abortedError(this.signal);
  }

  emit(phase: ExportPhase, fraction: number, detail?: string): void {
    if (!this.callback) return;
    const f = Math.min(1, Math.max(this.last, Number.isFinite(fraction) ? fraction : this.last));
    this.last = f;
    try {
      this.callback(detail === undefined ? { phase, fraction: f } : { phase, fraction: f, detail });
    } catch (error) {
      this.thrown = { error };
      throw error;
    }
  }

  /** Writer progress (phase-local fractions) mapped onto the overall scale. */
  readonly onWriterProgress = (p: WriteProgress): void => {
    const f = Math.min(1, Math.max(0, typeof p?.fraction === 'number' ? p.fraction : 0));
    if (p?.phase === 'zip') {
      this.zipStartedAt ??= now();
      this.emit('zip', ZIP_START + (1 - ZIP_START) * f, p.detail);
    } else if (p?.phase === 'write') {
      this.emit('write', WRITE_START + (ZIP_START - WRITE_START) * f, p.detail);
    }
  };

  /** Time since the writer reported the start of its zip phase (0 when it never did). */
  zipMs(): number {
    return this.zipStartedAt === null ? 0 : elapsed(this.zipStartedAt);
  }

  isCallbackError(error: unknown): boolean {
    return this.thrown !== null && this.thrown.error === error;
  }
}

function describe(list: readonly Diagnostic[]): string {
  return list.map((d) => `${d.code} at ${d.property}: ${d.message}`).join('; ');
}

/** Extraction + transformModel hook + theme overrides. Model warnings are fed into the collector. */
function prepareModel(source: Source, resolved: ResolvedOptions, collector: DiagnosticCollector): Prepared {
  const extract = {
    dataMode: resolved.dataMode,
    seriesVisibility: resolved.seriesVisibility,
    diagnostics: collector,
    ...(resolved.chartWidth !== undefined ? { chartWidth: resolved.chartWidth } : {}),
    ...(resolved.chartHeight !== undefined ? { chartHeight: resolved.chartHeight } : {}),
    ...(resolved.themeOverrides?.cssVariables ? { cssVariables: resolved.themeOverrides.cssVariables } : {}),
  };
  let model =
    source.kind === 'chart'
      ? extractChartModel(source.chart, { ...extract, highchartsVersion: getHighchartsVersion(source.chart) })
      : extractChartModelFromOptions(source.options, {
          ...extract,
          highchartsVersion: getHighchartsVersion(undefined),
        });
  const transform = resolved.hooks?.transformModel;
  if (transform) {
    const transformed = transform(model);
    if (transformed === null || typeof transformed !== 'object' || !Array.isArray((transformed as ChartModel).series)) {
      throw new ExportError('INVALID_OPTIONS', 'hooks.transformModel must return a ChartModel.', {
        chartId: model.meta.chartId,
      });
    }
    model = transformed;
  }
  model = applyThemeOverrides(model, resolved.themeOverrides);
  // Warnings added by the hook or the overrides go through the collector (dedupes, fires onWarning).
  collector.addAll(model.warnings);
  return { model, collector };
}

function translatePrepared(
  prepared: Prepared,
  resolved: ResolvedOptions,
  takenSheetNames: ReadonlySet<string>,
  referenceImage: ReferenceImage | null,
): Translated {
  const { model, collector } = prepared;
  const translation = translateChartModel(model, {
    chartSheetName: resolved.chartSheetName,
    dataSheetName: resolved.dataSheetName,
    includeSourceData: resolved.includeSourceData,
    fidelity: resolved.fidelity,
    diagnostics: collector,
    ...(resolved.chartWidth !== undefined ? { chartWidth: resolved.chartWidth } : {}),
    ...(resolved.chartHeight !== undefined ? { chartHeight: resolved.chartHeight } : {}),
    referenceImage,
    takenSheetNames,
  });
  const warnings = mergeWarnings(model.warnings, collector.items);
  const report = buildCompatibilityReport(
    model.meta.sourceChartType,
    translation.excelChartType,
    warnings,
    translation.supportedProperties,
  );
  return { ...prepared, translation, warnings, report };
}

/**
 * Throws CHART_NOT_EDITABLE when the chart cannot become a native chart (always), and in strict mode
 * also when any diagnostic reports a visible fidelity loss ('unsupported' at warning/error severity).
 */
function enforceEditable(t: Translated, strictMode: boolean, label: string): void {
  const chartId = t.model.meta.chartId;
  if (t.translation.blocking) {
    const blocking = t.warnings.filter((d) => d.outcome === 'blocking');
    throw new ExportError(
      'CHART_NOT_EDITABLE',
      `${label} cannot be exported as a native editable Excel chart: ${describe(blocking) || 'the chart type has no Excel equivalent.'}`,
      { chartId, diagnostics: blocking },
    );
  }
  if (!strictMode) return;
  const losses = t.warnings.filter(
    (d) =>
      d.outcome === 'unsupported' &&
      (d.severity === 'warning' || d.severity === 'error') &&
      d.property !== REFERENCE_IMAGE_PROPERTY,
  );
  if (losses.length > 0) {
    throw new ExportError('CHART_NOT_EDITABLE', `strictMode: ${label} has unsupported features: ${describe(losses)}`, {
      chartId,
      diagnostics: losses,
    });
  }
}

/**
 * Runs the writer with the run's signal and progress. Checks the signal before and after, so a
 * custom writer that ignores the context is still cancelled at the phase boundary.
 */
async function writeWorkbook(spec: WorkbookSpec, writer: ExcelWriter | undefined, run: ExportRun): Promise<Uint8Array> {
  run.checkpoint();
  run.emit('write', WRITE_START);
  try {
    const bytes = await (writer ?? createDefaultExcelWriter()).write(spec, {
      ...(run.signal ? { signal: run.signal } : {}),
      onProgress: run.onWriterProgress,
    });
    run.checkpoint();
    return bytes;
  } catch (error) {
    if (run.isCallbackError(error)) throw error;
    if (run.signal?.aborted) throw abortedError(run.signal);
    if (error instanceof ExportError) throw error;
    throw new ExportError(
      'WRITER_FAILURE',
      `Writing the XLSX package failed: ${error instanceof Error ? error.message : String(error)}`,
      {},
      { cause: error },
    );
  }
}

async function referenceImageFor(
  source: Source,
  model: ChartModel,
  resolved: ResolvedOptions,
  collector: DiagnosticCollector,
): Promise<ReferenceImage | null> {
  if (!resolved.includeReferenceImage) return null;
  if (source.kind === 'options') {
    reportReferenceImageUnavailable(collector, 'the reference image requires a browser and a rendered chart.', 'info');
    return null;
  }
  return renderReferenceImage(
    source.chart,
    resolved.chartWidth ?? model.width,
    resolved.chartHeight ?? model.height,
    collector,
  );
}

/**
 * Translates, enforces editability, and only then renders the optional reference image (a blocked
 * chart never pays for it). With an image the chart is translated again so the sheet carries it.
 */
async function translateWithImage(
  source: Source,
  prepared: Prepared,
  resolved: ResolvedOptions,
  taken: ReadonlySet<string>,
  label: string,
): Promise<{ translated: Translated; translateMs: number; imageMs: number }> {
  let t = now();
  let translated = translatePrepared(prepared, resolved, taken, null);
  let translateMs = elapsed(t);
  enforceEditable(translated, resolved.strictMode, label);
  if (!resolved.includeReferenceImage) return { translated, translateMs, imageMs: 0 };

  t = now();
  const image = await referenceImageFor(source, prepared.model, resolved, prepared.collector);
  const imageMs = elapsed(t);
  if (resolved.signal?.aborted) throw abortedError(resolved.signal);
  t = now();
  translated = translatePrepared(prepared, resolved, taken, image);
  translateMs += elapsed(t);
  enforceEditable(translated, resolved.strictMode, label);
  return { translated, translateMs, imageMs };
}

async function runSingle(source: Source, options: ExportOptions | undefined): Promise<ExportResult> {
  const t0 = now();
  const resolved = resolveExportOptions(options);
  validateRuntimeOptions(resolved);
  const run = new ExportRun(resolved.signal, resolved.onProgress);
  const collector = new DiagnosticCollector(resolved.onWarning);

  run.checkpoint();
  run.emit('extract', 0);
  const tExtract = now();
  const prepared = prepareModel(source, resolved, collector);
  const extractMs = elapsed(tExtract);
  await yieldToEventLoop();
  run.checkpoint();

  run.emit('translate', WRITE_START / 4);
  const { translated, translateMs, imageMs } = await translateWithImage(
    source,
    prepared,
    resolved,
    new Set<string>(),
    'The chart',
  );
  await yieldToEventLoop();

  const { model } = translated;
  const tWrite = now();
  const bytes = await writeWorkbook(
    {
      properties: {
        title: resolved.properties?.title ?? (model.title?.text || undefined),
        creator: resolved.properties?.creator ?? CREATOR,
      },
      sheets: translated.translation.sheets,
    },
    resolved.writer,
    run,
  );
  const writeMs = elapsed(tWrite);
  const zipMs = Math.min(writeMs, run.zipMs());

  const result: ExportResult = {
    bytes,
    filename: sanitizeFilename(resolved.filename ?? model.title?.text ?? 'chart'),
    mimeType: XLSX_MIME_TYPE,
    warnings: translated.warnings,
    report: translated.report,
    timings: { extractMs, translateMs, writeMs, zipMs, imageMs, totalMs: elapsed(t0) },
  };
  if (resolved.includeModel) result.model = model;
  run.emit('done', 1);
  return result;
}

/**
 * Exports a rendered Highcharts chart to an XLSX workbook (chart sheet + data sheet) containing a
 * native Excel chart that references the worksheet cells. Never triggers a download.
 *
 * In a browser the export yields to the event loop between the extract, translate and write
 * phases and, inside the built-in writer, between chunks of cells and chart cache points, so the
 * page can repaint (e.g. a spinner); the package is then zipped in Web Workers. Extraction and
 * translation are each synchronous. `signal` cancels it; `onProgress` reports it.
 *
 * @throws ExportError INVALID_OPTIONS for invalid options; INVALID_CHART for non-charts;
 *   CHART_NOT_EDITABLE when no native chart can be produced (or, with `strictMode`, when any
 *   feature is dropped); WRITER_FAILURE on writer errors (the writer's error is `error.cause`);
 *   ABORTED once `signal` is aborted (`error.cause` is `signal.reason`).
 *   Errors thrown by your own callbacks (`onWarning`, `onProgress`, `hooks.transformModel`)
 *   propagate unwrapped, as thrown.
 */
export async function exportHighchartsToXlsx(chart: unknown, options?: ExportOptions): Promise<ExportResult> {
  return runSingle({ kind: 'chart', chart }, options);
}

/**
 * Server-side / headless variant: exports from a plain Highcharts options object without rendering.
 * Uses the raw `series[i].data` and Highcharts' default styles where the options do not specify them.
 */
export async function exportHighchartsOptionsToXlsx(
  highchartsOptions: object,
  options?: ExportOptions,
): Promise<ExportResult> {
  return runSingle({ kind: 'options', options: highchartsOptions }, options);
}

/**
 * Synchronous dry run: extracts and translates the chart without writing a workbook and returns
 * the compatibility report. Blocked charts return `editable: false` instead of throwing.
 *
 * @throws ExportError INVALID_CHART for non-charts; INVALID_OPTIONS for invalid options. Errors
 *   thrown by `onWarning` or `hooks.transformModel` propagate unwrapped.
 */
export function analyzeChartCompatibility(chart: unknown, options?: ExportOptions): CompatibilityReport {
  const resolved = resolveExportOptions(options);
  const collector = new DiagnosticCollector(resolved.onWarning);
  const prepared = prepareModel({ kind: 'chart', chart }, resolved, collector);
  return translatePrepared(prepared, resolved, new Set<string>(), null).report;
}

/**
 * Exports several charts into one workbook. Each chart gets its own chart/data sheet pair
 * (default names "Chart 1"/"Data 1", "Chart 2"/"Data 2", …).
 *
 * An entry's `strictMode` overrides the workbook-wide `strictMode`. The reference image, when an
 * entry asks for one, is rendered only after that entry is known to be exportable.
 *
 * `signal` and `onProgress` are workbook-wide (second argument); in entry options they are ignored.
 *
 * @throws ExportError CHART_NOT_EDITABLE naming the entry index when a chart cannot be exported;
 *   INVALID_OPTIONS for invalid entries or options; ABORTED once `signal` is aborted. Errors
 *   thrown by your own callbacks propagate unwrapped.
 */
export async function exportChartsToWorkbook(
  entries: MultiChartExportEntry[],
  options: {
    filename?: string;
    properties?: ExportOptions['properties'];
    strictMode?: boolean;
    writer?: ExcelWriter;
    signal?: AbortSignal;
    onProgress?: (progress: ExportProgress) => void;
  } = {},
): Promise<MultiChartExportResult> {
  if (!Array.isArray(entries) || entries.length === 0) {
    throw new ExportError(
      'INVALID_OPTIONS',
      'exportChartsToWorkbook expects a non-empty array of { chart, options } entries.',
    );
  }
  // Workbook-level options are validated up front, before any chart work.
  resolveExportOptions({
    ...(options.filename !== undefined ? { filename: options.filename } : {}),
    ...(options.writer !== undefined ? { writer: options.writer } : {}),
  });
  validateRuntimeOptions(options);
  const run = new ExportRun(options.signal, options.onProgress);
  run.checkpoint();
  const taken = new Set<string>();
  const sheets: SheetSpec[] = [];
  const charts: MultiChartExportResult['charts'] = [];
  let firstTitle: string | undefined;

  for (const [i, entry] of entries.entries()) {
    if (entry === null || typeof entry !== 'object') {
      throw new ExportError('INVALID_OPTIONS', `Entry ${i} must be an object of the form { chart, options }.`);
    }
    const n = i + 1;
    // signal / onProgress are workbook-wide: the entry's own are replaced by the workbook's.
    const { signal: _signal, onProgress: _onProgress, ...entryOptions }: ExportOptions = { ...(entry.options ?? {}) };
    const resolved = resolveExportOptions({
      ...entryOptions,
      chartSheetName: entryOptions.chartSheetName ?? `Chart ${n}`,
      dataSheetName: entryOptions.dataSheetName ?? `Data ${n}`,
      strictMode: entryOptions.strictMode ?? options.strictMode ?? false,
      ...(options.signal ? { signal: options.signal } : {}),
    });
    const span = WRITE_START / entries.length;
    run.checkpoint();
    run.emit('extract', span * i, `chart entry ${i}`);
    const collector = new DiagnosticCollector(resolved.onWarning);
    const source: Source = { kind: 'chart', chart: entry.chart };
    let prepared: Prepared;
    try {
      prepared = prepareModel(source, resolved, collector);
    } catch (error) {
      if (error instanceof ExportError && error.code === 'INVALID_CHART') {
        throw new ExportError('INVALID_CHART', `Entry ${i}: ${error.message}`, { ...error.details }, { cause: error });
      }
      throw error;
    }
    await yieldToEventLoop();
    run.checkpoint();
    run.emit('translate', span * (i + 0.25), `chart entry ${i}`);
    const { translated } = await translateWithImage(source, prepared, resolved, taken, `Chart entry ${i}`);
    const { translation } = translated;
    taken.add(translation.chartSheetName);
    taken.add(translation.dataSheetName);
    sheets.push(...translation.sheets);
    charts.push({
      chartSheetName: translation.chartSheetName,
      dataSheetName: translation.dataSheetName,
      warnings: translated.warnings,
      report: translated.report,
    });
    firstTitle ??= translated.model.title?.text || undefined;
  }
  await yieldToEventLoop();

  const bytes = await writeWorkbook(
    {
      properties: { title: options.properties?.title ?? firstTitle, creator: options.properties?.creator ?? CREATOR },
      sheets,
    },
    options.writer,
    run,
  );
  run.emit('done', 1);
  return { bytes, filename: sanitizeFilename(options.filename ?? 'charts'), mimeType: XLSX_MIME_TYPE, charts };
}
