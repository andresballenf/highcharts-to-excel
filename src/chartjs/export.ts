/**
 * Chart.js export API: chart instance or configuration → XLSX bytes with a native, editable chart.
 *
 * Same pipeline as the Highcharts entry points (src/api/export.ts), with the Chart.js extractor:
 * resolve options → extract ChartModel → hooks.transformModel → theme overrides → translate →
 * (optional reference image) → write. Nothing here triggers a download except
 * `downloadChartJsAsXlsx`.
 */

import { triggerDownload } from '../browser/download';
import {
  REFERENCE_IMAGE_PROPERTY,
  reportReferenceImageUnavailable,
  type ReferenceImage,
} from '../browser/reference-image';
import { applyThemeOverrides } from '../core/theme-overrides';
import { translateChartModel, type TranslationResult } from '../core/translate-chart';
import { createDefaultExcelWriter } from '../excel/ooxml-writer';
import type { ExcelWriter, WorkbookSpec, WriteProgress } from '../excel/writer-interface';
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
} from '../types/public-api';
import { sanitizeFilename } from '../utils/filenames';
import { extractChartJsModel } from './extract-chart';
import { isChartJsChart, isRealBrowser } from './guards';

type ResolvedOptions = ReturnType<typeof resolveExportOptions>;

const CREATOR = 'highcharts-editable-excel';
const WRITE_START = 0.2;
const ZIP_START = 0.6;

function now(): number {
  return typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now();
}

function elapsed(from: number): number {
  return Math.max(0, now() - from);
}

function yieldToEventLoop(): Promise<void> {
  if (typeof window === 'undefined') return Promise.resolve();
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function aborted(signal: AbortSignal): ExportError {
  return new ExportError('ABORTED', 'The export was aborted (ExportOptions.signal).', {}, { cause: signal.reason });
}

function validateRuntimeOptions(o: { signal?: unknown; onProgress?: unknown }): void {
  const s = o.signal as { aborted?: unknown; addEventListener?: unknown } | undefined;
  if (s !== undefined && (typeof s !== 'object' || s === null || typeof s.aborted !== 'boolean')) {
    throw new ExportError('INVALID_OPTIONS', 'Invalid export option "signal": expected an AbortSignal.', {
      property: 'signal',
    });
  }
  if (o.onProgress !== undefined && typeof o.onProgress !== 'function') {
    throw new ExportError('INVALID_OPTIONS', 'Invalid export option "onProgress": expected a function.', {
      property: 'onProgress',
    });
  }
}

/** Monotonic progress events, abort checks, and the user's callback errors kept unwrapped. */
class Run {
  private last = 0;
  private zipStartedAt: number | null = null;
  private thrown: { error: unknown } | null = null;

  constructor(
    readonly signal: AbortSignal | undefined,
    private readonly callback: ((p: ExportProgress) => void) | undefined,
  ) {}

  checkpoint(): void {
    if (this.signal?.aborted) throw aborted(this.signal);
  }

  emit(phase: ExportPhase, fraction: number, detail?: string): void {
    if (!this.callback) return;
    const f = Math.min(1, Math.max(this.last, fraction));
    this.last = f;
    try {
      this.callback(detail === undefined ? { phase, fraction: f } : { phase, fraction: f, detail });
    } catch (error) {
      this.thrown = { error };
      throw error;
    }
  }

  readonly onWriterProgress = (p: WriteProgress): void => {
    const f = Math.min(1, Math.max(0, typeof p?.fraction === 'number' ? p.fraction : 0));
    if (p?.phase === 'zip') {
      this.zipStartedAt ??= now();
      this.emit('zip', ZIP_START + (1 - ZIP_START) * f, p.detail);
    } else if (p?.phase === 'write') {
      this.emit('write', WRITE_START + (ZIP_START - WRITE_START) * f, p.detail);
    }
  };

  zipMs(): number {
    return this.zipStartedAt === null ? 0 : elapsed(this.zipStartedAt);
  }

  isCallbackError(error: unknown): boolean {
    return this.thrown !== null && this.thrown.error === error;
  }
}

function mergeWarnings(...lists: ReadonlyArray<readonly Diagnostic[]>): Diagnostic[] {
  const seen = new Set<string>();
  const out: Diagnostic[] = [];
  for (const list of lists) {
    for (const d of list) {
      const key = `${d.code}|${d.property}|${d.seriesIndex ?? ''}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(d);
    }
  }
  return out;
}

function describe(list: readonly Diagnostic[]): string {
  return list.map((d) => `${d.code} at ${d.property}: ${d.message}`).join('; ');
}

/** Extraction + transformModel hook + theme overrides. */
function prepareModel(source: unknown, resolved: ResolvedOptions, collector: DiagnosticCollector): ChartModel {
  let model = extractChartJsModel(source, {
    dataMode: resolved.dataMode,
    seriesVisibility: resolved.seriesVisibility,
    diagnostics: collector,
    ...(resolved.chartWidth !== undefined ? { chartWidth: resolved.chartWidth } : {}),
    ...(resolved.chartHeight !== undefined ? { chartHeight: resolved.chartHeight } : {}),
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
  collector.addAll(model.warnings);
  return model;
}

interface Translated {
  translation: TranslationResult;
  warnings: Diagnostic[];
  report: CompatibilityReport;
}

function translate(
  model: ChartModel,
  resolved: ResolvedOptions,
  collector: DiagnosticCollector,
  referenceImage: ReferenceImage | null,
): Translated {
  const translation = translateChartModel(model, {
    chartSheetName: resolved.chartSheetName,
    dataSheetName: resolved.dataSheetName,
    includeSourceData: resolved.includeSourceData,
    fidelity: resolved.fidelity,
    diagnostics: collector,
    ...(resolved.chartWidth !== undefined ? { chartWidth: resolved.chartWidth } : {}),
    ...(resolved.chartHeight !== undefined ? { chartHeight: resolved.chartHeight } : {}),
    referenceImage,
    takenSheetNames: new Set<string>(),
  });
  const warnings = mergeWarnings(model.warnings, collector.items);
  const report = buildCompatibilityReport(
    model.meta.sourceChartType,
    translation.excelChartType,
    warnings,
    translation.supportedProperties,
  );
  return { translation, warnings, report };
}

function enforceEditable(t: Translated, model: ChartModel, strictMode: boolean): void {
  const chartId = model.meta.chartId;
  if (t.translation.blocking) {
    const blocking = t.warnings.filter((d) => d.outcome === 'blocking');
    throw new ExportError(
      'CHART_NOT_EDITABLE',
      `The chart cannot be exported as a native editable Excel chart: ${describe(blocking) || 'the chart type has no Excel equivalent.'}`,
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
    throw new ExportError('CHART_NOT_EDITABLE', `strictMode: the chart has unsupported features: ${describe(losses)}`, {
      chartId,
      diagnostics: losses,
    });
  }
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** The chart's own canvas as a PNG (live charts in a real browser only). Never throws. */
function canvasReferenceImage(
  source: unknown,
  model: ChartModel,
  collector: DiagnosticCollector,
): ReferenceImage | null {
  if (!isChartJsChart(source)) {
    reportReferenceImageUnavailable(collector, 'the reference image requires a rendered Chart.js chart.', 'info');
    return null;
  }
  const canvas = source.canvas as { toDataURL?: (type?: string) => string } | undefined;
  if (!isRealBrowser() || typeof canvas?.toDataURL !== 'function') {
    reportReferenceImageUnavailable(collector, 'the reference image requires a browser canvas.');
    return null;
  }
  try {
    const url = canvas.toDataURL('image/png');
    const comma = url.indexOf(',');
    if (!url.startsWith('data:image/png') || comma === -1) throw new Error('the canvas returned no PNG');
    return { png: base64ToBytes(url.slice(comma + 1)), widthPx: model.width, heightPx: model.height };
  } catch (error) {
    reportReferenceImageUnavailable(
      collector,
      `the canvas could not be read (${error instanceof Error ? error.message : String(error)}).`,
    );
    return null;
  }
}

async function writeWorkbook(spec: WorkbookSpec, writer: ExcelWriter | undefined, run: Run): Promise<Uint8Array> {
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
    if (run.signal?.aborted) throw aborted(run.signal);
    if (error instanceof ExportError) throw error;
    throw new ExportError(
      'WRITER_FAILURE',
      `Writing the XLSX package failed: ${error instanceof Error ? error.message : String(error)}`,
      {},
      { cause: error },
    );
  }
}

/**
 * Exports a Chart.js 4 chart (a `Chart` instance, or a plain `{ type, data, options }`
 * configuration) to an XLSX workbook (chart sheet + data sheet) holding a native Excel chart that
 * references the worksheet cells. Never mutates the chart or the configuration; never downloads.
 *
 * @throws ExportError INVALID_OPTIONS for invalid options; INVALID_CHART when the input is neither a
 *   chart nor a configuration; CHART_NOT_EDITABLE when no native chart can be produced (radar,
 *   polarArea, empty charts, or with `strictMode` any dropped feature); WRITER_FAILURE on writer
 *   errors; ABORTED once `signal` is aborted. Errors thrown by your own callbacks propagate unwrapped.
 */
export async function exportChartJsToXlsx(chartOrConfig: unknown, options?: ExportOptions): Promise<ExportResult> {
  const t0 = now();
  const resolved = resolveExportOptions(options);
  validateRuntimeOptions(resolved);
  const run = new Run(resolved.signal, resolved.onProgress);
  const collector = new DiagnosticCollector(resolved.onWarning);

  run.checkpoint();
  run.emit('extract', 0);
  const tExtract = now();
  const model = prepareModel(chartOrConfig, resolved, collector);
  const extractMs = elapsed(tExtract);
  await yieldToEventLoop();
  run.checkpoint();

  run.emit('translate', WRITE_START / 4);
  let t = now();
  let translated = translate(model, resolved, collector, null);
  let translateMs = elapsed(t);
  enforceEditable(translated, model, resolved.strictMode);
  let imageMs = 0;
  if (resolved.includeReferenceImage) {
    t = now();
    const image = canvasReferenceImage(chartOrConfig, model, collector);
    imageMs = elapsed(t);
    t = now();
    translated = translate(model, resolved, collector, image);
    translateMs += elapsed(t);
    enforceEditable(translated, model, resolved.strictMode);
  }
  await yieldToEventLoop();

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

  const result: ExportResult = {
    bytes,
    filename: sanitizeFilename(resolved.filename ?? model.title?.text ?? 'chart'),
    mimeType: XLSX_MIME_TYPE,
    warnings: translated.warnings,
    report: translated.report,
    timings: { extractMs, translateMs, writeMs, zipMs: Math.min(writeMs, run.zipMs()), imageMs, totalMs: elapsed(t0) },
  };
  if (resolved.includeModel) result.model = model;
  run.emit('done', 1);
  return result;
}

/**
 * Synchronous dry run: extracts and translates without writing a workbook and returns the
 * compatibility report. Blocked charts return `editable: false` instead of throwing.
 *
 * @throws ExportError INVALID_CHART / INVALID_OPTIONS.
 */
export function analyzeChartJsCompatibility(chartOrConfig: unknown, options?: ExportOptions): CompatibilityReport {
  const resolved = resolveExportOptions(options);
  const collector = new DiagnosticCollector(resolved.onWarning);
  const model = prepareModel(chartOrConfig, resolved, collector);
  return translate(model, resolved, collector, null).report;
}

/**
 * Exports the chart and immediately downloads the workbook (browser only).
 *
 * @throws ExportError as `exportChartJsToXlsx`, plus BROWSER_REQUIRED outside a browser.
 */
export async function downloadChartJsAsXlsx(chartOrConfig: unknown, options?: ExportOptions): Promise<ExportResult> {
  const result = await exportChartJsToXlsx(chartOrConfig, options);
  const signal = options?.signal;
  if (signal?.aborted) throw aborted(signal);
  triggerDownload(result.bytes, result.filename, result.mimeType);
  return result;
}
