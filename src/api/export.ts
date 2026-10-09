/**
 * Programmatic export API: Highcharts chart (or options) → XLSX bytes with a native, editable chart.
 *
 * Pipeline: resolve options → extract ChartModel → hooks.transformModel → theme overrides →
 * (optional reference image) → translate to sheet specs → write OOXML package.
 * Nothing here triggers a download (see src/browser/download.ts).
 */

import type { ChartModel } from '../types/chart-model';
import { DiagnosticCollector, buildCompatibilityReport, type CompatibilityReport, type Diagnostic } from '../types/diagnostics';
import {
  ExportError,
  XLSX_MIME_TYPE,
  resolveExportOptions,
  type ExportOptions,
  type ExportResult,
  type MultiChartExportEntry,
  type MultiChartExportResult,
} from '../types/public-api';
import type { SheetSpec, WorkbookSpec } from '../excel/writer-interface';
import { createDefaultExcelWriter } from '../excel/ooxml-writer';
import { extractChartModel, extractChartModelFromOptions } from '../highcharts/extract-chart';
import { getHighchartsVersion } from '../highcharts/guards';
import { applyThemeOverrides } from '../core/theme-overrides';
import { translateChartModel, type TranslationResult } from '../core/translate-chart';
import { sanitizeFilename } from '../utils/filenames';
import { REFERENCE_IMAGE_PROPERTY, renderReferenceImage, reportReferenceImageUnavailable, type ReferenceImage } from '../browser/reference-image';

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
  };
  let model =
    source.kind === 'chart'
      ? extractChartModel(source.chart, { ...extract, highchartsVersion: getHighchartsVersion(source.chart) })
      : extractChartModelFromOptions(source.options, { ...extract, highchartsVersion: getHighchartsVersion(undefined) });
  const transform = resolved.hooks?.transformModel;
  if (transform) {
    const transformed = transform(model);
    if (transformed === null || typeof transformed !== 'object' || !Array.isArray((transformed as ChartModel).series)) {
      throw new ExportError('INVALID_OPTIONS', 'hooks.transformModel must return a ChartModel.', { chartId: model.meta.chartId });
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
  const report = buildCompatibilityReport(model.meta.sourceChartType, translation.excelChartType, warnings, translation.supportedProperties);
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
    (d) => d.outcome === 'unsupported' && (d.severity === 'warning' || d.severity === 'error') && d.property !== REFERENCE_IMAGE_PROPERTY,
  );
  if (losses.length > 0) {
    throw new ExportError('CHART_NOT_EDITABLE', `strictMode: ${label} has unsupported features: ${describe(losses)}`, {
      chartId,
      diagnostics: losses,
    });
  }
}

async function writeWorkbook(spec: WorkbookSpec): Promise<Uint8Array> {
  try {
    return await createDefaultExcelWriter().write(spec);
  } catch (error) {
    if (error instanceof ExportError) throw error;
    throw new ExportError('WRITER_FAILURE', `Writing the XLSX package failed: ${error instanceof Error ? error.message : String(error)}`, {
      cause: error,
    });
  }
}

async function referenceImageFor(source: Source, model: ChartModel, resolved: ResolvedOptions, collector: DiagnosticCollector): Promise<ReferenceImage | null> {
  if (!resolved.includeReferenceImage) return null;
  if (source.kind === 'options') {
    reportReferenceImageUnavailable(collector, 'the reference image requires a browser and a rendered chart.', 'info');
    return null;
  }
  return renderReferenceImage(source.chart, resolved.chartWidth ?? model.width, resolved.chartHeight ?? model.height, collector);
}

async function runSingle(source: Source, options: ExportOptions | undefined): Promise<ExportResult> {
  const t0 = now();
  const resolved = resolveExportOptions(options);
  const collector = new DiagnosticCollector(resolved.onWarning);

  const tExtract = now();
  const prepared = prepareModel(source, resolved, collector);
  const image = await referenceImageFor(source, prepared.model, resolved, collector);
  const extractMs = elapsed(tExtract);

  const tTranslate = now();
  const translated = translatePrepared(prepared, resolved, new Set<string>(), image);
  const translateMs = elapsed(tTranslate);
  enforceEditable(translated, resolved.strictMode, 'The chart');

  const { model } = translated;
  const tWrite = now();
  const bytes = await writeWorkbook({
    properties: {
      title: resolved.properties?.title ?? (model.title?.text || undefined),
      creator: resolved.properties?.creator ?? CREATOR,
    },
    sheets: translated.translation.sheets,
  });
  const writeMs = elapsed(tWrite);

  const result: ExportResult = {
    bytes,
    filename: sanitizeFilename(resolved.filename ?? model.title?.text ?? 'chart'),
    mimeType: XLSX_MIME_TYPE,
    warnings: translated.warnings,
    report: translated.report,
    timings: { extractMs, translateMs, writeMs, totalMs: elapsed(t0) },
  };
  if (resolved.includeModel) result.model = model;
  return result;
}

/**
 * Exports a rendered Highcharts chart to an XLSX workbook (chart sheet + data sheet) containing a
 * native Excel chart that references the worksheet cells. Never triggers a download.
 *
 * @throws ExportError INVALID_CHART for non-charts; CHART_NOT_EDITABLE when no native chart can be
 *   produced (or, with `strictMode`, when any feature is dropped); WRITER_FAILURE on writer errors.
 */
export async function exportHighchartsToXlsx(chart: unknown, options?: ExportOptions): Promise<ExportResult> {
  return runSingle({ kind: 'chart', chart }, options);
}

/**
 * Server-side / headless variant: exports from a plain Highcharts options object without rendering.
 * Uses the raw `series[i].data` and Highcharts' default styles where the options do not specify them.
 */
export async function exportHighchartsOptionsToXlsx(highchartsOptions: object, options?: ExportOptions): Promise<ExportResult> {
  return runSingle({ kind: 'options', options: highchartsOptions }, options);
}

/**
 * Synchronous dry run: extracts and translates the chart without writing a workbook and returns
 * the compatibility report. Blocked charts return `editable: false`; only INVALID_CHART throws.
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
 * @throws ExportError CHART_NOT_EDITABLE naming the entry index when a chart cannot be exported.
 */
export async function exportChartsToWorkbook(
  entries: MultiChartExportEntry[],
  options: { filename?: string; properties?: ExportOptions['properties']; strictMode?: boolean } = {},
): Promise<MultiChartExportResult> {
  if (!Array.isArray(entries) || entries.length === 0) {
    throw new ExportError('INVALID_OPTIONS', 'exportChartsToWorkbook expects a non-empty array of { chart, options } entries.');
  }
  const taken = new Set<string>();
  const sheets: SheetSpec[] = [];
  const charts: MultiChartExportResult['charts'] = [];
  let firstTitle: string | undefined;

  for (const [i, entry] of entries.entries()) {
    if (entry === null || typeof entry !== 'object') {
      throw new ExportError('INVALID_OPTIONS', `Entry ${i} must be an object of the form { chart, options }.`);
    }
    const n = i + 1;
    const entryOptions: ExportOptions = { ...(entry.options ?? {}) };
    const resolved = resolveExportOptions({
      ...entryOptions,
      chartSheetName: entryOptions.chartSheetName ?? `Chart ${n}`,
      dataSheetName: entryOptions.dataSheetName ?? `Data ${n}`,
      strictMode: entryOptions.strictMode ?? options.strictMode ?? false,
    });
    const collector = new DiagnosticCollector(resolved.onWarning);
    const source: Source = { kind: 'chart', chart: entry.chart };
    let prepared: Prepared;
    try {
      prepared = prepareModel(source, resolved, collector);
    } catch (error) {
      if (error instanceof ExportError && error.code === 'INVALID_CHART') {
        throw new ExportError('INVALID_CHART', `Entry ${i}: ${error.message}`, { ...error.details, cause: error });
      }
      throw error;
    }
    const image = await referenceImageFor(source, prepared.model, resolved, collector);
    const translated = translatePrepared(prepared, resolved, taken, image);
    enforceEditable(translated, resolved.strictMode, `Chart entry ${i}`);
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

  const bytes = await writeWorkbook({
    properties: { title: options.properties?.title ?? firstTitle, creator: options.properties?.creator ?? CREATOR },
    sheets,
  });
  return { bytes, filename: sanitizeFilename(options.filename ?? 'charts'), mimeType: XLSX_MIME_TYPE, charts };
}
