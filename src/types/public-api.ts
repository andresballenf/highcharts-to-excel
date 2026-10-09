/**
 * Public option and result types.
 */

import type { ChartModel } from './chart-model';
import type { CompatibilityReport, Diagnostic } from './diagnostics';

/**
 * How hard the exporter tries to reproduce the source styling.
 * - 'best-effort' (default): translate every supported style; approximate the rest and report it.
 * - 'minimal': translate chart type, data, titles and axis types only; leave colors, fonts and
 *   layout to Excel's defaults. Useful when the Excel theme should win.
 */
export type FidelityMode = 'best-effort' | 'minimal';

/**
 * Which data to export.
 * - 'rendered' (default): the points the chart currently displays (after setData/update, after
 *   data grouping and cropping). Grouping/cropping is reported as a diagnostic.
 * - 'raw': the series' source data as last supplied (series.options.data), ignoring data grouping
 *   and zoom. Falls back to rendered data with a DATA_MODE_FALLBACK diagnostic if raw data is
 *   not recoverable.
 */
export type DataMode = 'rendered' | 'raw';

/**
 * Which series are exported.
 * - 'visible' (default): only series currently visible; hidden ones are reported.
 * - 'all': every series, including hidden ones (they appear in Excel as visible).
 */
export type SeriesVisibilityMode = 'visible' | 'all';

/** Partial font override; null-ish fields keep the extracted value. */
export interface FontOverride {
  family?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

/**
 * Explicit styling applied on top of (or, headless, instead of) extracted styles.
 * Colors are CSS color strings.
 */
export interface ThemeOverrides {
  /** Series palette, by series index. */
  colors?: string[];
  chartBackground?: string;
  plotBackground?: string;
  fontFamily?: string;
  title?: FontOverride;
  subtitle?: FontOverride;
  axisTitle?: FontOverride;
  axisLabels?: FontOverride;
  legend?: FontOverride;
  dataLabels?: FontOverride;
  gridLineColor?: string;
  gridLineWidth?: number;
  /** Per-series overrides keyed by series index. */
  series?: Record<number, { color?: string; lineWidth?: number; fillOpacity?: number }>;
}

export interface ExportHooks {
  /**
   * Called after extraction and normalization, before translation. May return a modified model.
   * Must not mutate the Highcharts chart.
   */
  transformModel?: (model: ChartModel) => ChartModel;
}

export interface ExportOptions {
  /** Output filename. ".xlsx" is appended if missing. Default: derived from the chart title or "chart.xlsx". */
  filename?: string;
  /** Name of the worksheet holding the chart. Default "Chart". Invalid characters are replaced and reported. */
  chartSheetName?: string;
  /** Name of the worksheet holding the data. Default "Data". */
  dataSheetName?: string;
  /**
   * When true (default) the data sheet is visible. When false it is hidden; the chart still
   * references it (the chart must reference cells to be editable).
   */
  includeSourceData?: boolean;
  fidelity?: FidelityMode;
  dataMode?: DataMode;
  seriesVisibility?: SeriesVisibilityMode;
  /** Receives each diagnostic as it is raised. */
  onWarning?: (diagnostic: Diagnostic) => void;
  themeOverrides?: ThemeOverrides;
  /** Override the exported chart size in CSS pixels. Defaults to the rendered chart size. */
  chartWidth?: number;
  chartHeight?: number;
  /**
   * When true, throw ExportError (code CHART_NOT_EDITABLE) instead of producing a workbook whenever
   * any diagnostic is 'blocking' or the chart type is unsupported. Default false.
   */
  strictMode?: boolean;
  /**
   * Opt-in: embed a PNG rendering of the source chart on the chart sheet next to the native chart,
   * for side-by-side comparison. Browser only; requires the Highcharts exporting module. Default false.
   */
  includeReferenceImage?: boolean;
  /** Include the normalized ChartModel in the result (for debugging). Default false. */
  includeModel?: boolean;
  hooks?: ExportHooks;
  /** Workbook document properties. */
  properties?: { title?: string; creator?: string };
}

export interface ExportTimings {
  extractMs: number;
  translateMs: number;
  writeMs: number;
  totalMs: number;
}

export interface ExportResult {
  bytes: Uint8Array;
  filename: string;
  /** MIME type to use for downloads. */
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  warnings: Diagnostic[];
  report: CompatibilityReport;
  timings: ExportTimings;
  /** Present when `includeModel` was set. */
  model?: ChartModel;
}

export interface MultiChartExportEntry {
  chart: unknown;
  /** Per-chart overrides; sheet names must be unique across the workbook. */
  options?: Omit<ExportOptions, 'filename' | 'properties'>;
}

export interface MultiChartExportResult {
  bytes: Uint8Array;
  filename: string;
  mimeType: ExportResult['mimeType'];
  charts: Array<{ chartSheetName: string; dataSheetName: string; warnings: Diagnostic[]; report: CompatibilityReport }>;
}

export interface InstallOptions {
  /** Menu item text. Default "Download editable Excel chart". */
  menuText?: string;
  /** Key used in exporting.menuItemDefinitions. Default "downloadEditableXLSX". */
  menuItemKey?: string;
  /** Default export options applied to every chart. */
  exportOptions?: ExportOptions;
  /**
   * Called when a download completes (or fails). When omitted, errors are rethrown to the console
   * via `console.error` and warnings are available via `exportOptions.onWarning`.
   */
  onExport?: (result: ExportResult, chart: unknown) => void;
  onError?: (error: unknown, chart: unknown) => void;
  /** Where to insert the item relative to Highcharts' default items. Default: after "downloadXLS"/"downloadCSV" if present, else at the end. */
  insertAfter?: string;
}

/**
 * Per-chart options read from `chart.options.exporting.editableExcel`
 * (or `exporting.editableExcel` in the chart config). Developers set this in their Highcharts
 * options to override installation defaults or disable the menu item for a single chart.
 */
export interface PerChartExportConfig extends ExportOptions {
  /** Set false to omit the menu item on this chart. */
  enabled?: boolean;
  menuText?: string;
}

export interface Installation {
  /** Removes the menu item definition and the global default menu entry. Charts created later are unaffected. */
  uninstall(): void;
  readonly options: Readonly<InstallOptions>;
}

export type ExportErrorCode =
  | 'INVALID_CHART'
  | 'CHART_NOT_EDITABLE'
  | 'UNSUPPORTED_CHART_TYPE'
  | 'ROW_LIMIT_EXCEEDED'
  | 'COLUMN_LIMIT_EXCEEDED'
  | 'EXPORTING_MODULE_MISSING'
  | 'BROWSER_REQUIRED'
  | 'WRITER_FAILURE'
  | 'INVALID_OPTIONS';

export class ExportError extends Error {
  override readonly name = 'HighchartsExcelExportError';
  constructor(
    readonly code: ExportErrorCode,
    message: string,
    readonly details: { chartId?: string | null; property?: string; diagnostics?: Diagnostic[]; cause?: unknown } = {},
  ) {
    super(message);
    if (details.cause !== undefined) (this as { cause?: unknown }).cause = details.cause;
  }
}

export const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' as const;

export const DEFAULT_MENU_TEXT = 'Download editable Excel chart';
export const DEFAULT_MENU_ITEM_KEY = 'downloadEditableXLSX';

export function resolveExportOptions(options: ExportOptions = {}): Required<
  Pick<
    ExportOptions,
    | 'chartSheetName'
    | 'dataSheetName'
    | 'includeSourceData'
    | 'fidelity'
    | 'dataMode'
    | 'seriesVisibility'
    | 'strictMode'
    | 'includeReferenceImage'
    | 'includeModel'
  >
> &
  ExportOptions {
  return {
    ...options,
    chartSheetName: options.chartSheetName ?? 'Chart',
    dataSheetName: options.dataSheetName ?? 'Data',
    includeSourceData: options.includeSourceData ?? true,
    fidelity: options.fidelity ?? 'best-effort',
    dataMode: options.dataMode ?? 'rendered',
    seriesVisibility: options.seriesVisibility ?? 'visible',
    strictMode: options.strictMode ?? false,
    includeReferenceImage: options.includeReferenceImage ?? false,
    includeModel: options.includeModel ?? false,
  };
}
