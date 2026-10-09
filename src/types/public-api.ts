/**
 * Public option and result types.
 */

import type { ChartModel } from './chart-model';
import type { CompatibilityReport, Diagnostic } from './diagnostics';
import type { ExcelWriter } from '../excel/writer-interface';

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
  /**
   * CSS custom property values for styled-mode charts, e.g. `{ '--highcharts-color-0': '#8e44ad' }`.
   * Consulted first when resolving `var(--…)` colors (headless and in a browser), so a server-side
   * export of a styled-mode chart can reproduce the app's theme; `STYLED_MODE_FALLBACK` is not
   * raised for series whose color variables all come from here.
   */
  cssVariables?: Record<string, string>;
}

export interface ExportHooks {
  /**
   * Called after extraction and normalization, before translation. May return a modified model.
   * Must not mutate the Highcharts chart.
   */
  transformModel?: (model: ChartModel) => ChartModel;
}

/** Export phases reported by `ExportOptions.onProgress`, in order. */
export type ExportPhase = 'extract' | 'translate' | 'write' | 'zip' | 'done';

/**
 * Progress of one export. `fraction` is the overall share done (0..1) and never decreases;
 * `detail` says what is being worked on (e.g. `sheet "Data"`, or the zip path taken).
 * Reported at least at the start of every phase and after every write chunk; 'done' (1) comes
 * once, last, after a successful export.
 */
export interface ExportProgress {
  phase: ExportPhase;
  fraction: number;
  detail?: string;
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
  /**
   * Override the exported chart size in CSS pixels (finite, 50 to 20000). Defaults to the rendered
   * chart size.
   */
  chartWidth?: number;
  chartHeight?: number;
  /**
   * Blocking diagnostics (unsupported chart type, polar, empty chart, row limits) always throw
   * ExportError CHART_NOT_EDITABLE: a degraded or image-based workbook is never produced.
   * When strictMode is true the export ALSO throws CHART_NOT_EDITABLE when any diagnostic has
   * outcome 'unsupported' with severity 'warning' or 'error' (a visible fidelity loss), instead of
   * silently dropping the feature. Default false.
   */
  strictMode?: boolean;
  /**
   * Custom Excel writer implementing the `ExcelWriter` interface (see `src/excel/writer-interface.ts`).
   * Defaults to the built-in OOXML writer. Advanced use only.
   */
  writer?: ExcelWriter;
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
  /**
   * Cancels the export. Checked between phases and between write chunks (and it terminates the
   * worker zip); once aborted the export rejects with ExportError ABORTED whose `cause` is
   * `signal.reason`. `downloadHighchartsAsXlsx` never downloads after an abort.
   */
  signal?: AbortSignal;
  /**
   * Receives progress events (see `ExportProgress`). Called synchronously; keep it cheap. Errors it
   * throws abort the export and propagate unwrapped.
   */
  onProgress?: (progress: ExportProgress) => void;
}

export interface ExportTimings {
  /** Extraction, `hooks.transformModel` and theme overrides (the reference image is excluded). */
  extractMs: number;
  translateMs: number;
  writeMs: number;
  /** Rendering the optional reference image (`includeReferenceImage`); 0 when none was rendered. */
  imageMs?: number;
  /**
   * Compressing the package, part of `writeMs` (which covers serializing and zipping). 0 when a
   * custom writer reports no 'zip' progress.
   */
  zipMs: number;
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
  /**
   * Per-chart overrides; sheet names must be unique across the workbook. An entry's `strictMode`
   * overrides the workbook-wide one. `writer` and `includeModel` are workbook-level concerns and
   * are not accepted here (pass the writer in `exportChartsToWorkbook`'s second argument).
   */
  options?: Omit<ExportOptions, 'filename' | 'properties' | 'writer' | 'includeModel' | 'signal' | 'onProgress'>;
}

export interface MultiChartExportResult {
  bytes: Uint8Array;
  filename: string;
  mimeType: ExportResult['mimeType'];
  charts: Array<{ chartSheetName: string; dataSheetName: string; warnings: Diagnostic[]; report: CompatibilityReport }>;
}

/**
 * Icon shown before the menu item text.
 * - `'excel'`: the built-in 14x14 spreadsheet glyph (`MENU_ICON_EXCEL_SVG`, drawn in `currentColor`).
 * - `{ svg }`: your own inline `<svg …>…</svg>` markup.
 * - `{ html }`: any other markup (an `<img>`, a `<span>` badge…).
 * - `null`: no icon (default).
 *
 * Markup is rendered by Highcharts' AST, which drops tags and attributes outside
 * `Highcharts.AST.allowedTags` / `allowedAttributes` (notably `viewBox` and `rx` are not allowed:
 * size the SVG with `width`/`height` and draw in pixel coordinates).
 */
export type MenuIconOption = 'excel' | { svg: string } | { html: string } | null;

/**
 * Global options for Highcharts' context (hamburger) button, written to
 * `exporting.buttons.contextButton` while installed and restored by `uninstall()`.
 * Applies to every chart; set `exporting.buttons.contextButton` in a chart's own options to
 * change a single chart.
 */
export interface ContextButtonOptions {
  /** Name of a symbol registered on the Highcharts renderer (`'menu'`, `'menuball'`, `'circle'`…). */
  symbol?: string;
  /**
   * A custom icon as an SVG path, registered as a renderer symbol named `symbol` (default
   * `'editableExcelButton'`). Either a `d` string or a flat array (`['M', 0, 0, 'L', 1, 1]`).
   * Commands M, L, H, V, C, Q, A and Z, absolute or relative (S, T are not supported). Coordinates
   * all within 0..1 are a unit box; otherwise the path's bounding box (made square, centered) is
   * scaled to the button's `symbolSize`.
   */
  svgPath?: string | Array<string | number>;
  /** Symbol fill (Highcharts default `#666666`). */
  symbolFill?: string;
  /** Symbol stroke color (Highcharts default `#666666`). */
  symbolStroke?: string;
  /** Symbol stroke width (Highcharts default 3). */
  symbolStrokeWidth?: number;
  /** Symbol box size in px (Highcharts default 14). */
  symbolSize?: number;
  /** SVG attributes of the button box (`fill`, `stroke`, `r`, `states.hover.fill`…). */
  theme?: Record<string, unknown>;
  /** Text drawn next to the symbol. */
  text?: string;
  /** Extra class name, added to Highcharts' `highcharts-contextbutton`. */
  className?: string;
  /** Button tooltip; sets `lang.contextButtonTitle`. */
  title?: string;
}

/** CSS declarations (camelCase keys) for the dropdown menu, as in Highcharts' `navigation.menuStyle`. */
export type MenuCssOptions = Record<string, string | number>;

export interface InstallOptions {
  /**
   * Menu item text for every chart. Default: none, the text then comes from
   * `Highcharts.getOptions().lang[langKey]` (default "Download editable Excel chart").
   * Resolution order: per-chart `exporting.editableExcel.menuText` → this option →
   * `lang[langKey]` → `DEFAULT_MENU_TEXT`. May contain markup allowed by Highcharts' AST.
   */
  menuText?: string;
  /** Key used in exporting.menuItemDefinitions. Default "downloadEditableXLSX". */
  menuItemKey?: string;
  /**
   * The `lang` key holding the item text (the definition's `textKey`). Default
   * `'downloadEditableXLSX'`. Translate with `Highcharts.setOptions({ lang: { [langKey]: '…' } })`
   * before or after install (charts created afterwards pick it up). When the key is absent, install
   * registers `DEFAULT_MENU_TEXT` under it and `uninstall()` removes it again.
   */
  langKey?: string;
  /** Icon before the menu item text. Default `null` (no icon). Per chart: `exporting.editableExcel.menuIcon`. */
  menuIcon?: MenuIconOption;
  /** Context-button branding (symbol, colors, custom SVG path, title). Default: Highcharts' button unchanged. */
  button?: ContextButtonOptions;
  /** Merged into `navigation.menuStyle` (the dropdown box). Ignored in styled mode (use CSS). */
  menuStyle?: MenuCssOptions;
  /** Merged into `navigation.menuItemStyle` (every menu entry). Ignored in styled mode. */
  menuItemStyle?: MenuCssOptions;
  /** Merged into `navigation.menuItemHoverStyle` (entry under the pointer). Ignored in styled mode. */
  menuItemHoverStyle?: MenuCssOptions;
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
  /** Item text for this chart; wins over the install `menuText` and `lang`. */
  menuText?: string;
  /** Icon for this chart; overrides the install `menuIcon` (`null` removes it). */
  menuIcon?: MenuIconOption;
}

export interface Installation {
  /**
   * Removes the menu item definition and the global default menu entry, and restores every global
   * option the install changed (`lang[langKey]` when it added it, `lang.contextButtonTitle`,
   * the context button, `navigation` menu styles, the custom symbol). Charts created later are
   * unaffected; charts already rendered keep their menu.
   */
  uninstall(): void;
  readonly options: Readonly<InstallOptions>;
  /** The `lang` key the item text is read from. */
  readonly langKey: string;
  /** The key in `exporting.menuItemDefinitions` / `menuItems`. */
  readonly menuItemKey: string;
}

export type ExportErrorCode =
  | 'INVALID_CHART'
  | 'CHART_NOT_EDITABLE'
  | 'EXPORTING_MODULE_MISSING'
  | 'BROWSER_REQUIRED'
  | 'WRITER_FAILURE'
  | 'INVALID_OPTIONS'
  /** `signal` was aborted; `error.cause` is `signal.reason`. */
  | 'ABORTED';

export interface ExportErrorDetails {
  chartId?: string | null;
  /** The option or chart property at fault (INVALID_OPTIONS names the option here). */
  property?: string;
  diagnostics?: Diagnostic[];
}

/**
 * Error thrown (or rejected) by the export API. The underlying error, when there is one, is the
 * standard `error.cause` (for example the writer's error for WRITER_FAILURE).
 */
export class ExportError extends Error {
  override readonly name = 'HighchartsExcelExportError';
  constructor(
    readonly code: ExportErrorCode,
    message: string,
    readonly details: ExportErrorDetails = {},
    options?: { cause?: unknown },
  ) {
    super(message, options?.cause !== undefined ? { cause: options.cause } : undefined);
  }
}

export const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' as const;

export const DEFAULT_MENU_TEXT = 'Download editable Excel chart';
export const DEFAULT_MENU_ITEM_KEY = 'downloadEditableXLSX';
/** Default `InstallOptions.langKey`: the `lang` entry holding the menu item text. */
export const DEFAULT_LANG_KEY = 'downloadEditableXLSX';
/** Renderer symbol name registered for `button.svgPath` when `button.symbol` is not given. */
export const DEFAULT_BUTTON_SYMBOL = 'editableExcelButton';
/**
 * The built-in `menuIcon: 'excel'` glyph: a 14x14 spreadsheet (frame, header row, first column) in
 * `currentColor`. Pixel coordinates and no `viewBox`, which Highcharts' AST does not allow.
 */
export const MENU_ICON_EXCEL_SVG =
  '<svg class="hc-excel-menu-icon" width="14" height="14" aria-hidden="true" ' +
  'style="vertical-align:-2px;margin-right:6px">' +
  '<path d="M2.5 1.5H11.5A1 1 0 0 1 12.5 2.5V11.5A1 1 0 0 1 11.5 12.5H2.5A1 1 0 0 1 1.5 11.5V2.5A1 1 0 0 1 2.5 1.5Z" ' +
  'fill="none" stroke="currentColor" stroke-width="1.3"/>' +
  '<path d="M1.5 5H12.5M1.5 8.75H12.5M5.25 5V12.5M8.75 5V12.5" fill="none" stroke="currentColor" stroke-width="1"/>' +
  '<path d="M2 2H12V5H2Z" fill="currentColor" opacity="0.35"/>' +
  '</svg>';

const FIDELITY_MODES: readonly FidelityMode[] = ['best-effort', 'minimal'];
const DATA_MODES: readonly DataMode[] = ['rendered', 'raw'];
const VISIBILITY_MODES: readonly SeriesVisibilityMode[] = ['visible', 'all'];
/** Bounds for `chartWidth` / `chartHeight`, in CSS pixels. */
export const MIN_CHART_SIZE_PX = 50;
export const MAX_CHART_SIZE_PX = 20_000;

function invalid(property: string, expected: string, value: unknown): ExportError {
  const shown =
    typeof value === 'string' ? JSON.stringify(value) : typeof value === 'number' ? String(value) : typeof value;
  return new ExportError(
    'INVALID_OPTIONS',
    `Invalid export option "${property}": expected ${expected}, got ${shown}.`,
    { property },
  );
}

function checkLiteral(options: Record<string, unknown>, property: string, allowed: readonly string[]): void {
  const v = options[property];
  if (v !== undefined && !allowed.includes(v as string))
    throw invalid(property, allowed.map((a) => `'${a}'`).join(' or '), v);
}

function checkString(options: Record<string, unknown>, property: string): void {
  const v = options[property];
  if (v !== undefined && typeof v !== 'string') throw invalid(property, 'a string', v);
}

function checkFunction(value: unknown, property: string): void {
  if (value !== undefined && typeof value !== 'function') throw invalid(property, 'a function', value);
}

function checkSize(options: Record<string, unknown>, property: string): void {
  const v = options[property];
  if (v === undefined) return;
  if (typeof v !== 'number' || !Number.isFinite(v) || v < MIN_CHART_SIZE_PX || v > MAX_CHART_SIZE_PX) {
    throw invalid(property, `a finite number of CSS pixels from ${MIN_CHART_SIZE_PX} to ${MAX_CHART_SIZE_PX}`, v);
  }
}

/**
 * Validates the options a JavaScript caller can get wrong without a type checker.
 *
 * @throws ExportError INVALID_OPTIONS naming the option (also in `details.property`).
 */
function validateExportOptions(options: unknown): asserts options is ExportOptions {
  if (typeof options !== 'object' || options === null || Array.isArray(options)) {
    throw new ExportError('INVALID_OPTIONS', 'Export options must be an object.', { property: 'options' });
  }
  const o = options as Record<string, unknown>;
  checkLiteral(o, 'fidelity', FIDELITY_MODES);
  checkLiteral(o, 'dataMode', DATA_MODES);
  checkLiteral(o, 'seriesVisibility', VISIBILITY_MODES);
  checkSize(o, 'chartWidth');
  checkSize(o, 'chartHeight');
  checkString(o, 'chartSheetName');
  checkString(o, 'dataSheetName');
  checkString(o, 'filename');
  checkFunction(o.onWarning, 'onWarning');
  if (o.hooks !== undefined) {
    if (typeof o.hooks !== 'object' || o.hooks === null) throw invalid('hooks', 'an object', o.hooks);
    checkFunction((o.hooks as Record<string, unknown>).transformModel, 'hooks.transformModel');
  }
  if (o.writer !== undefined) {
    const w = o.writer as { write?: unknown } | null;
    if (typeof w !== 'object' || w === null || typeof w.write !== 'function') {
      throw invalid('writer', 'an object with a write(spec) function', o.writer);
    }
  }
}

/**
 * Fills defaults after validating the options.
 *
 * @throws ExportError INVALID_OPTIONS when an option has the wrong type or value.
 */
export function resolveExportOptions(
  options: ExportOptions = {},
): Required<
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
  validateExportOptions(options);
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
