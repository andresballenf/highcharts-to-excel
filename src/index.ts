/**
 * highcharts-editable-excel — export Highcharts charts to XLSX workbooks containing native,
 * editable Excel charts linked to worksheet data.
 *
 * This entry point has no side effects: nothing is registered until `installHighchartsExcelExport`
 * is called with the application's Highcharts namespace.
 *
 * Stability: the runtime exports and the option / result / diagnostic types of this entry follow
 * semver. The `ChartModel` IR types are experimental (see README "Stability and versioning").
 */

// Programmatic API
export {
  exportHighchartsToXlsx,
  exportHighchartsOptionsToXlsx,
  analyzeChartCompatibility,
  exportChartsToWorkbook,
} from './api/export';
export { downloadHighchartsAsXlsx, triggerDownload } from './browser/download';
export { installHighchartsExcelExport, addEditableExcelMenuItem } from './highcharts/install-export-menu';

// Errors and constants
export { ExportError, XLSX_MIME_TYPE, DEFAULT_MENU_TEXT, DEFAULT_MENU_ITEM_KEY } from './types/public-api';
export { DEFAULT_LANG_KEY, DEFAULT_BUTTON_SYMBOL, MENU_ICON_EXCEL_SVG } from './types/public-api';

export { CHART_TYPE_MATRIX } from './core/chart-type-registry';

/*
 * The experimental pipeline (`extractChartModel`, `extractChartModelFromOptions`,
 * `applyThemeOverrides`, `translateChartModel`, `resolveChartType`, `createDefaultExcelWriter`,
 * `DiagnosticCollector`, `createDiagnostic`, `buildCompatibilityReport`) and the workbook-spec types
 * live in `highcharts-editable-excel/internals`. The TypeScript augmentation for
 * `exporting.editableExcel` is `highcharts-editable-excel/augment`.
 */

// Types
export type {
  FidelityMode,
  DataMode,
  SeriesVisibilityMode,
  FontOverride,
  ThemeOverrides,
  ExportHooks,
  ExportOptions,
  ExportTimings,
  ExportResult,
  MultiChartExportEntry,
  MultiChartExportResult,
  InstallOptions,
  ContextButtonOptions,
  MenuIconOption,
  MenuCssOptions,
  PerChartExportConfig,
  Installation,
  ExportErrorCode,
  ExportErrorDetails,
  ExportPhase,
  ExportProgress,
} from './types/public-api';
export type {
  DiagnosticSeverity,
  TranslationOutcome,
  DiagnosticCode,
  Diagnostic,
  CompatibilityReport,
} from './types/diagnostics';
/**
 * @experimental The IR types (`ChartModel` and its members) are exported because
 * `hooks.transformModel` receives and returns a `ChartModel`. They may change in minor versions.
 */
export type {
  /** @experimental */ Color,
  /** @experimental */ Fill,
  /** @experimental */ DashStyle,
  /** @experimental */ Stroke,
  /** @experimental */ Font,
  /** @experimental */ HorizontalAlign,
  /** @experimental */ VerticalAlign,
  /** @experimental */ TextBlock,
  /** @experimental */ NumberFormat,
  /** @experimental */ MarkerSymbol,
  /** @experimental */ MarkerStyle,
  /** @experimental */ DataLabelPosition,
  /** @experimental */ DataLabelStyle,
  /** @experimental */ SeriesKind,
  /** @experimental */ Stacking,
  /** @experimental */ PointModel,
  /** @experimental */ SeriesModel,
  /** @experimental */ SeriesDataSemantics,
  /** @experimental */ AxisKind,
  /** @experimental */ AxisLabelStyle,
  /** @experimental */ AxisModel,
  /** @experimental */ LegendPosition,
  /** @experimental */ LegendModel,
  /** @experimental */ PlotAreaModel,
  /** @experimental */ ChartMeta,
  /** @experimental */ ChartModel,
} from './types/chart-model';
/**
 * Stable: the interface of `ExportOptions.writer`. The `WorkbookSpec` it receives is experimental
 * (import it from `highcharts-editable-excel/internals`).
 */
export type { ExcelWriter, WriteContext, WriteProgress } from './excel/writer-interface';
export type { ChartTypeMatrixEntry } from './core/chart-type-registry';
