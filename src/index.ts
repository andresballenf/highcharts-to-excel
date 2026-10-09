/**
 * highcharts-editable-excel — export Highcharts charts to XLSX workbooks containing native,
 * editable Excel charts linked to worksheet data.
 *
 * This entry point has no side effects: nothing is registered until `installHighchartsExcelExport`
 * is called with the application's Highcharts namespace.
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

// Lower-level building blocks
export { CHART_TYPE_MATRIX, resolveChartType } from './core/chart-type-registry';
export { translateChartModel } from './core/translate-chart';
export { createDefaultExcelWriter } from './excel/ooxml-writer';
export { extractChartModel, extractChartModelFromOptions } from './highcharts/extract-chart';

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
  PerChartExportConfig,
  Installation,
  ExportErrorCode,
} from './types/public-api';
export type {
  DiagnosticSeverity,
  TranslationOutcome,
  DiagnosticCode,
  Diagnostic,
  CompatibilityReport,
} from './types/diagnostics';
export type {
  Color,
  Fill,
  DashStyle,
  Stroke,
  Font,
  HorizontalAlign,
  VerticalAlign,
  TextBlock,
  NumberFormat,
  MarkerSymbol,
  MarkerStyle,
  DataLabelPosition,
  DataLabelStyle,
  SeriesKind,
  Stacking,
  PointModel,
  SeriesModel,
  SeriesDataSemantics,
  AxisKind,
  AxisLabelStyle,
  AxisModel,
  LegendPosition,
  LegendModel,
  PlotAreaModel,
  ChartMeta,
  ChartModel,
} from './types/chart-model';
export type {
  ExcelWriter,
  WorkbookSpec,
  SheetSpec,
  RowSpec,
  CellValue,
  CellStyleSpec,
  CellSpec,
  AnchorSpec,
  DrawingSpec,
  ExcelChartSpec,
} from './excel/writer-interface';
export type { ChartTypeMatrixEntry, ChartTypeResolution, PlotGroupPlan } from './core/chart-type-registry';
export type { TranslateOptions, TranslationResult } from './core/translate-chart';
export type { ExtractOptions } from './highcharts/extract-chart';

export { DiagnosticCollector, buildCompatibilityReport, createDiagnostic } from './types/diagnostics';
