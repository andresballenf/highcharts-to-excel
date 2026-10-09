/**
 * `highcharts-editable-excel/internals` — EXPERIMENTAL lower-level building blocks.
 *
 * Everything exported here (the pipeline steps, the diagnostics helpers, the default writer and the
 * IR / workbook-spec types) is unstable: it may change in any minor version. Pin an exact package
 * version when you depend on it. The stable surface is the main entry, `highcharts-editable-excel`.
 *
 * Pipeline: `extractChartModel` / `extractChartModelFromOptions` → `ChartModel` →
 * `applyThemeOverrides` → `translateChartModel` → `WorkbookSpec` → any `ExcelWriter`
 * (`createDefaultExcelWriter()` is the built-in one).
 *
 * @experimental
 * @packageDocumentation
 */

// Pipeline steps
export { extractChartModel, extractChartModelFromOptions } from './highcharts/extract-chart';
export { applyThemeOverrides } from './core/theme-overrides';
export { resolveChartType } from './core/chart-type-registry';
export { translateChartModel } from './core/translate-chart';
export { createDefaultExcelWriter } from './excel/ooxml-writer';
export type { OoxmlWriterOptions, ZipMode } from './excel/ooxml-writer';

// Diagnostics helpers
export { DiagnosticCollector, buildCompatibilityReport, createDiagnostic } from './types/diagnostics';

// Pipeline option / result types
export type { ExtractOptions } from './highcharts/extract-chart';
export type { TranslateOptions, TranslationResult } from './core/translate-chart';
export type { ChartTypeResolution, PlotGroupPlan } from './core/chart-type-registry';

// IR types (also exported as types from the main entry, for `hooks.transformModel`)
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

// Workbook-spec types (`ExcelWriter` is also exported from the main entry, where it is stable)
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
