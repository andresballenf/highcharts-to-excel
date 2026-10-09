/**
 * `highcharts-editable-excel/augment` — opt-in TypeScript augmentation.
 *
 * `import 'highcharts-editable-excel/augment';` once in your app (next to your Highcharts import)
 * lets `Highcharts.Options` accept the per-chart `exporting.editableExcel` block read by the
 * library. The runtime module is empty. The type-only import below makes this file a module, so the
 * `declare module` block augments Highcharts instead of replacing its declarations.
 */
/// <reference types="highcharts" />
import type { PerChartExportConfig } from './types/public-api';

declare module 'highcharts' {
  interface ExportingOptions {
    /** Per-chart settings for the "Download editable Excel chart" menu item and its export. */
    editableExcel?: PerChartExportConfig;
  }
}
