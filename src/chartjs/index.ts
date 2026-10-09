/**
 * Chart.js adapter: export Chart.js 4 charts (instances or configurations) to XLSX workbooks with
 * native, editable Excel charts. Chart.js itself is never imported (optional peer dependency).
 */

export { analyzeChartJsCompatibility, downloadChartJsAsXlsx, exportChartJsToXlsx } from './export';
export { extractChartJsModel } from './extract-chart';
export { isChartJsChart, isChartJsConfig } from './guards';
export {
  CHARTJS_COLORS_PLUGIN_BACKGROUND,
  CHARTJS_COLORS_PLUGIN_BORDER,
  CHARTJS_DEFAULT_FONT,
} from './defaults';
export { dateDisplayFormatToExcel, intlFormatToExcel } from './number-format';
export type { ChartJsChartLike, ChartJsConfigLike, ChartJsExtractOptions } from './types';
