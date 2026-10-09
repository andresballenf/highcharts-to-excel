/**
 * ChartModel (neutral IR) → SheetSpecs with a native, cell-referencing Excel chart.
 *
 * Orchestration only: sheet names, chart-type resolution and the data layout, then the builders
 * (`axis-pairs`/`axis-builder`, `series-builder`, `plot-groups`, `chart-chrome`) sharing one
 * `TranslateContext`.
 */

import type { ChartModel } from '../types/chart-model';
import { createDiagnostic, type DiagnosticCollector } from '../types/diagnostics';
import type { FidelityMode } from '../types/public-api';
import type { ExcelChartSpec, ExcelSeriesSpec, PlotGroupSpec, SheetSpec } from '../excel/writer-interface';
import { sanitizeSheetName } from '../utils/filenames';
import { AXIS_IDS, dataDateFormat, hasTimeOfDay } from './axis-builder';
import { buildAxes } from './axis-pairs';
import { buildChartChrome, buildChartSheet, collectDeletedLegendEntries } from './chart-chrome';
import { resolveChartType, type ChartTypeResolution } from './chart-type-registry';
import { buildDataLayout, type SeriesRange } from './data-layout';
import { buildPlotGroup } from './plot-groups';
import { createSeriesBuilder } from './series-builder';
import { createTranslateContext } from './translate-context';

export { AXIS_IDS, axisDateFormat, baseTimeUnitOf, dataDateFormat } from './axis-builder';

export interface TranslateOptions {
  chartSheetName: string;
  dataSheetName: string;
  includeSourceData: boolean;
  fidelity: FidelityMode;
  diagnostics: DiagnosticCollector;
  chartWidth?: number;
  chartHeight?: number;
  referenceImage?: { png: Uint8Array; widthPx: number; heightPx: number } | null;
  takenSheetNames?: ReadonlySet<string>;
}

export interface TranslationResult {
  /** Chart sheet first, then the data sheet. Empty when blocking. */
  sheets: SheetSpec[];
  chartSheetName: string;
  dataSheetName: string;
  excelChartType: string | null;
  supportedProperties: string[];
  blocking: boolean;
  resolution: ChartTypeResolution;
}

/** Translates a ChartModel into the chart sheet and data sheet specs. Never mutates the model. */
export function translateChartModel(model: ChartModel, opts: TranslateOptions): TranslationResult {
  const ctx = createTranslateContext(model, opts.fidelity);
  const { out, support } = ctx;

  // --- Sheet names --------------------------------------------------------------------------------
  const taken = new Set<string>(opts.takenSheetNames ?? []);
  const sheetName = (requested: string, fallback: string, property: string): string => {
    const r = sanitizeSheetName(requested, taken, fallback);
    if (r.adjusted) {
      out.push(
        createDiagnostic(
          'SHEET_NAME_ADJUSTED',
          'approximated',
          property,
          `Sheet name "${requested}" was changed to "${r.name}" to satisfy Excel's naming rules.`,
          { severity: 'info', details: { requested, used: r.name } },
        ),
      );
    }
    return r.name;
  };
  const chartSheetName = sheetName(opts.chartSheetName, 'Chart', 'chartSheetName');
  taken.add(chartSheetName);
  const dataSheetName = sheetName(opts.dataSheetName, 'Data', 'dataSheetName');

  // --- Chart type -------------------------------------------------------------------------------------
  const resolution = resolveChartType(model);
  out.push(...resolution.diagnostics);
  const blockedResult = (): TranslationResult => {
    ctx.flush(opts.diagnostics);
    return {
      sheets: [],
      chartSheetName,
      dataSheetName,
      excelChartType: null,
      supportedProperties: [],
      blocking: true,
      resolution,
    };
  };
  if (resolution.blocking) return blockedResult();

  // --- Data layout --------------------------------------------------------------------------------------
  const exported = resolution.groups.flatMap((g) => g.seriesIndices.map((pos) => model.series[pos]!));
  const anyTime = exported.some((s) => s.points.some((p) => p.x !== null && Number.isFinite(p.x) && hasTimeOfDay(p.x)));
  const layout = buildDataLayout(model, resolution, {
    sheetName: dataSheetName,
    hidden: !opts.includeSourceData,
    dateFormatCode: (axis, span) => dataDateFormat(axis, span, anyTime),
  });
  out.push(...layout.diagnostics);
  if (layout.blocking) return blockedResult();
  const rangeByPos = new Map<number, SeriesRange>(layout.ranges.map((r) => [r.seriesIndex, r]));

  support('chart.type');
  if (opts.chartWidth === undefined) support('chart.width');
  if (opts.chartHeight === undefined) support('chart.height');

  // --- Axes ----------------------------------------------------------------------------------------------
  const { axes, axisIdsOf, isPieChart } = buildAxes(ctx, resolution.groups, layout.x);

  // --- Series & plot groups ------------------------------------------------------------------------------
  const buildSeries = createSeriesBuilder(ctx, rangeByPos);
  const plotGroups: PlotGroupSpec[] = [];
  /** Series idx whose legend entry is deleted (Highcharts showInLegend: false; pie legends list points). */
  const deletedLegendEntries: number[] = [];
  for (const g of resolution.groups) {
    const axisIds = axisIdsOf.get(g) ?? [AXIS_IDS.primaryCat, AXIS_IDS.primaryVal];
    if (g.range) {
      // Range series: hidden helper series (base, down) stacked with the visible range series.
      const series: ExcelSeriesSpec[] = [];
      for (const pos of g.seriesIndices) {
        const built = buildSeries.range(pos, g);
        series.push(...built.specs);
        deletedLegendEntries.push(...built.hiddenFromLegend);
        const s = model.series[pos]!;
        if (s.showInLegend === false) {
          const main = built.specs.find((x) => !built.hiddenFromLegend.includes(x.idx));
          if (main) deletedLegendEntries.push(main.idx);
          ctx.support(`series[${s.index}].showInLegend`);
        }
      }
      plotGroups.push(buildPlotGroup(ctx, g, series, axisIds));
      continue;
    }
    const series = g.seriesIndices.map((pos) => buildSeries(pos, g));
    collectDeletedLegendEntries(ctx, g, series, deletedLegendEntries);
    plotGroups.push(buildPlotGroup(ctx, g, series, axisIds));
  }

  // --- Title, chart area, plot area, legend, chart sheet ----------------------------------------------------
  // Pies and radars (polar) keep Excel's automatic plot area: their labels sit outside the circle,
  // so the source plot box would push the circle into the title.
  const radarOnly = resolution.groups.every((g) => g.kind === 'radar');
  const { title, textDefaults, chartArea, plotArea, legend } = buildChartChrome(
    ctx,
    isPieChart || radarOnly,
    deletedLegendEntries,
  );
  const chart: ExcelChartSpec = {
    title,
    textDefaults,
    chartArea,
    plotArea,
    plotGroups,
    axes: isPieChart ? [] : axes,
    legend,
    dispBlanksAs: 'gap',
    style: null,
  };
  const chartSheetSpec = buildChartSheet(ctx, chart, {
    name: chartSheetName,
    chartWidth: opts.chartWidth,
    chartHeight: opts.chartHeight,
    referenceImage: opts.referenceImage ?? null,
  });

  ctx.flush(opts.diagnostics);
  return {
    sheets: [chartSheetSpec, layout.sheet],
    chartSheetName,
    dataSheetName,
    excelChartType: resolution.excelChartType,
    supportedProperties: [...new Set(ctx.supported)],
    blocking: false,
    resolution,
  };
}
