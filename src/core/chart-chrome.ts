/**
 * Chart chrome: the title (with the subtitle merged in as a second line), the default text font,
 * chart and plot area fills/borders, the pinned plot-area manual layout, the legend (including the
 * deleted entries of series hidden from it), and the chart sheet itself: the title row, the chart
 * anchor and the placement of the optional reference image.
 */

import type {
  ExcelChartSpec,
  ExcelSeriesSpec,
  ExcelShapeStyle,
  ExcelTextSpec,
  SheetSpec,
} from '../excel/writer-interface';
import { createDiagnostic } from '../types/diagnostics';
import { stripControlChars } from '../utils/text';
import { clamp } from '../utils/units';
import { DEFAULT_HIGHCHARTS_FONT } from '../translators/typography-translator';
import type { PlotGroupPlan } from './chart-type-registry';
import { legendPosition, toExcelFill, toExcelFont, toExcelLine } from './style-mapping';
import { textOf, type TranslateContext } from './translate-context';

/** Default Excel column width in pixels (8.43 characters of Calibri 11). */
const DEFAULT_COLUMN_WIDTH_PX = 64;

/** Chart-level parts of the `ExcelChartSpec` (everything but plot groups and axes). */
export type ChartChrome = Pick<ExcelChartSpec, 'title' | 'textDefaults' | 'chartArea' | 'plotArea' | 'legend'>;

/**
 * Appends to `into` the idx of every series of `g` hidden from the legend (Highcharts
 * `showInLegend: false`). Pie legends list points, so pie/doughnut groups add nothing.
 */
export function collectDeletedLegendEntries(
  ctx: TranslateContext,
  g: PlotGroupPlan,
  series: readonly ExcelSeriesSpec[],
  into: number[],
): void {
  if (g.kind === 'pie' || g.kind === 'doughnut') return;
  g.seriesIndices.forEach((pos, i) => {
    const s = ctx.model.series[pos]!;
    if (s.showInLegend === false) {
      into.push(series[i]!.idx);
      ctx.support(`series[${s.index}].showInLegend`);
    }
  });
}

function buildTitle(ctx: TranslateContext): ExcelTextSpec | null {
  const { model, best, support } = ctx;
  if (model.title && textOf(model.title).length > 0) {
    const subtitleLines = model.subtitle ? textOf(model.subtitle) : [];
    const subtitleFont =
      model.subtitle && subtitleLines.length > 0 && best ? ctx.fontOf(model.subtitle.font, 'subtitle.style') : null;
    const title = ctx.textSpec(
      model.title,
      'title',
      subtitleLines.length > 0 ? { lines: subtitleLines, font: subtitleFont } : null,
    );
    support('title.text');
    if (best) support('title.style');
    if (subtitleLines.length > 0) {
      support('subtitle.text');
      ctx.out.push(
        createDiagnostic(
          'APPROXIMATED_LAYOUT',
          'approximated',
          'subtitle.text',
          'Excel charts have a single title; the subtitle is merged into the title as a second line in its own font.',
          { severity: 'info' },
        ),
      );
    }
    return title;
  }
  if (model.subtitle && textOf(model.subtitle).length > 0) {
    const title = ctx.textSpec(model.subtitle, 'subtitle');
    ctx.out.push(
      createDiagnostic(
        'APPROXIMATED_LAYOUT',
        'approximated',
        'subtitle.text',
        'The chart has no title; the subtitle is used as the Excel chart title.',
        { severity: 'info' },
      ),
    );
    return title;
  }
  return null;
}

function buildPlotArea(ctx: TranslateContext, isPieChart: boolean): ExcelChartSpec['plotArea'] {
  const { model, best, supportStyle } = ctx;
  let manualLayout: ExcelChartSpec['plotArea']['manualLayout'] = null;
  const box = model.plotArea.box;
  if (best && !isPieChart && box && model.width > 0 && model.height > 0) {
    const r = (v: number): number => Math.round(clamp(v, 0, 1) * 10000) / 10000;
    const x = r(box.left / model.width);
    const y = r(box.top / model.height);
    // Keep the pinned box inside the chart: x + w ≤ 1 and y + h ≤ 1.
    const w = r(Math.min(box.width / model.width, 1 - x));
    const h = r(Math.min(box.height / model.height, 1 - y));
    if (w > 0.2 && h > 0.2) {
      manualLayout = { x, y, w, h };
      ctx.out.push(
        createDiagnostic(
          'APPROXIMATED_LAYOUT',
          'approximated',
          'chart.plotArea',
          'Plot area pinned to source proportions; Excel positions titles and labels around it itself.',
          { severity: 'info', details: { ...manualLayout } },
        ),
      );
    }
  }
  const plotArea: ExcelChartSpec['plotArea'] = best
    ? { fill: toExcelFill(model.plotArea.background), line: toExcelLine(model.plotArea.border), manualLayout }
    : { fill: null, line: null, manualLayout: null };
  if (model.plotArea.background) supportStyle('chart.plotBackgroundColor');
  if (model.plotArea.border) supportStyle('chart.plotBorderWidth', 'chart.plotBorderColor');
  return plotArea;
}

function buildLegend(ctx: TranslateContext, deletedEntries: number[]): ExcelChartSpec['legend'] {
  const { model, best, support, supportStyle } = ctx;
  support('legend.enabled');
  if (!model.legend.enabled) return null;
  const lp = legendPosition(model.legend);
  if (lp.diagnostic) ctx.out.push(lp.diagnostic);
  else support('legend.align', 'legend.verticalAlign');
  if (model.legend.overlay) support('legend.floating');
  const legend: ExcelChartSpec['legend'] = {
    position: lp.position,
    overlay: lp.overlay,
    font: ctx.fontOf(model.legend.font, 'legend.itemStyle'),
    fill: best ? toExcelFill(model.legend.background) : null,
    line: best ? toExcelLine(model.legend.border) : null,
    deletedEntries,
  };
  if (model.legend.font) supportStyle('legend.itemStyle');
  if (model.legend.background) supportStyle('legend.backgroundColor');
  if (model.legend.border) supportStyle('legend.borderWidth', 'legend.borderColor');
  if (model.legend.reversed) {
    ctx.out.push(
      createDiagnostic(
        'UNSUPPORTED_STYLE',
        'unsupported',
        'legend.reversed',
        'Excel cannot reverse the legend order.',
        {
          severity: 'info',
        },
      ),
    );
  }
  return legend;
}

/** Title, text defaults, chart area, plot area and legend of the chart. */
export function buildChartChrome(
  ctx: TranslateContext,
  isPieChart: boolean,
  deletedLegendEntries: number[],
): ChartChrome {
  const { model, best, styleSink, supportStyle } = ctx;
  const title = buildTitle(ctx);
  const baseFamily = model.title?.font.family ?? model.legend.font?.family ?? DEFAULT_HIGHCHARTS_FONT.family;
  const textDefaults = best
    ? toExcelFont({ ...DEFAULT_HIGHCHARTS_FONT, family: baseFamily, size: 12 }, 'chart.style.fontFamily', styleSink)
    : null;
  const chartArea: ExcelShapeStyle = best
    ? { fill: toExcelFill(model.background), line: toExcelLine(model.border) }
    : { fill: null, line: null };
  if (model.background) supportStyle('chart.backgroundColor');
  if (model.border) supportStyle('chart.borderWidth', 'chart.borderColor');
  const plotArea = buildPlotArea(ctx, isPieChart);
  const legend = buildLegend(ctx, deletedLegendEntries);
  return { title, textDefaults, chartArea, plotArea, legend };
}

export interface ChartSheetOptions {
  name: string;
  chartWidth: number | undefined;
  chartHeight: number | undefined;
  referenceImage: { png: Uint8Array; widthPx: number; heightPx: number } | null;
}

/** The chart sheet: a bold title row, the chart below it and the reference image to its right. */
export function buildChartSheet(ctx: TranslateContext, chart: ExcelChartSpec, opts: ChartSheetOptions): SheetSpec {
  const { model } = ctx;
  const widthPx = Math.max(50, Math.round(opts.chartWidth ?? model.width));
  const heightPx = Math.max(50, Math.round(opts.chartHeight ?? model.height));
  const drawings: SheetSpec['drawings'] = [
    {
      kind: 'chart',
      chart,
      anchor: { col0: 0, row0: 2, colOffsetPx: 0, rowOffsetPx: 0, widthPx, heightPx },
      name: 'Chart 1',
    },
  ];
  if (opts.referenceImage) {
    const offset = widthPx + 20;
    const col0 = Math.floor(offset / DEFAULT_COLUMN_WIDTH_PX);
    drawings.push({
      kind: 'image',
      png: opts.referenceImage.png,
      anchor: {
        col0,
        row0: 2,
        colOffsetPx: offset - col0 * DEFAULT_COLUMN_WIDTH_PX,
        rowOffsetPx: 0,
        widthPx,
        heightPx,
      },
      name: 'Reference image',
    });
  }
  const title = chart.title;
  const titleText = title ? stripControlChars((model.title ? textOf(model.title) : title.lines).join(' ')) : '';
  return {
    name: opts.name,
    hidden: false,
    columns: [],
    rows: titleText
      ? [{ row0: 0, cells: [{ col0: 0, row0: 0, value: { type: 'string', value: titleText }, style: { bold: true } }] }]
      : [],
    freezeHeaderRow: false,
    drawings,
  };
}
