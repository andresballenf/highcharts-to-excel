/**
 * Plot-group parameters: one `PlotGroupSpec` per resolved plot group, with its grouping (stacking),
 * bar gap width and overlap derived from Highcharts' paddings, scatter style, bubble scale, pie
 * first-slice angle and doughnut hole size.
 */

import { createDiagnostic } from '../types/diagnostics';
import type { ExcelSeriesSpec, PlotGroupSpec } from '../excel/writer-interface';
import { clamp } from '../utils/units';
import type { PlotGroupPlan } from './chart-type-registry';
import type { TranslateContext } from './translate-context';

/** Excel bar gap width (% of a bar) and overlap (%) reproducing Highcharts' group/point paddings. */
export function barGeometry(
  groupPadding: number,
  pointPadding: number,
  seriesCount: number,
  stacked: boolean,
): { gapWidth: number; overlap: number } {
  const gp = clamp(groupPadding, 0, 1);
  const pp = clamp(pointPadding, 0, 1);
  // Bar geometry, category width = 1, n = bars side by side (1 when stacked).
  // Highcharts: slot = (1 - 2gp) / n, bar b = slot * (1 - 2pp), pp*slot of padding on each side of a bar.
  // Excel: category = n*b + (n-1)*(-overlap/100)*b + (gapWidth/100)*b.
  //   bars inside a cluster are 2pp*slot apart → overlap = -100 * 2pp / (1 - 2pp)
  //   between clusters lie 2gp + 2pp*slot      → gapWidth = 100 * (2gp*n + 2pp*(1 - 2gp)) / ((1 - 2gp)(1 - 2pp))
  // (defaults gp 0.2, pp 0.1: n=1 → 108, n=2 → 192, overlap -25).
  const n = stacked ? 1 : Math.max(1, seriesCount);
  const denom = (1 - 2 * gp) * (1 - 2 * pp);
  const gapWidth = denom <= 0 ? 500 : clamp(Math.round((100 * (2 * gp * n + 2 * pp * (1 - 2 * gp))) / denom), 0, 500);
  const overlap = stacked ? 100 : pp >= 0.5 ? -100 : clamp(Math.round((-100 * 2 * pp) / (1 - 2 * pp)), -100, 0) || 0; // || 0: no -0
  return { gapWidth, overlap };
}

/** Builds the plot group spec of `g` around its already built series. */
export function buildPlotGroup(
  ctx: TranslateContext,
  g: PlotGroupPlan,
  series: ExcelSeriesSpec[],
  axisIds: [number, number],
): PlotGroupSpec {
  const { model, support } = ctx;
  const first = model.series[g.seriesIndices[0]!]!;
  switch (g.kind) {
    case 'bar': {
      const bars = first.bars ?? { pointPadding: 0.1, groupPadding: 0.2, borderRadius: 0 };
      const { gapWidth, overlap } = barGeometry(
        bars.groupPadding,
        bars.pointPadding,
        g.seriesIndices.length,
        Boolean(g.stacking),
      );
      if (first.bars) support(`series[${first.index}].groupPadding`, `series[${first.index}].pointPadding`);
      return {
        kind: 'bar',
        barDir: g.barDir ?? 'col',
        grouping: g.stacking === 'percent' ? 'percentStacked' : g.stacking === 'normal' ? 'stacked' : 'clustered',
        gapWidth,
        overlap,
        varyColors: false,
        series,
        axisIds,
        dataLabels: null,
      };
    }
    case 'line':
    case 'area': {
      const grouping = g.stacking === 'percent' ? 'percentStacked' : g.stacking === 'normal' ? 'stacked' : 'standard';
      if (g.kind === 'line') {
        return { kind: 'line', grouping, varyColors: false, showMarkers: true, series, axisIds, dataLabels: null };
      }
      return { kind: 'area', grouping, varyColors: false, series, axisIds, dataLabels: null };
    }
    case 'scatter': {
      const lineW = first.line?.width ?? 0;
      const scatterStyle = !(lineW > 0) ? 'marker' : first.smooth ? 'smoothMarker' : 'lineMarker';
      return { kind: 'scatter', scatterStyle, varyColors: false, series, axisIds, dataLabels: null };
    }
    case 'bubble':
      return { kind: 'bubble', varyColors: false, bubbleScale: 100, series, axisIds, dataLabels: null };
    case 'radar':
      support('chart.polar');
      return {
        kind: 'radar',
        radarStyle: g.radarStyle ?? 'marker',
        varyColors: false,
        series,
        axisIds,
        dataLabels: null,
      };
    case 'pie':
    case 'doughnut': {
      const start = first.pie?.startAngle ?? 0;
      const firstSliceAngle = Math.round(((start % 360) + 360) % 360) % 360;
      if (first.pie) support(`series[${first.index}].startAngle`);
      const end = first.pie?.endAngle ?? null;
      if (end !== null && Math.abs(end - start - 360) > 0.5 && Math.abs(end - start) > 0.5) {
        ctx.out.push(
          createDiagnostic(
            'APPROXIMATED_CHART_TYPE',
            'approximated',
            `series[${first.index}].endAngle`,
            'Excel pies are always full circles; the partial (semi-circle) pie is drawn as a full circle.',
            { seriesIndex: first.index },
          ),
        );
      }
      if (g.kind === 'pie') return { kind: 'pie', varyColors: true, firstSliceAngle, series, dataLabels: null };
      const holeSize = clamp(Math.round((first.pie?.innerSize ?? 0.5) * 100), 10, 90);
      if (first.pie) support(`series[${first.index}].innerSize`);
      if (g.seriesIndices.length > 1) {
        ctx.out.push(
          createDiagnostic(
            'APPROXIMATED_LAYOUT',
            'approximated',
            `series[${model.series[g.seriesIndices[1]!]!.index}].size`,
            'Excel doughnut rings all have the same thickness; ring sizes are approximated.',
            { severity: 'info' },
          ),
        );
      }
      return { kind: 'doughnut', varyColors: true, firstSliceAngle, holeSize, series, dataLabels: null };
    }
  }
}
