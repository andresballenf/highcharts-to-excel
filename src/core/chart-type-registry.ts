/**
 * Centralized Highcharts → Excel chart-type mapping.
 *
 * `resolveChartType` decides, from the neutral ChartModel, which Excel plot groups are produced,
 * which series are dropped and which combinations are impossible in Excel.
 */

import type { ChartModel, SeriesKind, SeriesModel, Stacking } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import type { PlotGroupSpec } from '../excel/writer-interface';

export interface PlotGroupPlan {
  kind: PlotGroupSpec['kind'];
  /** Positions in `model.series`. */
  seriesIndices: number[];
  /**
   * Model y-axis index the group is plotted against. Groups never span more than two Excel value
   * axes: the lowest used y-axis index is the primary axis, every other one is collapsed onto the
   * secondary axis (the group then carries the first secondary model index).
   */
  yAxisIndex: number;
  stacking: Stacking;
  barDir?: 'col' | 'bar';
  smooth?: boolean;
}

export interface ChartTypeResolution {
  /** e.g. "line", "stackedColumn", "combo:column+line"; null when blocked. */
  excelChartType: string | null;
  groups: PlotGroupPlan[];
  /** Positions in `model.series` that are not exported to the chart. */
  droppedSeries: number[];
  diagnostics: Diagnostic[];
  blocking: boolean;
}

export interface ChartTypeMatrixEntry {
  highcharts: string;
  excel: string;
  support: 'native' | 'approximated' | 'unsupported';
  notes: string;
}

/** Documentation table (rendered in the README). */
export const CHART_TYPE_MATRIX: ReadonlyArray<ChartTypeMatrixEntry> = Object.freeze([
  { highcharts: 'line', excel: 'line', support: 'native', notes: 'Line chart; stacking maps to stacked/percent-stacked line.' },
  { highcharts: 'spline', excel: 'line (smoothed)', support: 'native', notes: 'Excel smoothing differs slightly from Highcharts splines.' },
  { highcharts: 'area', excel: 'area', support: 'native', notes: 'Stacking and percent stacking supported; fill opacity kept.' },
  { highcharts: 'areaspline', excel: 'area', support: 'approximated', notes: 'Excel area charts cannot be smoothed; drawn with straight segments.' },
  { highcharts: 'column', excel: 'column (clustered/stacked/100%)', support: 'native', notes: 'groupPadding/pointPadding map to gap width/overlap.' },
  { highcharts: 'bar', excel: 'bar (clustered/stacked/100%)', support: 'native', notes: 'Horizontal bars.' },
  { highcharts: 'pie', excel: 'pie', support: 'native', notes: 'One series per pie; slice colors and sliced points (explosion) kept.' },
  { highcharts: 'pie + innerSize', excel: 'doughnut', support: 'native', notes: 'Hole size from innerSize (10-90%); multiple rings share one hole size.' },
  { highcharts: 'scatter', excel: 'scatter', support: 'native', notes: 'X/Y columns per series; marker-only unless lineWidth > 0.' },
  { highcharts: 'bubble', excel: 'bubble', support: 'native', notes: 'X/Y/Size columns per series; cannot be combined with other types.' },
  { highcharts: 'columnrange', excel: '-', support: 'unsupported', notes: 'Excel has no floating range columns without helper series.' },
  { highcharts: 'arearange', excel: '-', support: 'unsupported', notes: 'Excel has no band/range area type.' },
  { highcharts: 'boxplot', excel: '-', support: 'unsupported', notes: 'Excel box & whisker is a chartex type that cannot reference this layout.' },
  { highcharts: 'heatmap', excel: '-', support: 'unsupported', notes: 'Excel has no heatmap chart (only conditional formatting).' },
  { highcharts: 'treemap', excel: '-', support: 'unsupported', notes: 'Excel treemap is a chartex type not produced by this library.' },
  { highcharts: 'waterfall', excel: '-', support: 'unsupported', notes: 'Excel waterfall is a chartex type not produced by this library.' },
  { highcharts: 'funnel', excel: '-', support: 'unsupported', notes: 'Excel funnel is a chartex type not produced by this library.' },
  { highcharts: 'gauge', excel: '-', support: 'unsupported', notes: 'Excel has no gauge chart.' },
  { highcharts: 'polar (any type)', excel: '-', support: 'unsupported', notes: 'Excel radar charts do not match polar/spider geometry.' },
  { highcharts: 'variablepie', excel: '-', support: 'unsupported', notes: 'Excel pies cannot vary slice radius.' },
  { highcharts: 'sankey', excel: '-', support: 'unsupported', notes: 'Excel has no flow diagrams.' },
  { highcharts: 'networkgraph', excel: '-', support: 'unsupported', notes: 'Excel has no network/graph layout chart.' },
  { highcharts: 'timeline', excel: '-', support: 'unsupported', notes: 'Excel has no timeline chart type.' },
  { highcharts: 'histogram', excel: '-', support: 'unsupported', notes: 'Excel histogram is a chartex type computed from raw data, not from this series.' },
  { highcharts: 'bellcurve', excel: '-', support: 'unsupported', notes: 'Derived series (computed in the browser) with no Excel equivalent.' },
  { highcharts: 'errorbar', excel: '-', support: 'unsupported', notes: 'Excel error bars are attached to another series, not standalone series.' },
  { highcharts: 'lollipop', excel: '-', support: 'unsupported', notes: 'No Excel equivalent without helper series and error bars.' },
  { highcharts: 'dumbbell', excel: '-', support: 'unsupported', notes: 'No Excel equivalent without helper series and high-low lines.' },
]);

type MappedKind = Exclude<SeriesKind, 'unknown'>;

const KIND_TO_GROUP: Readonly<Record<MappedKind, PlotGroupSpec['kind']>> = {
  line: 'line',
  spline: 'line',
  area: 'area',
  areaspline: 'area',
  column: 'bar',
  bar: 'bar',
  pie: 'pie',
  doughnut: 'doughnut',
  scatter: 'scatter',
  bubble: 'bubble',
};

const CATEGORY_GROUP_KINDS: ReadonlySet<PlotGroupSpec['kind']> = new Set(['bar', 'line', 'area']);

/** Excel chart type name for one plot group. */
export function groupTypeName(g: Pick<PlotGroupPlan, 'kind' | 'stacking' | 'barDir'>): string {
  const prefix = g.stacking === 'percent' ? 'percentStacked' : g.stacking === 'normal' ? 'stacked' : '';
  const withPrefix = (base: string): string => (prefix ? prefix + base[0]!.toUpperCase() + base.slice(1) : base);
  switch (g.kind) {
    case 'bar':
      return withPrefix(g.barDir === 'bar' ? 'bar' : 'column');
    case 'line':
    case 'area':
      return withPrefix(g.kind);
    default:
      return g.kind;
  }
}

function seriesPath(s: SeriesModel): string {
  return `series[${s.index}]`;
}

/** Resolves the Excel plot groups for a model. Never mutates the model. */
export function resolveChartType(model: ChartModel): ChartTypeResolution {
  const diagnostics: Diagnostic[] = [];
  const dropped = new Set<number>();
  const blocked = (): ChartTypeResolution => ({
    excelChartType: null,
    groups: [],
    droppedSeries: model.series.map((_, i) => i),
    diagnostics,
    blocking: true,
  });

  if (model.polar) {
    diagnostics.push(
      createDiagnostic('UNSUPPORTED_POLAR', 'blocking', 'chart.polar', 'Polar/spider charts have no editable Excel equivalent (Excel radar charts use different geometry).'),
    );
    return blocked();
  }

  if (model.series.length === 0) {
    diagnostics.push(createDiagnostic('EMPTY_CHART', 'blocking', 'series', 'The chart has no series to export.'));
    return blocked();
  }

  // 1. Map series kinds; drop unknown ones.
  interface Candidate {
    pos: number;
    s: SeriesModel;
    kind: PlotGroupSpec['kind'];
    barDir?: 'col' | 'bar';
  }
  const candidates: Candidate[] = [];
  model.series.forEach((s, pos) => {
    if (s.kind === 'unknown') {
      dropped.add(pos);
      diagnostics.push(
        createDiagnostic(
          'UNSUPPORTED_SERIES_TYPE',
          'unsupported',
          `${seriesPath(s)}.type`,
          `Series type "${s.sourceType}" has no Excel chart equivalent; the series is not exported to the chart.`,
          { seriesIndex: s.index, details: { sourceType: s.sourceType } },
        ),
      );
      return;
    }
    const kind = KIND_TO_GROUP[s.kind];
    const c: Candidate = { pos, s, kind };
    if (kind === 'bar') {
      // Highcharts `bar` is always horizontal (it forces chart.inverted); a column on an inverted chart is horizontal too.
      c.barDir = s.kind === 'bar' || model.inverted ? 'bar' : 'col';
    }
    candidates.push(c);
  });

  if (candidates.length === 0) {
    const types = [...new Set(model.series.map((s) => s.sourceType))].join(', ');
    diagnostics.push(
      createDiagnostic('UNSUPPORTED_CHART_TYPE', 'blocking', 'chart.type', `No series can be drawn as a native Excel chart (source types: ${types}).`, {
        details: { sourceTypes: types },
      }),
    );
    return blocked();
  }

  // 2. Pie family and bubble cannot be combined with anything else.
  const pieFamily = candidates.filter((c) => c.kind === 'pie' || c.kind === 'doughnut');
  const others = candidates.filter((c) => c.kind !== 'pie' && c.kind !== 'doughnut');
  if (pieFamily.length > 0 && others.length > 0) {
    diagnostics.push(
      createDiagnostic('UNSUPPORTED_CHART_TYPE', 'blocking', 'chart.type', 'Excel cannot combine pie with other chart types.', {
        details: { sourceTypes: [...new Set(candidates.map((c) => c.s.sourceType))] },
      }),
    );
    return blocked();
  }
  const bubbles = candidates.filter((c) => c.kind === 'bubble');
  if (bubbles.length > 0 && bubbles.length !== candidates.length) {
    diagnostics.push(
      createDiagnostic('UNSUPPORTED_CHART_TYPE', 'blocking', 'chart.type', 'Excel cannot combine bubble with other chart types.', {
        details: { sourceTypes: [...new Set(candidates.map((c) => c.s.sourceType))] },
      }),
    );
    return blocked();
  }

  let kept = candidates;
  if (pieFamily.length > 0) {
    const hasDoughnut = pieFamily.some((c) => c.kind === 'doughnut');
    const hasPie = pieFamily.some((c) => c.kind === 'pie');
    if (hasDoughnut && hasPie) {
      // Classic "pie with a donut ring around it": Excel draws all rings as a doughnut.
      for (const c of pieFamily) {
        if (c.kind === 'pie') {
          diagnostics.push(
            createDiagnostic('APPROXIMATED_CHART_TYPE', 'approximated', `${seriesPath(c.s)}.type`, 'Pie series combined with a doughnut is drawn as the inner doughnut ring.', {
              seriesIndex: c.s.index,
            }),
          );
          c.kind = 'doughnut';
        }
      }
    } else if (hasPie && pieFamily.length > 1) {
      kept = [pieFamily[0]!];
      for (const c of pieFamily.slice(1)) {
        dropped.add(c.pos);
        diagnostics.push(
          createDiagnostic(
            'UNSUPPORTED_SERIES_TYPE',
            'approximated',
            `${seriesPath(c.s)}.type`,
            `An Excel pie chart shows a single series; pie series "${c.s.name}" is not exported to the chart (its data stays on the data sheet).`,
            { seriesIndex: c.s.index },
          ),
        );
      }
    }
  }

  // 3. Inverted line/area/scatter.
  if (model.inverted) {
    for (const c of kept) {
      if (c.kind === 'line' || c.kind === 'area' || c.kind === 'scatter' || c.kind === 'bubble') {
        diagnostics.push(
          createDiagnostic(
            'APPROXIMATED_CHART_TYPE',
            'approximated',
            'chart.inverted',
            'Excel cannot invert line/area charts; the series keeps its normal (vertical) orientation.',
            { seriesIndex: c.s.index },
          ),
        );
      }
    }
  }

  // 4. Y-axis slots: lowest used index is primary, everything else collapses onto the secondary axis.
  const usedY = [...new Set(kept.map((c) => c.s.yAxisIndex))].sort((a, b) => a - b);
  const primaryY = usedY[0] ?? 0;
  const secondaryY = usedY[1];
  for (const extra of usedY.slice(2)) {
    diagnostics.push(
      createDiagnostic(
        'UNSUPPORTED_AXIS_FEATURE',
        'approximated',
        `yAxis[${extra}]`,
        `Excel charts have at most two value axes; series on y axis ${extra} are plotted against the secondary axis.`,
        { details: { yAxisIndex: extra, collapsedOnto: secondaryY } },
      ),
    );
  }
  const slotAxis = (y: number): number => (y === primaryY ? primaryY : (secondaryY ?? primaryY));

  // 5. Group by (kind, axis slot, stacking, barDir) preserving first appearance.
  const groups: PlotGroupPlan[] = [];
  const byKey = new Map<string, PlotGroupPlan>();
  const stackGroupsByKey = new Map<string, Set<string>>();
  for (const c of kept) {
    const y = slotAxis(c.s.yAxisIndex);
    const isPieLike = c.kind === 'pie' || c.kind === 'doughnut';
    const stacking: Stacking = isPieLike || c.kind === 'scatter' || c.kind === 'bubble' ? null : c.s.stacking;
    const key = `${c.kind}|${y}|${stacking ?? ''}|${c.barDir ?? ''}`;
    let g = byKey.get(key);
    if (!g) {
      g = { kind: c.kind, seriesIndices: [], yAxisIndex: y, stacking };
      if (c.barDir) g.barDir = c.barDir;
      if (c.kind === 'line' || c.kind === 'scatter') g.smooth = true;
      byKey.set(key, g);
      groups.push(g);
      stackGroupsByKey.set(key, new Set());
    }
    g.seriesIndices.push(c.pos);
    if (g.smooth !== undefined) g.smooth = g.smooth && c.s.smooth;
    if (stacking) {
      const sg = stackGroupsByKey.get(key)!;
      sg.add(c.s.stackGroup ?? '');
      if (sg.size > 1) {
        diagnostics.push(
          createDiagnostic(
            'APPROXIMATED_CHART_TYPE',
            'approximated',
            `${seriesPath(c.s)}.stack`,
            'Excel stacks all series of one type on an axis into a single stack; separate stack groups are merged.',
            { seriesIndex: c.s.index, details: { stack: c.s.stackGroup } },
          ),
        );
      }
    }
    if (c.s.kind === 'areaspline' && c.s.smooth) {
      diagnostics.push(
        createDiagnostic('APPROXIMATED_CHART_TYPE', 'approximated', `${seriesPath(c.s)}.type`, 'Excel area charts cannot be smoothed; the areaspline is drawn with straight segments.', {
          seriesIndex: c.s.index,
        }),
      );
    }
  }

  const names = [...new Set(groups.map(groupTypeName))];
  const excelChartType = names.length === 1 ? names[0]! : `combo:${names.join('+')}`;
  if (groups.length > 1) {
    const kinds = new Set(groups.map((g) => g.kind));
    if (kinds.size > 1) {
      diagnostics.push(
        createDiagnostic('MIXED_SERIES_TYPES', 'translated', 'series', `Series types are combined into an Excel combo chart (${names.join(' + ')}).`, {
          severity: 'info',
        }),
      );
    }
  }

  return {
    excelChartType,
    groups,
    droppedSeries: [...dropped].sort((a, b) => a - b),
    diagnostics,
    blocking: false,
  };
}

/** True for plot group kinds that use the shared category layout (column A = categories). */
export function isCategoryGroupKind(kind: PlotGroupSpec['kind']): boolean {
  return CATEGORY_GROUP_KINDS.has(kind);
}
