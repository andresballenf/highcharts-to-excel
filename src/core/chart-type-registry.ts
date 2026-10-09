/**
 * Centralized Highcharts → Excel chart-type mapping.
 *
 * `resolveChartType` decides, from the neutral ChartModel, which Excel plot groups are produced,
 * which series are dropped and which combinations are impossible in Excel.
 */

import type { ChartModel, SeriesKind, SeriesModel, Stacking } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import { EXCEL_MAX_SERIES_PER_CHART, type ExcelRadarStyle, type PlotGroupSpec } from '../excel/writer-interface';

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
  /** Radar groups (polar charts) only. */
  radarStyle?: ExcelRadarStyle;
  /**
   * A range series (columnrange/arearange) drawn as a stacked group with a hidden base series. Range
   * groups hold exactly one model series.
   */
  range?: boolean;
}

export interface ChartTypeResolution {
  /** e.g. "line", "stackedColumn", "combo:column+line"; null when blocked. */
  excelChartType: string | null;
  groups: PlotGroupPlan[];
  /** Positions in `model.series` that are not exported to the chart (error bars attached to a parent are not). */
  droppedSeries: number[];
  /** Error bar series (positions) drawn as Excel error bars on their parent series (positions). */
  errorBars: Array<{ parent: number; errorBar: number }>;
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
  {
    highcharts: 'line',
    excel: 'line',
    support: 'native',
    notes: 'Line chart; stacking maps to stacked/percent-stacked line.',
  },
  {
    highcharts: 'spline',
    excel: 'line (smoothed)',
    support: 'native',
    notes: 'Excel smoothing differs slightly from Highcharts splines.',
  },
  {
    highcharts: 'area',
    excel: 'area',
    support: 'native',
    notes: 'Stacking and percent stacking supported; fill opacity kept.',
  },
  {
    highcharts: 'areaspline',
    excel: 'area',
    support: 'approximated',
    notes: 'Excel area charts cannot be smoothed; drawn with straight segments.',
  },
  {
    highcharts: 'column',
    excel: 'column (clustered/stacked/100%)',
    support: 'native',
    notes: 'groupPadding/pointPadding map to gap width/overlap.',
  },
  { highcharts: 'bar', excel: 'bar (clustered/stacked/100%)', support: 'native', notes: 'Horizontal bars.' },
  {
    highcharts: 'pie',
    excel: 'pie',
    support: 'native',
    notes: 'One series per pie; slice colors and sliced points (explosion) kept.',
  },
  {
    highcharts: 'pie + innerSize',
    excel: 'doughnut',
    support: 'native',
    notes: 'Hole size from innerSize (10-90%); multiple rings share one hole size.',
  },
  {
    highcharts: 'scatter',
    excel: 'scatter',
    support: 'native',
    notes: 'X/Y columns per series; marker-only unless lineWidth > 0.',
  },
  {
    highcharts: 'bubble',
    excel: 'bubble',
    support: 'native',
    notes: 'X/Y/Size columns per series; cannot be combined with other types.',
  },
  {
    highcharts: 'columnrange',
    excel: 'stacked column (hidden base)',
    support: 'approximated',
    notes:
      'Floating bars: a hidden base series stacked under a Range column (=High-Low formula cells, so editing Low/High updates the chart). Ranges crossing zero use Base/Up/Down helper columns. One range series per axis.',
  },
  {
    highcharts: 'arearange',
    excel: 'stacked area (hidden base)',
    support: 'approximated',
    notes: 'A hidden Low area with a Range area (=High-Low formula cells) stacked on top. One range series per axis.',
  },
  {
    highcharts: 'boxplot',
    excel: '-',
    support: 'unsupported',
    notes: 'Excel box & whisker is a chartex type that cannot reference this layout.',
  },
  {
    highcharts: 'heatmap',
    excel: '-',
    support: 'unsupported',
    notes: 'Excel has no heatmap chart (only conditional formatting).',
  },
  {
    highcharts: 'treemap',
    excel: '-',
    support: 'unsupported',
    notes: 'Excel treemap is a chartex type not produced by this library.',
  },
  {
    highcharts: 'waterfall',
    excel: '-',
    support: 'unsupported',
    notes: 'Excel waterfall is a chartex type not produced by this library.',
  },
  {
    highcharts: 'funnel',
    excel: '-',
    support: 'unsupported',
    notes: 'Excel funnel is a chartex type not produced by this library.',
  },
  { highcharts: 'gauge', excel: '-', support: 'unsupported', notes: 'Excel has no gauge chart.' },
  {
    highcharts: 'polar (line/spline/area)',
    excel: 'radar (marker/standard/filled)',
    support: 'approximated',
    notes:
      'Categories are spaced evenly around the circle and lines are straight. Polar columns/bars and other types stay blocking (Excel has no polar columns).',
  },
  { highcharts: 'variablepie', excel: '-', support: 'unsupported', notes: 'Excel pies cannot vary slice radius.' },
  { highcharts: 'sankey', excel: '-', support: 'unsupported', notes: 'Excel has no flow diagrams.' },
  { highcharts: 'networkgraph', excel: '-', support: 'unsupported', notes: 'Excel has no network/graph layout chart.' },
  { highcharts: 'timeline', excel: '-', support: 'unsupported', notes: 'Excel has no timeline chart type.' },
  {
    highcharts: 'histogram',
    excel: '-',
    support: 'unsupported',
    notes: 'Excel histogram is a chartex type computed from raw data, not from this series.',
  },
  {
    highcharts: 'bellcurve',
    excel: '-',
    support: 'unsupported',
    notes: 'Derived series (computed in the browser) with no Excel equivalent.',
  },
  {
    highcharts: 'errorbar',
    excel: 'error bars (custom)',
    support: 'approximated',
    notes:
      'Drawn as custom Excel error bars on the linked parent series (bar, line, area, scatter, bubble); +err/-err columns on the data sheet. Unlinked error bars are not exported.',
  },
  {
    highcharts: 'lollipop',
    excel: '-',
    support: 'unsupported',
    notes: 'No Excel equivalent without helper series and error bars.',
  },
  {
    highcharts: 'dumbbell',
    excel: '-',
    support: 'unsupported',
    notes: 'No Excel equivalent without helper series and high-low lines.',
  },
]);

type MappedKind = Exclude<SeriesKind, 'unknown' | 'errorbar'>;

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
  columnrange: 'bar',
  arearange: 'area',
};

const CATEGORY_GROUP_KINDS: ReadonlySet<PlotGroupSpec['kind']> = new Set(['bar', 'line', 'area', 'radar']);

/** Group kinds whose series can carry Excel error bars. */
const ERROR_BAR_GROUP_KINDS: ReadonlySet<PlotGroupSpec['kind']> = new Set(['bar', 'line', 'area', 'scatter', 'bubble']);

/** Range series kinds (points carry low/high). */
export function isRangeKind(kind: SeriesKind): kind is 'columnrange' | 'arearange' {
  return kind === 'columnrange' || kind === 'arearange';
}

/**
 * A column range with a negative low cannot be drawn as "hidden Low + visible High-Low" (Excel
 * stacks negative values separately below the axis): it needs Base/Up/Down helper series.
 */
export function rangeNeedsSplit(s: SeriesModel): boolean {
  return s.kind === 'columnrange' && s.points.some((p) => typeof p.low === 'number' && p.low < 0);
}

/** Excel series a model series becomes (range series add hidden helpers). */
export function excelSeriesCount(s: SeriesModel): number {
  if (!isRangeKind(s.kind)) return 1;
  return rangeNeedsSplit(s) ? 3 : 2;
}

/** Excel chart type name for one plot group. */
export function groupTypeName(g: Pick<PlotGroupPlan, 'kind' | 'stacking' | 'barDir' | 'radarStyle'>): string {
  const prefix = g.stacking === 'percent' ? 'percentStacked' : g.stacking === 'normal' ? 'stacked' : '';
  const withPrefix = (base: string): string => (prefix ? prefix + base[0]!.toUpperCase() + base.slice(1) : base);
  switch (g.kind) {
    case 'bar':
      return withPrefix(g.barDir === 'bar' ? 'bar' : 'column');
    case 'line':
    case 'area':
      return withPrefix(g.kind);
    case 'radar':
      return g.radarStyle === 'filled' ? 'filledRadar' : 'radar';
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
    errorBars: [],
    diagnostics,
    blocking: true,
  });

  if (model.series.length === 0) {
    diagnostics.push(createDiagnostic('EMPTY_CHART', 'blocking', 'series', 'The chart has no series to export.'));
    return blocked();
  }

  if (model.polar) return resolvePolar(model, diagnostics, blocked);

  // 1. Map series kinds; drop unknown ones. Error bars are attached to their parent at the end.
  interface Candidate {
    pos: number;
    s: SeriesModel;
    kind: PlotGroupSpec['kind'];
    barDir?: 'col' | 'bar';
    range?: boolean;
  }
  const candidates: Candidate[] = [];
  const errorBarSeries: Array<{ pos: number; s: SeriesModel }> = [];
  model.series.forEach((s, pos) => {
    if (s.kind === 'unknown') {
      dropped.add(pos);
      diagnostics.push(unknownSeriesDiagnostic(s));
      return;
    }
    if (s.kind === 'errorbar') {
      errorBarSeries.push({ pos, s });
      return;
    }
    const kind = KIND_TO_GROUP[s.kind];
    const c: Candidate = { pos, s, kind };
    if (kind === 'bar') {
      // Highcharts `bar` is always horizontal (it forces chart.inverted); a column on an inverted chart is horizontal too.
      c.barDir = s.kind === 'bar' || model.inverted ? 'bar' : 'col';
    }
    if (isRangeKind(s.kind)) c.range = true;
    candidates.push(c);
  });

  if (candidates.length === 0) {
    const types = [...new Set(model.series.map((s) => s.sourceType))].join(', ');
    diagnostics.push(
      createDiagnostic(
        'UNSUPPORTED_CHART_TYPE',
        'blocking',
        'chart.type',
        `No series can be drawn as a native Excel chart (source types: ${types}).`,
        {
          details: { sourceTypes: types },
        },
      ),
    );
    return blocked();
  }

  // 2. Pie family and bubble cannot be combined with anything else.
  const pieFamily = candidates.filter((c) => c.kind === 'pie' || c.kind === 'doughnut');
  const others = candidates.filter((c) => c.kind !== 'pie' && c.kind !== 'doughnut');
  if (pieFamily.length > 0 && others.length > 0) {
    diagnostics.push(
      createDiagnostic(
        'UNSUPPORTED_CHART_TYPE',
        'blocking',
        'chart.type',
        'Excel cannot combine pie with other chart types.',
        {
          details: { sourceTypes: [...new Set(candidates.map((c) => c.s.sourceType))] },
        },
      ),
    );
    return blocked();
  }
  const bubbles = candidates.filter((c) => c.kind === 'bubble');
  if (bubbles.length > 0 && bubbles.length !== candidates.length) {
    diagnostics.push(
      createDiagnostic(
        'UNSUPPORTED_CHART_TYPE',
        'blocking',
        'chart.type',
        'Excel cannot combine bubble with other chart types.',
        {
          details: { sourceTypes: [...new Set(candidates.map((c) => c.s.sourceType))] },
        },
      ),
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
            createDiagnostic(
              'APPROXIMATED_CHART_TYPE',
              'approximated',
              `${seriesPath(c.s)}.type`,
              'Pie series combined with a doughnut is drawn as the inner doughnut ring.',
              {
                seriesIndex: c.s.index,
              },
            ),
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
  const keyOf = (c: Candidate): string => `${c.kind}|${slotAxis(c.s.yAxisIndex)}|${c.barDir ?? ''}`;

  // 4b. A range series needs a stacked group of its own (hidden base + range): Excel draws one group
  // per chart type and axis, so it cannot share it with other series of that type.
  const plainKeys = new Set(kept.filter((c) => !c.range).map(keyOf));
  const rangeKeys = new Map<string, Candidate>();
  kept = kept.filter((c) => {
    if (!c.range) return true;
    const key = keyOf(c);
    const other = rangeKeys.get(key);
    if (plainKeys.has(key) || other) {
      dropped.add(c.pos);
      diagnostics.push(
        createDiagnostic(
          'UNSUPPORTED_SERIES_TYPE',
          'unsupported',
          `${seriesPath(c.s)}.type`,
          `Excel draws a range series as its own stacked ${c.kind === 'bar' ? 'column' : 'area'} group, and one such group per axis; range series "${c.s.name}" shares its axis with other ${c.kind === 'bar' ? 'column/bar' : 'area'} series and is not exported to the chart.`,
          { severity: 'warning', seriesIndex: c.s.index },
        ),
      );
      return false;
    }
    rangeKeys.set(key, c);
    diagnostics.push(
      createDiagnostic(
        'APPROXIMATED_CHART_TYPE',
        'approximated',
        `${seriesPath(c.s)}.type`,
        `Range series exported as a stacked ${c.kind === 'bar' ? (c.barDir === 'bar' ? 'bar' : 'column') : 'area'} with a hidden base series.`,
        { severity: 'info', seriesIndex: c.s.index },
      ),
    );
    return true;
  });

  // 5. Group by (kind, axis slot, barDir) preserving first appearance. Excel has one grouping per
  // chart-type group on an axis pair, so series of one kind that differ only by stacking share a group
  // whose stacking is the majority's (ties → stacked).
  const groups: PlotGroupPlan[] = [];
  const byKey = new Map<string, PlotGroupPlan>();
  const membersByGroup = new Map<PlotGroupPlan, Array<{ c: Candidate; stacking: Stacking }>>();
  for (const c of kept) {
    const y = slotAxis(c.s.yAxisIndex);
    const isPieLike = c.kind === 'pie' || c.kind === 'doughnut';
    const stacking: Stacking = c.range
      ? 'normal'
      : isPieLike || c.kind === 'scatter' || c.kind === 'bubble'
        ? null
        : c.s.stacking;
    const key = keyOf(c);
    let g = byKey.get(key);
    if (!g) {
      g = { kind: c.kind, seriesIndices: [], yAxisIndex: y, stacking: null };
      if (c.barDir) g.barDir = c.barDir;
      if (c.range) g.range = true;
      if (c.kind === 'line' || c.kind === 'scatter') g.smooth = true;
      byKey.set(key, g);
      groups.push(g);
      membersByGroup.set(g, []);
    }
    g.seriesIndices.push(c.pos);
    membersByGroup.get(g)!.push({ c, stacking });
    if (g.smooth !== undefined) g.smooth = g.smooth && c.s.smooth;
    if (c.s.kind === 'areaspline' && c.s.smooth) {
      diagnostics.push(
        createDiagnostic(
          'APPROXIMATED_CHART_TYPE',
          'approximated',
          `${seriesPath(c.s)}.type`,
          'Excel area charts cannot be smoothed; the areaspline is drawn with straight segments.',
          {
            seriesIndex: c.s.index,
          },
        ),
      );
    }
  }
  for (const g of groups) {
    const members = membersByGroup.get(g)!;
    g.stacking = majorityStacking(members.map((m) => m.stacking));
    const stackGroups = new Set<string>();
    for (const { c, stacking } of members) {
      if (stacking !== g.stacking) {
        diagnostics.push(
          createDiagnostic(
            'APPROXIMATED_CHART_TYPE',
            'approximated',
            `${seriesPath(c.s)}.stacking`,
            `Excel uses one stacking mode per chart type and axis; this series is drawn ${g.stacking === null ? 'unstacked' : g.stacking === 'percent' ? 'percent-stacked' : 'stacked'} like the other ${g.kind === 'bar' ? 'bar/column' : g.kind} series.`,
            { seriesIndex: c.s.index, details: { stacking, used: g.stacking } },
          ),
        );
      }
      if (g.stacking && !g.range) {
        stackGroups.add(c.s.stackGroup ?? '');
        if (stackGroups.size > 1) {
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
    }
  }

  // 6. Excel has two axis groups (primary, secondary). Category groups on both value axes take both,
  // so XY (scatter) groups, which need their own value-axis pair, cannot be drawn.
  const catSlots = new Set(groups.filter((g) => CATEGORY_GROUP_KINDS.has(g.kind)).map((g) => g.yAxisIndex));
  if (catSlots.size > 1) {
    for (let i = groups.length - 1; i >= 0; i--) {
      const g = groups[i]!;
      if (g.kind !== 'scatter' && g.kind !== 'bubble') continue;
      groups.splice(i, 1);
      for (const pos of g.seriesIndices) {
        dropped.add(pos);
        const s = model.series[pos]!;
        diagnostics.push(
          createDiagnostic(
            'UNSUPPORTED_SERIES_TYPE',
            'unsupported',
            `${seriesPath(s)}.type`,
            'Excel allows two axis groups; scatter series dropped.',
            {
              severity: 'warning',
              seriesIndex: s.index,
            },
          ),
        );
      }
    }
  }

  // 7. Excel draws at most 255 series per chart: later series (in drawing order) are dropped.
  limitSeriesCount(model, groups, dropped, diagnostics);

  // 8. Error bars ride on their parent series.
  const errorBars = attachErrorBars(model, groups, errorBarSeries, dropped, diagnostics);

  const names = [...new Set(groups.map(groupTypeName))];
  const excelChartType = names.length === 1 ? names[0]! : `combo:${names.join('+')}`;
  if (groups.length > 1) {
    const kinds = new Set(groups.map((g) => g.kind));
    if (kinds.size > 1) {
      diagnostics.push(
        createDiagnostic(
          'MIXED_SERIES_TYPES',
          'translated',
          'series',
          `Series types are combined into an Excel combo chart (${names.join(' + ')}).`,
          {
            severity: 'info',
          },
        ),
      );
    }
  }

  return {
    excelChartType,
    groups,
    droppedSeries: [...dropped].sort((a, b) => a - b),
    errorBars,
    diagnostics,
    blocking: false,
  };
}

function unknownSeriesDiagnostic(s: SeriesModel): Diagnostic {
  return createDiagnostic(
    'UNSUPPORTED_SERIES_TYPE',
    'unsupported',
    `${seriesPath(s)}.type`,
    `Series type "${s.sourceType}" has no Excel chart equivalent; the series is not exported to the chart.`,
    { seriesIndex: s.index, details: { sourceType: s.sourceType } },
  );
}

/** Drops the series (in drawing order) beyond Excel's 255 series per chart; range helpers count. */
function limitSeriesCount(
  model: ChartModel,
  groups: PlotGroupPlan[],
  dropped: Set<number>,
  diagnostics: Diagnostic[],
): void {
  let used = 0;
  const overflow: number[] = [];
  for (const g of groups) {
    const keep: number[] = [];
    for (const pos of g.seriesIndices) {
      const n = excelSeriesCount(model.series[pos]!);
      if (overflow.length === 0 && used + n <= EXCEL_MAX_SERIES_PER_CHART) {
        keep.push(pos);
        used += n;
      } else {
        overflow.push(pos);
      }
    }
    g.seriesIndices = keep;
  }
  if (overflow.length === 0) return;
  for (let i = groups.length - 1; i >= 0; i--) if (groups[i]!.seriesIndices.length === 0) groups.splice(i, 1);
  for (const pos of overflow) dropped.add(pos);
  const first = model.series[overflow[0]!]!;
  diagnostics.push(
    createDiagnostic(
      'WRITER_LIMITATION',
      'unsupported',
      `${seriesPath(first)}`,
      `Excel charts hold at most ${EXCEL_MAX_SERIES_PER_CHART} series; ${overflow.length} series from "${first.name}" on are not drawn.`,
      { severity: 'warning', details: { dropped: overflow.length, limit: EXCEL_MAX_SERIES_PER_CHART } },
    ),
  );
}

/**
 * Pairs every errorbar series with its parent (`linkedTo`): the parent must be drawn in a group whose
 * Excel series can carry error bars, and an Excel series holds one set of error bars. Anything else
 * is dropped with UNSUPPORTED_SERIES_TYPE.
 */
function attachErrorBars(
  model: ChartModel,
  groups: readonly PlotGroupPlan[],
  errorBarSeries: ReadonlyArray<{ pos: number; s: SeriesModel }>,
  dropped: Set<number>,
  diagnostics: Diagnostic[],
): Array<{ parent: number; errorBar: number }> {
  const groupOf = new Map<number, PlotGroupPlan>();
  for (const g of groups) for (const pos of g.seriesIndices) groupOf.set(pos, g);
  const out: Array<{ parent: number; errorBar: number }> = [];
  const taken = new Set<number>();
  for (const { pos, s } of errorBarSeries) {
    const linked = s.linkedTo ?? null;
    const parent = linked === null ? -1 : model.series.findIndex((p) => p.id === linked && p.kind !== 'errorbar');
    const g = parent >= 0 ? groupOf.get(parent) : undefined;
    let property = `${seriesPath(s)}.type`;
    let reason: string | null = null;
    if (parent < 0) {
      property = `${seriesPath(s)}.linkedTo`;
      reason = `Error bar series "${s.name}" is not linked to another series; Excel error bars belong to a series, so it is not exported.`;
    } else if (!g) {
      reason = `The parent series of error bar series "${s.name}" is not drawn, so its error bars are not exported.`;
    } else if (g.range || !ERROR_BAR_GROUP_KINDS.has(g.kind)) {
      reason = `Excel cannot draw error bars on ${g.range ? 'range' : g.kind} series; error bar series "${s.name}" is not exported.`;
    } else if (taken.has(parent)) {
      reason = `An Excel series holds one set of error bars; error bar series "${s.name}" is not exported.`;
    }
    if (reason !== null) {
      dropped.add(pos);
      diagnostics.push(
        createDiagnostic('UNSUPPORTED_SERIES_TYPE', 'unsupported', property, reason, {
          severity: 'warning',
          seriesIndex: s.index,
          details: { linkedTo: linked },
        }),
      );
      continue;
    }
    taken.add(parent);
    out.push({ parent, errorBar: pos });
  }
  return out;
}

/**
 * Polar charts → one Excel radar group: line/spline series as a (marker) radar, area/areaspline as a
 * filled radar. Columns/bars and the other types have no radar form and keep the chart blocking.
 */
function resolvePolar(
  model: ChartModel,
  diagnostics: Diagnostic[],
  blocked: () => ChartTypeResolution,
): ChartTypeResolution {
  const dropped = new Set<number>();
  const kept: Array<{ pos: number; s: SeriesModel; filled: boolean }> = [];
  for (const [pos, s] of model.series.entries()) {
    if (s.kind === 'unknown' || s.kind === 'errorbar') {
      dropped.add(pos);
      diagnostics.push(
        s.kind === 'unknown'
          ? unknownSeriesDiagnostic(s)
          : createDiagnostic(
              'UNSUPPORTED_SERIES_TYPE',
              'unsupported',
              `${seriesPath(s)}.type`,
              'Excel radar charts have no error bars; the error bar series is not exported.',
              { severity: 'warning', seriesIndex: s.index },
            ),
      );
      continue;
    }
    const lineLike = s.kind === 'line' || s.kind === 'spline';
    const areaLike = s.kind === 'area' || s.kind === 'areaspline';
    if (!lineLike && !areaLike) {
      diagnostics.push(
        createDiagnostic(
          'UNSUPPORTED_POLAR',
          'blocking',
          'chart.polar',
          `Polar ${s.sourceType} series have no editable Excel equivalent (Excel radar charts draw lines and filled areas only; Excel has no polar columns).`,
          { seriesIndex: s.index, details: { sourceType: s.sourceType } },
        ),
      );
      return blocked();
    }
    kept.push({ pos, s, filled: areaLike });
  }
  if (kept.length === 0) {
    diagnostics.push(
      createDiagnostic(
        'UNSUPPORTED_POLAR',
        'blocking',
        'chart.polar',
        'No series of this polar chart can be drawn as an Excel radar chart.',
      ),
    );
    return blocked();
  }
  diagnostics.push(
    createDiagnostic(
      'APPROXIMATED_CHART_TYPE',
      'approximated',
      'chart.polar',
      'Polar chart exported as an Excel radar chart: categories are spaced evenly around the circle and lines are straight.',
      { severity: 'info' },
    ),
  );
  // One radar group (one radar style): the first series decides; the others follow it.
  const filled = kept[0]!.filled;
  const radarStyle: ExcelRadarStyle = filled
    ? 'filled'
    : kept.every((k) => k.filled || k.s.marker?.enabled === false)
      ? 'standard'
      : 'marker';
  const primaryY = Math.min(...kept.map((k) => k.s.yAxisIndex));
  for (const { s, filled: f } of kept) {
    if (f !== filled) {
      diagnostics.push(
        createDiagnostic(
          'APPROXIMATED_CHART_TYPE',
          'approximated',
          `${seriesPath(s)}.type`,
          `Excel draws one radar style per chart; this series is drawn ${filled ? 'filled' : 'as a line'} like the first series.`,
          { seriesIndex: s.index },
        ),
      );
    }
    if (s.yAxisIndex !== primaryY) {
      diagnostics.push(
        createDiagnostic(
          'UNSUPPORTED_AXIS_FEATURE',
          'approximated',
          `${seriesPath(s)}.yAxis`,
          'Excel radar charts have a single value axis; the series is plotted against it.',
          { seriesIndex: s.index },
        ),
      );
    }
    if (s.stacking !== null) {
      diagnostics.push(
        createDiagnostic(
          'APPROXIMATED_CHART_TYPE',
          'approximated',
          `${seriesPath(s)}.stacking`,
          'Excel radar charts cannot stack series; the series is drawn unstacked.',
          { seriesIndex: s.index },
        ),
      );
    }
    if (s.smooth) {
      diagnostics.push(
        createDiagnostic(
          'APPROXIMATED_CHART_TYPE',
          'approximated',
          `${seriesPath(s)}.type`,
          'Excel radar lines cannot be smoothed; the series is drawn with straight segments.',
          { severity: 'info', seriesIndex: s.index },
        ),
      );
    }
  }
  const groups: PlotGroupPlan[] = [
    { kind: 'radar', seriesIndices: kept.map((k) => k.pos), yAxisIndex: primaryY, stacking: null, radarStyle },
  ];
  limitSeriesCount(model, groups, dropped, diagnostics);
  return {
    excelChartType: groupTypeName(groups[0]!),
    groups,
    droppedSeries: [...dropped].sort((a, b) => a - b),
    errorBars: [],
    diagnostics,
    blocking: false,
  };
}

/** Most common stacking among a group's series; ties prefer stacked ('normal' over 'percent'). */
function majorityStacking(list: Stacking[]): Stacking {
  const count = (v: Stacking): number => list.filter((x) => x === v).length;
  let best: Stacking = 'normal';
  for (const v of ['percent', null] as Stacking[]) if (count(v) > count(best)) best = v;
  return best;
}

/** True for plot group kinds that use the shared category layout (column A = categories). */
export function isCategoryGroupKind(kind: PlotGroupSpec['kind']): boolean {
  return CATEGORY_GROUP_KINDS.has(kind);
}
