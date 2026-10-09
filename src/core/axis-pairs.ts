/**
 * Axis-pair wiring: the category/value axis pair (plus the hidden secondary category axis), the XY
 * axis pair of scatter/bubble charts (plus the hidden axes of XY series inside a combo), the
 * multiple-x-axes and secondary-axis diagnostics, and the axis ids of every plot group.
 */

import type { AxisModel } from '../types/chart-model';
import { createDiagnostic } from '../types/diagnostics';
import type { ExcelAxisSpec } from '../excel/writer-interface';
import {
  AXIS_IDS,
  type AxisParams,
  axisDateFormat,
  baseTimeUnitOf,
  buildAxis,
  categoryCrosses,
  hasTimeOfDay,
  valuePlacement,
} from './axis-builder';
import { isCategoryGroupKind, type PlotGroupPlan } from './chart-type-registry';
import type { DataLayoutXInfo } from './data-layout';
import type { TranslateContext } from './translate-context';

export interface AxesResult {
  /** Every axis, in writing order (pie-only charts have none). */
  axes: ExcelAxisSpec[];
  /** Axis id pair of each plot group (groups missing from the map use the primary category pair). */
  axisIdsOf: Map<PlotGroupPlan, [number, number]>;
  /** No category and no XY group: the chart has no axes. */
  isPieChart: boolean;
}

interface AxisWiring {
  ctx: TranslateContext;
  x: DataLayoutXInfo;
  xModel: AxisModel | null;
  slotOf: (g: PlotGroupPlan) => 0 | 1;
  yModelOf: (g: PlotGroupPlan) => AxisModel | null;
  yPath: (g: PlotGroupPlan) => string;
  axes: ExcelAxisSpec[];
  axisIdsOf: Map<PlotGroupPlan, [number, number]>;
}

/** Builds every axis of the chart and the axis ids of each plot group. */
export function buildAxes(ctx: TranslateContext, groups: readonly PlotGroupPlan[], x: DataLayoutXInfo): AxesResult {
  const { model } = ctx;
  const catGroups = groups.filter((g) => isCategoryGroupKind(g.kind));
  const xyGroups = groups.filter((g) => g.kind === 'scatter' || g.kind === 'bubble');
  const usedY = [...new Set(groups.map((g) => g.yAxisIndex))].sort((a, b) => a - b);
  const primaryY = usedY[0] ?? 0;
  const secondaryY = usedY[1] ?? null;
  const isPieChart = catGroups.length === 0 && xyGroups.length === 0;
  const w: AxisWiring = {
    ctx,
    x,
    xModel: model.xAxes[0] ?? null,
    slotOf: (g) => (g.yAxisIndex === primaryY ? 0 : 1),
    yModelOf: (g) => model.yAxes[g.yAxisIndex] ?? null,
    yPath: (g) => `yAxis[${g.yAxisIndex}]`,
    axes: [],
    axisIdsOf: new Map(),
  };

  if (!isPieChart && model.xAxes.length > 1) {
    ctx.out.push(
      createDiagnostic(
        'MULTIPLE_X_AXES',
        'approximated',
        'xAxis[1]',
        'Excel charts share one category/X axis per axis group; the first x axis is used for every series.',
        { details: { count: model.xAxes.length } },
      ),
    );
  }
  if (secondaryY !== null && !isPieChart) {
    ctx.out.push(
      createDiagnostic(
        'SECONDARY_AXIS',
        'translated',
        `yAxis[${secondaryY}]`,
        'Series on a second y axis are plotted against an Excel secondary value axis.',
        { severity: 'info' },
      ),
    );
  }
  if (catGroups.length > 0) buildCategoryAxes(w, catGroups);
  if (xyGroups.length > 0) buildXyAxes(w, xyGroups, catGroups.length > 0);
  return { axes: w.axes, axisIdsOf: w.axisIdsOf, isPieChart };
}

/** Category axis (date, numeric or text) + value axis, plus the hidden secondary category/value pair. */
function buildCategoryAxes(w: AxisWiring, catGroups: PlotGroupPlan[]): void {
  const { ctx, x, xModel, axes } = w;
  const { support } = ctx;
  const horizontal = catGroups.some((g) => g.barDir === 'bar');
  if (horizontal && catGroups.some((g) => g.kind !== 'bar')) {
    ctx.out.push(
      createDiagnostic(
        'APPROXIMATED_CHART_TYPE',
        'approximated',
        'chart.type',
        'Horizontal bars and line/area series share rotated axes in Excel; lines are drawn on the bar orientation.',
        { severity: 'info' },
      ),
    );
  }
  // Category axis kind.
  let catKind: ExcelAxisSpec['kind'] = 'cat';
  let catFormat: ExcelAxisSpec['numberFormat'] = null;
  let baseTimeUnit: 'days' | 'months' | 'years' | undefined;
  const xk = x.kind;
  const numericKeys = x.keys.filter((k): k is number => typeof k === 'number');
  if (xk === 'datetime') {
    const span = x.min !== null && x.max !== null ? x.max - x.min : 0;
    catFormat = { code: axisDateFormat(xModel, span), sourceLinked: false };
    if (numericKeys.some(hasTimeOfDay)) {
      ctx.out.push(
        createDiagnostic(
          'APPROXIMATED_DATETIME',
          'approximated',
          'xAxis[0].type',
          'Excel date axes have a resolution of one day; intraday timestamps are shown on an evenly spaced category axis.',
        ),
      );
    } else {
      catKind = 'date';
      baseTimeUnit = baseTimeUnitOf(numericKeys);
      support('xAxis[0].type');
    }
  } else if (xk === 'linear' || xk === 'logarithmic') {
    const gaps = numericKeys.slice(1).map((k, i) => k - numericKeys[i]!);
    const uneven = gaps.some((g) => Math.abs(g - (gaps[0] ?? 0)) > 1e-9 * Math.max(1, Math.abs(gaps[0] ?? 1)));
    if (uneven || xk === 'logarithmic') {
      ctx.out.push(
        createDiagnostic(
          'APPROXIMATED_AXIS_SCALE',
          'approximated',
          'xAxis[0].type',
          xk === 'logarithmic'
            ? 'A logarithmic x axis on a line/column chart is shown as evenly spaced categories in Excel.'
            : 'Numeric x values are not evenly spaced; Excel category axes space them evenly.',
        ),
      );
    } else {
      support('xAxis[0].type');
    }
  } else {
    support('xAxis[0].categories');
  }
  // Radar charts (polar) take a plain category axis around the circle: no date axis, labels next to it.
  const radar = catGroups.every((g) => g.kind === 'radar');
  if (radar && catKind === 'date') {
    catKind = 'cat';
    baseTimeUnit = undefined;
  }
  const catPos: ExcelAxisSpec['position'] = horizontal ? (xModel?.opposite ? 'r' : 'l') : xModel?.opposite ? 't' : 'b';
  if (xModel?.opposite) support('xAxis[0].opposite');
  const primary = catGroups.filter((g) => w.slotOf(g) === 0);
  const secondary = catGroups.filter((g) => w.slotOf(g) === 1);
  const hasPrimary = primary.length > 0;
  // Primary pair (always present when there are category groups; secondary-only charts use it as primary).
  const primaryYModel = w.yModelOf(hasPrimary ? primary[0]! : secondary[0]!);
  const primaryPath = w.yPath(hasPrimary ? primary[0]! : secondary[0]!);
  const catParams: AxisParams = {
    id: AXIS_IDS.primaryCat,
    kind: catKind,
    position: catPos,
    crossAxisId: AXIS_IDS.primaryVal,
    axis: xModel,
    path: 'xAxis[0]',
    crosses: categoryCrosses(xModel, primaryYModel),
    numberFormat: catFormat,
    serial: catKind === 'date',
    labelsLow: !radar,
    crossAxisReversed: primaryYModel?.reversed === true,
  };
  if (baseTimeUnit) catParams.baseTimeUnit = baseTimeUnit;
  axes.push(buildAxis(ctx, catParams));
  const catReversed = xModel?.reversed === true;
  const pv = valuePlacement(primaryYModel, horizontal, false, catReversed);
  if (primaryYModel?.opposite) support(`${primaryPath}.opposite`);
  axes.push(
    buildAxis(ctx, {
      id: AXIS_IDS.primaryVal,
      kind: 'val',
      position: pv.position,
      crossAxisId: AXIS_IDS.primaryCat,
      axis: primaryYModel,
      path: primaryPath,
      crosses: pv.crosses,
    }),
  );
  for (const g of catGroups) w.axisIdsOf.set(g, [AXIS_IDS.primaryCat, AXIS_IDS.primaryVal]);
  if (secondary.length > 0 && hasPrimary) {
    const sModel = w.yModelOf(secondary[0]!);
    const sv = valuePlacement(sModel, horizontal, true, catReversed);
    const secondaryCat: AxisParams = {
      id: AXIS_IDS.secondaryCat,
      kind: catKind,
      position: catPos,
      crossAxisId: AXIS_IDS.secondaryVal,
      axis: xModel,
      path: 'xAxis[0]',
      forceDeleted: true,
      numberFormat: catFormat,
      serial: catKind === 'date',
      reportSupport: false,
    };
    if (baseTimeUnit) secondaryCat.baseTimeUnit = baseTimeUnit;
    axes.push(buildAxis(ctx, secondaryCat));
    axes.push(
      buildAxis(ctx, {
        id: AXIS_IDS.secondaryVal,
        kind: 'val',
        position: sv.position,
        crossAxisId: AXIS_IDS.secondaryCat,
        axis: sModel,
        path: w.yPath(secondary[0]!),
        crosses: sv.crosses,
      }),
    );
    for (const g of secondary) w.axisIdsOf.set(g, [AXIS_IDS.secondaryCat, AXIS_IDS.secondaryVal]);
  }
}

/** X/Y value axes of scatter/bubble groups (hidden auxiliary axes when they share a combo with categories). */
function buildXyAxes(w: AxisWiring, xyGroups: PlotGroupPlan[], inCombo: boolean): void {
  const { ctx, x, xModel, axes } = w;
  const xIsDate = x.kind === 'datetime';
  const span = x.min !== null && x.max !== null ? x.max - x.min : 0;
  const xFormat: ExcelAxisSpec['numberFormat'] = xIsDate
    ? { code: axisDateFormat(xModel, span), sourceLinked: false }
    : null;
  const hiddenX = (id: number, crossAxisId: number): ExcelAxisSpec =>
    buildAxis(ctx, {
      id,
      kind: 'val',
      position: 'b',
      crossAxisId,
      axis: xModel,
      path: 'xAxis[0]',
      forceDeleted: true,
      numberFormat: xFormat,
      serial: xIsDate,
      reportSupport: false,
    });
  if (inCombo) {
    // Excel cannot put XY series on category axes: give them their own hidden value axes.
    ctx.out.push(
      createDiagnostic(
        'APPROXIMATED_AXIS_SCALE',
        'approximated',
        'xAxis[0]',
        'Scatter series in a combo chart are plotted on their own hidden X/Y axes; their scale may differ from the category axis.',
      ),
    );
    axes.push(hiddenX(AXIS_IDS.scatterX, AXIS_IDS.scatterAux));
    axes.push(
      buildAxis(ctx, {
        id: AXIS_IDS.scatterAux,
        kind: 'val',
        position: 'l',
        crossAxisId: AXIS_IDS.scatterX,
        axis: w.yModelOf(xyGroups[0]!),
        path: w.yPath(xyGroups[0]!),
        forceDeleted: true,
        reportSupport: false,
      }),
    );
    for (const g of xyGroups) w.axisIdsOf.set(g, [AXIS_IDS.scatterX, AXIS_IDS.scatterAux]);
    return;
  }
  if (xIsDate) ctx.support('xAxis[0].type');
  const primary = xyGroups.filter((g) => w.slotOf(g) === 0);
  const secondary = xyGroups.filter((g) => w.slotOf(g) === 1);
  const pModel = w.yModelOf(primary[0] ?? secondary[0]!);
  const pPath = w.yPath(primary[0] ?? secondary[0]!);
  const pv = valuePlacement(pModel, false, false, xModel?.reversed === true);
  axes.push(
    buildAxis(ctx, {
      id: AXIS_IDS.scatterX,
      kind: 'val',
      position: xModel?.opposite ? 't' : 'b',
      crossAxisId: AXIS_IDS.primaryVal,
      axis: xModel,
      path: 'xAxis[0]',
      crosses: categoryCrosses(xModel, pModel),
      numberFormat: xFormat,
      serial: xIsDate,
      labelsLow: true,
      crossAxisReversed: pModel?.reversed === true,
    }),
  );
  axes.push(
    buildAxis(ctx, {
      id: AXIS_IDS.primaryVal,
      kind: 'val',
      position: pv.position,
      crossAxisId: AXIS_IDS.scatterX,
      axis: pModel,
      path: pPath,
      crosses: pv.crosses,
    }),
  );
  for (const g of xyGroups) w.axisIdsOf.set(g, [AXIS_IDS.scatterX, AXIS_IDS.primaryVal]);
  if (secondary.length > 0 && primary.length > 0) {
    axes.push(hiddenX(AXIS_IDS.scatterAux, AXIS_IDS.secondaryVal));
    axes.push(
      buildAxis(ctx, {
        id: AXIS_IDS.secondaryVal,
        kind: 'val',
        position: 'r',
        crossAxisId: AXIS_IDS.scatterAux,
        axis: w.yModelOf(secondary[0]!),
        path: w.yPath(secondary[0]!),
        crosses: xModel?.reversed === true ? 'min' : 'max',
      }),
    );
    for (const g of secondary) w.axisIdsOf.set(g, [AXIS_IDS.scatterAux, AXIS_IDS.secondaryVal]);
  }
}
