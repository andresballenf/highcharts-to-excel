/**
 * ChartModel (neutral IR) → SheetSpecs with a native, cell-referencing Excel chart.
 */

import type { AxisModel, ChartModel, Color, Fill, MarkerStyle, SeriesModel, TextBlock } from '../types/chart-model';
import { createDiagnostic, type Diagnostic, type DiagnosticCode, type DiagnosticCollector } from '../types/diagnostics';
import type { FidelityMode } from '../types/public-api';
import type {
  AnchorSpec,
  DrawingSpec,
  ExcelAxisSpec,
  ExcelChartSpec,
  ExcelDataLabelsSpec,
  ExcelFontSpec,
  ExcelLineSpec,
  ExcelMarkerSpec,
  ExcelSeriesSpec,
  ExcelShapeStyle,
  ExcelTextSpec,
  PlotGroupSpec,
  SheetSpec,
} from '../excel/writer-interface';
import { HIGHCHARTS_DEFAULT_PALETTE, parseColor } from '../utils/colors';
import { guessExcelDateFormatForRange, highchartsDateFormatToExcel, msToExcelSerial } from '../utils/dates';
import { sanitizeSheetName } from '../utils/filenames';
import { stripControlChars } from '../utils/text';
import { clamp } from '../utils/units';
import { DEFAULT_HIGHCHARTS_FONT } from '../translators/typography-translator';
import { fillToSolidColor } from '../translators/color-translator';
import { isCategoryGroupKind, resolveChartType, type ChartTypeResolution, type PlotGroupPlan } from './chart-type-registry';
import { buildDataLayout, type SeriesRange } from './data-layout';
import {
  NO_LINE,
  hiddenDataLabels,
  legendPosition,
  solidFillOf,
  toExcelDataLabels,
  toExcelFill,
  toExcelFont,
  toExcelLine,
  toExcelMarker,
} from './style-mapping';

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

/** Axis ids. */
export const AXIS_IDS = Object.freeze({
  primaryCat: 1000,
  secondaryCat: 1001,
  primaryVal: 2000,
  secondaryVal: 2001,
  scatterX: 3000,
  /** Pure scatter charts: hidden secondary X; scatter inside a combo: its own (hidden) Y axis. */
  scatterAux: 3001,
});

const MS_PER_DAY = 86_400_000;
/** Default Excel column width in pixels (8.43 characters of Calibri 11). */
const DEFAULT_COLUMN_WIDTH_PX = 64;

/** Diagnostic codes that only concern styling (dropped in 'minimal' fidelity). */
const STYLE_CODES: ReadonlySet<DiagnosticCode> = new Set<DiagnosticCode>([
  'APPROXIMATED_FONT',
  'APPROXIMATED_FONT_SIZE',
  'APPROXIMATED_MARKER',
  'APPROXIMATED_COLOR',
  'UNSUPPORTED_GRADIENT',
  'APPROXIMATED_DASH_STYLE',
  'UNSUPPORTED_STYLE',
]);

function hasTimeOfDay(ms: number): boolean {
  return ((ms % MS_PER_DAY) + MS_PER_DAY) % MS_PER_DAY !== 0;
}

/** Excel date format for the data cells holding datetime x values. */
export function dataDateFormat(axis: AxisModel, spanMs: number, withTime: boolean): string {
  if (axis.dateFormat) {
    const f = highchartsDateFormatToExcel(axis.dateFormat);
    if (f.kind === 'excel' && /y/.test(f.code) && /d/.test(f.code)) return f.code;
  }
  return withTime || spanMs < MS_PER_DAY ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd';
}

/** Excel number format for a datetime axis' tick labels. */
export function axisDateFormat(axis: AxisModel | null, spanMs: number): string {
  if (axis?.labels.format?.kind === 'excel') return axis.labels.format.code;
  if (axis?.dateFormat) {
    const f = highchartsDateFormatToExcel(axis.dateFormat);
    if (f.kind === 'excel') return f.code;
  }
  return guessExcelDateFormatForRange(spanMs);
}

/** Base time unit for an Excel date axis from the x values (ms): month/year starts → months/years. */
export function baseTimeUnitOf(keysMs: readonly number[]): 'days' | 'months' | 'years' {
  if (keysMs.length < 2) return 'days';
  const dates = keysMs.map((k) => new Date(k));
  const allMonthStarts = keysMs.every((k, i) => !hasTimeOfDay(k) && dates[i]!.getUTCDate() === 1);
  if (!allMonthStarts) return 'days';
  return dates.every((d) => d.getUTCMonth() === 0) ? 'years' : 'months';
}

function textOf(tb: TextBlock): string[] {
  return stripControlChars(tb.text)
    .split(/\r?\n|<br\s*\/?>/i)
    .map((l) => l.trim())
    .filter((l) => l !== '');
}

/** Translates a ChartModel into the chart sheet and data sheet specs. Never mutates the model. */
export function translateChartModel(model: ChartModel, opts: TranslateOptions): TranslationResult {
  const best = opts.fidelity !== 'minimal';
  const out: Diagnostic[] = [];
  const styleSink: Diagnostic[] = [];
  const supported: string[] = [];
  const support = (...paths: string[]): void => {
    for (const p of paths) supported.push(p);
  };
  const supportStyle = (...paths: string[]): void => {
    if (best) support(...paths);
  };
  const flush = (): void => {
    opts.diagnostics.addAll(out);
    if (best) opts.diagnostics.addAll(styleSink.filter(Boolean));
    else opts.diagnostics.addAll(styleSink.filter((d) => !STYLE_CODES.has(d.code)));
    out.length = 0;
    styleSink.length = 0;
  };

  // --- Sheet names --------------------------------------------------------------------------------
  const taken = new Set<string>(opts.takenSheetNames ?? []);
  const chartSheet = sanitizeSheetName(opts.chartSheetName, taken, 'Chart');
  if (chartSheet.adjusted) {
    out.push(
      createDiagnostic('SHEET_NAME_ADJUSTED', 'approximated', 'chartSheetName', `Sheet name "${opts.chartSheetName}" was changed to "${chartSheet.name}" to satisfy Excel's naming rules.`, {
        severity: 'info',
        details: { requested: opts.chartSheetName, used: chartSheet.name },
      }),
    );
  }
  taken.add(chartSheet.name);
  const dataSheet = sanitizeSheetName(opts.dataSheetName, taken, 'Data');
  if (dataSheet.adjusted) {
    out.push(
      createDiagnostic('SHEET_NAME_ADJUSTED', 'approximated', 'dataSheetName', `Sheet name "${opts.dataSheetName}" was changed to "${dataSheet.name}" to satisfy Excel's naming rules.`, {
        severity: 'info',
        details: { requested: opts.dataSheetName, used: dataSheet.name },
      }),
    );
  }

  // --- Chart type -------------------------------------------------------------------------------------
  const resolution = resolveChartType(model);
  out.push(...resolution.diagnostics);
  const blockedResult = (): TranslationResult => {
    flush();
    return {
      sheets: [],
      chartSheetName: chartSheet.name,
      dataSheetName: dataSheet.name,
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
    sheetName: dataSheet.name,
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
  const groups = resolution.groups;
  const catGroups = groups.filter((g) => isCategoryGroupKind(g.kind));
  const xyGroups = groups.filter((g) => g.kind === 'scatter' || g.kind === 'bubble');
  const usedY = [...new Set(groups.map((g) => g.yAxisIndex))].sort((a, b) => a - b);
  const primaryY = usedY[0] ?? 0;
  const secondaryY = usedY[1] ?? null;
  const slotOf = (g: PlotGroupPlan): 0 | 1 => (g.yAxisIndex === primaryY ? 0 : 1);
  const xModel = model.xAxes[0] ?? null;
  const axes: ExcelAxisSpec[] = [];
  const axisIdsOf = new Map<PlotGroupPlan, [number, number]>();
  const isPieChart = catGroups.length === 0 && xyGroups.length === 0;

  if (!isPieChart && model.xAxes.length > 1) {
    out.push(
      createDiagnostic('MULTIPLE_X_AXES', 'approximated', 'xAxis[1]', 'Excel charts share one category/X axis per axis group; the first x axis is used for every series.', {
        details: { count: model.xAxes.length },
      }),
    );
  }
  if (secondaryY !== null && !isPieChart) {
    out.push(
      createDiagnostic('SECONDARY_AXIS', 'translated', `yAxis[${secondaryY}]`, 'Series on a second y axis are plotted against an Excel secondary value axis.', {
        severity: 'info',
      }),
    );
  }

  const fontOf = (font: Parameters<typeof toExcelFont>[0], property: string): ExcelFontSpec | null =>
    best ? toExcelFont(font, property, styleSink) : null;
  const textSpec = (tb: TextBlock | null, property: string, extra: { lines: string[]; font: ExcelFontSpec | null } | null = null): ExcelTextSpec | null => {
    if (!tb) return null;
    const own = textOf(tb);
    const lines = [...own, ...(extra?.lines ?? [])];
    if (lines.length === 0) return null;
    const spec: ExcelTextSpec = { lines, font: fontOf(tb.font, `${property}.style`), overlay: false };
    if (extra && extra.lines.length > 0 && extra.font) {
      spec.lineFonts = [...own.map(() => null), ...extra.lines.map(() => extra.font)];
    }
    return spec;
  };

  interface AxisParams {
    id: number;
    kind: ExcelAxisSpec['kind'];
    position: ExcelAxisSpec['position'];
    crossAxisId: number;
    axis: AxisModel | null;
    path: string;
    forceDeleted?: boolean;
    crosses?: ExcelAxisSpec['crosses'];
    numberFormat?: ExcelAxisSpec['numberFormat'];
    /** Values on this axis are Excel date serials (min/max converted from ms). */
    serial?: boolean;
    baseTimeUnit?: 'days' | 'months' | 'years';
    /** Pin labels to the low end of the perpendicular axis ('high' when that axis is reversed). */
    labelsLow?: boolean;
    crossAxisReversed?: boolean;
    reportSupport?: boolean;
  }

  const buildAxis = (p: AxisParams): ExcelAxisSpec => {
    const a = p.axis;
    const report = p.reportSupport !== false && !p.forceDeleted && a !== null;
    const isValueLike = p.kind === 'val' || p.kind === 'date';
    const deleted = p.forceDeleted === true || (a !== null && !a.visible);
    const toScale = (v: number | null): number | null => (v === null || !Number.isFinite(v) ? null : p.serial ? msToExcelSerial(v) : v);
    let min: number | null = null;
    let max: number | null = null;
    if (a && isValueLike) {
      min = toScale(a.min);
      max = toScale(a.max);
      if (report && a.min !== null) support(`${p.path}.min`);
      if (report && a.max !== null) support(`${p.path}.max`);
    } else if (a && (a.min !== null || a.max !== null) && report) {
      out.push(
        createDiagnostic('APPROXIMATED_AXIS_SCALE', 'approximated', `${p.path}.min`, 'Excel category axes cannot be cropped with min/max; all categories are shown.', {
          severity: 'info',
        }),
      );
    }
    let logBase: number | null = null;
    if (a?.kind === 'logarithmic' && p.kind === 'val') {
      logBase = a.logBase && a.logBase >= 2 ? a.logBase : 10;
      if (report) support(`${p.path}.type`);
    }
    let majorUnit: number | null = null;
    let minorUnit: number | null = null;
    if (a && a.tickInterval !== null && a.tickInterval > 0 && logBase === null) {
      if (p.kind === 'val') majorUnit = a.tickInterval;
      else if (p.kind === 'date' && p.baseTimeUnit === 'days') majorUnit = a.tickInterval / MS_PER_DAY;
      if (majorUnit !== null && report) support(`${p.path}.tickInterval`);
    }
    if (a && p.kind === 'val' && a.minorTickInterval !== null && a.minorTickInterval > 0 && logBase === null) {
      minorUnit = a.minorTickInterval;
      if (report) support(`${p.path}.minorTickInterval`);
    }
    let numberFormat = p.numberFormat ?? null;
    if (numberFormat === null && a?.labels.format?.kind === 'excel') {
      numberFormat = { code: a.labels.format.code, sourceLinked: false };
    }
    if (report && a?.labels.format?.kind === 'excel') support(`${p.path}.labels.format`);
    const gridOf = (s: AxisModel['gridLines'], key: string): ExcelLineSpec | null => {
      if (!s || !(s.width > 0)) return null;
      if (report) supportStyle(`${key}Width`, `${key}Color`);
      return best ? toExcelLine(s) : { widthPx: 1, hex: null, alpha: 1, dash: 'solid', noFill: false };
    };
    let rotation: number | null = null;
    if (a && Number.isFinite(a.labels.rotation) && a.labels.rotation !== 0) {
      rotation = clamp(Math.round(a.labels.rotation), -90, 90);
      if (report) support(`${p.path}.labels.rotation`);
    }
    const labelsHidden = a !== null && !a.labels.enabled;
    if (report) {
      support(`${p.path}.visible`);
      if (labelsHidden) support(`${p.path}.labels.enabled`);
      if (a.reversed) support(`${p.path}.reversed`);
      if (a.title) support(`${p.path}.title.text`);
      if (a.title && best) support(`${p.path}.title.style`);
      if (a.labels.font) supportStyle(`${p.path}.labels.style`);
      if (a.axisLine) supportStyle(`${p.path}.lineWidth`, `${p.path}.lineColor`);
      if (a.tickMarks) support(`${p.path}.tickWidth`);
    }
    const spec: ExcelAxisSpec = {
      id: p.id,
      kind: p.kind,
      position: p.position,
      crossAxisId: p.crossAxisId,
      deleted,
      title: p.forceDeleted ? null : textSpec(a?.title ?? null, `${p.path}.title`),
      numberFormat,
      majorGridlines: p.forceDeleted ? null : gridOf(a?.gridLines ?? null, `${p.path}.gridLine`),
      minorGridlines: p.forceDeleted ? null : gridOf(a?.minorGridLines ?? null, `${p.path}.minorGridLine`),
      axisLine: best && a?.axisLine ? toExcelLine(a.axisLine) : null,
      labels: {
        position: labelsHidden ? 'none' : p.labelsLow ? (p.crossAxisReversed ? 'high' : 'low') : 'nextTo',
        font: fontOf(a?.labels.font ?? null, `${p.path}.labels.style`),
        rotation,
      },
      scaling: { min, max, orientation: a?.reversed ? 'maxMin' : 'minMax', logBase },
      majorUnit,
      minorUnit,
      crosses: p.crosses ?? 'autoZero',
      majorTickMark: a?.tickMarks && a.tickMarks.width > 0 ? 'out' : 'none',
      dateAxis: p.kind === 'date' ? { baseTimeUnit: p.baseTimeUnit ?? null } : null,
    };
    return spec;
  };

  const yModelOf = (g: PlotGroupPlan): AxisModel | null => model.yAxes[g.yAxisIndex] ?? null;
  const yPath = (g: PlotGroupPlan): string => `yAxis[${g.yAxisIndex}]`;
  /**
   * Value axis placement. Horizontal bar charts whose category axis is reversed (Highcharts bars list the
   * first category at the top: the IR x axis arrives with `reversed: true`) need the standard Excel
   * "categories in reverse order + value axis crosses at maximum category" combination to keep the
   * value axis at the bottom; a secondary value axis then crosses at the minimum (top).
   */
  const valuePlacement = (
    axis: AxisModel | null,
    horizontal: boolean,
    secondary: boolean,
    catReversed = false,
  ): { position: ExcelAxisSpec['position']; crosses: ExcelAxisSpec['crosses'] } => {
    // In Excel an axis sits where it crosses its perpendicular axis, so a reversed perpendicular axis
    // flips "crosses at min/max". The general rule for an axis A crossing axis B is:
    // A is on the far side  <=>  (A.opposite XOR B.reversed)  <=>  crosses 'max'.
    const far = (secondary || axis?.opposite === true) !== catReversed;
    if (secondary || axis?.opposite) return { position: horizontal ? 't' : 'r', crosses: far ? 'max' : 'min' };
    if (axis?.crossing !== null && axis?.crossing !== undefined && Number.isFinite(axis.crossing)) {
      return { position: horizontal ? 'b' : 'l', crosses: { at: axis.crossing } };
    }
    return { position: horizontal ? 'b' : 'l', crosses: far ? 'max' : 'autoZero' };
  };
  /** Category/X axis crossing: far side when (opposite XOR the value axis is reversed). */
  const categoryCrosses = (x: AxisModel | null, y: AxisModel | null): ExcelAxisSpec['crosses'] =>
    (x?.opposite === true) !== (y?.reversed === true) ? 'max' : 'autoZero';

  if (catGroups.length > 0) {
    const horizontal = catGroups.some((g) => g.barDir === 'bar');
    if (horizontal && catGroups.some((g) => g.kind !== 'bar')) {
      out.push(
        createDiagnostic('APPROXIMATED_CHART_TYPE', 'approximated', 'chart.type', 'Horizontal bars and line/area series share rotated axes in Excel; lines are drawn on the bar orientation.', {
          severity: 'info',
        }),
      );
    }
    // Category axis kind.
    let catKind: ExcelAxisSpec['kind'] = 'cat';
    let catFormat: ExcelAxisSpec['numberFormat'] = null;
    let baseTimeUnit: 'days' | 'months' | 'years' | undefined;
    const xk = layout.x.kind;
    const numericKeys = layout.x.keys.filter((k): k is number => typeof k === 'number');
    if (xk === 'datetime') {
      const span = layout.x.min !== null && layout.x.max !== null ? layout.x.max - layout.x.min : 0;
      catFormat = { code: axisDateFormat(xModel, span), sourceLinked: false };
      if (numericKeys.some(hasTimeOfDay)) {
        out.push(
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
        out.push(
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
    const catPos: ExcelAxisSpec['position'] = horizontal ? (xModel?.opposite ? 'r' : 'l') : xModel?.opposite ? 't' : 'b';
    if (xModel?.opposite) support('xAxis[0].opposite');
    const primary = catGroups.filter((g) => slotOf(g) === 0);
    const secondary = catGroups.filter((g) => slotOf(g) === 1);
    const hasPrimary = primary.length > 0;
    // Primary pair (always present when there are category groups; secondary-only charts use it as primary).
    const primaryYModel = yModelOf(hasPrimary ? primary[0]! : secondary[0]!);
    const catCrosses: ExcelAxisSpec['crosses'] = categoryCrosses(xModel, primaryYModel);
    const primaryPath = yPath(hasPrimary ? primary[0]! : secondary[0]!);
    const catParams: AxisParams = {
      id: AXIS_IDS.primaryCat,
      kind: catKind,
      position: catPos,
      crossAxisId: AXIS_IDS.primaryVal,
      axis: xModel,
      path: 'xAxis[0]',
      crosses: catCrosses,
      numberFormat: catFormat,
      serial: catKind === 'date',
      labelsLow: true,
      crossAxisReversed: primaryYModel?.reversed === true,
    };
    if (baseTimeUnit) catParams.baseTimeUnit = baseTimeUnit;
    axes.push(buildAxis(catParams));
    const catReversed = xModel?.reversed === true;
    const pv = valuePlacement(primaryYModel, horizontal, false, catReversed);
    if (primaryYModel?.opposite) support(`${primaryPath}.opposite`);
    axes.push(
      buildAxis({
        id: AXIS_IDS.primaryVal,
        kind: 'val',
        position: pv.position,
        crossAxisId: AXIS_IDS.primaryCat,
        axis: primaryYModel,
        path: primaryPath,
        crosses: pv.crosses,
      }),
    );
    for (const g of catGroups) axisIdsOf.set(g, [AXIS_IDS.primaryCat, AXIS_IDS.primaryVal]);
    if (secondary.length > 0 && hasPrimary) {
      const sModel = yModelOf(secondary[0]!);
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
      axes.push(buildAxis(secondaryCat));
      axes.push(
        buildAxis({
          id: AXIS_IDS.secondaryVal,
          kind: 'val',
          position: sv.position,
          crossAxisId: AXIS_IDS.secondaryCat,
          axis: sModel,
          path: yPath(secondary[0]!),
          crosses: sv.crosses,
        }),
      );
      for (const g of secondary) axisIdsOf.set(g, [AXIS_IDS.secondaryCat, AXIS_IDS.secondaryVal]);
    }
  }

  if (xyGroups.length > 0) {
    const xIsDate = layout.x.kind === 'datetime';
    const span = layout.x.min !== null && layout.x.max !== null ? layout.x.max - layout.x.min : 0;
    const xFormat: ExcelAxisSpec['numberFormat'] = xIsDate ? { code: axisDateFormat(xModel, span), sourceLinked: false } : null;
    if (catGroups.length > 0) {
      // Excel cannot put XY series on category axes: give them their own hidden value axes.
      out.push(
        createDiagnostic(
          'APPROXIMATED_AXIS_SCALE',
          'approximated',
          'xAxis[0]',
          'Scatter series in a combo chart are plotted on their own hidden X/Y axes; their scale may differ from the category axis.',
        ),
      );
      axes.push(
        buildAxis({
          id: AXIS_IDS.scatterX,
          kind: 'val',
          position: 'b',
          crossAxisId: AXIS_IDS.scatterAux,
          axis: xModel,
          path: 'xAxis[0]',
          forceDeleted: true,
          numberFormat: xFormat,
          serial: xIsDate,
          reportSupport: false,
        }),
      );
      axes.push(
        buildAxis({
          id: AXIS_IDS.scatterAux,
          kind: 'val',
          position: 'l',
          crossAxisId: AXIS_IDS.scatterX,
          axis: yModelOf(xyGroups[0]!),
          path: yPath(xyGroups[0]!),
          forceDeleted: true,
          reportSupport: false,
        }),
      );
      for (const g of xyGroups) axisIdsOf.set(g, [AXIS_IDS.scatterX, AXIS_IDS.scatterAux]);
    } else {
      if (xIsDate) support('xAxis[0].type');
      const primary = xyGroups.filter((g) => slotOf(g) === 0);
      const secondary = xyGroups.filter((g) => slotOf(g) === 1);
      const pModel = yModelOf(primary[0] ?? secondary[0]!);
      const pPath = yPath(primary[0] ?? secondary[0]!);
      const pv = valuePlacement(pModel, false, false, xModel?.reversed === true);
      axes.push(
        buildAxis({
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
        buildAxis({
          id: AXIS_IDS.primaryVal,
          kind: 'val',
          position: pv.position,
          crossAxisId: AXIS_IDS.scatterX,
          axis: pModel,
          path: pPath,
          crosses: pv.crosses,
        }),
      );
      for (const g of xyGroups) axisIdsOf.set(g, [AXIS_IDS.scatterX, AXIS_IDS.primaryVal]);
      if (secondary.length > 0 && primary.length > 0) {
        axes.push(
          buildAxis({
            id: AXIS_IDS.scatterAux,
            kind: 'val',
            position: 'b',
            crossAxisId: AXIS_IDS.secondaryVal,
            axis: xModel,
            path: 'xAxis[0]',
            forceDeleted: true,
            numberFormat: xFormat,
            serial: xIsDate,
            reportSupport: false,
          }),
        );
        axes.push(
          buildAxis({
            id: AXIS_IDS.secondaryVal,
            kind: 'val',
            position: 'r',
            crossAxisId: AXIS_IDS.scatterAux,
            axis: yModelOf(secondary[0]!),
            path: yPath(secondary[0]!),
            crosses: xModel?.reversed === true ? 'min' : 'max',
          }),
        );
        for (const g of secondary) axisIdsOf.set(g, [AXIS_IDS.scatterAux, AXIS_IDS.secondaryVal]);
      }
    }
  }

  // --- Series & plot groups ------------------------------------------------------------------------------
  const palette: Color[] = model.colors.length > 0 ? model.colors : HIGHCHARTS_DEFAULT_PALETTE.map((c) => parseColor(c)!).filter(Boolean);
  const paletteColor = (i: number): Color | null => palette[((i % palette.length) + palette.length) % palette.length] ?? null;
  const seriesColorOf = (s: SeriesModel): Color | null => s.color ?? fillToSolidColor(s.fill) ?? paletteColor(s.index);
  const minDim = Math.max(1, Math.min(model.width, model.height));
  let nextIdx = 0;

  const fillOrColor = (fill: Fill | null, color: Color | null): Fill | null => fill ?? (color ? { type: 'solid', color } : null);

  const buildSeries = (pos: number, g: PlotGroupPlan): ExcelSeriesSpec => {
    const s = model.series[pos]!;
    const range = rangeByPos.get(pos)!;
    const path = `series[${s.index}]`;
    const color = seriesColorOf(s);
    const kind = g.kind;
    const stacked = g.stacking !== null;
    support(`${path}.type`, `${path}.name`, `${path}.data`);
    if (g.stacking) support(`${path}.stacking`);
    if (s.yAxisIndex !== 0 && s.yAxisIndex === g.yAxisIndex) support(`${path}.yAxis`);

    // Shape
    let shape: ExcelShapeStyle = { fill: null, line: null };
    let marker: ExcelMarkerSpec | null = null;
    if (best) {
      switch (kind) {
        case 'bar':
        case 'area':
        case 'bubble':
          shape.fill = toExcelFill(fillOrColor(s.fill, color), s.fillOpacity);
          supportStyle(`${path}.color`);
          if (kind === 'area' && s.fillOpacity !== 1) supportStyle(`${path}.fillOpacity`);
          break;
        default:
          break;
      }
      switch (kind) {
        case 'line':
          shape.line = toExcelLine(s.line ?? { color, width: 2, dash: 'solid' }, color);
          supportStyle(`${path}.color`, `${path}.lineWidth`);
          if (s.line && s.line.dash !== 'solid') supportStyle(`${path}.dashStyle`);
          break;
        case 'area':
          shape.line = s.line ? toExcelLine(s.line, color) : null;
          if (s.line) supportStyle(`${path}.lineWidth`);
          break;
        case 'scatter':
          shape.line = s.line && s.line.width > 0 ? toExcelLine(s.line, color) : { ...NO_LINE };
          supportStyle(`${path}.color`, `${path}.lineWidth`);
          break;
        case 'bar':
        case 'pie':
        case 'doughnut':
          shape.line = toExcelLine(s.border);
          if (s.border) supportStyle(`${path}.borderWidth`, `${path}.borderColor`);
          break;
        case 'bubble':
          shape.line = toExcelLine(s.border ?? s.line, color);
          break;
      }
      if (kind === 'line' || kind === 'scatter') {
        const m: Partial<MarkerStyle> | null =
          s.marker ?? (kind === 'scatter' ? { enabled: true, symbol: 'circle', radius: 4, fill: null, stroke: null, strokeWidth: 0 } : null);
        marker = toExcelMarker(m, color, `${path}.marker`, styleSink);
        if (s.marker) supportStyle(`${path}.marker`);
      }
    } else {
      shape = { fill: null, line: null };
      if (kind === 'scatter' && !(s.line && s.line.width > 0)) shape.line = { ...NO_LINE };
    }
    if (best && (kind === 'bar') && s.bars && s.bars.borderRadius > 0) {
      styleSink.push(
        createDiagnostic('UNSUPPORTED_STYLE', 'approximated', `${path}.borderRadius`, 'Excel bars have square corners; borderRadius is ignored.', {
          severity: 'info',
          seriesIndex: s.index,
        }),
      );
    }

    // Data labels
    const dlPath = `${path}.dataLabels`;
    const dataLabels = toExcelDataLabels(s.dataLabels, kind, stacked, dlPath, out, s.index);
    if (dataLabels) support(dlPath);
    if (dataLabels && !best) {
      dataLabels.font = null;
      dataLabels.fill = null;
      dataLabels.line = null;
    }

    // Data points
    const dataPoints: ExcelSeriesSpec['dataPoints'] = [];
    const isPie = kind === 'pie' || kind === 'doughnut';
    s.points.forEach((p, j) => {
      const idx = range.pointOffsets[j] ?? -1;
      if (idx < 0) return;
      const pPath = `${path}.data[${j}]`;
      let dpShape: ExcelShapeStyle | null = null;
      let dpMarker: ExcelMarkerSpec | null = null;
      let explosion: number | null = null;
      let dpLabels: ExcelDataLabelsSpec | null = null;
      if (isPie) {
        const pc = p.color ?? paletteColor(j);
        if (best) {
          dpShape = {
            fill: pc ? solidFillOf(pc, s.fillOpacity) : null,
            line: toExcelLine(p.border ?? s.border),
          };
          if (p.color) supportStyle(`${pPath}.color`);
        }
        if (p.sliced !== null) {
          explosion = p.sliced > 0 ? clamp(Math.round((p.sliced / minDim) * 100), 1, 100) : 10;
          support(`${pPath}.sliced`);
        }
        if (!p.visible) {
          out.push(
            createDiagnostic('HIDDEN_POINT', 'approximated', `${pPath}.visible`, 'Hidden pie slice is exported and shown in Excel.', {
              severity: 'info',
              seriesIndex: s.index,
            }),
          );
        }
      } else if (best) {
        if (kind === 'bar' || kind === 'bubble') {
          if (p.color || p.border) {
            dpShape = {
              fill: p.color ? solidFillOf(p.color, s.fillOpacity) : null,
              line: p.border ? toExcelLine(p.border) : null,
            };
            if (p.color) supportStyle(`${pPath}.color`);
          }
        }
        if ((kind === 'line' || kind === 'scatter') && (p.marker || p.color)) {
          const base: Partial<MarkerStyle> = s.marker ?? { enabled: true, symbol: 'circle', radius: 4, fill: null, stroke: null, strokeWidth: 0 };
          const merged: Partial<MarkerStyle> = { ...base, ...(p.marker ?? {}) };
          if (p.color && !p.marker?.fill) merged.fill = p.color;
          dpMarker = toExcelMarker(merged, p.color ?? color, `${pPath}.marker`, styleSink);
          supportStyle(`${pPath}.marker`);
        }
      }
      if (p.dataLabels) {
        const merged = { ...(s.dataLabels ?? {}), ...p.dataLabels } as Parameters<typeof toExcelDataLabels>[0];
        if (merged && !merged.enabled) {
          if (dataLabels) dpLabels = hiddenDataLabels();
        } else {
          dpLabels = toExcelDataLabels(merged, kind, stacked, `${pPath}.dataLabels`, out, s.index);
          if (dpLabels && !best) {
            dpLabels.font = null;
            dpLabels.fill = null;
            dpLabels.line = null;
          }
        }
        if (dpLabels) support(`${pPath}.dataLabels`);
      }
      if (dpShape || dpMarker || explosion !== null || dpLabels) {
        dataPoints.push({ idx, shape: dpShape, marker: dpMarker, explosion, dataLabels: dpLabels });
      }
    });

    const name: ExcelSeriesSpec['name'] =
      range.literalName !== null ? { kind: 'literal', text: range.literalName } : { kind: 'ref', formula: range.nameRef.formula, cache: range.nameRef.cache };
    const categories: ExcelSeriesSpec['categories'] = range.categories
      ? {
          formula: range.categories.formula,
          cache: range.categories.cache.slice(),
          kind: range.categories.kind,
          ...(range.categories.formatCode !== undefined ? { formatCode: range.categories.formatCode } : {}),
        }
      : null;
    const idx = nextIdx++;
    return {
      idx,
      order: idx,
      name,
      categories,
      values: { formula: range.values.formula, cache: range.values.cache.slice(), formatCode: range.values.formatCode },
      bubbleSizes: range.bubbleSizes ? { formula: range.bubbleSizes.formula, cache: range.bubbleSizes.cache.slice(), formatCode: range.bubbleSizes.formatCode } : null,
      shape,
      marker: best ? marker : null,
      smooth: kind === 'line' || kind === 'scatter' ? s.smooth : null,
      dataPoints,
      dataLabels,
      invertIfNegative: false,
    };
  };

  const plotGroups: PlotGroupSpec[] = [];
  for (const g of groups) {
    const series = g.seriesIndices.map((pos) => buildSeries(pos, g));
    const first = model.series[g.seriesIndices[0]!]!;
    const axisIds = axisIdsOf.get(g) ?? [AXIS_IDS.primaryCat, AXIS_IDS.primaryVal];
    switch (g.kind) {
      case 'bar': {
        const bars = first.bars ?? { pointPadding: 0.1, groupPadding: 0.2, borderRadius: 0 };
        const gp = clamp(bars.groupPadding, 0, 1);
        const gapWidth = gp >= 0.5 ? 500 : clamp(Math.round(((gp * 2) / (1 - gp * 2)) * 100), 0, 500);
        const overlap = g.stacking ? 100 : clamp(0 - Math.round(clamp(bars.pointPadding, 0, 1) * 100), -100, 0);
        if (first.bars) support(`series[${first.index}].groupPadding`, `series[${first.index}].pointPadding`);
        plotGroups.push({
          kind: 'bar',
          barDir: g.barDir ?? 'col',
          grouping: g.stacking === 'percent' ? 'percentStacked' : g.stacking === 'normal' ? 'stacked' : 'clustered',
          gapWidth,
          overlap,
          varyColors: false,
          series,
          axisIds,
          dataLabels: null,
        });
        break;
      }
      case 'line':
      case 'area': {
        const grouping = g.stacking === 'percent' ? 'percentStacked' : g.stacking === 'normal' ? 'stacked' : 'standard';
        if (g.kind === 'line') {
          plotGroups.push({ kind: 'line', grouping, varyColors: false, showMarkers: true, series, axisIds, dataLabels: null });
        } else {
          plotGroups.push({ kind: 'area', grouping, varyColors: false, series, axisIds, dataLabels: null });
        }
        break;
      }
      case 'scatter': {
        const lineW = first.line?.width ?? 0;
        const scatterStyle = !(lineW > 0) ? 'marker' : first.smooth ? 'smoothMarker' : 'lineMarker';
        plotGroups.push({ kind: 'scatter', scatterStyle, varyColors: false, series, axisIds, dataLabels: null });
        break;
      }
      case 'bubble':
        plotGroups.push({ kind: 'bubble', varyColors: false, bubbleScale: 100, series, axisIds, dataLabels: null });
        break;
      case 'pie':
      case 'doughnut': {
        const start = first.pie?.startAngle ?? 0;
        const firstSliceAngle = Math.round((((start % 360) + 360) % 360)) % 360;
        if (first.pie) support(`series[${first.index}].startAngle`);
        const end = first.pie?.endAngle ?? null;
        if (end !== null && Math.abs(end - start - 360) > 0.5 && Math.abs(end - start) > 0.5) {
          out.push(
            createDiagnostic('APPROXIMATED_CHART_TYPE', 'approximated', `series[${first.index}].endAngle`, 'Excel pies are always full circles; the partial (semi-circle) pie is drawn as a full circle.', {
              seriesIndex: first.index,
            }),
          );
        }
        if (g.kind === 'pie') {
          plotGroups.push({ kind: 'pie', varyColors: true, firstSliceAngle, series, dataLabels: null });
        } else {
          const holeSize = clamp(Math.round((first.pie?.innerSize ?? 0.5) * 100), 10, 90);
          if (first.pie) support(`series[${first.index}].innerSize`);
          if (g.seriesIndices.length > 1) {
            out.push(
              createDiagnostic('APPROXIMATED_LAYOUT', 'approximated', `series[${model.series[g.seriesIndices[1]!]!.index}].size`, 'Excel doughnut rings all have the same thickness; ring sizes are approximated.', {
                severity: 'info',
              }),
            );
          }
          plotGroups.push({ kind: 'doughnut', varyColors: true, firstSliceAngle, holeSize, series, dataLabels: null });
        }
        break;
      }
    }
  }

  // --- Title, chart area, plot area, legend ----------------------------------------------------------------
  let title: ExcelTextSpec | null = null;
  if (model.title && textOf(model.title).length > 0) {
    const subtitleLines = model.subtitle ? textOf(model.subtitle) : [];
    const subtitleFont = model.subtitle && subtitleLines.length > 0 && best ? fontOf(model.subtitle.font, 'subtitle.style') : null;
    title = textSpec(model.title, 'title', subtitleLines.length > 0 ? { lines: subtitleLines, font: subtitleFont } : null);
    support('title.text');
    if (best) support('title.style');
    if (subtitleLines.length > 0) {
      support('subtitle.text');
      out.push(
        createDiagnostic('APPROXIMATED_LAYOUT', 'approximated', 'subtitle.text', 'Excel charts have a single title; the subtitle is merged into the title as a second line in its own font.', {
          severity: 'info',
        }),
      );
    }
  } else if (model.subtitle && textOf(model.subtitle).length > 0) {
    title = textSpec(model.subtitle, 'subtitle');
    out.push(
      createDiagnostic('APPROXIMATED_LAYOUT', 'approximated', 'subtitle.text', 'The chart has no title; the subtitle is used as the Excel chart title.', { severity: 'info' }),
    );
  }

  const baseFamily = model.title?.font.family ?? model.legend.font?.family ?? DEFAULT_HIGHCHARTS_FONT.family;
  const textDefaults = best ? toExcelFont({ ...DEFAULT_HIGHCHARTS_FONT, family: baseFamily, size: 12 }, 'chart.style.fontFamily', styleSink) : null;

  const chartArea: ExcelShapeStyle = best
    ? { fill: toExcelFill(model.background), line: toExcelLine(model.border) }
    : { fill: null, line: null };
  if (model.background) supportStyle('chart.backgroundColor');
  if (model.border) supportStyle('chart.borderWidth', 'chart.borderColor');

  let manualLayout: ExcelChartSpec['plotArea']['manualLayout'] = null;
  const box = model.plotArea.box;
  if (best && !isPieChart && box && model.width > 0 && model.height > 0) {
    const w = box.width / model.width;
    const h = box.height / model.height;
    if (w > 0.2 && h > 0.2 && w <= 1.0001 && h <= 1.0001) {
      const r = (v: number): number => Math.round(clamp(v, 0, 1) * 10000) / 10000;
      manualLayout = { x: r(box.left / model.width), y: r(box.top / model.height), w: r(w), h: r(h) };
      out.push(
        createDiagnostic('APPROXIMATED_LAYOUT', 'approximated', 'chart.plotArea', 'Plot area pinned to source proportions; Excel positions titles and labels around it itself.', {
          severity: 'info',
          details: { ...manualLayout },
        }),
      );
    }
  }
  const plotArea: ExcelChartSpec['plotArea'] = best
    ? { fill: toExcelFill(model.plotArea.background), line: toExcelLine(model.plotArea.border), manualLayout }
    : { fill: null, line: null, manualLayout: null };
  if (model.plotArea.background) supportStyle('chart.plotBackgroundColor');
  if (model.plotArea.border) supportStyle('chart.plotBorderWidth', 'chart.plotBorderColor');

  let legend: ExcelChartSpec['legend'] = null;
  support('legend.enabled');
  if (model.legend.enabled) {
    const lp = legendPosition(model.legend);
    if (lp.diagnostic) out.push(lp.diagnostic);
    else support('legend.align', 'legend.verticalAlign');
    if (model.legend.overlay) support('legend.floating');
    legend = {
      position: lp.position,
      overlay: lp.overlay,
      font: fontOf(model.legend.font, 'legend.itemStyle'),
      fill: best ? toExcelFill(model.legend.background) : null,
      line: best ? toExcelLine(model.legend.border) : null,
    };
    if (model.legend.font) supportStyle('legend.itemStyle');
    if (model.legend.background) supportStyle('legend.backgroundColor');
    if (model.legend.border) supportStyle('legend.borderWidth', 'legend.borderColor');
    if (model.legend.reversed) {
      out.push(createDiagnostic('UNSUPPORTED_STYLE', 'unsupported', 'legend.reversed', 'Excel cannot reverse the legend order.', { severity: 'info' }));
    }
  }

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

  // --- Chart sheet --------------------------------------------------------------------------------------------
  const widthPx = Math.max(50, Math.round(opts.chartWidth ?? model.width));
  const heightPx = Math.max(50, Math.round(opts.chartHeight ?? model.height));
  const chartAnchor: AnchorSpec = { col0: 0, row0: 2, colOffsetPx: 0, rowOffsetPx: 0, widthPx, heightPx };
  const drawings: DrawingSpec[] = [{ kind: 'chart', chart, anchor: chartAnchor, name: 'Chart 1' }];
  if (opts.referenceImage) {
    const offset = widthPx + 20;
    const col0 = Math.floor(offset / DEFAULT_COLUMN_WIDTH_PX);
    drawings.push({
      kind: 'image',
      png: opts.referenceImage.png,
      anchor: { col0, row0: 2, colOffsetPx: offset - col0 * DEFAULT_COLUMN_WIDTH_PX, rowOffsetPx: 0, widthPx, heightPx },
      name: 'Reference image',
    });
  }
  const titleText = title ? stripControlChars((model.title ? textOf(model.title) : title.lines).join(' ')) : '';
  const chartSheetSpec: SheetSpec = {
    name: chartSheet.name,
    hidden: false,
    columns: [],
    rows: titleText ? [{ row0: 0, cells: [{ col0: 0, row0: 0, value: { type: 'string', value: titleText }, style: { bold: true } }] }] : [],
    freezeHeaderRow: false,
    drawings,
  };

  flush();
  return {
    sheets: [chartSheetSpec, layout.sheet],
    chartSheetName: chartSheet.name,
    dataSheetName: dataSheet.name,
    excelChartType: resolution.excelChartType,
    supportedProperties: [...new Set(supported)],
    blocking: false,
    resolution,
  };
}

