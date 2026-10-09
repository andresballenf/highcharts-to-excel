/**
 * Single-axis assembly: `buildAxis` (scaling, units, number format, gridlines, labels, title and the
 * supported-property report of one Excel axis), value and category axis placement and crossing, the
 * axis ids, and the Excel date formats and base time unit of datetime x values. The axis pairs of a
 * chart and the axis ids of each plot group are wired up in `axis-pairs.ts`.
 */

import type { AxisModel } from '../types/chart-model';
import { createDiagnostic } from '../types/diagnostics';
import type { ExcelAxisSpec, ExcelLineSpec } from '../excel/writer-interface';
import { guessExcelDateFormatForRange, highchartsDateFormatToExcel, msToExcelSerial } from '../utils/dates';
import { isValidExcelFormatCode } from '../utils/format-code';
import { clamp } from '../utils/units';
import { toExcelLine } from './style-mapping';
import type { TranslateContext } from './translate-context';

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

/** True when a timestamp (ms) is not at midnight UTC. */
export function hasTimeOfDay(ms: number): boolean {
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
  if (axis?.labels.format?.kind === 'excel' && isValidExcelFormatCode(axis.labels.format.code))
    return axis.labels.format.code;
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

export interface AxisParams {
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

/** One Excel axis from its IR axis and placement; reports supported properties unless hidden/aux. */
export function buildAxis(ctx: TranslateContext, p: AxisParams): ExcelAxisSpec {
  const { best, support, supportStyle } = ctx;
  const a = p.axis;
  const report = p.reportSupport !== false && !p.forceDeleted && a !== null;
  const isValueLike = p.kind === 'val' || p.kind === 'date';
  const deleted = p.forceDeleted === true || (a !== null && !a.visible);
  const toScale = (v: number | null): number | null =>
    v === null || !Number.isFinite(v) ? null : p.serial ? msToExcelSerial(v) : v;
  let min: number | null = null;
  let max: number | null = null;
  if (a && isValueLike) {
    min = toScale(a.min);
    max = toScale(a.max);
    if (report && a.min !== null) support(`${p.path}.min`);
    if (report && a.max !== null) support(`${p.path}.max`);
  } else if (a && (a.min !== null || a.max !== null) && report) {
    ctx.out.push(
      createDiagnostic(
        'APPROXIMATED_AXIS_SCALE',
        'approximated',
        `${p.path}.min`,
        'Excel category axes cannot be cropped with min/max; all categories are shown.',
        { severity: 'info' },
      ),
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
  if (numberFormat === null && a?.labels.format?.kind === 'excel' && isValidExcelFormatCode(a.labels.format.code)) {
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
  return {
    id: p.id,
    kind: p.kind,
    position: p.position,
    crossAxisId: p.crossAxisId,
    deleted,
    title: p.forceDeleted ? null : ctx.textSpec(a?.title ?? null, `${p.path}.title`),
    numberFormat,
    majorGridlines: p.forceDeleted ? null : gridOf(a?.gridLines ?? null, `${p.path}.gridLine`),
    minorGridlines: p.forceDeleted ? null : gridOf(a?.minorGridLines ?? null, `${p.path}.minorGridLine`),
    axisLine: best && a?.axisLine ? toExcelLine(a.axisLine) : null,
    labels: {
      position: labelsHidden ? 'none' : p.labelsLow ? (p.crossAxisReversed ? 'high' : 'low') : 'nextTo',
      font: ctx.fontOf(a?.labels.font ?? null, `${p.path}.labels.style`),
      rotation,
    },
    scaling: { min, max, orientation: a?.reversed ? 'maxMin' : 'minMax', logBase },
    majorUnit,
    minorUnit,
    crosses: p.crosses ?? 'autoZero',
    majorTickMark: a?.tickMarks && a.tickMarks.width > 0 ? 'out' : 'none',
    dateAxis: p.kind === 'date' ? { baseTimeUnit: p.baseTimeUnit ?? null } : null,
  };
}

/**
 * Value axis placement. Horizontal bar charts whose category axis is reversed (Highcharts bars list the
 * first category at the top: the IR x axis arrives with `reversed: true`) need the standard Excel
 * "categories in reverse order + value axis crosses at maximum category" combination to keep the
 * value axis at the bottom; a secondary value axis then crosses at the minimum (top).
 */
export function valuePlacement(
  axis: AxisModel | null,
  horizontal: boolean,
  secondary: boolean,
  catReversed = false,
): { position: ExcelAxisSpec['position']; crosses: ExcelAxisSpec['crosses'] } {
  // In Excel an axis sits where it crosses its perpendicular axis, so a reversed perpendicular axis
  // flips "crosses at min/max". The general rule for an axis A crossing axis B is:
  // A is on the far side  <=>  (A.opposite XOR B.reversed)  <=>  crosses 'max'.
  const far = (secondary || axis?.opposite === true) !== catReversed;
  if (secondary || axis?.opposite) return { position: horizontal ? 't' : 'r', crosses: far ? 'max' : 'min' };
  if (axis?.crossing !== null && axis?.crossing !== undefined && Number.isFinite(axis.crossing)) {
    return { position: horizontal ? 'b' : 'l', crosses: { at: axis.crossing } };
  }
  return { position: horizontal ? 'b' : 'l', crosses: far ? 'max' : 'autoZero' };
}

/** Category/X axis crossing: far side when (opposite XOR the value axis is reversed). */
export function categoryCrosses(x: AxisModel | null, y: AxisModel | null): ExcelAxisSpec['crosses'] {
  return (x?.opposite === true) !== (y?.reversed === true) ? 'max' : 'autoZero';
}
