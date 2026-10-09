/**
 * Series and data-point styling: one `ExcelSeriesSpec` per exported series, with its cell references,
 * shape (fill/line per group kind), marker, smoothing, series data labels, and the per-point
 * overrides (slice colors and explosion, bar/bubble colors, marker merges, point data labels).
 */

import type { Color, Fill, MarkerStyle, SeriesModel } from '../types/chart-model';
import { createDiagnostic } from '../types/diagnostics';
import type { ExcelDataLabelsSpec, ExcelMarkerSpec, ExcelSeriesSpec, ExcelShapeStyle } from '../excel/writer-interface';
import { HIGHCHARTS_DEFAULT_PALETTE, parseColor } from '../utils/colors';
import { clamp } from '../utils/units';
import { fillToSolidColor } from '../translators/color-translator';
import type { PlotGroupPlan } from './chart-type-registry';
import type { SeriesRange } from './data-layout';
import {
  NO_LINE,
  hiddenDataLabels,
  solidFillOf,
  toExcelDataLabels,
  toExcelFill,
  toExcelLine,
  toExcelMarker,
} from './style-mapping';
import type { TranslateContext } from './translate-context';

/**
 * Builds the series spec of `model.series[pos]` inside plot group `g`; `range` builds the Excel
 * series of a range series (hidden base, visible range, and the Down part of a split column range).
 */
export type SeriesBuilder = ((pos: number, g: PlotGroupPlan) => ExcelSeriesSpec) & {
  range(pos: number, g: PlotGroupPlan): RangeSeriesSpecs;
};

/** Excel series of one range series, in stacking order; `hiddenFromLegend` lists helper idx values. */
export interface RangeSeriesSpecs {
  specs: ExcelSeriesSpec[];
  hiddenFromLegend: number[];
}

/** Marker used by scatter series (and point marker merges) when the source has none. */
const DEFAULT_MARKER: Readonly<Partial<MarkerStyle>> = {
  enabled: true,
  symbol: 'circle',
  radius: 4,
  fill: null,
  stroke: null,
  strokeWidth: 0,
};

const fillOrColor = (fill: Fill | null, color: Color | null): Fill | null =>
  fill ?? (color ? { type: 'solid', color } : null);

/** Drops the styling of data labels (minimal fidelity keeps only what they show and where). */
function unstyled(dl: ExcelDataLabelsSpec | null, best: boolean): ExcelDataLabelsSpec | null {
  if (dl && !best) {
    dl.font = null;
    dl.fill = null;
    dl.line = null;
  }
  return dl;
}

/**
 * Returns a series builder bound to the run. Series indices (`idx`/`order`) are assigned in call
 * order, so call it once per exported series in plot-group order.
 */
export function createSeriesBuilder(
  ctx: TranslateContext,
  rangeByPos: ReadonlyMap<number, SeriesRange>,
): SeriesBuilder {
  const { model, best, out, styleSink, support, supportStyle } = ctx;
  const palette: Color[] =
    model.colors.length > 0 ? model.colors : HIGHCHARTS_DEFAULT_PALETTE.map((c) => parseColor(c)!).filter(Boolean);
  const paletteColor = (i: number): Color | null =>
    palette[((i % palette.length) + palette.length) % palette.length] ?? null;
  const seriesColorOf = (s: SeriesModel): Color | null => s.color ?? fillToSolidColor(s.fill) ?? paletteColor(s.index);
  const minDim = Math.max(1, Math.min(model.width, model.height));
  let nextIdx = 0;

  const build = (pos: number, g: PlotGroupPlan): ExcelSeriesSpec => {
    const s = model.series[pos]!;
    const range = rangeByPos.get(pos)!;
    const path = `series[${s.index}]`;
    const color = seriesColorOf(s);
    const kind = g.kind;
    const stacked = g.stacking !== null;
    const filledRadar = kind === 'radar' && g.radarStyle === 'filled';
    /** Line-like series: a line with markers (line groups, unfilled radar). */
    const lineLike = kind === 'line' || (kind === 'radar' && !filledRadar);
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
        case 'radar':
          if (kind === 'radar' && !filledRadar) break;
          shape.fill = toExcelFill(fillOrColor(s.fill, color), s.fillOpacity);
          supportStyle(`${path}.color`);
          if ((kind === 'area' || filledRadar) && s.fillOpacity !== 1) supportStyle(`${path}.fillOpacity`);
          break;
        default:
          break;
      }
      switch (kind) {
        case 'radar':
          if (filledRadar) {
            shape.line = s.line ? toExcelLine(s.line, color) : null;
            break;
          }
          shape.line = toExcelLine(s.line ?? { color, width: 2, dash: 'solid' }, color);
          supportStyle(`${path}.color`, `${path}.lineWidth`);
          break;
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
      if (lineLike || kind === 'scatter') {
        const m: Partial<MarkerStyle> | null = s.marker ?? (kind === 'scatter' ? { ...DEFAULT_MARKER } : null);
        marker = toExcelMarker(m, color, `${path}.marker`, styleSink);
        if (s.marker) supportStyle(`${path}.marker`);
      }
    } else {
      shape = { fill: null, line: null };
      if (kind === 'scatter' && !(s.line && s.line.width > 0)) shape.line = { ...NO_LINE };
    }
    if (best && kind === 'bar' && s.bars && s.bars.borderRadius > 0) {
      styleSink.push(
        createDiagnostic(
          'UNSUPPORTED_STYLE',
          'approximated',
          `${path}.borderRadius`,
          'Excel bars have square corners; borderRadius is ignored.',
          { severity: 'info', seriesIndex: s.index },
        ),
      );
    }

    // Data labels
    const dlPath = `${path}.dataLabels`;
    const dataLabels = toExcelDataLabels(s.dataLabels, kind, stacked, dlPath, out, s.index);
    if (dataLabels) support(dlPath);
    unstyled(dataLabels, best);

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
            createDiagnostic(
              'HIDDEN_POINT',
              'approximated',
              `${pPath}.visible`,
              'Hidden pie slice is exported and shown in Excel.',
              { severity: 'info', seriesIndex: s.index },
            ),
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
        if ((lineLike || kind === 'scatter') && (p.marker || p.color)) {
          const base: Partial<MarkerStyle> = s.marker ?? { ...DEFAULT_MARKER };
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
          dpLabels = unstyled(toExcelDataLabels(merged, kind, stacked, `${pPath}.dataLabels`, out, s.index), best);
        }
        if (dpLabels) support(`${pPath}.dataLabels`);
      }
      if (dpShape || dpMarker || explosion !== null || dpLabels) {
        dataPoints.push({ idx, shape: dpShape, marker: dpMarker, explosion, dataLabels: dpLabels });
      }
    });

    const name: ExcelSeriesSpec['name'] =
      range.literalName !== null
        ? { kind: 'literal', text: range.literalName }
        : { kind: 'ref', formula: range.nameRef.formula, cache: range.nameRef.cache };
    const categories: ExcelSeriesSpec['categories'] = range.categories
      ? {
          formula: range.categories.formula,
          cache: range.categories.cache.slice(),
          kind: range.categories.kind,
          ...(range.categories.formatCode !== undefined ? { formatCode: range.categories.formatCode } : {}),
        }
      : null;
    const idx = nextIdx++;
    const spec: ExcelSeriesSpec = {
      idx,
      order: idx,
      name,
      categories,
      values: { formula: range.values.formula, cache: range.values.cache.slice(), formatCode: range.values.formatCode },
      bubbleSizes: range.bubbleSizes
        ? {
            formula: range.bubbleSizes.formula,
            cache: range.bubbleSizes.cache.slice(),
            formatCode: range.bubbleSizes.formatCode,
          }
        : null,
      shape,
      marker: best ? marker : null,
      smooth: kind === 'line' || kind === 'scatter' ? s.smooth : null,
      dataPoints,
      dataLabels,
      invertIfNegative: false,
    };
    if (range.errorBars) {
      const e = model.series[range.errorBars.seriesIndex]!;
      const ePath = `series[${e.index}]`;
      const ref = (r: { formula: string; cache: (number | null)[]; formatCode: string }) => ({
        formula: r.formula,
        cache: r.cache.slice(),
        formatCode: r.formatCode,
      });
      spec.errorBars = {
        plus: ref(range.errorBars.plus),
        minus: ref(range.errorBars.minus),
        line: best ? toExcelLine(e.line ?? { color: e.color, width: 1, dash: 'solid' }, e.color) : null,
      };
      support(`${ePath}.type`, `${ePath}.linkedTo`, `${ePath}.data`);
      if (best) supportStyle(`${ePath}.color`);
    }
    return spec;
  };

  const range = (pos: number, g: PlotGroupPlan): RangeSeriesSpecs => {
    const s = model.series[pos]!;
    const parts = rangeByPos.get(pos)!.rangeParts!;
    const baseIdx = nextIdx++;
    const main = build(pos, g);
    // Labels would show the Range/Up helper values, not the low and high ends.
    if (main.dataLabels || main.dataPoints.some((dp) => dp.dataLabels)) {
      out.push(
        createDiagnostic(
          'APPROXIMATED_DATA_LABELS',
          'approximated',
          `series[${s.index}].dataLabels`,
          'Excel labels one value per point; the low/high labels of the range series are not exported.',
          { seriesIndex: s.index },
        ),
      );
    }
    main.dataLabels = null;
    main.dataPoints = main.dataPoints
      .map((dp) => ({ ...dp, dataLabels: null }))
      .filter((dp) => dp.shape || dp.marker || dp.explosion !== null);
    const cloneCats = (): ExcelSeriesSpec['categories'] =>
      main.categories ? { ...main.categories, cache: main.categories.cache.slice() } : null;
    const column = (c: { formula: string; cache: (number | null)[]; formatCode: string }) => ({
      formula: c.formula,
      cache: c.cache.slice(),
      formatCode: c.formatCode,
    });
    // The base is invisible in every fidelity: Excel's automatic fill would show it.
    const base: ExcelSeriesSpec = {
      idx: baseIdx,
      order: baseIdx,
      name: { kind: 'ref', formula: parts.base.name.formula, cache: parts.base.name.cache },
      categories: cloneCats(),
      values: column(parts.base),
      bubbleSizes: null,
      shape: { fill: { type: 'none' }, line: { ...NO_LINE } },
      marker: null,
      smooth: null,
      dataPoints: [],
      dataLabels: null,
      invertIfNegative: false,
    };
    const specs = [base, main];
    const hiddenFromLegend = [baseIdx];
    if (parts.down && best) {
      // Up and Down meet at zero: a border there would draw a seam across the bar.
      main.shape = { ...main.shape, line: { ...NO_LINE } };
    }
    if (parts.down) {
      const downIdx = nextIdx++;
      specs.push({
        ...main,
        idx: downIdx,
        order: downIdx,
        categories: cloneCats(),
        values: column(parts.down),
        dataPoints: main.dataPoints.map((dp) => ({ ...dp })),
      });
      hiddenFromLegend.push(downIdx);
    }
    return { specs, hiddenFromLegend };
  };

  return Object.assign(build, { range });
}
