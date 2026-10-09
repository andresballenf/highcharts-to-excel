/**
 * IR style → Excel spec helpers (fills, lines, fonts, markers, data labels, legend).
 */

import type { Color, DataLabelStyle, Fill, Font, LegendModel, MarkerStyle, Stroke } from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import type {
  ExcelDataLabelPosition,
  ExcelDataLabelsSpec,
  ExcelFillSpec,
  ExcelFontSpec,
  ExcelLineSpec,
  ExcelMarkerSpec,
  OoxmlDash,
  PlotGroupSpec,
} from '../excel/writer-interface';
import { colorToHex } from '../utils/colors';
import { clamp } from '../utils/units';
import { isValidExcelFormatCode } from '../utils/format-code';
import { fontToOoxml } from '../translators/typography-translator';
import { dashStyleToOoxml, markerRadiusToOoxmlSize, markerSymbolToOoxml } from '../translators/style-translator';

function alphaOf(a: number, opacity = 1): number {
  const v = (Number.isFinite(a) ? a : 1) * (Number.isFinite(opacity) ? opacity : 1);
  return Math.round(clamp(v, 0, 1) * 10000) / 10000;
}

/** Solid fill for a color (alpha × opacity); fully transparent → no fill. */
export function solidFillOf(color: Color, opacity = 1): ExcelFillSpec {
  const alpha = alphaOf(color.a, opacity);
  if (alpha === 0) return { type: 'none' };
  return { type: 'solid', hex: colorToHex(color), alpha };
}

/** Fill → Excel fill. `opacity` multiplies every color's alpha. */
export function toExcelFill(fill: Fill | null, opacity = 1): ExcelFillSpec | null {
  if (!fill) return null;
  switch (fill.type) {
    case 'none':
      return { type: 'none' };
    case 'solid':
      return solidFillOf(fill.color, opacity);
    case 'gradient': {
      if (fill.stops.length === 0) return null;
      return {
        type: 'gradient',
        stops: fill.stops.map((s) => ({
          pos: clamp(s.offset, 0, 1),
          hex: colorToHex(s.color),
          alpha: alphaOf(s.color.a, opacity),
        })),
        angle: Number.isFinite(fill.angle) ? fill.angle : 0,
      };
    }
    default:
      return null;
  }
}

export const NO_LINE: Readonly<ExcelLineSpec> = Object.freeze({
  widthPx: 0,
  hex: null,
  alpha: 1,
  dash: 'solid',
  noFill: true,
});

/**
 * Stroke → Excel line. Width 0 (or a fully transparent color) → explicit "no line".
 * `fallbackColor` is used when the stroke has no color of its own (e.g. series lines).
 */
export function toExcelLine(stroke: Stroke | null, fallbackColor: Color | null = null): ExcelLineSpec | null {
  if (!stroke) return null;
  const color = stroke.color ?? fallbackColor;
  if (!(stroke.width > 0) || (color !== null && color.a === 0)) return { ...NO_LINE };
  return {
    widthPx: stroke.width,
    hex: color ? colorToHex(color) : null,
    alpha: color ? alphaOf(color.a) : 1,
    dash: dashStyleToOoxml(stroke.dash) as OoxmlDash,
    noFill: false,
  };
}

/** Font → Excel font (via fontToOoxml); its diagnostic, if any, goes to `sink`. */
export function toExcelFont(
  font: Font | null,
  property = 'style.fontFamily',
  sink?: Diagnostic[],
): ExcelFontSpec | null {
  if (!font) return null;
  const r = fontToOoxml(font, property);
  if (r.diagnostic && sink) sink.push(r.diagnostic);
  return {
    typeface: r.typeface,
    sizeHundredthsPt: r.sizeHundredthsPt,
    bold: r.bold,
    italic: r.italic,
    colorHex: r.colorHex,
  };
}

/** Marker → Excel marker. Disabled markers → symbol 'none'. */
export function toExcelMarker(
  marker: Partial<MarkerStyle> | null,
  seriesColor: Color | null,
  property = 'marker',
  sink?: Diagnostic[],
): ExcelMarkerSpec | null {
  if (!marker) return null;
  if (marker.enabled === false) return { symbol: 'none', size: 5, fill: null, line: null };
  const sym = markerSymbolToOoxml(marker.symbol ?? 'circle', `${property}.symbol`);
  if (sym.diagnostic && sink) sink.push(sym.diagnostic);
  const fillColor = marker.fill ?? seriesColor;
  const strokeWidth = marker.strokeWidth ?? 0;
  const strokeColor = marker.stroke ?? seriesColor;
  return {
    symbol: sym.symbol,
    size: markerRadiusToOoxmlSize(marker.radius ?? 4),
    fill: fillColor ? solidFillOf(fillColor) : null,
    line: strokeWidth > 0 ? toExcelLine({ color: strokeColor, width: strokeWidth, dash: 'solid' }) : { ...NO_LINE },
  };
}

const VALID_POSITIONS: Record<PlotGroupSpec['kind'] | 'barStacked', ReadonlySet<ExcelDataLabelPosition>> = {
  bar: new Set(['ctr', 'inEnd', 'inBase', 'outEnd']),
  barStacked: new Set(['ctr', 'inEnd', 'inBase']),
  line: new Set(['ctr', 'l', 'r', 't', 'b']),
  scatter: new Set(['ctr', 'l', 'r', 't', 'b']),
  bubble: new Set(['ctr', 'l', 'r', 't', 'b']),
  pie: new Set(['ctr', 'inEnd', 'outEnd', 'bestFit']),
  area: new Set(),
  doughnut: new Set(),
};

const POSITION_MAP: Readonly<Record<DataLabelStyle['position'], ExcelDataLabelPosition | null>> = {
  auto: null,
  outsideEnd: 'outEnd',
  insideEnd: 'inEnd',
  insideBase: 'inBase',
  center: 'ctr',
  above: 't',
  below: 'b',
  left: 'l',
  right: 'r',
  bestFit: 'bestFit',
};

/**
 * Data labels → Excel data labels. Returns null when labels are disabled (or absent).
 * Positions Excel does not allow for the chart type become null (Excel default) with a diagnostic.
 */
export function toExcelDataLabels(
  dl: Partial<DataLabelStyle> | null,
  groupKind: PlotGroupSpec['kind'],
  isStacked: boolean,
  property = 'dataLabels',
  sink?: Diagnostic[],
  seriesIndex?: number,
): ExcelDataLabelsSpec | null {
  if (!dl?.enabled) return null;
  const extra = seriesIndex !== undefined ? { seriesIndex } : {};
  let position: ExcelDataLabelPosition | null = null;
  const requested = dl.position ?? 'auto';
  if (groupKind !== 'area' && groupKind !== 'doughnut') {
    position = POSITION_MAP[requested] ?? null;
    if (position === 'outEnd' && groupKind === 'bar' && isStacked) {
      position = 'inEnd';
      sink?.push(
        createDiagnostic(
          'APPROXIMATED_DATA_LABELS',
          'approximated',
          `${property}.position`,
          'Excel cannot place labels outside stacked bars; labels are placed at the inside end.',
          extra,
        ),
      );
    }
    const valid = VALID_POSITIONS[groupKind === 'bar' && isStacked ? 'barStacked' : groupKind];
    if (position !== null && !valid.has(position)) {
      sink?.push(
        createDiagnostic(
          'APPROXIMATED_DATA_LABELS',
          'approximated',
          `${property}.position`,
          `Label position "${requested}" is not available for this Excel chart type; Excel's default is used.`,
          extra,
        ),
      );
      position = null;
    }
  } else if (requested !== 'auto') {
    sink?.push(
      createDiagnostic(
        'APPROXIMATED_DATA_LABELS',
        'approximated',
        `${property}.position`,
        'Excel does not allow positioning labels on area/doughnut charts; the default position is used.',
        {
          ...extra,
          severity: 'info',
        },
      ),
    );
  }
  let numberFormat: string | null = null;
  if (dl.format && dl.format.kind === 'excel' && isValidExcelFormatCode(dl.format.code)) numberFormat = dl.format.code;
  let showValue = dl.showValue ?? true;
  let showPercent = dl.showPercentage ?? false;
  if (showPercent && groupKind !== 'pie' && groupKind !== 'doughnut') {
    // Excel ignores showPercent outside pie/doughnut charts.
    showPercent = false;
    const onlyPercent = !showValue && !(dl.showCategoryName ?? false) && !(dl.showSeriesName ?? false);
    if (onlyPercent) {
      showValue = true;
      numberFormat = null; // the format was written for percentages, not for the values
    }
    sink?.push(
      createDiagnostic(
        'APPROXIMATED_DATA_LABELS',
        'approximated',
        `${property}.format`,
        onlyPercent
          ? 'Excel shows percentages only on pie charts; values are shown instead.'
          : 'Excel shows percentages only on pie charts; the percentage part of the label is dropped.',
        extra,
      ),
    );
  }
  return {
    showValue,
    showCategoryName: dl.showCategoryName ?? false,
    showSeriesName: dl.showSeriesName ?? false,
    showPercent,
    position,
    numberFormat,
    font: toExcelFont(dl.font ?? null, `${property}.style`, sink),
    fill: toExcelFill(dl.background ?? null),
    line: toExcelLine(dl.border ?? null),
  };
}

/** Data labels spec that hides a single point's label. */
export function hiddenDataLabels(): ExcelDataLabelsSpec {
  return {
    showValue: false,
    showCategoryName: false,
    showSeriesName: false,
    showPercent: false,
    position: null,
    numberFormat: null,
    font: null,
    fill: null,
    line: null,
  };
}

export function legendPosition(legend: LegendModel): {
  position: 'b' | 't' | 'l' | 'r' | 'tr';
  overlay: boolean;
  diagnostic?: Diagnostic;
} {
  const approx = (
    pos: 'b' | 't',
    where: string,
  ): { position: 'b' | 't'; overlay: boolean; diagnostic: Diagnostic } => ({
    position: pos,
    overlay: legend.overlay,
    diagnostic: createDiagnostic(
      'APPROXIMATED_LEGEND_POSITION',
      'approximated',
      'legend.align',
      `Excel has no ${where} legend position; the legend is centered at the ${pos === 'b' ? 'bottom' : 'top'}.`,
      { details: { position: legend.position } },
    ),
  });
  switch (legend.position) {
    case 'top':
      return { position: 't', overlay: legend.overlay };
    case 'bottom':
      return { position: 'b', overlay: legend.overlay };
    case 'left':
      return { position: 'l', overlay: legend.overlay };
    case 'right':
      return { position: 'r', overlay: legend.overlay };
    case 'topRight':
      return { position: 'tr', overlay: legend.overlay };
    case 'topLeft':
      return approx('t', 'top-left');
    case 'bottomLeft':
      return approx('b', 'bottom-left');
    case 'bottomRight':
      return approx('b', 'bottom-right');
    default:
      return { position: 'b', overlay: legend.overlay };
  }
}
