/**
 * Applies explicit ThemeOverrides to a ChartModel (returns a new model; the input is not mutated).
 */

import type {
  AxisModel,
  Color,
  DataLabelStyle,
  Font,
  PointModel,
  SeriesModel,
  Stroke,
  TextBlock,
} from '../types/chart-model';
import { createDiagnostic, type Diagnostic } from '../types/diagnostics';
import type { ChartModel } from '../types/chart-model';
import type { FontOverride, ThemeOverrides } from '../types/public-api';
import { parseColor } from '../utils/colors';
import { DEFAULT_HIGHCHARTS_FONT } from '../translators/typography-translator';

const DEFAULT_GRID_COLOR: Color = { r: 0xe6, g: 0xe6, b: 0xe6, a: 1, source: '#e6e6e6' };

function cloneFont(f: Font): Font {
  return { ...f, color: f.color ? { ...f.color } : null };
}

function defaultFont(): Font {
  return cloneFont(DEFAULT_HIGHCHARTS_FONT);
}

/** Returns the model with `overrides` applied. Without overrides the same model is returned. */
export function applyThemeOverrides(model: ChartModel, overrides: ThemeOverrides | undefined): ChartModel {
  if (!overrides) return model;
  const warnings: Diagnostic[] = [...model.warnings];
  const color = (value: string | undefined, property: string): Color | null => {
    if (value === undefined) return null;
    const c = parseColor(value);
    if (!c) {
      warnings.push(
        createDiagnostic(
          'UNRESOLVED_COLOR',
          'approximated',
          property,
          `Theme override color "${value}" could not be parsed and is ignored.`,
          {
            details: { value },
          },
        ),
      );
    }
    return c;
  };
  const mergeFont = (
    base: Font | null,
    o: FontOverride | undefined,
    property: string,
    create: boolean,
  ): Font | null => {
    const family = overrides.fontFamily;
    if (!o && family === undefined) return base;
    if (!base && !create && !o) return base;
    const f = base ? cloneFont(base) : defaultFont();
    if (family !== undefined && family.trim() !== '') f.family = family;
    if (o) {
      if (o.family !== undefined && o.family.trim() !== '') f.family = o.family;
      if (o.size !== undefined && Number.isFinite(o.size) && o.size > 0) f.size = o.size;
      if (o.bold !== undefined) f.bold = o.bold;
      if (o.italic !== undefined) f.italic = o.italic;
      if (o.color !== undefined) {
        const c = color(o.color, `${property}.color`);
        if (c) f.color = c;
      }
    }
    return f;
  };
  const mergeText = (tb: TextBlock | null, o: FontOverride | undefined, property: string): TextBlock | null => {
    if (!tb) return null;
    const font = mergeFont(tb.font, o, property, true);
    return font === tb.font ? tb : { ...tb, font: font! };
  };

  const next: ChartModel = { ...model, warnings };

  // Palette.
  let palette: Color[] | null = null;
  if (overrides.colors && overrides.colors.length > 0) {
    palette = overrides.colors
      .map((c, i) => color(c, `themeOverrides.colors[${i}]`))
      .filter((c): c is Color => c !== null);
    if (palette.length > 0) next.colors = palette;
    else palette = null;
  }

  // Backgrounds.
  const chartBg = color(overrides.chartBackground, 'themeOverrides.chartBackground');
  if (chartBg) next.background = { type: 'solid', color: chartBg };
  const plotBg = color(overrides.plotBackground, 'themeOverrides.plotBackground');
  next.plotArea = { ...model.plotArea, box: model.plotArea.box ? { ...model.plotArea.box } : null };
  if (plotBg) next.plotArea.background = { type: 'solid', color: plotBg };

  // Titles and legend.
  next.title = mergeText(model.title, overrides.title, 'themeOverrides.title');
  next.subtitle = mergeText(model.subtitle, overrides.subtitle, 'themeOverrides.subtitle');
  next.legend = {
    ...model.legend,
    font: mergeFont(model.legend.font, overrides.legend, 'themeOverrides.legend', true),
  };

  // Axes.
  const gridColor = color(overrides.gridLineColor, 'themeOverrides.gridLineColor');
  const gridWidth =
    overrides.gridLineWidth !== undefined && Number.isFinite(overrides.gridLineWidth) && overrides.gridLineWidth >= 0
      ? overrides.gridLineWidth
      : null;
  const mapAxis = (a: AxisModel, isY: boolean): AxisModel => {
    const out: AxisModel = {
      ...a,
      categories: a.categories ? [...a.categories] : null,
      title: mergeText(a.title, overrides.axisTitle, 'themeOverrides.axisTitle'),
      labels: { ...a.labels, font: mergeFont(a.labels.font, overrides.axisLabels, 'themeOverrides.axisLabels', true) },
    };
    if (out.gridLines) {
      const g: Stroke = { ...out.gridLines };
      if (gridColor) g.color = gridColor;
      if (gridWidth !== null) g.width = gridWidth;
      out.gridLines = g;
    } else if (isY && gridWidth !== null && gridWidth > 0) {
      out.gridLines = { color: gridColor ?? DEFAULT_GRID_COLOR, width: gridWidth, dash: 'solid' };
    }
    return out;
  };
  next.xAxes = model.xAxes.map((a) => mapAxis(a, false));
  next.yAxes = model.yAxes.map((a) => mapAxis(a, true));

  // Series.
  const mapLabels = <T extends Partial<DataLabelStyle> | null>(dl: T, property: string): T => {
    if (!dl) return dl;
    if (!overrides.dataLabels && overrides.fontFamily === undefined) return dl;
    const font = mergeFont((dl as Partial<DataLabelStyle>).font ?? null, overrides.dataLabels, property, true);
    return { ...dl, font };
  };
  next.series = model.series.map((s, pos) => {
    let out: SeriesModel = {
      ...s,
      dataLabels: mapLabels(s.dataLabels, 'themeOverrides.dataLabels'),
      points: s.points,
    };
    const isPieLike = s.kind === 'pie' || s.kind === 'doughnut';
    if (palette) {
      if (isPieLike) {
        out.points = s.points.map((p, j) => ({ ...p, color: palette![j % palette!.length]! }));
      } else {
        out = recolor(out, palette[pos % palette.length]!);
      }
    }
    const so = overrides.series?.[pos] ?? overrides.series?.[s.index];
    if (so) {
      if (so.color !== undefined) {
        const c = color(so.color, `themeOverrides.series[${pos}].color`);
        if (c) out = recolor(out, c);
      }
      if (so.lineWidth !== undefined && Number.isFinite(so.lineWidth) && so.lineWidth >= 0) {
        out.line = out.line
          ? { ...out.line, width: so.lineWidth }
          : { color: out.color, width: so.lineWidth, dash: 'solid' };
      }
      if (so.fillOpacity !== undefined && Number.isFinite(so.fillOpacity)) {
        out.fillOpacity = Math.min(1, Math.max(0, so.fillOpacity));
      }
    }
    if (overrides.dataLabels || overrides.fontFamily !== undefined) {
      out.points = out.points.map((p) =>
        p.dataLabels ? { ...p, dataLabels: mapLabels(p.dataLabels, 'themeOverrides.dataLabels') } : p,
      );
    }
    return out;
  });

  return next;
}

/** New series with its color (and the fill/line/marker colors derived from it) replaced. */
function recolor(s: SeriesModel, c: Color): SeriesModel {
  const old = s.color;
  const same = (x: Color | null): boolean =>
    x === null || (old !== null && x.r === old.r && x.g === old.g && x.b === old.b);
  const out: SeriesModel = { ...s, color: c };
  if (!s.fill || s.fill.type === 'solid') out.fill = { type: 'solid', color: { ...c } };
  if (s.line) out.line = { ...s.line, color: same(s.line.color) ? c : s.line.color };
  if (s.marker) out.marker = { ...s.marker, fill: same(s.marker.fill) ? c : s.marker.fill };
  const isPieLike = s.kind === 'pie' || s.kind === 'doughnut';
  if (!isPieLike && s.points.some((p) => p.color !== null && same(p.color))) {
    out.points = s.points.map((p: PointModel) => (p.color !== null && same(p.color) ? { ...p, color: c } : p));
  }
  return out;
}
