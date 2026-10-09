/**
 * xl/charts/chartN.xml builder (DrawingML chart, ECMA-376 CT_ChartSpace).
 *
 * Element order follows the schema sequences exactly; Excel rejects the file (or silently drops
 * the chart) when it does not. Absent optional elements are omitted.
 */
import {
  EXCEL_MAX_SERIES_PER_CHART,
  type ExcelAxisSpec,
  type ExcelChartSpec,
  type ExcelDataLabelPosition,
  type ExcelDataLabelsSpec,
  type ExcelMarkerSpec,
  type ExcelSeriesRef,
  type ExcelSeriesSpec,
  type PlotGroupSpec,
} from './writer-interface';
import { fillXml, lineXml, spPrXml, titleXml, txPrXml } from './drawingml-xml';
import { assertExcelFormatCode } from '../utils/format-code';
import {
  clampInt,
  clampNum,
  escapeAttr,
  escapeXml,
  escapeXstring,
  finite,
  formatNumber,
  valEl,
  xmlDocument,
} from './xml';

/** Largest value of an ST_UnsignedInt idx/order. */
const MAX_UNSIGNED_INT = 4_294_967_295;
/** Largest tickLblSkip / tickMarkSkip Excel accepts (ST_Skip is 1..31999). */
const MAX_SKIP = 31_999;

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

type GroupKind = PlotGroupSpec['kind'];

// ---------------------------------------------------------------------------
// Data references
// ---------------------------------------------------------------------------

function formulaXml(formula: string, what: string): string {
  const f = String(formula ?? '')
    .trim()
    .replace(/^=/, '');
  if (f === '') throw new Error(`Chart ${what} reference has an empty formula`);
  return `<c:f>${escapeXml(f)}</c:f>`;
}

/**
 * Excel 2007 capped series at 32,000 points; Excel 2010+ is bounded by memory only. The translator
 * reports ROW_LIMIT_EXCEEDED as a warning above that guidance, so the writer does not refuse here.
 */
function checkPointCount(n: number, what: string): void {
  if (!Number.isFinite(n) || n < 0) throw new Error(`Chart ${what} has an invalid point count`);
}

/** Number for a cache point, or null when the point must be skipped. */
function cacheNumber(v: unknown): number | null {
  if (typeof v === 'number') return Number.isFinite(v) ? (v === 0 ? 0 : v) : null;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v);
    return Number.isFinite(n) ? (n === 0 ? 0 : n) : null;
  }
  return null;
}

export type CacheKind = 'num' | 'str';

/**
 * Cache points (`<c:pt>` elements) already serialized for a reference, e.g. by a writer that
 * builds large caches in chunks (see `cachePointsXml`). Returning undefined builds them in place.
 */
export type CachePointsLookup = (ref: ExcelSeriesRef<unknown>, kind: CacheKind) => string | undefined;

/**
 * The `<c:pt>` elements for cache indices `from` (inclusive) to `to` (exclusive). Concatenating
 * consecutive ranges gives exactly the full cache, so a writer can build it in chunks.
 */
export function cachePointsXml(ref: ExcelSeriesRef<unknown>, kind: CacheKind, from: number, to: number): string {
  const cache = ref.cache;
  const pts: string[] = [];
  const end = Math.min(to, cache.length);
  for (let i = from; i < end; i++) {
    // Holes of sparse arrays read as undefined and are skipped below, as forEach skipped them.
    const v = cache[i];
    if (kind === 'num') {
      const n = cacheNumber(v);
      if (n !== null) pts.push(`<c:pt idx="${i}"><c:v>${formatNumber(n)}</c:v></c:pt>`);
      continue;
    }
    if (v === null || v === undefined) continue;
    if (typeof v === 'number') {
      if (Number.isFinite(v)) pts.push(`<c:pt idx="${i}"><c:v>${formatNumber(v)}</c:v></c:pt>`);
      continue;
    }
    pts.push(`<c:pt idx="${i}"><c:v>${escapeXstring(String(v))}</c:v></c:pt>`);
  }
  return pts.join('');
}

/**
 * Every cache reference `buildChartXml` writes points for, in document order: categories / X
 * values, values, bubble sizes. Tolerates malformed specs (they are rejected by `buildChartXml`).
 */
export function chartCacheRefs(spec: ExcelChartSpec): Array<{ ref: ExcelSeriesRef<unknown>; kind: CacheKind }> {
  const out: Array<{ ref: ExcelSeriesRef<unknown>; kind: CacheKind }> = [];
  const add = (ref: ExcelSeriesRef<unknown> | null | undefined, kind: CacheKind): void => {
    if (ref && typeof ref === 'object' && Array.isArray(ref.cache)) out.push({ ref, kind });
  };
  for (const g of Array.isArray(spec?.plotGroups) ? spec.plotGroups : []) {
    for (const s of Array.isArray(g?.series) ? g.series : []) {
      if (!s || typeof s !== 'object') continue;
      if (s.errorBars && typeof s.errorBars === 'object') {
        add(s.errorBars.plus, 'num');
        add(s.errorBars.minus, 'num');
      }
      add(s.categories, s.categories?.kind === 'num' ? 'num' : 'str');
      add(s.values, 'num');
      if (g.kind === 'bubble') add(s.bubbleSizes, 'num');
    }
  }
  return out;
}

export function numRefXml(ref: ExcelSeriesRef<unknown>, what: string, points?: string): string {
  checkPointCount(ref.cache.length, what);
  const pts = points ?? cachePointsXml(ref, 'num', 0, ref.cache.length);
  const code =
    ref.formatCode && ref.formatCode !== '' ? assertExcelFormatCode(ref.formatCode, `chart ${what} cache`) : 'General';
  return (
    '<c:numRef>' +
    formulaXml(ref.formula, what) +
    `<c:numCache><c:formatCode>${escapeXml(code)}</c:formatCode><c:ptCount val="${ref.cache.length}"/>${pts}</c:numCache>` +
    '</c:numRef>'
  );
}

export function strRefXml(ref: ExcelSeriesRef<unknown>, what: string, points?: string): string {
  checkPointCount(ref.cache.length, what);
  const pts = points ?? cachePointsXml(ref, 'str', 0, ref.cache.length);
  return (
    '<c:strRef>' +
    formulaXml(ref.formula, what) +
    `<c:strCache><c:ptCount val="${ref.cache.length}"/>${pts}</c:strCache>` +
    '</c:strRef>'
  );
}

function seriesTxXml(name: ExcelSeriesSpec['name']): string {
  if (name.kind === 'ref') {
    return (
      '<c:tx><c:strRef>' +
      formulaXml(name.formula, 'series name') +
      `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXstring(name.cache ?? '')}</c:v></c:pt></c:strCache>` +
      '</c:strRef></c:tx>'
    );
  }
  return `<c:tx><c:v>${escapeXstring(name.text ?? '')}</c:v></c:tx>`;
}

function categoriesXml(
  tag: 'c:cat' | 'c:xVal',
  cats: ExcelSeriesSpec['categories'],
  lookup: CachePointsLookup | null,
): string {
  if (!cats) return '';
  const inner =
    cats.kind === 'num' ? numRefXml(cats, tag, lookup?.(cats, 'num')) : strRefXml(cats, tag, lookup?.(cats, 'str'));
  return `<${tag}>${inner}</${tag}>`;
}

// ---------------------------------------------------------------------------
// Data labels
// ---------------------------------------------------------------------------

const VALID_DLBL_POS: Record<GroupKind | 'barStacked', ReadonlySet<ExcelDataLabelPosition>> = {
  bar: new Set(['ctr', 'inEnd', 'inBase', 'outEnd']),
  barStacked: new Set(['ctr', 'inEnd', 'inBase']),
  line: new Set(['ctr', 'l', 'r', 't', 'b']),
  scatter: new Set(['ctr', 'l', 'r', 't', 'b']),
  bubble: new Set(['ctr', 'l', 'r', 't', 'b']),
  pie: new Set(['ctr', 'inEnd', 'outEnd', 'bestFit']),
  area: new Set(),
  doughnut: new Set(),
  radar: new Set(),
};

interface GroupCtx {
  kind: GroupKind;
  stacked: boolean;
  /** Line group with showMarkers=false, or a radar group whose style draws no markers. */
  hideMarkers: boolean;
  scatterStyle: string | null;
  /** Pre-built cache points, when the writer built them in chunks. */
  cachePoints: CachePointsLookup | null;
}

function dLblPosXml(pos: ExcelDataLabelPosition | null, ctx: GroupCtx): string {
  if (pos === null || pos === undefined) return '';
  const key = ctx.kind === 'bar' && ctx.stacked ? 'barStacked' : ctx.kind;
  return VALID_DLBL_POS[key].has(pos) ? valEl('c:dLblPos', pos) : '';
}

/** The shared numFmt/spPr/txPr/dLblPos/show* sequence of CT_DLbls and CT_DLbl. */
function dLblBodyXml(spec: ExcelDataLabelsSpec, ctx: GroupCtx): string {
  return (
    (spec.numberFormat
      ? `<c:numFmt formatCode="${escapeAttr(assertExcelFormatCode(spec.numberFormat, 'data labels'))}" sourceLinked="0"/>`
      : '') +
    spPrXml(spec.fill, spec.line) +
    txPrXml(spec.font) +
    dLblPosXml(spec.position, ctx) +
    valEl('c:showLegendKey', 0) +
    valEl('c:showVal', spec.showValue ? 1 : 0) +
    valEl('c:showCatName', spec.showCategoryName ? 1 : 0) +
    valEl('c:showSerName', spec.showSeriesName ? 1 : 0) +
    valEl('c:showPercent', spec.showPercent ? 1 : 0) +
    valEl('c:showBubbleSize', 0)
  );
}

function anyShown(spec: ExcelDataLabelsSpec): boolean {
  return spec.showValue || spec.showCategoryName || spec.showSeriesName || spec.showPercent;
}

const NO_LABELS: ExcelDataLabelsSpec = {
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

function dLblsXml(
  spec: ExcelDataLabelsSpec | null,
  perPoint: Array<{ idx: number; labels: ExcelDataLabelsSpec }>,
  ctx: GroupCtx,
): string {
  if (!spec && perPoint.length === 0) return '';
  const dLbl = perPoint
    .sort((a, b) => a.idx - b.idx)
    .map(({ idx, labels }) =>
      anyShown(labels)
        ? `<c:dLbl>${valEl('c:idx', idx)}${dLblBodyXml(labels, ctx)}</c:dLbl>`
        : `<c:dLbl>${valEl('c:idx', idx)}${valEl('c:delete', 1)}</c:dLbl>`,
    )
    .join('');
  return `<c:dLbls>${dLbl}${dLblBodyXml(spec ?? NO_LABELS, ctx)}</c:dLbls>`;
}

// ---------------------------------------------------------------------------
// Markers & data points
// ---------------------------------------------------------------------------

function markerXml(marker: ExcelMarkerSpec | null, ctx: GroupCtx): string {
  if (!marker) {
    // Honour group-level intent where Excel would otherwise draw automatic markers.
    const noMarkers =
      ((ctx.kind === 'line' || ctx.kind === 'radar') && ctx.hideMarkers) ||
      (ctx.kind === 'scatter' &&
        (ctx.scatterStyle === 'line' || ctx.scatterStyle === 'smooth' || ctx.scatterStyle === 'none'));
    return noMarkers ? '<c:marker><c:symbol val="none"/></c:marker>' : '';
  }
  if (marker.symbol === 'none') return '<c:marker><c:symbol val="none"/></c:marker>';
  return (
    '<c:marker>' +
    valEl('c:symbol', marker.symbol) +
    valEl('c:size', clampInt(marker.size, 2, 72, 5)) +
    spPrXml(marker.fill, marker.line) +
    '</c:marker>'
  );
}

function dPtXml(dp: ExcelSeriesSpec['dataPoints'][number], ctx: GroupCtx): string {
  const hasMarker = ctx.kind === 'line' || ctx.kind === 'scatter' || ctx.kind === 'radar';
  const marker = hasMarker && dp.marker ? markerXml(dp.marker, ctx) : '';
  const explosion =
    (ctx.kind === 'pie' || ctx.kind === 'doughnut') && dp.explosion !== null && Number.isFinite(dp.explosion)
      ? valEl('c:explosion', clampInt(dp.explosion, 0, 400, 0))
      : '';
  const spPr = dp.shape ? spPrXml(dp.shape.fill, dp.shape.line) : '';
  if (marker === '' && explosion === '' && spPr === '') return '';
  return `<c:dPt>${valEl('c:idx', dp.idx)}${marker}${explosion}${spPr}</c:dPt>`;
}

function checkIdx(n: number, what: string): number {
  if (!Number.isInteger(n) || n < 0 || n > MAX_UNSIGNED_INT) {
    throw new Error(`Chart ${what} must be an integer in 0..${MAX_UNSIGNED_INT}, got ${String(n)}`);
  }
  return n;
}

/**
 * Series-level checks Excel enforces when loading: series idx/order unique across the chart, and
 * dPt / dLbl idx unique per series and below the point count.
 */
function validateSeries(spec: ExcelChartSpec): void {
  const total = spec.plotGroups.reduce((n, g) => n + g.series.length, 0);
  if (total > EXCEL_MAX_SERIES_PER_CHART) {
    throw new Error(`Chart has ${total} series; Excel allows at most ${EXCEL_MAX_SERIES_PER_CHART} per chart`);
  }
  const idxs = new Set<number>();
  const orders = new Set<number>();
  for (const g of spec.plotGroups) {
    for (const s of g.series) {
      checkIdx(s.idx, 'series idx');
      checkIdx(s.order, 'series order');
      if (idxs.has(s.idx)) throw new Error(`Chart has duplicate series idx ${s.idx}`);
      if (orders.has(s.order)) throw new Error(`Chart has duplicate series order ${s.order}`);
      idxs.add(s.idx);
      orders.add(s.order);
      const ptCount = s.values.cache.length;
      const seen = new Set<number>();
      for (const dp of s.dataPoints) {
        checkIdx(dp.idx, `series ${s.idx} data point idx`);
        if (seen.has(dp.idx)) throw new Error(`Chart series ${s.idx}: duplicate data point idx ${dp.idx}`);
        seen.add(dp.idx);
        if (dp.idx >= ptCount) {
          throw new Error(
            `Chart series ${s.idx}: data point idx ${dp.idx} (dPt/dLbl) is outside the series' ${ptCount} points`,
          );
        }
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Series
// ---------------------------------------------------------------------------

function seriesShapeXml(s: ExcelSeriesSpec, ctx: GroupCtx): string {
  let line = s.shape.line;
  // Excel ignores scatterStyle "marker"; the connecting line must be removed explicitly.
  if (ctx.kind === 'scatter' && !line && (ctx.scatterStyle === 'marker' || ctx.scatterStyle === 'none')) {
    line = { widthPx: 0, hex: null, alpha: 1, dash: 'solid', noFill: true };
  }
  return spPrXml(s.shape.fill, line);
}

function bubbleSizeXml(s: ExcelSeriesSpec, lookup: CachePointsLookup | null): string {
  if (s.bubbleSizes)
    return `<c:bubbleSize>${numRefXml(s.bubbleSizes, 'bubbleSize', lookup?.(s.bubbleSizes, 'num'))}</c:bubbleSize>`;
  // No sizes supplied: equal-size bubbles via a literal, so the series still renders.
  const n = s.values.cache.length;
  const pts = Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>1</c:v></c:pt>`).join('');
  return `<c:bubbleSize><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${pts}</c:numLit></c:bubbleSize>`;
}

/**
 * CT_ErrBars (errDir?, errBarType, errValType, noEndCap?, plus?, minus?, val?, spPr?): custom
 * values in both directions. `errDir` is written for scatter/bubble only (Y bars); on bar, line and
 * area series the direction is implied by the value axis.
 */
function errBarsXml(s: ExcelSeriesSpec, ctx: GroupCtx): string {
  const eb = s.errorBars;
  if (!eb) return '';
  const supported =
    ctx.kind === 'bar' || ctx.kind === 'line' || ctx.kind === 'area' || ctx.kind === 'scatter' || ctx.kind === 'bubble';
  if (!supported) throw new Error(`Chart series ${s.idx}: error bars are not available on ${ctx.kind} charts`);
  const lookup = ctx.cachePoints;
  return (
    '<c:errBars>' +
    (ctx.kind === 'scatter' || ctx.kind === 'bubble' ? valEl('c:errDir', 'y') : '') +
    valEl('c:errBarType', 'both') +
    valEl('c:errValType', 'cust') +
    valEl('c:noEndCap', 0) +
    `<c:plus>${numRefXml(eb.plus, 'errBars plus', lookup?.(eb.plus, 'num'))}</c:plus>` +
    `<c:minus>${numRefXml(eb.minus, 'errBars minus', lookup?.(eb.minus, 'num'))}</c:minus>` +
    spPrXml(null, eb.line) +
    '</c:errBars>'
  );
}

function smoothXml(s: ExcelSeriesSpec, ctx: GroupCtx): string {
  if (s.smooth !== null && s.smooth !== undefined) return valEl('c:smooth', s.smooth ? 1 : 0);
  if (ctx.kind === 'scatter')
    return valEl('c:smooth', ctx.scatterStyle === 'smooth' || ctx.scatterStyle === 'smoothMarker' ? 1 : 0);
  return valEl('c:smooth', 0);
}

export function seriesXml(s: ExcelSeriesSpec, ctx: GroupCtx): string {
  const idx = checkIdx(s.idx, 'series idx');
  const order = checkIdx(s.order, 'series order');
  const head = valEl('c:idx', idx) + valEl('c:order', order) + seriesTxXml(s.name) + seriesShapeXml(s, ctx);
  const dPts = [...s.dataPoints]
    .sort((a, b) => a.idx - b.idx)
    .map((dp) => {
      checkIdx(dp.idx, 'data point idx');
      return dPtXml(dp, ctx);
    })
    .join('');
  const perPoint = s.dataPoints
    .filter((dp) => dp.dataLabels !== null && dp.dataLabels !== undefined)
    .map((dp) => ({ idx: dp.idx, labels: dp.dataLabels! }));
  const dLbls = dLblsXml(s.dataLabels, perPoint, ctx);
  // Schema order: … dLbls, trendline*, errBars, cat/xVal …
  const errBars = errBarsXml(s, ctx);
  const invert = valEl('c:invertIfNegative', s.invertIfNegative ? 1 : 0);
  const lookup = ctx.cachePoints;
  const valuePoints = lookup?.(s.values, 'num');
  const val = `<c:val>${numRefXml(s.values, 'values', valuePoints)}</c:val>`;

  let body: string;
  switch (ctx.kind) {
    case 'bar':
      body = head + invert + dPts + dLbls + errBars + categoriesXml('c:cat', s.categories, lookup) + val;
      break;
    case 'line':
      body =
        head +
        markerXml(s.marker, ctx) +
        dPts +
        dLbls +
        errBars +
        categoriesXml('c:cat', s.categories, lookup) +
        val +
        smoothXml(s, ctx);
      break;
    case 'area':
      body = head + dPts + dLbls + errBars + categoriesXml('c:cat', s.categories, lookup) + val;
      break;
    case 'radar':
      // CT_RadarSer: idx, order, tx, spPr, marker, dPt*, dLbls, cat, val (no smoothing).
      body = head + markerXml(s.marker, ctx) + dPts + dLbls + categoriesXml('c:cat', s.categories, lookup) + val;
      break;
    case 'pie':
    case 'doughnut':
      body = head + dPts + dLbls + categoriesXml('c:cat', s.categories, lookup) + val;
      break;
    case 'scatter':
      body =
        head +
        markerXml(s.marker, ctx) +
        dPts +
        dLbls +
        errBars +
        categoriesXml('c:xVal', s.categories, lookup) +
        `<c:yVal>${numRefXml(s.values, 'yVal', valuePoints)}</c:yVal>` +
        smoothXml(s, ctx);
      break;
    case 'bubble':
      body =
        head +
        invert +
        dPts +
        dLbls +
        errBars +
        categoriesXml('c:xVal', s.categories, lookup) +
        `<c:yVal>${numRefXml(s.values, 'yVal', valuePoints)}</c:yVal>` +
        bubbleSizeXml(s, lookup) +
        valEl('c:bubble3D', 0);
      break;
    default: {
      const never: never = ctx.kind;
      throw new Error(`Unknown plot group kind ${String(never)}`);
    }
  }
  return `<c:ser>${body}</c:ser>`;
}

// ---------------------------------------------------------------------------
// Plot groups
// ---------------------------------------------------------------------------

function axIdsXml(ids: [number, number]): string {
  return valEl('c:axId', checkIdx(ids[0], 'axis id')) + valEl('c:axId', checkIdx(ids[1], 'axis id'));
}

export function plotGroupXml(g: PlotGroupSpec, cachePoints: CachePointsLookup | null = null): string {
  const ctx: GroupCtx = {
    kind: g.kind,
    stacked:
      (g.kind === 'bar' || g.kind === 'line' || g.kind === 'area') &&
      g.grouping !== 'clustered' &&
      g.grouping !== 'standard',
    hideMarkers: (g.kind === 'line' && !g.showMarkers) || (g.kind === 'radar' && g.radarStyle !== 'marker'),
    scatterStyle: g.kind === 'scatter' ? g.scatterStyle : null,
    cachePoints,
  };
  const ser = g.series.map((s) => seriesXml(s, ctx)).join('');
  const dLbls = dLblsXml(g.dataLabels, [], ctx);
  const vary = valEl('c:varyColors', g.varyColors ? 1 : 0);

  switch (g.kind) {
    case 'bar': {
      const overlap = ctx.stacked ? 100 : g.overlap === null ? null : clampInt(g.overlap, -100, 100, 0);
      return (
        '<c:barChart>' +
        valEl('c:barDir', g.barDir === 'bar' ? 'bar' : 'col') +
        valEl('c:grouping', g.grouping) +
        vary +
        ser +
        dLbls +
        valEl('c:gapWidth', clampInt(g.gapWidth, 0, 500, 150)) +
        (overlap === null ? '' : valEl('c:overlap', overlap)) +
        axIdsXml(g.axisIds) +
        '</c:barChart>'
      );
    }
    case 'line':
      return (
        '<c:lineChart>' +
        valEl('c:grouping', g.grouping) +
        vary +
        ser +
        dLbls +
        valEl('c:marker', 1) +
        axIdsXml(g.axisIds) +
        '</c:lineChart>'
      );
    case 'area':
      return `<c:areaChart>${valEl('c:grouping', g.grouping)}${vary}${ser}${dLbls}${axIdsXml(g.axisIds)}</c:areaChart>`;
    case 'scatter':
      return (
        '<c:scatterChart>' +
        valEl('c:scatterStyle', g.scatterStyle) +
        vary +
        ser +
        dLbls +
        axIdsXml(g.axisIds) +
        '</c:scatterChart>'
      );
    case 'bubble':
      return (
        '<c:bubbleChart>' +
        vary +
        ser +
        dLbls +
        valEl('c:bubble3D', 0) +
        valEl('c:bubbleScale', clampInt(g.bubbleScale, 0, 300, 100)) +
        valEl('c:showNegBubbles', 0) +
        axIdsXml(g.axisIds) +
        '</c:bubbleChart>'
      );
    case 'radar':
      // CT_RadarChart: radarStyle, varyColors, ser*, dLbls, axId, axId.
      return (
        '<c:radarChart>' +
        valEl('c:radarStyle', g.radarStyle === 'filled' || g.radarStyle === 'marker' ? g.radarStyle : 'standard') +
        vary +
        ser +
        dLbls +
        axIdsXml(g.axisIds) +
        '</c:radarChart>'
      );
    case 'pie':
      return (
        '<c:pieChart>' +
        vary +
        ser +
        dLbls +
        valEl('c:firstSliceAng', clampInt(g.firstSliceAngle, 0, 360, 0)) +
        '</c:pieChart>'
      );
    case 'doughnut':
      return (
        '<c:doughnutChart>' +
        vary +
        ser +
        dLbls +
        valEl('c:firstSliceAng', clampInt(g.firstSliceAngle, 0, 360, 0)) +
        valEl('c:holeSize', clampInt(g.holeSize, 10, 90, 50)) +
        '</c:doughnutChart>'
      );
    default: {
      const never: never = g;
      throw new Error(`Unknown plot group ${JSON.stringify(never)}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Axes
// ---------------------------------------------------------------------------

function scalingXml(ax: ExcelAxisSpec): string {
  const s = ax.scaling;
  const log = s.logBase !== null && Number.isFinite(s.logBase) ? clampNum(s.logBase, 2, 1000, 10) : null;
  let min = s.min !== null && Number.isFinite(s.min) ? finite(s.min, 0) : null;
  let max = s.max !== null && Number.isFinite(s.max) ? finite(s.max, 0) : null;
  if (log !== null) {
    if (min !== null && min <= 0) min = null;
    if (max !== null && max <= 0) max = null;
  }
  if (min !== null && max !== null && min >= max) {
    min = null;
    max = null;
  }
  return (
    '<c:scaling>' +
    (log !== null ? valEl('c:logBase', log) : '') +
    valEl('c:orientation', s.orientation === 'maxMin' ? 'maxMin' : 'minMax') +
    (max !== null ? valEl('c:max', max) : '') +
    (min !== null ? valEl('c:min', min) : '') +
    '</c:scaling>'
  );
}

function gridlinesXml(tag: 'c:majorGridlines' | 'c:minorGridlines', line: ExcelAxisSpec['majorGridlines']): string {
  if (!line) return '';
  const ln = lineXml(line);
  return ln === '' ? `<${tag}/>` : `<${tag}><c:spPr>${ln}</c:spPr></${tag}>`;
}

function crossesXml(c: ExcelAxisSpec['crosses']): string {
  if (typeof c === 'object' && c !== null) return valEl('c:crossesAt', finite(c.at, 0));
  return valEl('c:crosses', c === 'min' || c === 'max' ? c : 'autoZero');
}

function positiveOrNull(n: number | null): number | null {
  return n !== null && Number.isFinite(n) && n > 0 ? n : null;
}

/** Shared axis head: axId … crossAx, crosses. */
function axisCommonXml(ax: ExcelAxisSpec): string {
  const labelPos = ax.labels.position;
  return (
    valEl('c:axId', checkIdx(ax.id, 'axis id')) +
    scalingXml(ax) +
    valEl('c:delete', ax.deleted ? 1 : 0) +
    valEl('c:axPos', ax.position) +
    gridlinesXml('c:majorGridlines', ax.majorGridlines) +
    gridlinesXml('c:minorGridlines', ax.minorGridlines) +
    titleXml(ax.title) +
    (ax.numberFormat
      ? `<c:numFmt formatCode="${escapeAttr(ax.numberFormat.code ? assertExcelFormatCode(ax.numberFormat.code, `axis ${ax.id}`) : 'General')}" sourceLinked="${ax.numberFormat.sourceLinked ? 1 : 0}"/>`
      : '') +
    valEl('c:majorTickMark', ax.majorTickMark) +
    valEl('c:minorTickMark', 'none') +
    valEl('c:tickLblPos', labelPos) +
    (ax.axisLine ? `<c:spPr>${fillXml(null)}${lineXml(ax.axisLine)}</c:spPr>` : '') +
    txPrXml(ax.labels.font, ax.labels.rotation) +
    valEl('c:crossAx', checkIdx(ax.crossAxisId, 'cross axis id')) +
    crossesXml(ax.crosses)
  );
}

export function axisXml(ax: ExcelAxisSpec, crossBetween: 'between' | 'midCat'): string {
  const common = axisCommonXml(ax);
  const major = positiveOrNull(ax.majorUnit);
  const minor = positiveOrNull(ax.minorUnit);
  switch (ax.kind) {
    case 'cat': {
      // On a category axis the "unit" is a label/tick skip count.
      const skip = major !== null ? clampInt(major, 1, MAX_SKIP, 1) : null;
      return (
        '<c:catAx>' +
        common +
        valEl('c:auto', 1) +
        valEl('c:lblAlgn', 'ctr') +
        valEl('c:lblOffset', 100) +
        (skip !== null && skip > 1 ? valEl('c:tickLblSkip', skip) + valEl('c:tickMarkSkip', skip) : '') +
        valEl('c:noMultiLvlLbl', 0) +
        '</c:catAx>'
      );
    }
    case 'date': {
      const base = ax.dateAxis?.baseTimeUnit ?? null;
      const unit = base ?? 'days';
      // Date axis units count whole base time units: anything else is omitted (Excel's automatic unit).
      const dateMajor = major !== null && Number.isInteger(major) && major >= 1 ? major : null;
      const dateMinor = minor !== null && Number.isInteger(minor) && minor >= 1 ? minor : null;
      return (
        '<c:dateAx>' +
        common +
        valEl('c:auto', 1) +
        valEl('c:lblOffset', 100) +
        (base !== null ? valEl('c:baseTimeUnit', base) : '') +
        (dateMajor !== null ? valEl('c:majorUnit', dateMajor) + valEl('c:majorTimeUnit', unit) : '') +
        (dateMinor !== null ? valEl('c:minorUnit', dateMinor) + valEl('c:minorTimeUnit', unit) : '') +
        '</c:dateAx>'
      );
    }
    case 'val':
      return (
        '<c:valAx>' +
        common +
        valEl('c:crossBetween', crossBetween) +
        (major !== null ? valEl('c:majorUnit', major) : '') +
        (minor !== null ? valEl('c:minorUnit', minor) : '') +
        '</c:valAx>'
      );
    default: {
      const never: never = ax.kind;
      throw new Error(`Unknown axis kind ${String(never)}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Chart space
// ---------------------------------------------------------------------------

function validateAxes(spec: ExcelChartSpec): Map<number, ExcelAxisSpec> {
  const byId = new Map<number, ExcelAxisSpec>();
  for (const ax of spec.axes) {
    checkIdx(ax.id, 'axis id');
    if (byId.has(ax.id)) throw new Error(`Chart has duplicate axis id ${ax.id}`);
    byId.set(ax.id, ax);
  }
  for (const ax of spec.axes) {
    const partner = byId.get(ax.crossAxisId);
    if (!partner) throw new Error(`Axis ${ax.id} crosses unknown axis ${ax.crossAxisId}`);
    if (partner.crossAxisId !== ax.id) {
      throw new Error(
        `Axis ${ax.id} crosses axis ${ax.crossAxisId}, which crosses ${partner.crossAxisId}: crossAx must be reciprocal`,
      );
    }
  }
  const pairs = new Set<string>();
  const used = new Set<number>();
  for (const g of spec.plotGroups) {
    if (g.kind === 'pie' || g.kind === 'doughnut') continue;
    for (const id of g.axisIds) {
      if (!byId.has(id)) throw new Error(`Plot group "${g.kind}" references unknown axis id ${id}`);
      used.add(id);
    }
    pairs.add([...g.axisIds].sort((a, b) => a - b).join('/'));
  }
  if (pairs.size > 2) {
    throw new Error(
      `Chart uses ${pairs.size} axis pairs (${[...pairs].join(', ')}); Excel has only two axis groups (primary and secondary)`,
    );
  }
  for (const ax of spec.axes) {
    if (!used.has(ax.id)) throw new Error(`Axis ${ax.id} is not used by any plot group`);
  }
  return byId;
}

function manualLayoutXml(ml: { x: number; y: number; w: number; h: number } | null): string {
  if (!ml) return '<c:layout/>';
  return (
    '<c:layout><c:manualLayout>' +
    valEl('c:layoutTarget', 'inner') +
    valEl('c:xMode', 'edge') +
    valEl('c:yMode', 'edge') +
    valEl('c:x', clampNum(ml.x, 0, 1, 0)) +
    valEl('c:y', clampNum(ml.y, 0, 1, 0)) +
    valEl('c:w', clampNum(ml.w, 0, 1, 1)) +
    valEl('c:h', clampNum(ml.h, 0, 1, 1)) +
    '</c:manualLayout></c:layout>'
  );
}

function legendXml(legend: ExcelChartSpec['legend']): string {
  if (!legend) return '';
  return (
    '<c:legend>' +
    valEl('c:legendPos', legend.position) +
    // Schema order: legendPos, legendEntry*, layout, overlay, spPr, txPr.
    [...new Set(legend.deletedEntries ?? [])]
      .filter((i) => Number.isInteger(i) && i >= 0)
      .sort((a, b) => a - b)
      .map((i) => `<c:legendEntry>${valEl('c:idx', i)}${valEl('c:delete', 1)}</c:legendEntry>`)
      .join('') +
    '<c:layout/>' +
    valEl('c:overlay', legend.overlay ? 1 : 0) +
    spPrXml(legend.fill, legend.line) +
    txPrXml(legend.font) +
    '</c:legend>'
  );
}

/**
 * @param options.cachePoints Pre-built `<c:pt>` runs per cache reference (see `cachePointsXml`);
 *   the output is byte-identical with or without them.
 */
export function buildChartXml(spec: ExcelChartSpec, options: { cachePoints?: CachePointsLookup } = {}): string {
  if (spec.plotGroups.length === 0) throw new Error('Chart has no plot groups');
  validateSeries(spec);
  const axesById = validateAxes(spec);
  const groupsByAxis = new Map<number, GroupKind[]>();
  for (const g of spec.plotGroups) {
    if (g.kind === 'pie' || g.kind === 'doughnut') continue;
    for (const id of g.axisIds) groupsByAxis.set(id, [...(groupsByAxis.get(id) ?? []), g.kind]);
  }
  const crossBetweenFor = (ax: ExcelAxisSpec): 'between' | 'midCat' => {
    const partner = axesById.get(ax.crossAxisId);
    if (partner && partner.kind === 'val') return 'midCat';
    const kinds = groupsByAxis.get(ax.id) ?? [];
    if (kinds.length > 0 && kinds.every((k) => k === 'scatter' || k === 'bubble')) return 'midCat';
    return 'between';
  };

  const hasTitle = spec.title !== null && spec.title.lines.length > 0;
  const plotArea =
    '<c:plotArea>' +
    manualLayoutXml(spec.plotArea.manualLayout) +
    spec.plotGroups.map((g) => plotGroupXml(g, options.cachePoints ?? null)).join('') +
    spec.axes.map((ax) => axisXml(ax, crossBetweenFor(ax))).join('') +
    spPrXml(spec.plotArea.fill, spec.plotArea.line) +
    '</c:plotArea>';

  const chart =
    '<c:chart>' +
    titleXml(spec.title) +
    valEl('c:autoTitleDeleted', hasTitle ? 0 : 1) +
    plotArea +
    legendXml(spec.legend) +
    valEl('c:plotVisOnly', 1) +
    valEl('c:dispBlanksAs', spec.dispBlanksAs) +
    '</c:chart>';

  const style =
    spec.style !== null && Number.isFinite(spec.style) ? valEl('c:style', clampInt(spec.style, 1, 48, 2)) : '';

  return xmlDocument(
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
      valEl('c:date1904', 0) +
      valEl('c:lang', 'en-US') +
      valEl('c:roundedCorners', 0) +
      style +
      chart +
      spPrXml(spec.chartArea.fill, spec.chartArea.line) +
      txPrXml(spec.textDefaults) +
      '</c:chartSpace>',
  );
}
