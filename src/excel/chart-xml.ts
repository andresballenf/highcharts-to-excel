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
import { clampInt, clampNum, escapeAttr, escapeXml, finite, formatNumber, valEl, xmlDocument } from './xml';

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

type GroupKind = PlotGroupSpec['kind'];

// ---------------------------------------------------------------------------
// Data references
// ---------------------------------------------------------------------------

function formulaXml(formula: string, what: string): string {
  const f = String(formula ?? '').trim().replace(/^=/, '');
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

export function numRefXml(ref: ExcelSeriesRef<unknown>, what: string): string {
  checkPointCount(ref.cache.length, what);
  const pts: string[] = [];
  ref.cache.forEach((v, i) => {
    const n = cacheNumber(v);
    if (n !== null) pts.push(`<c:pt idx="${i}"><c:v>${formatNumber(n)}</c:v></c:pt>`);
  });
  const code = ref.formatCode && ref.formatCode !== '' ? ref.formatCode : 'General';
  return (
    '<c:numRef>' +
    formulaXml(ref.formula, what) +
    `<c:numCache><c:formatCode>${escapeXml(code)}</c:formatCode><c:ptCount val="${ref.cache.length}"/>${pts.join('')}</c:numCache>` +
    '</c:numRef>'
  );
}

export function strRefXml(ref: ExcelSeriesRef<unknown>, what: string): string {
  checkPointCount(ref.cache.length, what);
  const pts: string[] = [];
  ref.cache.forEach((v, i) => {
    if (v === null || v === undefined) return;
    if (typeof v === 'number') {
      if (!Number.isFinite(v)) return;
      pts.push(`<c:pt idx="${i}"><c:v>${formatNumber(v)}</c:v></c:pt>`);
      return;
    }
    pts.push(`<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`);
  });
  return (
    '<c:strRef>' +
    formulaXml(ref.formula, what) +
    `<c:strCache><c:ptCount val="${ref.cache.length}"/>${pts.join('')}</c:strCache>` +
    '</c:strRef>'
  );
}

function seriesTxXml(name: ExcelSeriesSpec['name']): string {
  if (name.kind === 'ref') {
    return (
      '<c:tx><c:strRef>' +
      formulaXml(name.formula, 'series name') +
      `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(name.cache ?? '')}</c:v></c:pt></c:strCache>` +
      '</c:strRef></c:tx>'
    );
  }
  return `<c:tx><c:v>${escapeXml(name.text ?? '')}</c:v></c:tx>`;
}

function categoriesXml(tag: 'c:cat' | 'c:xVal', cats: ExcelSeriesSpec['categories']): string {
  if (!cats) return '';
  const inner = cats.kind === 'num' ? numRefXml(cats, tag) : strRefXml(cats, tag);
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
};

interface GroupCtx {
  kind: GroupKind;
  stacked: boolean;
  /** Line group with showMarkers=false. */
  hideMarkers: boolean;
  scatterStyle: string | null;
}

function dLblPosXml(pos: ExcelDataLabelPosition | null, ctx: GroupCtx): string {
  if (pos === null || pos === undefined) return '';
  const key = ctx.kind === 'bar' && ctx.stacked ? 'barStacked' : ctx.kind;
  return VALID_DLBL_POS[key].has(pos) ? valEl('c:dLblPos', pos) : '';
}

/** The shared numFmt/spPr/txPr/dLblPos/show* sequence of CT_DLbls and CT_DLbl. */
function dLblBodyXml(spec: ExcelDataLabelsSpec, ctx: GroupCtx): string {
  return (
    (spec.numberFormat ? `<c:numFmt formatCode="${escapeAttr(spec.numberFormat)}" sourceLinked="0"/>` : '') +
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
      (ctx.kind === 'line' && ctx.hideMarkers) ||
      (ctx.kind === 'scatter' && (ctx.scatterStyle === 'line' || ctx.scatterStyle === 'smooth' || ctx.scatterStyle === 'none'));
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
  const hasMarker = ctx.kind === 'line' || ctx.kind === 'scatter';
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
  if (!Number.isInteger(n) || n < 0) throw new Error(`Chart ${what} must be a non-negative integer, got ${String(n)}`);
  return n;
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

function bubbleSizeXml(s: ExcelSeriesSpec): string {
  if (s.bubbleSizes) return `<c:bubbleSize>${numRefXml(s.bubbleSizes, 'bubbleSize')}</c:bubbleSize>`;
  // No sizes supplied: equal-size bubbles via a literal, so the series still renders.
  const n = s.values.cache.length;
  const pts = Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>1</c:v></c:pt>`).join('');
  return `<c:bubbleSize><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${pts}</c:numLit></c:bubbleSize>`;
}

function smoothXml(s: ExcelSeriesSpec, ctx: GroupCtx): string {
  if (s.smooth !== null && s.smooth !== undefined) return valEl('c:smooth', s.smooth ? 1 : 0);
  if (ctx.kind === 'scatter') return valEl('c:smooth', ctx.scatterStyle === 'smooth' || ctx.scatterStyle === 'smoothMarker' ? 1 : 0);
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
  const invert = valEl('c:invertIfNegative', s.invertIfNegative ? 1 : 0);
  const val = `<c:val>${numRefXml(s.values, 'values')}</c:val>`;

  let body: string;
  switch (ctx.kind) {
    case 'bar':
      body = head + invert + dPts + dLbls + categoriesXml('c:cat', s.categories) + val;
      break;
    case 'line':
      body = head + markerXml(s.marker, ctx) + dPts + dLbls + categoriesXml('c:cat', s.categories) + val + smoothXml(s, ctx);
      break;
    case 'area':
    case 'pie':
    case 'doughnut':
      body = head + dPts + dLbls + categoriesXml('c:cat', s.categories) + val;
      break;
    case 'scatter':
      body =
        head +
        markerXml(s.marker, ctx) +
        dPts +
        dLbls +
        categoriesXml('c:xVal', s.categories) +
        `<c:yVal>${numRefXml(s.values, 'yVal')}</c:yVal>` +
        smoothXml(s, ctx);
      break;
    case 'bubble':
      body =
        head +
        invert +
        dPts +
        dLbls +
        categoriesXml('c:xVal', s.categories) +
        `<c:yVal>${numRefXml(s.values, 'yVal')}</c:yVal>` +
        bubbleSizeXml(s) +
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

export function plotGroupXml(g: PlotGroupSpec): string {
  if (g.series.length > EXCEL_MAX_SERIES_PER_CHART) {
    throw new Error(`Chart group has ${g.series.length} series; Excel allows at most ${EXCEL_MAX_SERIES_PER_CHART}`);
  }
  const ctx: GroupCtx = {
    kind: g.kind,
    stacked: (g.kind === 'bar' || g.kind === 'line' || g.kind === 'area') && g.grouping !== 'clustered' && g.grouping !== 'standard',
    hideMarkers: g.kind === 'line' && !g.showMarkers,
    scatterStyle: g.kind === 'scatter' ? g.scatterStyle : null,
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
      return '<c:areaChart>' + valEl('c:grouping', g.grouping) + vary + ser + dLbls + axIdsXml(g.axisIds) + '</c:areaChart>';
    case 'scatter':
      return '<c:scatterChart>' + valEl('c:scatterStyle', g.scatterStyle) + vary + ser + dLbls + axIdsXml(g.axisIds) + '</c:scatterChart>';
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
    case 'pie':
      return '<c:pieChart>' + vary + ser + dLbls + valEl('c:firstSliceAng', clampInt(g.firstSliceAngle, 0, 360, 0)) + '</c:pieChart>';
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
      ? `<c:numFmt formatCode="${escapeAttr(ax.numberFormat.code || 'General')}" sourceLinked="${ax.numberFormat.sourceLinked ? 1 : 0}"/>`
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
      const skip = major !== null ? Math.max(1, Math.round(major)) : null;
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
      return (
        '<c:dateAx>' +
        common +
        valEl('c:auto', 1) +
        valEl('c:lblOffset', 100) +
        (base !== null ? valEl('c:baseTimeUnit', base) : '') +
        (major !== null ? valEl('c:majorUnit', major) + valEl('c:majorTimeUnit', unit) : '') +
        (minor !== null ? valEl('c:minorUnit', minor) + valEl('c:minorTimeUnit', unit) : '') +
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
    if (!byId.has(ax.crossAxisId)) throw new Error(`Axis ${ax.id} crosses unknown axis ${ax.crossAxisId}`);
  }
  for (const g of spec.plotGroups) {
    if (g.kind === 'pie' || g.kind === 'doughnut') continue;
    for (const id of g.axisIds) {
      if (!byId.has(id)) throw new Error(`Plot group "${g.kind}" references unknown axis id ${id}`);
    }
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
    '<c:layout/>' +
    valEl('c:overlay', legend.overlay ? 1 : 0) +
    spPrXml(legend.fill, legend.line) +
    txPrXml(legend.font) +
    '</c:legend>'
  );
}

export function buildChartXml(spec: ExcelChartSpec): string {
  if (spec.plotGroups.length === 0) throw new Error('Chart has no plot groups');
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
    spec.plotGroups.map(plotGroupXml).join('') +
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

  const style = spec.style !== null && Number.isFinite(spec.style) ? valEl('c:style', clampInt(spec.style, 1, 48, 2)) : '';

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
