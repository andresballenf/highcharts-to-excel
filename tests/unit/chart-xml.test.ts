import { describe, expect, it } from 'vitest';
import { XMLValidator } from 'fast-xml-parser';
import { buildChartXml } from '../../src/excel/chart-xml';
import { elementOrder } from '../helpers/inspect-xlsx';
import { CAT_F, NAME1_F, NAME2_F, VAL1_F, VAL2_F, axis, chart, chartSpecs, labels, series } from '../fixtures/writer-specs';

// Canonical schema sequences (ECMA-376 dml-chart.xsd). Observed children must be a subsequence.
const ORDER = {
  chartSpace: ['c:date1904', 'c:lang', 'c:roundedCorners', 'c:style', 'c:clrMapOvr', 'c:pivotSource', 'c:protection', 'c:chart', 'c:spPr', 'c:txPr', 'c:externalData', 'c:printSettings', 'c:userShapes'],
  chart: ['c:title', 'c:autoTitleDeleted', 'c:pivotFmts', 'c:view3D', 'c:floor', 'c:sideWall', 'c:backWall', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs'],
  barChart: ['c:barDir', 'c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:gapWidth', 'c:overlap', 'c:serLines', 'c:axId'],
  lineChart: ['c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:hiLowLines', 'c:upDownBars', 'c:marker', 'c:smooth', 'c:axId'],
  areaChart: ['c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:axId'],
  scatterChart: ['c:scatterStyle', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId'],
  bubbleChart: ['c:varyColors', 'c:ser', 'c:dLbls', 'c:bubble3D', 'c:bubbleScale', 'c:showNegBubbles', 'c:sizeRepresents', 'c:axId'],
  pieChart: ['c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng'],
  doughnutChart: ['c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng', 'c:holeSize'],
  barSer: ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:pictureOptions', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:shape'],
  lineSer: ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:smooth'],
  areaSer: ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:pictureOptions', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val'],
  scatterSer: ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:xVal', 'c:yVal', 'c:smooth'],
  bubbleSer: ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:xVal', 'c:yVal', 'c:bubbleSize', 'c:bubble3D'],
  pieSer: ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:explosion', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val'],
  dPt: ['c:idx', 'c:invertIfNegative', 'c:marker', 'c:bubble3D', 'c:explosion', 'c:spPr', 'c:pictureOptions'],
  dLbls: ['c:dLbl', 'c:delete', 'c:numFmt', 'c:spPr', 'c:txPr', 'c:dLblPos', 'c:showLegendKey', 'c:showVal', 'c:showCatName', 'c:showSerName', 'c:showPercent', 'c:showBubbleSize', 'c:separator', 'c:showLeaderLines'],
  catAx: ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:auto', 'c:lblAlgn', 'c:lblOffset', 'c:tickLblSkip', 'c:tickMarkSkip', 'c:noMultiLvlLbl'],
  valAx: ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:crossBetween', 'c:majorUnit', 'c:minorUnit', 'c:dispUnits'],
  dateAx: ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:auto', 'c:lblOffset', 'c:baseTimeUnit', 'c:majorUnit', 'c:majorTimeUnit', 'c:minorUnit', 'c:minorTimeUnit'],
  scaling: ['c:logBase', 'c:orientation', 'c:max', 'c:min'],
  marker: ['c:symbol', 'c:size', 'c:spPr'],
  legend: ['c:legendPos', 'c:legendEntry', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr'],
  title: ['c:tx', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr'],
  spPr: ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:ln'],
  ln: ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill', 'a:prstDash', 'a:custDash', 'a:round', 'a:bevel', 'a:miter', 'a:headEnd', 'a:tailEnd'],
} as const;

const REPEATABLE = new Set(['c:ser', 'c:dPt', 'c:dLbl', 'c:axId', 'c:legendEntry', 'c:trendline', 'c:errBars']);

function expectSchemaOrder(observed: string[], canonical: readonly string[]): void {
  let last = -1;
  let prev = '';
  for (const name of observed) {
    const i = canonical.indexOf(name);
    expect(i, `<${name}> is not allowed here (observed ${observed.join(', ')})`).toBeGreaterThanOrEqual(0);
    expect(i >= last, `<${name}> appears after <${prev}> (observed ${observed.join(', ')})`).toBe(true);
    if (i === last) expect(REPEATABLE.has(name), `<${name}> repeated`).toBe(true);
    last = i;
    prev = name;
  }
}

/** First `<tag …>…</tag>` substring (non-nested tags only). */
function section(xml: string, tag: string, nth = 0): string {
  const re = new RegExp(`<${tag}(?=[\\s>/])[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g');
  const all = xml.match(re) ?? [];
  const s = all[nth];
  if (!s) throw new Error(`no <${tag}> #${nth}`);
  return s;
}

function attr(xml: string, tag: string, name = 'val'): string | undefined {
  return new RegExp(`<${tag}\\s[^>]*?\\b${name}="([^"]*)"`).exec(xml)?.[1];
}

function formulas(xml: string): string[] {
  return [...xml.matchAll(/<c:f>([\s\S]*?)<\/c:f>/g)].map((m) => m[1]!.replace(/&apos;/g, "'"));
}

function checkCommon(xml: string): void {
  expect(XMLValidator.validate(xml)).toBe(true);
  expect(xml).not.toMatch(/>null</);
  expect(xml).not.toMatch(/NaN|Infinity|undefined/);
  expectSchemaOrder(elementOrder(xml, 'c:chartSpace'), ORDER.chartSpace);
  expectSchemaOrder(elementOrder(xml, 'c:chart'), ORDER.chart);
  for (const tag of ['catAx', 'valAx', 'dateAx'] as const) {
    let i = 0;
    while (xml.includes(`<c:${tag}>`) && i < 10) {
      let s: string;
      try {
        s = section(xml, `c:${tag}`, i);
      } catch {
        break;
      }
      expectSchemaOrder(elementOrder(s, `c:${tag}`), ORDER[tag]);
      expectSchemaOrder(elementOrder(s, 'c:scaling'), ORDER.scaling);
      i++;
    }
  }
  if (xml.includes('<c:legend>')) expectSchemaOrder(elementOrder(xml, 'c:legend'), ORDER.legend);
  if (xml.includes('<c:title>')) expectSchemaOrder(elementOrder(xml, 'c:title'), ORDER.title);
  for (const m of xml.matchAll(/<c:spPr>[\s\S]*?<\/c:spPr>/g)) expectSchemaOrder(elementOrder(m[0], 'c:spPr'), ORDER.spPr);
  for (const m of xml.matchAll(/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/g)) expectSchemaOrder(elementOrder(m[0], 'a:ln'), ORDER.ln);
  for (const m of xml.matchAll(/<c:marker>[\s\S]*?<\/c:marker>/g)) expectSchemaOrder(elementOrder(m[0], 'c:marker'), ORDER.marker);
  for (const m of xml.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)) expectSchemaOrder(elementOrder(m[0], 'c:dPt'), ORDER.dPt);
  for (const m of xml.matchAll(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g)) {
    const kids = elementOrder(m[0], 'c:dLbls');
    expectSchemaOrder(kids, ORDER.dLbls);
    for (const flag of ['c:showLegendKey', 'c:showVal', 'c:showCatName', 'c:showSerName', 'c:showPercent', 'c:showBubbleSize']) {
      expect(kids).toContain(flag);
    }
  }
}

describe('buildChartXml – structure per chart type', () => {
  it('bar clustered (col) with per-point dPt and outEnd labels', () => {
    const xml = buildChartXml(chartSpecs.barClustered());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:chartSpace')).toEqual(['c:date1904', 'c:lang', 'c:roundedCorners', 'c:chart', 'c:spPr', 'c:txPr']);
    expect(elementOrder(xml, 'c:chart')).toEqual(['c:title', 'c:autoTitleDeleted', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs']);
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:barChart', 'c:catAx', 'c:valAx']);
    expect(elementOrder(xml, 'c:barChart')).toEqual(['c:barDir', 'c:grouping', 'c:varyColors', 'c:ser', 'c:ser', 'c:gapWidth', 'c:overlap', 'c:axId', 'c:axId']);
    expect(elementOrder(xml, 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val']);
    expect(elementOrder(xml, 'c:dLbls')).toEqual([
      'c:dLbl', 'c:numFmt', 'c:txPr', 'c:dLblPos', 'c:showLegendKey', 'c:showVal', 'c:showCatName', 'c:showSerName', 'c:showPercent', 'c:showBubbleSize',
    ]);
    expect(elementOrder(xml, 'c:dLbl')).toEqual(['c:idx', 'c:delete']);
    expect(elementOrder(xml, 'c:dPt')).toEqual(['c:idx', 'c:spPr']);
    expect(attr(section(xml, 'c:dPt'), 'c:idx')).toBe('2');
    expect(attr(xml, 'c:dLblPos')).toBe('outEnd');
    expect(attr(xml, 'c:overlap')).toBe('-10');
    expect(attr(xml, 'c:numFmt', 'formatCode')).toBe('0.0');
    expect(elementOrder(xml, 'c:catAx')).toEqual([
      'c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:crossAx', 'c:crosses', 'c:auto', 'c:lblAlgn', 'c:lblOffset', 'c:noMultiLvlLbl',
    ]);
    expect(elementOrder(xml, 'c:valAx')).toEqual([
      'c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:crossAx', 'c:crosses', 'c:crossBetween',
    ]);
    expect(attr(xml, 'c:crossBetween')).toBe('between');
    expect(formulas(xml)).toEqual([NAME1_F, CAT_F, VAL1_F, NAME2_F, CAT_F, VAL2_F]);
    // nulls/NaN skipped from the cache; ptCount keeps the full length
    const val = section(xml, 'c:val');
    expect(attr(val, 'c:ptCount')).toBe('4');
    expect([...val.matchAll(/<c:pt idx="(\d+)">/g)].map((m) => m[1])).toEqual(['0', '2']);
    const cat = section(xml, 'c:cat');
    expect(attr(cat, 'c:ptCount')).toBe('4');
    expect([...cat.matchAll(/<c:pt idx="(\d+)">/g)].map((m) => m[1])).toEqual(['0', '1', '3']);
    // alpha 0.8 → 80000
    expect(xml).toContain('<a:srgbClr val="4472C4"><a:alpha val="80000"/></a:srgbClr>');
    expect(xml).toContain('<c:autoTitleDeleted val="0"/>');
  });

  it('bar stacked forces overlap 100 and drops dLblPos outEnd', () => {
    const xml = buildChartXml(chartSpecs.barStacked());
    checkCommon(xml);
    expect(attr(xml, 'c:grouping')).toBe('stacked');
    expect(attr(xml, 'c:overlap')).toBe('100');
    expect(attr(xml, 'c:gapWidth')).toBe('80');
    expect(xml).not.toContain('<c:dLblPos');
    // group-level dLbls sits between the last ser and gapWidth
    expect(elementOrder(xml, 'c:barChart')).toEqual(['c:barDir', 'c:grouping', 'c:varyColors', 'c:ser', 'c:ser', 'c:dLbls', 'c:gapWidth', 'c:overlap', 'c:axId', 'c:axId']);
  });

  it('bar with barDir "bar" clamps gapWidth', () => {
    const xml = buildChartXml(chartSpecs.barHorizontal());
    checkCommon(xml);
    expect(attr(xml, 'c:barDir')).toBe('bar');
    expect(attr(xml, 'c:gapWidth')).toBe('500');
    expect(attr(xml, 'c:overlap')).toBe('0');
    expect(attr(section(xml, 'c:catAx'), 'c:axPos')).toBe('l');
  });

  it('line with markers, smooth and dash', () => {
    const xml = buildChartXml(chartSpecs.line());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:lineChart')).toEqual(['c:grouping', 'c:varyColors', 'c:ser', 'c:ser', 'c:marker', 'c:axId', 'c:axId']);
    expect(elementOrder(xml, 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val', 'c:smooth']);
    expect(elementOrder(xml, 'c:marker')).toEqual(['c:symbol', 'c:size', 'c:spPr']);
    expect(elementOrder(xml, 'c:dPt')).toEqual(['c:idx', 'c:marker']);
    expect(section(xml, 'c:dPt')).toContain('<c:size val="72"/>'); // clamped from 100
    expect(xml).toContain('<a:ln w="19050"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln>');
    expect(attr(section(xml, 'c:ser', 0), 'c:smooth')).toBe('1');
    expect(attr(xml, 'c:dLblPos')).toBe('t');
    // second series: symbol none
    expect(section(xml, 'c:ser', 1)).toContain('<c:marker><c:symbol val="none"/></c:marker>');
    expect(attr(section(xml, 'c:ser', 1), 'c:smooth')).toBe('0');
    expect(elementOrder(section(xml, 'c:ser', 1), 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:cat', 'c:val', 'c:smooth']);
  });

  it('area stacked never writes dLblPos', () => {
    const xml = buildChartXml(chartSpecs.areaStacked());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:areaChart')).toEqual(['c:grouping', 'c:varyColors', 'c:ser', 'c:ser', 'c:dLbls', 'c:axId', 'c:axId']);
    expect(elementOrder(xml, 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:dLbls', 'c:cat', 'c:val']);
    expect(attr(xml, 'c:grouping')).toBe('stacked');
    expect(xml).not.toContain('c:dLblPos');
  });

  it('scatter lineMarker uses xVal/yVal and midCat', () => {
    const xml = buildChartXml(chartSpecs.scatter());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:scatterChart', 'c:valAx', 'c:valAx']);
    expect(elementOrder(xml, 'c:scatterChart')).toEqual(['c:scatterStyle', 'c:varyColors', 'c:ser', 'c:axId', 'c:axId']);
    expect(elementOrder(xml, 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:xVal', 'c:yVal', 'c:smooth']);
    expect(attr(xml, 'c:scatterStyle')).toBe('lineMarker');
    expect(formulas(xml)).toEqual([NAME1_F, "'Data'!$D$2:$D$5", VAL1_F]);
    const xVal = section(xml, 'c:xVal');
    expect(xVal).toContain('<c:numRef>');
    expect(xVal).toContain('<c:formatCode>General</c:formatCode><c:ptCount val="4"/>');
    expect([...xVal.matchAll(/<c:pt idx="(\d+)">/g)].map((m) => m[1])).toEqual(['0', '1', '3']);
    expect([...xml.matchAll(/<c:crossBetween val="(\w+)"/g)].map((m) => m[1])).toEqual(['midCat', 'midCat']);
    // marker line 1px → 9525 EMU
    expect(section(xml, 'c:marker')).toContain('<a:ln w="9525">');
  });

  it('bubble chart', () => {
    const xml = buildChartXml(chartSpecs.bubble());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:bubbleChart')).toEqual(['c:varyColors', 'c:ser', 'c:bubble3D', 'c:bubbleScale', 'c:showNegBubbles', 'c:axId', 'c:axId']);
    expect(elementOrder(xml, 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:xVal', 'c:yVal', 'c:bubbleSize', 'c:bubble3D']);
    expect(attr(xml, 'c:bubbleScale')).toBe('300'); // clamped from 400
    const sizes = section(xml, 'c:bubbleSize');
    expect(sizes).toContain("<c:f>'Data'!$E$2:$E$5</c:f>");
    expect(attr(sizes, 'c:ptCount')).toBe('4');
    expect([...sizes.matchAll(/<c:pt idx="(\d+)">/g)].map((m) => m[1])).toEqual(['0', '1', '3']);
  });

  it('bubble without sizes falls back to a literal of equal sizes', () => {
    const spec = chartSpecs.bubble();
    const g = spec.plotGroups[0]!;
    g.series[0]!.bubbleSizes = null;
    const xml = buildChartXml(spec);
    checkCommon(xml);
    expect(section(xml, 'c:bubbleSize')).toContain('<c:numLit>');
  });

  it('pie with explosion and bestFit labels', () => {
    const xml = buildChartXml(chartSpecs.pie());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:pieChart']);
    expect(elementOrder(xml, 'c:pieChart')).toEqual(['c:varyColors', 'c:ser', 'c:firstSliceAng']);
    expect(elementOrder(xml, 'c:ser')).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val']);
    expect(elementOrder(xml, 'c:dPt')).toEqual(['c:idx', 'c:explosion', 'c:spPr']);
    expect(attr(xml, 'c:explosion')).toBe('15');
    expect(attr(xml, 'c:dLblPos')).toBe('bestFit');
    expect(attr(xml, 'c:showPercent')).toBe('1');
    expect(attr(xml, 'c:firstSliceAng')).toBe('90');
    expect(attr(xml, 'c:varyColors')).toBe('1');
  });

  it('doughnut clamps holeSize/firstSliceAng and never writes dLblPos', () => {
    const xml = buildChartXml(chartSpecs.doughnut());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:doughnutChart')).toEqual(['c:varyColors', 'c:ser', 'c:firstSliceAng', 'c:holeSize']);
    expect(attr(xml, 'c:holeSize')).toBe('10');
    expect(attr(xml, 'c:firstSliceAng')).toBe('360');
    expect(xml).not.toContain('c:dLblPos');
  });

  it('combo: bar + line sharing axes', () => {
    const xml = buildChartXml(chartSpecs.combo());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:barChart', 'c:lineChart', 'c:catAx', 'c:valAx']);
    const lineGroup = section(xml, 'c:lineChart');
    expectSchemaOrder(elementOrder(lineGroup, 'c:lineChart'), ORDER.lineChart);
    expectSchemaOrder(elementOrder(lineGroup, 'c:ser'), ORDER.lineSer);
    expect([...xml.matchAll(/<c:axId val="(\d+)"\/>/g)].map((m) => m[1])).toEqual(['100', '200', '100', '200', '100', '200']);
    // showMarkers=false with marker=null → series explicitly hides markers
    expect(lineGroup).toContain('<c:marker><c:symbol val="none"/></c:marker>');
    expect(formulas(xml)).toEqual([NAME1_F, CAT_F, VAL1_F, NAME2_F, CAT_F, VAL2_F]);
  });

  it('secondary axis: two valAx, right one crosses max, its cat axis deleted', () => {
    const xml = buildChartXml(chartSpecs.secondary());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:barChart', 'c:lineChart', 'c:catAx', 'c:valAx', 'c:catAx', 'c:valAx']);
    const val2 = section(xml, 'c:valAx', 1);
    expect(attr(val2, 'c:axId')).toBe('201');
    expect(attr(val2, 'c:axPos')).toBe('r');
    expect(attr(val2, 'c:crosses')).toBe('max');
    expect(attr(val2, 'c:crossAx')).toBe('101');
    const cat2 = section(xml, 'c:catAx', 1);
    expect(attr(cat2, 'c:delete')).toBe('1');
    // axis title with two lines → two a:p
    const val1 = section(xml, 'c:valAx', 0);
    expect(elementOrder(val1, 'c:valAx')).toEqual([
      'c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:title', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:crossAx', 'c:crosses', 'c:crossBetween',
    ]);
    expect(elementOrder(val1, 'c:rich')).toEqual(['a:bodyPr', 'a:lstStyle', 'a:p', 'a:p']);
    expect(lineAxIds(xml)).toEqual(['101', '201']);
  });

  it('date axis', () => {
    const xml = buildChartXml(chartSpecs.dateAxis());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:lineChart', 'c:dateAx', 'c:valAx']);
    expect(elementOrder(xml, 'c:dateAx')).toEqual([
      'c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:auto', 'c:lblOffset', 'c:baseTimeUnit', 'c:majorUnit', 'c:majorTimeUnit',
    ]);
    const d = section(xml, 'c:dateAx');
    expect(attr(d, 'c:baseTimeUnit')).toBe('months');
    expect(attr(d, 'c:majorTimeUnit')).toBe('months');
    expect(attr(d, 'c:tickLblPos')).toBe('low');
    expect(d).toContain('<c:numFmt formatCode="mmm yy" sourceLinked="0"/>');
    // -45° → -2700000 (60000ths of a degree)
    expect(d).toContain('<a:bodyPr rot="-2700000" vert="horz"/>');
    expect(section(xml, 'c:cat')).toContain('<c:formatCode>yyyy-mm-dd</c:formatCode>');
    expect(attr(section(xml, 'c:valAx'), 'c:crossBetween')).toBe('between');
  });

  it('log axis with reversed orientation', () => {
    const xml = buildChartXml(chartSpecs.logAxis());
    checkCommon(xml);
    const v = section(xml, 'c:valAx');
    expect(elementOrder(v, 'c:scaling')).toEqual(['c:logBase', 'c:orientation', 'c:max', 'c:min']);
    expect(attr(v, 'c:logBase')).toBe('10');
    expect(attr(v, 'c:orientation')).toBe('maxMin');
    expect(attr(v, 'c:max')).toBe('10000');
    expect(attr(v, 'c:min')).toBe('1');
    expect(attr(v, 'c:crossesAt')).toBe('1');
    expect(elementOrder(v, 'c:valAx')).toEqual([
      'c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:crossAx', 'c:crossesAt', 'c:crossBetween', 'c:majorUnit',
    ]);
    // category axis majorUnit → label/tick skip
    const c = section(xml, 'c:catAx');
    expect(attr(c, 'c:tickLblSkip')).toBe('2');
    expectSchemaOrder(elementOrder(c, 'c:catAx'), ORDER.catAx);
  });

  it('gradient fill, manual layout, style, no title, no legend', () => {
    const xml = buildChartXml(chartSpecs.gradientLayout());
    checkCommon(xml);
    expect(elementOrder(xml, 'c:chartSpace')).toEqual(['c:date1904', 'c:lang', 'c:roundedCorners', 'c:style', 'c:chart', 'c:spPr', 'c:txPr']);
    expect(elementOrder(xml, 'c:chart')).toEqual(['c:autoTitleDeleted', 'c:plotArea', 'c:plotVisOnly', 'c:dispBlanksAs']);
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>');
    expect(attr(xml, 'c:dispBlanksAs')).toBe('span');
    expect(elementOrder(xml, 'c:plotArea')).toEqual(['c:layout', 'c:barChart', 'c:catAx', 'c:valAx', 'c:spPr']);
    expect(elementOrder(xml, 'c:manualLayout')).toEqual(['c:layoutTarget', 'c:xMode', 'c:yMode', 'c:x', 'c:y', 'c:w', 'c:h']);
    expect(attr(xml, 'c:layoutTarget')).toBe('inner');
    expect(attr(xml, 'c:h')).toBe('1'); // clamped from 1.5
    expect(attr(xml, 'c:x')).toBe('0.1');
    // percentStacked forces overlap 100
    expect(attr(xml, 'c:overlap')).toBe('100');
    const grad = section(xml, 'a:gradFill');
    expect(elementOrder(grad, 'a:gradFill')).toEqual(['a:gsLst', 'a:lin']);
    // stops sorted by position; 45° → 2700000
    expect([...grad.matchAll(/<a:gs pos="(\d+)">/g)].map((m) => m[1])).toEqual(['0', '100000']);
    expect(grad).toContain('<a:alpha val="50000"/>');
    expect(grad).toContain('<a:lin ang="2700000" scaled="1"/>');
    // 2px → 19050 EMU, alpha 0.8 → 80000, sysDot dash
    expect(xml).toContain('<a:ln w="19050"><a:solidFill><a:srgbClr val="1F4E79"><a:alpha val="80000"/></a:srgbClr></a:solidFill><a:prstDash val="sysDot"/><a:round/></a:ln>');
    expectSchemaOrder(elementOrder(section(xml, 'c:ser'), 'c:ser'), ORDER.barSer);
  });
});

function lineAxIds(xml: string): string[] {
  return [...section(xml, 'c:lineChart').matchAll(/<c:axId val="(\d+)"\/>/g)].map((m) => m[1]!);
}

describe('buildChartXml – every fixture is schema ordered', () => {
  for (const [name, make] of Object.entries(chartSpecs)) {
    it(name, () => {
      const xml = buildChartXml(make());
      checkCommon(xml);
      const kinds = elementOrder(xml, 'c:plotArea').filter((n) => n.endsWith('Chart'));
      for (const [i, k] of kinds.entries()) {
        const key = k.replace('c:', '') as keyof typeof ORDER;
        const group = section(xml, k, kinds.slice(0, i).filter((x) => x === k).length);
        expectSchemaOrder(elementOrder(group, k), ORDER[key]);
        const serKey = `${key.replace('Chart', '').replace('doughnut', 'pie')}Ser` as keyof typeof ORDER;
        expectSchemaOrder(elementOrder(group, 'c:ser'), ORDER[serKey]);
      }
    });
  }
});

describe('buildChartXml – edge cases', () => {
  const base = (): ReturnType<typeof chartSpecs.barClustered> => chartSpecs.barClustered();

  it('escapes formulas and text', () => {
    const spec = base();
    const s = spec.plotGroups[0]!.series[0]!;
    s.name = { kind: 'literal', text: 'A & B <c>' };
    s.values = { formula: "'R&D <x>'!$B$2:$B$3", cache: [1, 2] };
    const xml = buildChartXml(spec);
    expect(XMLValidator.validate(xml)).toBe(true);
    expect(xml).toContain('<c:tx><c:v>A &amp; B &lt;c&gt;</c:v></c:tx>');
    expect(xml).toContain("<c:f>'R&amp;D &lt;x&gt;'!$B$2:$B$3</c:f>");
  });

  it('coerces -0 and omits dLblPos invalid for the group', () => {
    const spec = chart(
      [
        {
          kind: 'line',
          grouping: 'standard',
          varyColors: false,
          showMarkers: true,
          axisIds: [1, 2],
          dataLabels: labels({ position: 'outEnd' }),
          series: [series(0, { values: { formula: VAL1_F, cache: [-0, 1, Number.POSITIVE_INFINITY, null] } })],
        },
      ],
      [axis(1, 'cat', 2), axis(2, 'val', 1, { scaling: { min: -0, max: null, orientation: 'minMax', logBase: null } })],
    );
    const xml = buildChartXml(spec);
    checkCommon(xml);
    expect(xml).not.toContain('c:dLblPos');
    expect(xml).toContain('<c:pt idx="0"><c:v>0</c:v></c:pt>');
    expect(xml).not.toContain('-0');
    expect(attr(xml, 'c:min')).toBe('0');
  });

  it('throws on impossible input', () => {
    const missingAxis = base();
    missingAxis.axes = missingAxis.axes.slice(0, 1);
    expect(() => buildChartXml(missingAxis)).toThrow(/unknown axis/);
    const empty = base();
    empty.plotGroups = [];
    expect(() => buildChartXml(empty)).toThrow(/no plot groups/);
    const badColor = base();
    badColor.chartArea = { fill: { type: 'solid', hex: 'red', alpha: 1 }, line: null };
    expect(() => buildChartXml(badColor)).toThrow(/RRGGBB/);
    const emptyFormula = base();
    emptyFormula.plotGroups[0]!.series[0]!.values.formula = '';
    expect(() => buildChartXml(emptyFormula)).toThrow(/empty formula/);
  });
});

describe('title line fonts', () => {
  it('applies per-line fonts to merged subtitle lines', async () => {
    const { richTextXml } = await import('../../src/excel/drawingml-xml');
    const xml = richTextXml({
      lines: ['Title', 'Subtitle'],
      font: { typeface: 'Georgia', sizeHundredthsPt: 1400, bold: true, italic: false, colorHex: '333333' },
      lineFonts: [null, { typeface: 'Georgia', sizeHundredthsPt: 900, bold: false, italic: false, colorHex: '666666' }],
      overlay: false,
    });
    const paras = xml.split('<a:p>').slice(1);
    expect(paras).toHaveLength(2);
    expect(paras[0]).toContain('sz="1400"');
    expect(paras[0]).toContain('b="1"');
    expect(paras[1]).toContain('sz="900"');
    expect(paras[1]).toContain('b="0"');
    expect(paras[1]).toContain('<a:t>Subtitle</a:t>');
  });
});
