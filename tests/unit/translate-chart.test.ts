import { describe, expect, it } from 'vitest';
import { translateChartModel, baseTimeUnitOf, type TranslateOptions, type TranslationResult } from '../../src/core/translate-chart';
import { DiagnosticCollector } from '../../src/types/diagnostics';
import type { ChartModel } from '../../src/types/chart-model';
import type { ExcelChartSpec, PlotGroupSpec, SheetSpec } from '../../src/excel/writer-interface';
import * as F from '../fixtures/chart-models';

function run(model: ChartModel, extra: Partial<TranslateOptions> = {}): { r: TranslationResult; d: DiagnosticCollector; chart: ExcelChartSpec } {
  const d = new DiagnosticCollector();
  const r = translateChartModel(model, {
    chartSheetName: 'Chart',
    dataSheetName: 'Data',
    includeSourceData: true,
    fidelity: 'best-effort',
    diagnostics: d,
    ...extra,
  });
  const drawing = r.sheets[0]?.drawings[0];
  const chart = (drawing && drawing.kind === 'chart' ? drawing.chart : null) as ExcelChartSpec;
  return { r, d, chart };
}

const codes = (d: DiagnosticCollector) => d.items.map((x) => x.code);
const headerRow = (sheet: SheetSpec) => sheet.rows[0]!.cells.map((c) => (c.value.type === 'string' ? c.value.value : '?'));
function group<K extends PlotGroupSpec['kind']>(chart: ExcelChartSpec, kind: K): Extract<PlotGroupSpec, { kind: K }> {
  return chart.plotGroups.find((g) => g.kind === kind) as Extract<PlotGroupSpec, { kind: K }>;
}

describe('translateChartModel: structure', () => {
  it('produces the chart sheet then the data sheet', () => {
    const { r, chart } = run(F.lineModel());
    expect(r.blocking).toBe(false);
    expect(r.excelChartType).toBe('line');
    expect(r.sheets.map((s) => s.name)).toEqual(['Chart', 'Data']);
    const chartSheet = r.sheets[0]!;
    expect(chartSheet.freezeHeaderRow).toBe(false);
    expect(chartSheet.rows).toEqual([{ row0: 0, cells: [{ col0: 0, row0: 0, value: { type: 'string', value: 'Monthly sales' }, style: { bold: true } }] }]);
    expect(chartSheet.drawings[0]).toMatchObject({
      kind: 'chart',
      name: 'Chart 1',
      anchor: { col0: 0, row0: 2, colOffsetPx: 0, rowOffsetPx: 0, widthPx: 600, heightPx: 400 },
    });
    expect(chart.title).toEqual({ lines: ['Monthly sales'], font: expect.objectContaining({ typeface: 'Helvetica', sizeHundredthsPt: 1350 }), overlay: false });
    expect(chart.textDefaults).toEqual({ typeface: 'Helvetica', sizeHundredthsPt: 900, bold: false, italic: false, colorHex: '333333' });
    expect(chart.dispBlanksAs).toBe('gap');
    expect(chart.style).toBeNull();
    expect(chart.legend).toMatchObject({ position: 'b', overlay: false });
  });

  it('references worksheet cells for every series', () => {
    const { chart } = run(F.lineModel());
    const s = group(chart, 'line').series[0]!;
    expect(s.name).toEqual({ kind: 'ref', formula: 'Data!$B$1', cache: 'Sales' });
    expect(s.categories).toEqual({ formula: 'Data!$A$2:$A$5', kind: 'str', cache: ['Jan', 'Feb', 'Mar', 'Apr'] });
    expect(s.values).toEqual({ formula: 'Data!$B$2:$B$5', cache: [1, 2, 3, 4], formatCode: 'General' });
    expect(s.shape.line).toMatchObject({ widthPx: 2, hex: '2CAFFE', noFill: false });
    expect(s.marker).toMatchObject({ symbol: 'circle', size: 6, fill: { type: 'solid', hex: '2CAFFE', alpha: 1 } });
    expect(s.invertIfNegative).toBe(false);
  });

  it('honours chartWidth/chartHeight, hidden data sheet and the reference image', () => {
    const png = new Uint8Array([1, 2, 3]);
    const { r } = run(F.lineModel(), { chartWidth: 300, chartHeight: 200, includeSourceData: false, referenceImage: { png, widthPx: 600, heightPx: 400 } });
    expect(r.sheets[1]!.hidden).toBe(true);
    const [chartD, img] = r.sheets[0]!.drawings;
    expect(chartD!.anchor).toMatchObject({ widthPx: 300, heightPx: 200 });
    // 300 + 20 px = 5 default columns (64 px) + 0 px.
    expect(img).toMatchObject({ kind: 'image', png, anchor: { col0: 5, row0: 2, colOffsetPx: 0, widthPx: 300, heightPx: 200 } });
  });

  it('deletes the title when the model has none', () => {
    const m = F.comboModel();
    const { r, chart } = run(m);
    expect(chart.title).toBeNull();
    expect(r.sheets[0]!.rows).toEqual([]);
  });

  it('does not mutate the model and leaves model.warnings alone', () => {
    const m = F.styledModel();
    const before = JSON.stringify(m);
    run(m);
    run(m, { fidelity: 'minimal' });
    expect(JSON.stringify(m)).toBe(before);
    expect(m.warnings).toEqual([]);
  });
});

describe('translateChartModel: combos and axes', () => {
  it('combo column+line shares the primary axes', () => {
    const { r, chart } = run(F.comboModel());
    expect(r.excelChartType).toBe('combo:column+line');
    expect(chart.plotGroups.map((g) => g.kind)).toEqual(['bar', 'line']);
    const bar = group(chart, 'bar');
    const line = group(chart, 'line');
    expect(bar.axisIds).toEqual([1000, 2000]);
    expect(line.axisIds).toEqual([1000, 2000]);
    expect(headerRow(r.sheets[1]!)).toEqual(['Category', 'Rainfall', 'Average']);
    expect(bar.series[0]!.values.formula).toBe('Data!$B$2:$B$5');
    expect(line.series[0]!.values.formula).toBe('Data!$C$2:$C$5');
    expect(line.series[0]!.categories!.formula).toBe('Data!$A$2:$A$5');
    expect([bar.series[0]!.idx, line.series[0]!.idx]).toEqual([0, 1]);
    expect([bar.series[0]!.order, line.series[0]!.order]).toEqual([0, 1]);
    expect(line.series[0]!.shape.line).toMatchObject({ widthPx: 3, dash: 'dash' });
    expect(chart.axes.map((a) => [a.id, a.kind, a.position, a.crossAxisId, a.crosses])).toEqual([
      [1000, 'cat', 'b', 2000, 'autoZero'],
      [2000, 'val', 'l', 1000, 'autoZero'],
    ]);
  });

  it('secondary axis: deleted cat axis 1001 and value axis 2001 on the right crossing at max', () => {
    const { chart, d, r } = run(F.secondaryAxisModel());
    expect(group(chart, 'bar').axisIds).toEqual([1000, 2000]);
    expect(group(chart, 'line').axisIds).toEqual([1001, 2001]);
    const cat2 = chart.axes.find((a) => a.id === 1001)!;
    const val2 = chart.axes.find((a) => a.id === 2001)!;
    expect(cat2).toMatchObject({ kind: 'cat', deleted: true, crossAxisId: 2001 });
    expect(val2).toMatchObject({ kind: 'val', position: 'r', crosses: 'max', crossAxisId: 1001, deleted: false });
    expect(val2.title!.lines).toEqual(['Temperature (°C)']);
    expect(chart.axes.find((a) => a.id === 2000)!.numberFormat).toEqual({ code: '0" mm"', sourceLinked: false });
    expect(d.items.find((x) => x.code === 'SECONDARY_AXIS')).toMatchObject({ outcome: 'translated', severity: 'info', property: 'yAxis[1]' });
    expect(group(chart, 'line').series[0]!.smooth).toBe(true);
    expect(r.supportedProperties).toContain('series[1].yAxis');
  });

  it('horizontal bars swap axis positions and map padding to gap/overlap', () => {
    const { chart } = run(F.barModel());
    const bar = group(chart, 'bar');
    expect(bar).toMatchObject({ barDir: 'bar', grouping: 'clustered', gapWidth: 25, overlap: 0, varyColors: false });
    expect(chart.axes.map((a) => [a.id, a.position, a.scaling.orientation, a.crosses])).toEqual([
      [1000, 'l', 'maxMin', 'autoZero'],
      [2000, 'b', 'minMax', 'max'],
    ]);
    expect(bar.series[0]!.shape.line).toEqual({ widthPx: 0, hex: null, alpha: 1, dash: 'solid', noFill: true });
  });

  it('reversed horizontal bars put a secondary value axis on top (crosses min)', () => {
    const m = F.barModel();
    m.yAxes.push(F.axis({ index: 1, kind: 'linear' }));
    m.series.push(F.series({ kind: 'column', index: 1, yAxisIndex: 1, points: F.pts([1, 2, 3, 4]) }));
    const { chart } = run(m);
    expect(chart.axes.find((a) => a.id === 2000)!.crosses).toBe('max');
    expect(chart.axes.find((a) => a.id === 2001)).toMatchObject({ position: 't', crosses: 'min' });
    expect(chart.axes.find((a) => a.id === 1001)!.scaling.orientation).toBe('maxMin');
  });

  it('a reversed x axis on a vertical chart keeps the value axis on the left via crosses max', () => {
    const m = F.comboModel();
    m.xAxes = [{ ...m.xAxes[0]!, reversed: true }];
    const { chart } = run(m);
    expect(chart.axes.map((a) => [a.scaling.orientation, a.crosses])).toEqual([
      ['maxMin', 'autoZero'],
      ['minMax', 'max'],
    ]);
  });

  it('a reversed + opposite y axis keeps the category axis at the bottom and the value axis on the right', () => {
    const m = F.comboModel();
    m.yAxes = [{ ...m.yAxes[0]!, reversed: true, opposite: true }];
    const { chart } = run(m);
    const [cat, val] = chart.axes;
    expect(cat!.crosses).toBe('max'); // value axis is maxMin, so "max" is at the bottom
    expect(val!.position).toBe('r');
    expect(val!.scaling.orientation).toBe('maxMin');
    expect(val!.crosses).toBe('max');
  });

  it('a reversed-only y axis still keeps the category axis at the bottom', () => {
    const m = F.comboModel();
    m.yAxes = [{ ...m.yAxes[0]!, reversed: true }];
    const { chart } = run(m);
    const [cat, val] = chart.axes;
    expect(cat!.crosses).toBe('max');
    expect(val!.crosses).toBe('autoZero');
  });

  it('stacked columns: overlap 100, gap from groupPadding, labels never outEnd', () => {
    const { chart, d } = run(F.columnStackedModel());
    const bar = group(chart, 'bar');
    expect(bar).toMatchObject({ grouping: 'stacked', overlap: 100, gapWidth: 67 });
    expect(bar.series[0]!.dataLabels!.position).toBe('inEnd');
    expect(d.items.find((x) => x.code === 'APPROXIMATED_DATA_LABELS')).toMatchObject({ property: 'series[0].dataLabels.position', seriesIndex: 0 });
    expect(d.items.find((x) => x.code === 'UNSUPPORTED_STYLE')).toMatchObject({ property: 'series[0].borderRadius' });
  });

  it('area percent stacking keeps fill opacity', () => {
    const { chart } = run(F.areaPercentModel());
    const area = group(chart, 'area');
    expect(area.grouping).toBe('percentStacked');
    expect(area.series[0]!.shape.fill).toEqual({ type: 'solid', hex: '2CAFFE', alpha: 0.75 });
  });

  it('datetime x → date axis with serial categories and a day-based major unit', () => {
    const { chart, r } = run(F.datetimeModel());
    const cat = chart.axes.find((a) => a.id === 1000)!;
    expect(cat).toMatchObject({ kind: 'date', dateAxis: { baseTimeUnit: 'days' }, majorUnit: 2, numberFormat: { code: 'ddd d mmm', sourceLinked: false } });
    const s = group(chart, 'line').series[0]!;
    expect(s.categories).toMatchObject({ kind: 'num', cache: [45292, 45293, 45294, 45295], formatCode: 'yyyy-mm-dd' });
    expect(r.sheets[1]!.rows[1]!.cells[0]).toEqual({ col0: 0, row0: 1, value: { type: 'number', value: 45292 }, style: { numberFormat: 'yyyy-mm-dd' } });
  });

  it('monthly data uses a months base unit; intraday data falls back to a category axis', () => {
    expect(baseTimeUnitOf([Date.UTC(2024, 0, 1), Date.UTC(2024, 1, 1), Date.UTC(2024, 2, 1)])).toBe('months');
    expect(baseTimeUnitOf([Date.UTC(2023, 0, 1), Date.UTC(2024, 0, 1)])).toBe('years');
    expect(baseTimeUnitOf([Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2)])).toBe('days');
    const m = F.datetimeModel();
    m.series[0] = { ...m.series[0]!, points: m.series[0]!.points.map((p, i) => ({ ...p, x: Date.UTC(2024, 0, 1, i) })) };
    const { chart, d } = run(m);
    expect(chart.axes[0]!.kind).toBe('cat');
    expect(codes(d)).toContain('APPROXIMATED_DATETIME');
    expect(group(chart, 'line').series[0]!.categories!.formatCode).toBe('yyyy-mm-dd hh:mm');
  });

  it('scatter uses value axes 3000/2000 and literal names', () => {
    const { chart } = run(F.scatterModel());
    const sc = group(chart, 'scatter');
    expect(sc).toMatchObject({ scatterStyle: 'marker', axisIds: [3000, 2000] });
    expect(sc.series[0]!.name).toEqual({ kind: 'literal', text: 'Men' });
    expect(sc.series[0]!.categories).toMatchObject({ formula: 'Data!$A$2:$A$4', kind: 'num' });
    expect(sc.series[0]!.shape.line!.noFill).toBe(true);
    expect(sc.series[0]!.marker).toMatchObject({ symbol: 'circle' });
    expect(chart.axes.map((a) => [a.id, a.kind, a.position, a.crossAxisId])).toEqual([
      [3000, 'val', 'b', 2000],
      [2000, 'val', 'l', 3000],
    ]);
    expect(chart.axes[0]!.title!.lines).toEqual(['Height']);
  });

  it('scatter inside a combo gets its own hidden axis pair', () => {
    const m = F.comboModel();
    m.series.push(F.series({ kind: 'scatter', index: 2, name: 'Pts', points: [F.point({ x: 0.5, y: 3 })], line: { color: null, width: 1, dash: 'solid' }, smooth: true }));
    const { chart, d } = run(m);
    expect(group(chart, 'scatter')).toMatchObject({ axisIds: [3000, 3001], scatterStyle: 'smoothMarker' });
    expect(chart.axes.filter((a) => a.id >= 3000).every((a) => a.deleted)).toBe(true);
    expect(codes(d)).toContain('APPROXIMATED_AXIS_SCALE');
  });

  it('bubble group has bubble sizes', () => {
    const { chart } = run(F.bubbleModel());
    const b = group(chart, 'bubble');
    expect(b.bubbleScale).toBe(100);
    expect(b.series[0]!.bubbleSizes!.formula).toBe('Data!$C$2:$C$3');
    expect(b.series[0]!.shape.fill).toEqual({ type: 'solid', hex: '2CAFFE', alpha: 0.5 });
  });

  it('reports multiple x axes', () => {
    const m = F.comboModel();
    m.xAxes.push(F.axis({ index: 1, kind: 'linear' }));
    expect(run(m).d.items.find((x) => x.code === 'MULTIPLE_X_AXES')).toMatchObject({ outcome: 'approximated', property: 'xAxis[1]' });
  });
});

describe('translateChartModel: pie and doughnut', () => {
  it('pie writes one dPt per point with its color, explosion for sliced points, no axes', () => {
    const { chart, r } = run(F.pieModel());
    expect(r.excelChartType).toBe('pie');
    const pie = group(chart, 'pie');
    expect(pie).toMatchObject({ varyColors: true, firstSliceAngle: 270 });
    expect(chart.axes).toEqual([]);
    expect(chart.legend).toBeNull();
    expect(chart.plotArea.manualLayout).toBeNull();
    const s = pie.series[0]!;
    expect(s.dataPoints.map((p) => [p.idx, p.shape?.fill])).toEqual([
      [0, { type: 'solid', hex: '2CAFFE', alpha: 1 }],
      [1, { type: 'solid', hex: '544FC5', alpha: 1 }],
      [2, { type: 'solid', hex: '00E272', alpha: 1 }],
      [3, { type: 'solid', hex: 'FE6A35', alpha: 1 }],
    ]);
    expect(s.dataPoints.map((p) => p.explosion)).toEqual([null, 3, null, null]);
    expect(s.dataPoints[0]!.shape!.line).toMatchObject({ hex: 'FFFFFF', widthPx: 1 });
    expect(s.dataLabels).toMatchObject({ showCategoryName: true, showValue: false, position: 'outEnd' });
    expect(r.supportedProperties).toEqual(expect.arrayContaining(['series[0].data[1].sliced', 'series[0].data[0].color']));
  });

  it('doughnut rings map points to the union rows', () => {
    const { chart, d } = run(F.doughnutModel());
    const g = group(chart, 'doughnut');
    expect(g.holeSize).toBe(40);
    expect(g.series).toHaveLength(2);
    expect(g.series[1]!.dataPoints.map((p) => [p.idx, p.shape?.fill])).toEqual([
      [0, { type: 'solid', hex: '00E272', alpha: 1 }],
      [2, { type: 'solid', hex: 'FE6A35', alpha: 1 }],
    ]);
    expect(codes(d)).toContain('APPROXIMATED_LAYOUT');
  });
});

describe('translateChartModel: styling', () => {
  it('translates fonts, backgrounds, gridlines, legend, data labels and layout', () => {
    const { chart, d, r } = run(F.styledModel());
    expect(chart.title!.lines).toEqual(['Styled', 'Second line']);
    expect(chart.title!.font).toEqual({ typeface: 'Open Sans', sizeHundredthsPt: 1500, bold: true, italic: false, colorHex: '112233' });
    expect(d.items.find((x) => x.property === 'subtitle.text')).toMatchObject({ code: 'APPROXIMATED_LAYOUT', severity: 'info' });
    expect(chart.chartArea.fill).toEqual({
      type: 'gradient',
      angle: 90,
      stops: [
        { pos: 0, hex: 'FFFFFF', alpha: 1 },
        { pos: 1, hex: 'DDDDDD', alpha: 1 },
      ],
    });
    expect(chart.chartArea.line).toMatchObject({ widthPx: 2, hex: '333333' });
    expect(chart.plotArea.fill).toEqual({ type: 'solid', hex: 'FAFAFA', alpha: 1 });
    expect(chart.plotArea.manualLayout).toEqual({ x: 0.1, y: 0.12, w: 0.75, h: 0.7 });
    expect(d.items.find((x) => x.property === 'chart.plotArea')).toMatchObject({ code: 'APPROXIMATED_LAYOUT' });
    const val = chart.axes.find((a) => a.id === 2000)!;
    expect(val.majorGridlines).toEqual({ widthPx: 2, hex: 'CCCCCC', alpha: 1, dash: 'dash', noFill: false });
    expect(val.scaling).toEqual({ min: 0, max: 10, orientation: 'minMax', logBase: null });
    expect(val.majorUnit).toBe(2);
    expect(val.title!.font!.sizeHundredthsPt).toBe(1050);
    expect(val.labels.font!.sizeHundredthsPt).toBe(825);
    expect(chart.axes.find((a) => a.id === 1000)!).toMatchObject({ majorTickMark: 'out', majorGridlines: null });
    expect(chart.legend).toEqual({
      position: 'r',
      overlay: false,
      font: expect.objectContaining({ bold: true, sizeHundredthsPt: 825 }),
      fill: { type: 'solid', hex: 'EEEEEE', alpha: 1 },
      line: expect.objectContaining({ hex: '999999', widthPx: 1 }),
    });
    const s = group(chart, 'bar').series[0]!;
    expect(s.shape).toEqual({ fill: { type: 'solid', hex: 'FF0000', alpha: 1 }, line: expect.objectContaining({ hex: '000000', widthPx: 1 }) });
    expect(s.dataPoints).toEqual([{ idx: 2, shape: { fill: { type: 'solid', hex: '00FF00', alpha: 1 }, line: null }, marker: null, explosion: null, dataLabels: null }]);
    expect(s.dataLabels).toMatchObject({ position: 'outEnd', numberFormat: '0.0', font: expect.objectContaining({ bold: true }) });
    expect(r.supportedProperties).toEqual(
      expect.arrayContaining([
        'chart.type',
        'title.text',
        'title.style',
        'series[0].name',
        'series[0].color',
        'series[0].data',
        'xAxis[0].categories',
        'yAxis[0].title.text',
        'yAxis[0].min',
        'yAxis[0].max',
        'yAxis[0].gridLineWidth',
        'legend.align',
        'legend.itemStyle',
        'chart.backgroundColor',
        'chart.plotBackgroundColor',
        'series[0].dataLabels',
      ]),
    );
  });

  it('maps legend positions', () => {
    const pos = (position: ChartModel['legend']['position']) => {
      const m = F.lineModel();
      m.legend = { ...m.legend, position };
      return run(m);
    };
    expect(pos('topRight').chart.legend!.position).toBe('tr');
    expect(pos('left').chart.legend!.position).toBe('l');
    const tl = pos('topLeft');
    expect(tl.chart.legend!.position).toBe('t');
    expect(codes(tl.d)).toContain('APPROXIMATED_LEGEND_POSITION');
    const br = pos('bottomRight');
    expect(br.chart.legend!.position).toBe('b');
    expect(br.d.items.find((x) => x.code === 'APPROXIMATED_LEGEND_POSITION')!.outcome).toBe('approximated');
  });

  it('minimal fidelity drops styles but keeps data, types, titles, formats and legend position', () => {
    const { chart, d } = run(F.styledModel(), { fidelity: 'minimal' });
    expect(chart.textDefaults).toBeNull();
    expect(chart.chartArea).toEqual({ fill: null, line: null });
    expect(chart.plotArea).toEqual({ fill: null, line: null, manualLayout: null });
    expect(chart.title).toEqual({ lines: ['Styled', 'Second line'], font: null, overlay: false });
    expect(chart.legend).toEqual({ position: 'r', overlay: false, font: null, fill: null, line: null });
    const s = group(chart, 'bar').series[0]!;
    expect(s.shape).toEqual({ fill: null, line: null });
    expect(s.marker).toBeNull();
    expect(s.dataPoints).toEqual([]);
    expect(s.values.formula).toBe('Data!$B$2:$B$5');
    expect(s.dataLabels).toMatchObject({ numberFormat: '0.0', font: null, fill: null, line: null });
    for (const a of chart.axes) {
      expect(a.labels.font).toBeNull();
      expect(a.axisLine).toBeNull();
      expect(a.title?.font ?? null).toBeNull();
    }
    expect(chart.axes.find((a) => a.id === 2000)!.majorGridlines).toMatchObject({ hex: null, noFill: false });
    expect(codes(d)).not.toContain('APPROXIMATED_FONT');
  });

  it('minimal pie keeps explosion but drops slice colors', () => {
    const { chart } = run(F.pieModel(), { fidelity: 'minimal' });
    const s = group(chart, 'pie').series[0]!;
    expect(s.dataPoints).toEqual([{ idx: 1, shape: null, marker: null, explosion: 3, dataLabels: null }]);
  });
});

describe('translateChartModel: data concerns', () => {
  it('writes blanks for nulls and keeps negative values', () => {
    const { chart, d } = run(F.nullNegativeModel());
    expect(group(chart, 'bar').series[0]!.values.cache).toEqual([5, null, -3, 7]);
    expect(codes(d)).toContain('NULL_VALUES');
    expect(chart.axes[0]!.labels.position).toBe('low');
  });

  it('keeps formula-like strings as strings', () => {
    const { r, chart, d } = run(F.formulaLikeModel());
    expect(r.sheets[1]!.rows[1]!.cells[0]!.value).toEqual({ type: 'string', value: '=SUM(A1)' });
    expect(r.sheets[1]!.rows[0]!.cells[1]!.value).toEqual({ type: 'string', value: '+foo' });
    expect(group(chart, 'bar').series[0]!.name).toEqual({ kind: 'ref', formula: 'Data!$B$1', cache: '+foo' });
    expect(codes(d)).toContain('FORMULA_LIKE_TEXT_ESCAPED');
  });

  it('reports hidden series included in the export', () => {
    const { d, chart } = run(F.hiddenSeriesModel());
    expect(group(chart, 'line').series).toHaveLength(2);
    expect(d.items.find((x) => x.code === 'HIDDEN_SERIES_INCLUDED')!.seriesIndex).toBe(1);
  });
});

describe('translateChartModel: sheet names', () => {
  it('sanitizes names and keeps them unique', () => {
    const { r, d, chart } = run(F.lineModel(), { chartSheetName: 'My:Chart?', dataSheetName: 'mychart', takenSheetNames: new Set(['Other']) });
    expect(r.chartSheetName).toBe('MyChart');
    expect(r.dataSheetName).toBe('mychart (2)');
    expect(r.sheets.map((s) => s.name)).toEqual(['MyChart', 'mychart (2)']);
    expect(d.items.filter((x) => x.code === 'SHEET_NAME_ADJUSTED').map((x) => x.property)).toEqual(['chartSheetName', 'dataSheetName']);
    expect(group(chart, 'line').series[0]!.values.formula).toBe("'mychart (2)'!$B$2:$B$5");
  });

  it('avoids names already taken in the workbook', () => {
    const { r } = run(F.lineModel(), { takenSheetNames: new Set(['Chart', 'Data']) });
    expect([r.chartSheetName, r.dataSheetName]).toEqual(['Chart (2)', 'Data (2)']);
  });
});

describe('translateChartModel: blocking', () => {
  it.each([
    ['polarModel', 'UNSUPPORTED_POLAR'],
    ['unknownTypeModel', 'UNSUPPORTED_CHART_TYPE'],
    ['emptyModel', 'EMPTY_CHART'],
  ] as const)('%s is blocking with %s', (name, code) => {
    const { r, d } = run(F[name]());
    expect(r.blocking).toBe(true);
    expect(r.sheets).toEqual([]);
    expect(r.excelChartType).toBeNull();
    expect(r.chartSheetName).toBe('Chart');
    expect(d.items.find((x) => x.code === code)!.outcome).toBe('blocking');
    expect(d.hasBlocking()).toBe(true);
  });

  it('pie combined with other types is blocking', () => {
    const m = F.comboModel();
    m.series.push(F.series({ kind: 'pie', index: 2, points: [F.point({ name: 'a', y: 1 })] }));
    expect(run(m).r.blocking).toBe(true);
  });
});
