/**
 * Fidelity backlog (roadmap item 5): per-ring doughnut categories, polar → radar, error bars,
 * range series as floating bars/areas, and the CSS-variable overrides of headless styled mode.
 * Core (IR → WorkbookSpec) and writer (WorkbookSpec → XML) level; the live-chart side is covered
 * by the shared extractor / export / menu suites.
 */

import { describe, expect, it } from 'vitest';
import { XMLValidator } from 'fast-xml-parser';
import { resolveChartType, CHART_TYPE_MATRIX } from '../../src/core/chart-type-registry';
import { buildDataLayout, type DataLayout } from '../../src/core/data-layout';
import { translateChartModel, type TranslationResult } from '../../src/core/translate-chart';
import { buildChartXml } from '../../src/excel/chart-xml';
import { buildWorksheetXml } from '../../src/excel/worksheet-xml';
import { buildStyles } from '../../src/excel/styles-xml';
import { createCssVariableResolver } from '../../src/highcharts/css-resolver';
import type { ChartModel, PointModel, SeriesModel } from '../../src/types/chart-model';
import { DiagnosticCollector } from '../../src/types/diagnostics';
import type { CellValue, ExcelChartSpec, PlotGroupSpec, SheetSpec } from '../../src/excel/writer-interface';
import { parseColor } from '../../src/utils/colors';
import { elementOrder } from '../helpers/inspect-xlsx';
import * as F from '../fixtures/chart-models';
import { axis, chart as chartSpec, series as serSpec } from '../fixtures/writer-specs';

function layoutOf(model: ChartModel): DataLayout {
  return buildDataLayout(model, resolveChartType(model), {
    sheetName: 'Data',
    hidden: false,
    dateFormatCode: () => 'yyyy-mm-dd',
  });
}

function run(model: ChartModel, fidelity: 'best-effort' | 'minimal' = 'best-effort') {
  const d = new DiagnosticCollector();
  const r: TranslationResult = translateChartModel(model, {
    chartSheetName: 'Chart',
    dataSheetName: 'Data',
    includeSourceData: true,
    fidelity,
    diagnostics: d,
  });
  const drawing = r.sheets[0]?.drawings[0];
  const chart = (drawing && drawing.kind === 'chart' ? drawing.chart : null) as ExcelChartSpec;
  return { r, d, chart };
}

function cell(sheet: SheetSpec, col0: number, row0: number): CellValue | undefined {
  return sheet.rows.find((r) => r.row0 === row0)?.cells.find((c) => c.col0 === col0)?.value;
}

function headerRow(sheet: SheetSpec): string[] {
  return (sheet.rows.find((r) => r.row0 === 0)?.cells ?? []).map((c) =>
    c.value.type === 'string' ? c.value.value : '?',
  );
}

function group<K extends PlotGroupSpec['kind']>(chart: ExcelChartSpec, kind: K): Extract<PlotGroupSpec, { kind: K }> {
  return chart.plotGroups.find((g) => g.kind === kind) as Extract<PlotGroupSpec, { kind: K }>;
}

const rangePoint = (x: number, low: number | null, high: number | null, extra: Partial<PointModel> = {}): PointModel =>
  F.point({ x, y: null, low, high, isNull: low === null || high === null, ...extra });

// ---------------------------------------------------------------------------
// 1. Per-ring doughnut categories
// ---------------------------------------------------------------------------

describe('1. doughnut rings with their own categories', () => {
  it('gives each ring its own category column when the name sets differ', () => {
    const l = layoutOf(F.doughnutModel());
    expect(headerRow(l.sheet)).toEqual(['Category 1', 'Inner', 'Category 2', 'Outer']);
    const [inner, outer] = l.ranges;
    expect(inner!.categories).toEqual({ formula: 'Data!$A$2:$A$3', kind: 'str', cache: ['A', 'B'] });
    expect(inner!.values).toMatchObject({ formula: 'Data!$B$2:$B$3', cache: [1, 2] });
    expect(outer!.categories).toEqual({ formula: 'Data!$C$2:$C$3', kind: 'str', cache: ['A', 'C'] });
    expect(outer!.values).toMatchObject({ formula: 'Data!$D$2:$D$3', cache: [5, 6] });
    expect(outer!.pointOffsets).toEqual([0, 1]);
    expect(outer!.nameRef).toEqual({ formula: 'Data!$D$1', cache: 'Outer' });
    expect(l.diagnostics.filter((d) => d.code === 'APPROXIMATED_LAYOUT')).toEqual([]);
  });

  it('keeps each ring in its own slice order when the rings share names in a different order', () => {
    const m = F.doughnutModel();
    m.series[1] = {
      ...m.series[1]!,
      points: ['B', 'A'].map((name, i) => F.point({ name, y: i + 5 })),
    };
    m.series[0] = { ...m.series[0]!, points: ['A', 'B'].map((name, i) => F.point({ name, y: i + 1 })) };
    const l = layoutOf(m);
    expect(headerRow(l.sheet)).toEqual(['Category 1', 'Inner', 'Category 2', 'Outer']);
    expect(l.ranges[1]!.categories!.cache).toEqual(['B', 'A']);
    expect(l.ranges[1]!.values.cache).toEqual([5, 6]);
  });

  it('keeps the shared category column when every ring lists the same names in the same order', () => {
    const m = F.doughnutModel();
    m.series[1] = { ...m.series[1]!, points: ['A', 'B'].map((name, i) => F.point({ name, y: i + 5 })) };
    const l = layoutOf(m);
    expect(headerRow(l.sheet)).toEqual(['Category', 'Inner', 'Outer']);
    expect(l.ranges[1]!.categories!.formula).toBe('Data!$A$2:$A$3');
  });

  it('translates per-ring categories into each doughnut series c:cat', () => {
    const { chart, d } = run(F.doughnutModel());
    const g = group(chart, 'doughnut');
    expect(g.series.map((s) => s.categories?.formula)).toEqual(['Data!$A$2:$A$3', 'Data!$C$2:$C$3']);
    expect(d.items.filter((x) => x.code === 'APPROXIMATED_LAYOUT' && x.property === 'series[1].data')).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 2. Radar for polar charts
// ---------------------------------------------------------------------------

describe('2. polar charts as Excel radar charts', () => {
  it('maps polar line series to a marker radar group with an info diagnostic', () => {
    const r = resolveChartType(F.polarModel());
    expect(r.blocking).toBe(false);
    expect(r.excelChartType).toBe('radar');
    expect(r.groups).toEqual([expect.objectContaining({ kind: 'radar', radarStyle: 'marker', seriesIndices: [0] })]);
    expect(r.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'APPROXIMATED_CHART_TYPE',
        outcome: 'approximated',
        severity: 'info',
        property: 'chart.polar',
      }),
    );
  });

  it('uses the standard style without markers and the filled style for areas', () => {
    const noMarkers = F.polarModel();
    noMarkers.series[0] = { ...noMarkers.series[0]!, marker: { ...noMarkers.series[0]!.marker!, enabled: false } };
    expect(resolveChartType(noMarkers).groups[0]!.radarStyle).toBe('standard');
    const area = F.polarModel();
    area.series[0] = { ...area.series[0]!, kind: 'area', fillOpacity: 0.5 };
    const r = resolveChartType(area);
    expect(r.groups[0]!.radarStyle).toBe('filled');
    expect(r.excelChartType).toBe('filledRadar');
  });

  it('keeps polar columns blocking (Excel has no polar columns)', () => {
    const m = F.polarModel();
    m.series[0] = { ...m.series[0]!, kind: 'column' };
    const r = resolveChartType(m);
    expect(r.blocking).toBe(true);
    expect(r.diagnostics).toContainEqual(expect.objectContaining({ code: 'UNSUPPORTED_POLAR', outcome: 'blocking' }));
  });

  it('keeps polar charts without a drawable series blocking on chart.polar', () => {
    const m = F.polarModel();
    m.series[0] = { ...m.series[0]!, kind: 'unknown', sourceType: 'radar' };
    const r = resolveChartType(m);
    expect(r.blocking).toBe(true);
    expect(r.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'UNSUPPORTED_POLAR', outcome: 'blocking', property: 'chart.polar' }),
    );
  });

  it('builds a radar plot group on a category/value axis pair', () => {
    const { r, chart } = run(F.polarModel());
    expect(r.blocking).toBe(false);
    const g = group(chart, 'radar');
    expect(g).toMatchObject({ kind: 'radar', radarStyle: 'marker', varyColors: false, axisIds: [1000, 2000] });
    expect(g.series[0]!.values.formula).toBe('Data!$B$2:$B$5');
    expect(g.series[0]!.categories!.formula).toBe('Data!$A$2:$A$5');
    expect(g.series[0]!.smooth).toBeNull();
    // Labels sit outside the circle: Excel lays the radar out itself (no pinned plot box).
    expect(chart.plotArea.manualLayout).toBeNull();
    expect(chart.axes.map((a) => [a.id, a.kind])).toEqual([
      [1000, 'cat'],
      [2000, 'val'],
    ]);
  });

  it('writes c:radarChart in schema order', () => {
    const { chart } = run(F.polarModel());
    const xml = buildChartXml(chart);
    expect(XMLValidator.validate(xml)).toBe(true);
    expect(elementOrder(xml, 'c:radarChart')).toEqual(['c:radarStyle', 'c:varyColors', 'c:ser', 'c:axId', 'c:axId']);
    expect(xml).toContain('<c:radarStyle val="marker"/>');
    const ser = elementOrder(xml, 'c:ser');
    expect(ser).toEqual(['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:cat', 'c:val']);
  });

  it('the documentation matrix lists polar as approximated', () => {
    expect(CHART_TYPE_MATRIX.find((e) => e.highcharts.startsWith('polar'))).toMatchObject({
      support: 'approximated',
    });
  });
});

// ---------------------------------------------------------------------------
// 3. Error bars
// ---------------------------------------------------------------------------

function errorBarModel(parentKind: SeriesModel['kind'] = 'column'): ChartModel {
  const parentPoints = F.pts([10, 20, 30, 40]);
  return F.baseModel({
    series: [
      F.series({ kind: parentKind, index: 0, id: 'rain', name: 'Rain', points: parentPoints }),
      F.series({
        kind: 'errorbar',
        sourceType: 'errorbar',
        index: 1,
        name: 'Rain error',
        linkedTo: 'rain',
        color: F.rgb('#000000'),
        line: { color: F.rgb('#000000'), width: 2, dash: 'solid' },
        points: [rangePoint(0, 8, 13), rangePoint(1, 18, 21), rangePoint(2, null, null), rangePoint(3, 35, 44)],
      }),
    ],
  });
}

describe('3. error bars', () => {
  it('attaches a linked errorbar series to its parent instead of drawing it', () => {
    const r = resolveChartType(errorBarModel());
    expect(r.blocking).toBe(false);
    expect(r.groups.map((g) => [g.kind, g.seriesIndices])).toEqual([['bar', [0]]]);
    expect(r.errorBars).toEqual([{ parent: 0, errorBar: 1 }]);
    expect(r.droppedSeries).toEqual([]);
  });

  it('reports an unlinked errorbar series as unsupported', () => {
    const m = errorBarModel();
    m.series[1] = { ...m.series[1]!, linkedTo: null };
    const r = resolveChartType(m);
    expect(r.errorBars).toEqual([]);
    expect(r.droppedSeries).toEqual([1]);
    expect(r.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'UNSUPPORTED_SERIES_TYPE', property: 'series[1].linkedTo', seriesIndex: 1 }),
    );
  });

  it('writes plus/minus columns after the parent column', () => {
    const l = layoutOf(errorBarModel());
    expect(headerRow(l.sheet)).toEqual(['Category', 'Rain', 'Rain +err', 'Rain -err']);
    const r = l.ranges[0]!;
    expect(r.errorBars).toMatchObject({
      seriesIndex: 1,
      plus: { formula: 'Data!$C$2:$C$5', cache: [3, 1, null, 4] },
      minus: { formula: 'Data!$D$2:$D$5', cache: [2, 2, null, 5] },
    });
    expect(cell(l.sheet, 2, 1)).toEqual({ type: 'number', value: 3 });
    expect(cell(l.sheet, 3, 3)).toBeUndefined();
  });

  it('carries the error bars onto the parent series spec with the errorbar line', () => {
    const { chart, r } = run(errorBarModel());
    expect(r.excelChartType).toBe('column');
    const s = group(chart, 'bar').series;
    expect(s).toHaveLength(1);
    expect(s[0]!.errorBars).toMatchObject({
      plus: { formula: 'Data!$C$2:$C$5' },
      minus: { formula: 'Data!$D$2:$D$5' },
      line: { hex: '000000', noFill: false },
    });
  });

  it('writes c:errBars between c:dLbls and c:cat (bar) and adds errDir for scatter', () => {
    const { chart } = run(errorBarModel());
    const xml = buildChartXml(chart);
    expect(XMLValidator.validate(xml)).toBe(true);
    expect(elementOrder(xml, 'c:ser')).toEqual([
      'c:idx',
      'c:order',
      'c:tx',
      'c:spPr',
      'c:invertIfNegative',
      'c:errBars',
      'c:cat',
      'c:val',
    ]);
    expect(elementOrder(xml, 'c:errBars')).toEqual([
      'c:errBarType',
      'c:errValType',
      'c:noEndCap',
      'c:plus',
      'c:minus',
      'c:spPr',
    ]);
    expect(xml).toContain('<c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="0"/>');
    expect(xml).toContain('<c:plus><c:numRef><c:f>Data!$C$2:$C$5</c:f>');

    const line = buildChartXml(run(errorBarModel('line')).chart);
    const lineSer = elementOrder(line, 'c:ser');
    expect(lineSer.indexOf('c:errBars')).toBe(lineSer.indexOf('c:cat') - 1);

    const m = errorBarModel('scatter');
    m.xAxes = [F.axis({ index: 0, kind: 'linear' })];
    const sc = run(m);
    expect(headerRow(sc.r.sheets[1]!)).toEqual(['Rain X', 'Rain Y', 'Rain +err', 'Rain -err']);
    const sxml = buildChartXml(sc.chart);
    expect(elementOrder(sxml, 'c:errBars')[0]).toBe('c:errDir');
    const sser = elementOrder(sxml, 'c:ser');
    expect(sser.indexOf('c:errBars')).toBe(sser.indexOf('c:xVal') - 1);
  });
});

// ---------------------------------------------------------------------------
// 4. Range series as floating bars / areas
// ---------------------------------------------------------------------------

function rangeModel(kind: 'columnrange' | 'arearange', lows: Array<number | null>, highs: Array<number | null>) {
  return F.baseModel({
    series: [
      F.series({
        kind,
        sourceType: kind,
        index: 0,
        name: 'Temp',
        fillOpacity: kind === 'arearange' ? 0.75 : 1,
        bars: kind === 'columnrange' ? { pointPadding: 0.1, groupPadding: 0.2, borderRadius: 0 } : null,
        points: lows.map((low, i) => rangePoint(i, low, highs[i] ?? null)),
      }),
    ],
  });
}

describe('4. range series', () => {
  it('resolves a column range to a stacked column group and reports the approximation', () => {
    const r = resolveChartType(rangeModel('columnrange', [1, 2, 3, 4], [5, 6, 7, 8]));
    expect(r.blocking).toBe(false);
    expect(r.excelChartType).toBe('stackedColumn');
    expect(r.groups).toEqual([
      expect.objectContaining({ kind: 'bar', stacking: 'normal', range: true, barDir: 'col', seriesIndices: [0] }),
    ]);
    expect(r.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'APPROXIMATED_CHART_TYPE',
        severity: 'info',
        property: 'series[0].type',
        message: expect.stringContaining('hidden base series'),
      }),
    );
  });

  it('writes Low, High and a Range formula column', () => {
    const l = layoutOf(rangeModel('columnrange', [1, 2, null, 4], [5, 6, 7, 9]));
    expect(headerRow(l.sheet)).toEqual(['Category', 'Temp Low', 'Temp High', 'Temp Range']);
    expect(cell(l.sheet, 3, 1)).toEqual({ type: 'formula', formula: 'C2-B2', cached: 4 });
    expect(cell(l.sheet, 3, 4)).toEqual({ type: 'formula', formula: 'C5-B5', cached: 5 });
    // A point without a low has no range: nothing is fabricated.
    expect(cell(l.sheet, 1, 3)).toBeUndefined();
    expect(cell(l.sheet, 3, 3)).toBeUndefined();
    const r = l.ranges[0]!;
    expect(r.values).toMatchObject({ formula: 'Data!$D$2:$D$5', cache: [4, 4, null, 5] });
    expect(r.rangeParts).toMatchObject({
      base: { formula: 'Data!$B$2:$B$5', cache: [1, 2, null, 4], name: { formula: 'Data!$B$1', cache: 'Temp Low' } },
      down: null,
    });
    expect(r.literalName).toBe('Temp');
  });

  it('builds a hidden base series and the visible range series; the base has no legend entry', () => {
    for (const fidelity of ['best-effort', 'minimal'] as const) {
      const { chart } = run(rangeModel('columnrange', [1, 2, 3, 4], [5, 6, 7, 8]), fidelity);
      const g = group(chart, 'bar');
      expect(g.grouping).toBe('stacked');
      expect(g.overlap).toBe(100);
      expect(g.series).toHaveLength(2);
      const [base, range] = g.series;
      expect(base!.values.formula).toBe('Data!$B$2:$B$5');
      expect(base!.shape).toEqual({
        fill: { type: 'none' },
        line: expect.objectContaining({ noFill: true }),
      });
      expect(range!.values.formula).toBe('Data!$D$2:$D$5');
      expect(range!.name).toEqual({ kind: 'literal', text: 'Temp' });
      expect(chart.legend!.deletedEntries).toEqual([base!.idx]);
    }
    const { chart } = run(rangeModel('columnrange', [1, 2, 3, 4], [5, 6, 7, 8]));
    expect(group(chart, 'bar').series[1]!.shape.fill).toMatchObject({ type: 'solid', hex: '2CAFFE' });
  });

  it('splits a column range that crosses zero into base/up/down helper columns', () => {
    const l = layoutOf(rangeModel('columnrange', [-9.5, 2, -7], [8, 6, -1]));
    expect(headerRow(l.sheet)).toEqual(['Category', 'Temp Low', 'Temp High', 'Temp Base', 'Temp Up', 'Temp Down']);
    expect(cell(l.sheet, 3, 1)).toEqual({ type: 'formula', formula: 'MAX(0,B2)+MIN(0,C2)', cached: 0 });
    expect(cell(l.sheet, 4, 1)).toEqual({ type: 'formula', formula: 'MAX(0,C2)-MAX(0,B2)', cached: 8 });
    expect(cell(l.sheet, 5, 1)).toEqual({ type: 'formula', formula: 'MIN(0,B2)-MIN(0,C2)', cached: -9.5 });
    expect(cell(l.sheet, 3, 2)).toMatchObject({ cached: 2 });
    expect(cell(l.sheet, 4, 2)).toMatchObject({ cached: 4 });
    expect(cell(l.sheet, 5, 2)).toMatchObject({ cached: 0 });
    expect(cell(l.sheet, 3, 3)).toMatchObject({ cached: -1 });
    expect(cell(l.sheet, 4, 3)).toMatchObject({ cached: 0 });
    expect(cell(l.sheet, 5, 3)).toMatchObject({ cached: -6 });
    const { chart } = run(rangeModel('columnrange', [-9.5, 2, -7], [8, 6, -1]));
    const g = group(chart, 'bar');
    expect(g.series.map((s) => s.values.formula)).toEqual(['Data!$D$2:$D$4', 'Data!$E$2:$E$4', 'Data!$F$2:$F$4']);
    expect(chart.legend!.deletedEntries).toEqual([g.series[0]!.idx, g.series[2]!.idx]);
  });

  it('maps an area range to a stacked area group with a hidden base', () => {
    const { r, chart } = run(rangeModel('arearange', [-3, -2, -1, 0], [3, 4, 5, 6]));
    expect(r.excelChartType).toBe('stackedArea');
    const g = group(chart, 'area');
    expect(g.grouping).toBe('stacked');
    expect(g.series.map((s) => s.values.formula)).toEqual(['Data!$B$2:$B$5', 'Data!$D$2:$D$5']);
    expect(g.series[0]!.shape.fill).toEqual({ type: 'none' });
    expect(headerRow(r.sheets[1]!)).toEqual(['Category', 'Temp Low', 'Temp High', 'Temp Range']);
  });

  it('writes formula cells with a cached value and rejects anything but library formulas', () => {
    const sheet: SheetSpec = {
      name: 'Data',
      hidden: false,
      columns: [],
      freezeHeaderRow: false,
      drawings: [],
      rows: [
        {
          row0: 1,
          cells: [
            { col0: 3, row0: 1, value: { type: 'formula', formula: 'C2-B2', cached: 17.5 } },
            { col0: 4, row0: 1, value: { type: 'formula', formula: 'MAX(0,C2)-MAX(0,B2)', cached: null } },
          ],
        },
      ],
    };
    const styles = buildStyles([]);
    const xml = buildWorksheetXml(sheet, styles, { tabSelected: false, drawingRelId: null });
    expect(xml).toContain('<c r="D2"><f>C2-B2</f><v>17.5</v></c>');
    expect(xml).toContain('<c r="E2"><f>MAX(0,C2)-MAX(0,B2)</f></c>');
    for (const bad of ['HYPERLINK("http://x")', 'C2-B2;1', 'cmd|x', 'SUM(A1:A9)', '=C2-B2', 'C2--B2', '']) {
      const evil: SheetSpec = {
        ...sheet,
        rows: [{ row0: 1, cells: [{ col0: 0, row0: 1, value: { type: 'formula', formula: bad, cached: 1 } }] }],
      };
      expect(() => buildWorksheetXml(evil, buildStyles([]), { tabSelected: false, drawingRelId: null })).toThrow(
        /formula/i,
      );
    }
  });

  it('the documentation matrix lists columnrange, arearange and errorbar as approximated', () => {
    for (const t of ['columnrange', 'arearange', 'errorbar']) {
      expect(CHART_TYPE_MATRIX.find((e) => e.highcharts === t)?.support, t).toBe('approximated');
    }
  });
});

// ---------------------------------------------------------------------------
// Writer specs: radar + errBars validate as XML on their own
// ---------------------------------------------------------------------------

describe('writer: radar and error bar specs', () => {
  it('clamps nothing and keeps the radar style', () => {
    const xml = buildChartXml(
      chartSpec(
        [
          {
            kind: 'radar',
            radarStyle: 'filled',
            varyColors: false,
            series: [serSpec(0)],
            axisIds: [100, 200],
            dataLabels: null,
          },
        ],
        [axis(100, 'cat', 200), axis(200, 'val', 100)],
      ),
    );
    expect(XMLValidator.validate(xml)).toBe(true);
    expect(xml).toContain('<c:radarChart><c:radarStyle val="filled"/><c:varyColors val="0"/>');
  });
});

// ---------------------------------------------------------------------------
// 7. CSS variable overrides
// ---------------------------------------------------------------------------

describe('7. CSS variable overrides', () => {
  it('the resolver consults the overrides first, headless too', () => {
    const resolve = createCssVariableResolver(undefined, { '--highcharts-color-0': '#8e44ad' });
    expect(resolve('--highcharts-color-0')).toBe('#8e44ad');
    expect(resolve('--highcharts-color-1')).toBeUndefined();
    const c = parseColor('var(--highcharts-color-0)', resolve);
    expect(c).toMatchObject({ r: 0x8e, g: 0x44, b: 0xad });
  });
});

describe('themeOverrides.cssVariables reaches the extractor through exportHighchartsToXlsx', () => {
  it('uses the override palette for a styled-mode chart exported headless', async () => {
    const { default: Highcharts } = await import('highcharts');
    await import('highcharts/modules/exporting');
    const { renderChart } = await import('../helpers/render-chart');
    const { styledMode } = await import('../fixtures/highcharts-options');
    const { exportHighchartsToXlsx } = await import('../../src/index');
    const { inspectXlsx } = await import('../helpers/inspect-xlsx');
    const chart = renderChart(Highcharts, styledMode);
    const result = await exportHighchartsToXlsx(chart, {
      themeOverrides: {
        cssVariables: {
          '--highcharts-color-0': '#8e44ad',
          '--highcharts-color-1': '#16a085',
          '--highcharts-color-2': '#d35400',
        },
      },
    });
    const insp = await inspectXlsx(result.bytes);
    expect(insp.text(insp.chartPaths()[0]!)).toContain('8E44AD');
    expect(result.warnings.map((w) => w.code)).not.toContain('STYLED_MODE_FALLBACK');
  });
});
