import { describe, expect, it } from 'vitest';
import { CHART_TYPE_MATRIX, resolveChartType } from '../../src/core/chart-type-registry';
import * as F from '../fixtures/chart-models';

const kinds = (m: ReturnType<typeof F.lineModel>) => resolveChartType(m).groups.map((g) => g.kind);

describe('resolveChartType: single types', () => {
  it.each([
    ['lineModel', 'line', ['line']],
    ['multiSeriesLineModel', 'line', ['line']],
    ['columnStackedModel', 'stackedColumn', ['bar']],
    ['barModel', 'bar', ['bar']],
    ['areaPercentModel', 'percentStackedArea', ['area']],
    ['pieModel', 'pie', ['pie']],
    ['doughnutModel', 'doughnut', ['doughnut']],
    ['scatterModel', 'scatter', ['scatter']],
    ['bubbleModel', 'bubble', ['bubble']],
    ['datetimeModel', 'line', ['line']],
    ['comboModel', 'combo:column+line', ['bar', 'line']],
    ['secondaryAxisModel', 'combo:column+line', ['bar', 'line']],
    ['nullNegativeModel', 'column', ['bar']],
  ] as const)('%s → %s', (name, type, groupKinds) => {
    // biome-ignore lint/performance/noDynamicNamespaceImportAccess: the table names fixtures by key
    const r = resolveChartType(F[name]());
    expect(r.blocking).toBe(false);
    expect(r.excelChartType).toBe(type);
    expect(r.groups.map((g) => g.kind)).toEqual(groupKinds);
  });

  it('column is vertical, inverted column is horizontal, bar is always horizontal', () => {
    const col = resolveChartType(F.columnStackedModel()).groups[0]!;
    expect(col).toMatchObject({ kind: 'bar', barDir: 'col', stacking: 'normal', seriesIndices: [0, 1] });
    expect(resolveChartType(F.barModel()).groups[0]!.barDir).toBe('bar');
    const bar = F.lineModel();
    bar.series = [F.series({ kind: 'bar', index: 0, points: F.pts([1]) })];
    expect(resolveChartType(bar).groups[0]!.barDir).toBe('bar');
    expect(resolveChartType({ ...bar, inverted: true }).groups[0]!.barDir).toBe('bar');
  });

  it('marks spline groups smooth; mixed line/spline is not smooth at group level', () => {
    const m = F.lineModel();
    m.series = [F.series({ kind: 'spline', index: 0, smooth: true, points: F.pts([1, 2]) })];
    expect(resolveChartType(m).groups[0]!.smooth).toBe(true);
    expect(resolveChartType(F.multiSeriesLineModel()).groups[0]!.smooth).toBe(false);
  });

  it('names stacked types', () => {
    const m = F.columnStackedModel();
    m.series = m.series.map((s) => ({ ...s, stacking: 'percent' as const }));
    expect(resolveChartType(m).excelChartType).toBe('percentStackedColumn');
    const b = F.barModel();
    b.series = b.series.map((s) => ({ ...s, stacking: 'normal' as const }));
    expect(resolveChartType(b).excelChartType).toBe('stackedBar');
    const a = F.areaPercentModel();
    a.series = a.series.map((s) => ({ ...s, stacking: 'normal' as const }));
    expect(resolveChartType(a).excelChartType).toBe('stackedArea');
  });
});

describe('resolveChartType: grouping and combos', () => {
  it('groups by kind and axis, preserving first appearance', () => {
    const r = resolveChartType(F.secondaryAxisModel());
    expect(r.groups.map((g) => [g.kind, g.yAxisIndex, g.seriesIndices])).toEqual([
      ['bar', 0, [0]],
      ['line', 1, [1]],
    ]);
    expect(r.diagnostics.map((d) => d.code)).toContain('MIXED_SERIES_TYPES');
  });

  it('collapses a third y axis onto the secondary axis', () => {
    const m = F.secondaryAxisModel();
    m.yAxes.push(F.axis({ index: 2, kind: 'linear' }));
    m.series.push(F.series({ kind: 'line', index: 2, yAxisIndex: 2, points: F.pts([1, 2, 3, 4]) }));
    const r = resolveChartType(m);
    expect(r.groups.map((g) => [g.kind, g.yAxisIndex, g.seriesIndices])).toEqual([
      ['bar', 0, [0]],
      ['line', 1, [1, 2]],
    ]);
    const d = r.diagnostics.find((x) => x.code === 'UNSUPPORTED_AXIS_FEATURE');
    expect(d).toMatchObject({ outcome: 'approximated', property: 'yAxis[2]' });
  });

  it('merges different stack groups with an approximation diagnostic', () => {
    const m = F.columnStackedModel();
    m.series[1] = { ...m.series[1]!, stackGroup: 'other' };
    const r = resolveChartType(m);
    expect(r.groups).toHaveLength(1);
    expect(r.diagnostics.find((d) => d.code === 'APPROXIMATED_CHART_TYPE')).toMatchObject({
      property: 'series[1].stack',
      seriesIndex: 1,
    });
  });

  it('allows scatter in a combo with category types', () => {
    const m = F.comboModel();
    m.series.push(F.series({ kind: 'scatter', index: 2, points: [F.point({ x: 1, y: 1 })] }));
    const r = resolveChartType(m);
    expect(r.blocking).toBe(false);
    expect(r.excelChartType).toBe('combo:column+line+scatter');
  });

  it('pie cannot combine with other kinds (blocking)', () => {
    const m = F.comboModel();
    m.series.push(F.series({ kind: 'pie', index: 2, points: [F.point({ name: 'a', y: 1 })] }));
    const r = resolveChartType(m);
    expect(r.blocking).toBe(true);
    expect(r.excelChartType).toBeNull();
    expect(r.diagnostics.find((d) => d.code === 'UNSUPPORTED_CHART_TYPE')).toMatchObject({
      outcome: 'blocking',
      message: 'Excel cannot combine pie with other chart types.',
    });
  });

  it('bubble cannot combine with other kinds (blocking)', () => {
    const m = F.bubbleModel();
    m.series.push(F.series({ kind: 'line', index: 1, points: F.pts([1]) }));
    expect(resolveChartType(m).blocking).toBe(true);
  });

  it('keeps only the first pie series', () => {
    const m = F.pieModel();
    m.series.push({ ...m.series[0]!, index: 1, id: 's1', name: 'Second' });
    const r = resolveChartType(m);
    expect(r.groups[0]!.seriesIndices).toEqual([0]);
    expect(r.droppedSeries).toEqual([1]);
    expect(r.diagnostics.find((d) => d.code === 'UNSUPPORTED_SERIES_TYPE')).toMatchObject({
      outcome: 'approximated',
      property: 'series[1].type',
      seriesIndex: 1,
    });
  });

  it('pie + doughnut become doughnut rings', () => {
    const m = F.doughnutModel();
    m.series[0] = { ...m.series[0]!, kind: 'pie' };
    const r = resolveChartType(m);
    expect(r.excelChartType).toBe('doughnut');
    expect(r.groups[0]!.seriesIndices).toEqual([0, 1]);
  });

  it('reports inverted line charts', () => {
    const r = resolveChartType({ ...F.lineModel(), inverted: true });
    expect(r.blocking).toBe(false);
    expect(r.diagnostics.find((d) => d.code === 'APPROXIMATED_CHART_TYPE')).toMatchObject({
      outcome: 'approximated',
      property: 'chart.inverted',
      message: expect.stringContaining('Excel cannot invert line/area charts'),
    });
  });
});

describe('resolveChartType: unsupported and blocking', () => {
  it('drops unknown series and keeps the rest', () => {
    const m = F.lineModel();
    m.series.push(F.series({ kind: 'unknown', sourceType: 'heatmap', index: 1, points: F.pts([1]) }));
    const r = resolveChartType(m);
    expect(r.blocking).toBe(false);
    expect(r.droppedSeries).toEqual([1]);
    expect(r.diagnostics[0]).toMatchObject({
      code: 'UNSUPPORTED_SERIES_TYPE',
      outcome: 'unsupported',
      property: 'series[1].type',
      seriesIndex: 1,
      message: expect.stringContaining('heatmap'),
    });
  });

  it('unknown-only charts are blocking UNSUPPORTED_CHART_TYPE naming the source types', () => {
    const r = resolveChartType(F.unknownTypeModel());
    expect(r.blocking).toBe(true);
    const codes = r.diagnostics.map((d) => d.code);
    expect(codes).toEqual(['UNSUPPORTED_SERIES_TYPE', 'UNSUPPORTED_CHART_TYPE']);
    expect(r.diagnostics[1]!.message).toContain('treemap');
    expect(r.diagnostics[1]!.outcome).toBe('blocking');
  });

  it('polar is blocking', () => {
    const r = resolveChartType(F.polarModel());
    expect(r.blocking).toBe(true);
    expect(r.diagnostics).toEqual([
      expect.objectContaining({ code: 'UNSUPPORTED_POLAR', outcome: 'blocking', property: 'chart.polar' }),
    ]);
  });

  it('empty charts are blocking', () => {
    const r = resolveChartType(F.emptyModel());
    expect(r.blocking).toBe(true);
    expect(kinds(F.emptyModel())).toEqual([]);
    expect(r.diagnostics[0]).toMatchObject({ code: 'EMPTY_CHART', outcome: 'blocking' });
  });

  it('does not mutate the model', () => {
    const m = F.doughnutModel();
    m.series[0] = { ...m.series[0]!, kind: 'pie' };
    const before = JSON.stringify(m);
    resolveChartType(m);
    expect(JSON.stringify(m)).toBe(before);
  });
});

describe('CHART_TYPE_MATRIX', () => {
  it('documents every listed Highcharts type with a reason', () => {
    const names = CHART_TYPE_MATRIX.map((e) => e.highcharts.split(' ')[0]);
    for (const t of [
      'line',
      'spline',
      'area',
      'areaspline',
      'column',
      'bar',
      'pie',
      'scatter',
      'bubble',
      'columnrange',
      'arearange',
      'boxplot',
      'heatmap',
      'treemap',
      'waterfall',
      'funnel',
      'gauge',
      'polar',
      'variablepie',
      'sankey',
      'networkgraph',
      'timeline',
      'histogram',
      'bellcurve',
      'errorbar',
      'lollipop',
      'dumbbell',
    ]) {
      expect(names).toContain(t);
    }
    expect(CHART_TYPE_MATRIX.find((e) => e.highcharts === 'pie + innerSize')?.excel).toBe('doughnut');
    for (const e of CHART_TYPE_MATRIX) expect(e.notes.length).toBeGreaterThan(10);
    expect(CHART_TYPE_MATRIX.filter((e) => e.support === 'unsupported').length).toBeGreaterThanOrEqual(18);
  });
});

describe('resolveChartType: mixed stacking (C4)', () => {
  it('merges stacked + unstacked columns on the same axes into one group (majority stacking, ties → stacked)', () => {
    const m = F.baseModel();
    m.series = [
      F.series({ kind: 'column', index: 0, stacking: 'normal', points: F.pts([1, 2, 3, 4]) }),
      F.series({ kind: 'column', index: 1, stacking: null, points: F.pts([1, 2, 3, 4]) }),
    ];
    const r = resolveChartType(m);
    expect(r.groups).toHaveLength(1);
    expect(r.groups[0]).toMatchObject({ kind: 'bar', stacking: 'normal', seriesIndices: [0, 1] });
    expect(r.excelChartType).toBe('stackedColumn');
    expect(r.diagnostics.find((d) => d.code === 'APPROXIMATED_CHART_TYPE')).toMatchObject({
      outcome: 'approximated',
      severity: 'warning',
      property: 'series[1].stacking',
      seriesIndex: 1,
    });
  });

  it('majority unstacked wins', () => {
    const m = F.baseModel();
    m.series = [
      F.series({ kind: 'column', index: 0, stacking: 'normal', points: F.pts([1, 2, 3, 4]) }),
      F.series({ kind: 'column', index: 1, points: F.pts([1, 2, 3, 4]) }),
      F.series({ kind: 'column', index: 2, points: F.pts([1, 2, 3, 4]) }),
    ];
    const r = resolveChartType(m);
    expect(r.groups.map((g) => [g.kind, g.stacking, g.seriesIndices])).toEqual([['bar', null, [0, 1, 2]]]);
    expect(r.diagnostics.filter((d) => d.code === 'APPROXIMATED_CHART_TYPE').map((d) => d.property)).toEqual([
      'series[0].stacking',
    ]);
  });
});
