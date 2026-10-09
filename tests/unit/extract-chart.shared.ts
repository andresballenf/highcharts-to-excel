/**
 * Extractor assertions shared by the per-version test files. Each version file imports its own
 * Highcharts build (and modules) and calls `runExtractorSuite`.
 */

import { afterEach, describe, expect, it } from 'vitest';
import { extractChartModel, extractChartModelFromOptions, type ExtractOptions } from '../../src/highcharts';
import type { ChartModel } from '../../src/types/chart-model';
import { DiagnosticCollector } from '../../src/types/diagnostics';
import { colorToHex } from '../../src/utils/colors';
import * as F from '../fixtures/highcharts-options';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

type AnyChart = any; // eslint-disable-line @typescript-eslint/no-explicit-any

function opts(over: Partial<ExtractOptions> = {}): ExtractOptions {
  return { dataMode: 'rendered', seriesVisibility: 'visible', diagnostics: new DiagnosticCollector(), ...over };
}

const hex = (c: { r: number; g: number; b: number; a: number } | null | undefined): string | null => (c ? `#${colorToHex(c).toLowerCase()}` : null);
const codes = (m: ChartModel): string[] => m.warnings.map((w) => w.code);

export function runExtractorSuite(Highcharts: HighchartsLike, label: string): void {
  const H = Highcharts;
  const extract = (fixture: Parameters<typeof renderChart>[1], over: Partial<ExtractOptions> = {}): { chart: AnyChart; model: ChartModel } => {
    const chart = renderChart(H, fixture);
    return { chart, model: extractChartModel(chart, opts({ highchartsVersion: H.version ?? null, ...over })) };
  };

  describe(`extractChartModel (${label})`, () => {
    afterEach(() => destroyAll());

    it.each(Object.entries(F.allFixtures))('%s extracts and is JSON-serializable', (_name, fixture) => {
      const { model } = extract(fixture);
      expect(model.meta.sourceLibrary).toBe('highcharts');
      expect(model.meta.sourceVersion).toBe(H.version);
      expect(model.meta.extraction).toBe('headless');
      expect(JSON.parse(JSON.stringify(model))).toEqual(model);
    });

    it('rejects non-charts', () => {
      expect(() => extractChartModel({}, opts())).toThrow(/Highcharts chart/);
    });

    it('maps series kinds and counts', () => {
      const kinds = (f: Parameters<typeof renderChart>[1]): string[] => extract(f).model.series.map((s) => s.kind);
      expect(kinds(F.multiLine)).toEqual(['line', 'line', 'line']);
      expect(kinds(F.splineChart)).toEqual(['spline']);
      expect(kinds(F.columnChart)).toEqual(['column', 'column']);
      expect(kinds(F.barChart)).toEqual(['bar']);
      expect(kinds(F.areaChart)).toEqual(['area']);
      expect(kinds(F.comboChart)).toEqual(['column', 'line', 'spline']);
      expect(kinds(F.bubbleChart)).toEqual(['bubble', 'bubble']);
      expect(kinds(F.pieChart)).toEqual(['pie']);
      expect(kinds(F.doughnutChart)).toEqual(['doughnut']);
      const { model } = extract(F.unsupportedType);
      expect(model.series[0]!.kind).toBe('unknown');
      expect(model.series[0]!.sourceType).toBe('columnrange');
    });

    it('reads categories, points, nulls and negatives', () => {
      const { model } = extract(F.nullNegative);
      expect(model.xAxes[0]!.kind).toBe('category');
      expect(model.xAxes[0]!.categories).toEqual(['A', 'B', 'C', 'D', 'E']);
      const [line, col] = model.series;
      expect(line!.points.map((p) => p.y)).toEqual([1, null, 3, -4, 5]);
      expect(line!.points.map((p) => p.isNull)).toEqual([false, true, false, false, false]);
      expect(line!.points.map((p) => p.name)).toEqual(['A', 'B', 'C', 'D', 'E']);
      expect(line!.points.map((p) => p.x)).toEqual([0, 1, 2, 3, 4]);
      expect(col!.points.map((p) => p.y)).toEqual([-1, -2, 0, 2, null]);
    });

    it('keeps per-point and per-series colors', () => {
      const { model } = extract(F.customColors);
      expect(hex(model.series[0]!.color)).toBe('#aa3333');
      expect(hex(model.series[0]!.points[1]!.color)).toBe('#ff0000');
      expect(model.series[0]!.points[0]!.color).toBeNull();
      expect(hex(model.series[1]!.color)).toBe('#008000');
      expect(model.background?.type).toBe('gradient');
    });

    it('resolves the default palette (v13 CSS variables) to hex', () => {
      const { chart, model } = extract(F.simpleLine);
      expect(hex(model.series[0]!.color)).toBe('#2caffe');
      expect(hex(model.colors[0]!)).toBe('#2caffe');
      if (String(chart.series[0].color).startsWith('var(')) {
        expect(model.series[0]!.color?.source).toBe('var(--highcharts-color-0)');
      }
    });

    it('excludes hidden series in "visible" mode and includes them in "all" mode', () => {
      const visible = extract(F.hiddenSeries).model;
      expect(visible.series.map((s) => s.name)).toEqual(['Shown']);
      expect(codes(visible)).toContain('HIDDEN_SERIES_EXCLUDED');
      const all = extract(F.hiddenSeries, { seriesVisibility: 'all' }).model;
      expect(all.series.map((s) => [s.name, s.visible])).toEqual([
        ['Shown', true],
        ['Hidden', false],
      ]);
      expect(all.series[1]!.points.map((p) => p.y)).toEqual([3, 2, 1]);
      expect(codes(all)).toContain('HIDDEN_SERIES_INCLUDED');
    });

    it('reflects setData, addSeries and update in rendered mode; raw mode returns options.data', () => {
      const chart: AnyChart = renderChart(H, F.simpleLine);
      chart.series[0].setData([5, 6, 7]);
      chart.addSeries({ type: 'area', name: 'Added', data: [1, 2, 3] });
      chart.update({ title: { text: 'Updated' } });
      const model = extractChartModel(chart, opts());
      expect(model.title?.text).toBe('Updated');
      expect(model.series.map((s) => s.name)).toEqual(['Sales', 'Added']);
      expect(model.series[0]!.points.map((p) => p.y)).toEqual([5, 6, 7]);
      expect(model.series[1]!.kind).toBe('area');
      const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
      expect(raw.series[0]!.dataSemantics.mode).toBe('raw');
      expect(raw.series[0]!.points.map((p) => p.y)).toEqual([5, 6, 7]);
      expect(raw.series[1]!.points.map((p) => [p.x, p.y])).toEqual([
        [0, 1],
        [1, 2],
        [2, 3],
      ]);
    });

    it('raw mode parses the source data of every fixture shape', () => {
      const { model } = extract(F.pieChart, { dataMode: 'raw' });
      const pts = model.series[0]!.points;
      expect(pts.map((p) => p.name)).toEqual(['Chrome', 'Edge', 'Firefox', 'Safari']);
      expect(pts.map((p) => p.y)).toEqual([61.4, 11.8, 10.9, 4.6]);
      expect(pts[0]!.sliced).toBe(10);
      expect(hex(pts[1]!.color)).toBe('#ff0000');
      const scatter = extract(F.scatterChart, { dataMode: 'raw' }).model;
      expect(scatter.series[1]!.points.map((p) => p.x)).toEqual([174.0, 175.3, 193.5]);
    });

    it('falls back to rendered data when raw data is unavailable', () => {
      const chart: AnyChart = renderChart(H, F.simpleLine);
      // A stand-in that hides options.data (as with data-module/dataTable-fed series); the chart is untouched.
      const s = chart.series[0];
      const proxySeries = new Proxy(s, {
        get(target, key) {
          if (key === 'options') return { ...target.options, data: undefined };
          if (key === 'userOptions') return { ...target.userOptions, data: undefined };
          return Reflect.get(target, key) as unknown;
        },
      });
      const proxyChart = new Proxy(chart, { get: (t, k) => (k === 'series' ? [proxySeries] : (Reflect.get(t, k) as unknown)) });
      const model = extractChartModel(proxyChart, opts({ dataMode: 'raw' }));
      expect(model.series[0]!.dataSemantics.mode).toBe('rendered');
      expect(model.series[0]!.points.map((p) => p.y)).toEqual([1, 3, 2, 4, 6, 5]);
      expect(codes(model)).toContain('DATA_MODE_FALLBACK');
    });

    it('datetime x values are ms epoch and the explicit label format is kept', () => {
      const { model } = extract(F.datetimeChart);
      expect(model.xAxes[0]!.kind).toBe('datetime');
      expect(model.xAxes[0]!.dateFormat).toBe('%b %e');
      expect(model.xAxes[0]!.labels.format?.kind).toBe('excel');
      expect(model.series[0]!.points.map((p) => p.x)).toEqual([Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2), Date.UTC(2024, 0, 3), Date.UTC(2024, 0, 4)]);
      expect(model.series[0]!.points[0]!.name).toBeNull();
    });

    it('pie → doughnut when innerSize, slices carry colors and sliced offset', () => {
      const pie = extract(F.pieChart).model.series[0]!;
      expect(pie.kind).toBe('pie');
      expect(pie.pie).toEqual({ innerSize: 0, startAngle: 0, endAngle: null });
      expect(pie.points.map((p) => p.name)).toEqual(['Chrome', 'Edge', 'Firefox', 'Safari']);
      expect(pie.points.map((p) => hex(p.color))).toEqual(['#2caffe', '#ff0000', '#00e272', '#fe6a35']);
      expect(pie.points.map((p) => p.sliced)).toEqual([10, null, null, null]);
      expect(pie.dataLabels?.enabled).toBe(true);
      expect(pie.dataLabels?.showCategoryName).toBe(true);
      const doughnut = extract(F.doughnutChart).model.series[0]!;
      expect(doughnut.kind).toBe('doughnut');
      expect(doughnut.pie?.innerSize).toBeCloseTo(0.55, 2);
      expect(extract(F.pieChart).model.xAxes).toEqual([]);
    });

    it('scatter points keep their own x', () => {
      const { model } = extract(F.scatterChart);
      expect(model.series[0]!.points.map((p) => p.x)).toEqual([161.2, 167.5, 159.5, 157.0]);
      expect(model.series[1]!.points.map((p) => p.x)).toEqual([174.0, 175.3, 193.5]);
      expect(model.series[0]!.line?.width).toBe(0);
      expect(model.series[0]!.marker?.enabled).toBe(true);
      expect(model.xAxes[0]!.kind).toBe('linear');
    });

    it('extracts axis kinds, bounds, orientation and lines', () => {
      const { model } = extract(F.customAxes);
      const x = model.xAxes[0]!;
      const y = model.yAxes[0]!;
      expect(y.kind).toBe('linear');
      expect([y.min, y.max, y.tickInterval]).toEqual([0, 100, 25]);
      expect(y.reversed).toBe(true);
      expect(y.opposite).toBe(true);
      expect(y.gridLines).toMatchObject({ width: 2, dash: 'dash' });
      expect(hex(y.gridLines!.color)).toBe('#cccccc');
      expect(y.title?.text).toBe('Value');
      expect(y.title?.font.bold).toBe(true);
      expect(hex(y.title!.font.color)).toBe('#654321');
      expect(y.dataMin).toBe(10);
      expect(y.dataMax).toBe(90);
      expect(x.gridLines).toBeNull();
      expect(x.axisLine).toMatchObject({ width: 2 });
      expect(hex(x.axisLine!.color)).toBe('#0000ff');
      expect(x.title?.font.size).toBe(14);
      expect(x.reversed).toBe(false);
      // Defaults: y grid 1px #e6e6e6, y axis line hidden.
      const d = extract(F.simpleLine).model.yAxes[0]!;
      expect(d.gridLines?.width).toBe(1);
      expect(hex(d.gridLines!.color)).toBe('#e6e6e6');
      expect(d.axisLine?.width).toBe(0);
      expect(d.title?.text).toBe('Values');
    });

    it('binds series to secondary axes', () => {
      const { model } = extract(F.secondaryAxis);
      expect(model.yAxes.map((a) => a.opposite)).toEqual([false, true]);
      expect(model.series.map((s) => s.yAxisIndex)).toEqual([0, 1]);
      expect(model.yAxes[1]!.title?.text).toBe('Temperature');
    });

    it('maps legend layout and position; title fonts; sizes', () => {
      const { model } = extract(F.customStyling);
      expect(model.width).toBe(900);
      expect(model.height).toBe(500);
      expect(model.legend).toMatchObject({ enabled: true, position: 'right', layout: 'vertical', overlay: false });
      expect(model.title?.font).toMatchObject({ family: 'Georgia, serif', size: 24 });
      expect(hex(model.title!.font.color)).toBe('#112233');
      expect(model.plotArea.background).toMatchObject({ type: 'solid' });
      expect(model.plotArea.box).not.toBeNull();
      expect(model.series[0]!.dataLabels?.enabled).toBe(true);
      expect(model.series[0]!.dataLabels?.position).toBe('outsideEnd');
      const def = extract(F.simpleLine).model;
      expect(def.legend.position).toBe('bottom');
      expect(def.title?.font.size).toBeCloseTo(19.2, 3);
      expect(def.title?.font.bold).toBe(true);
      expect(def.legend.font?.size).toBeCloseTo(12.8, 3);
    });

    it('translates label formats', () => {
      const { model } = extract(F.percentChart);
      expect(model.yAxes[0]!.labels.format).toMatchObject({ kind: 'excel' });
      expect(model.series[0]!.dataLabels).toMatchObject({ enabled: true, showValue: true });
      expect(model.series[0]!.dataLabels?.format).toMatchObject({ kind: 'excel', source: '{point.y:.1f}%' });
    });

    it('stacking and stack groups', () => {
      const s = extract(F.stackedColumn).model.series;
      expect(s.map((x) => [x.stacking, x.stackGroup])).toEqual([
        ['normal', 'left'],
        ['normal', 'left'],
        ['normal', 'right'],
      ]);
      expect(extract(F.percentStackedColumn).model.series[0]!.stacking).toBe('percent');
      const area = extract(F.areaChart).model.series[0]!;
      expect(area.fillOpacity).toBe(0.75);
      expect(area.fill?.type).toBe('solid');
      expect(area.line?.width).toBe(2);
      const col = extract(F.columnChart).model.series[0]!;
      expect(col.bars).toMatchObject({ pointPadding: 0.1, groupPadding: 0.2 });
      expect(col.border?.width).toBe(1);
      expect(col.marker).toBeNull();
    });

    it('styledMode falls back to palette by colorIndex in jsdom', () => {
      const { model } = extract(F.styledMode);
      expect(model.meta.styledMode).toBe(true);
      expect(codes(model).filter((c) => c === 'STYLED_MODE_FALLBACK')).toHaveLength(1);
      expect(model.series.map((s) => hex(s.color))).toEqual(['#2caffe', '#544fc5']);
    });

    it('stock data grouping is reported and raw mode bypasses it', () => {
      if (!H.stockChart) throw new Error('stock module not loaded');
      const data: number[][] = [];
      for (let i = 0; i < 400; i++) data.push([Date.UTC(2024, 0, 1) + i * 86_400_000, i]);
      const fixture = { series: [{ type: 'line' as const, name: 'Daily', data, dataGrouping: { forced: true, enabled: true, units: [['month', [1]]] as Array<[string, number[]]> } }] };
      const chart = renderChart(H, fixture, 'stockChart');
      const model = extractChartModel(chart, opts());
      expect(model.series).toHaveLength(1); // navigator series excluded
      expect(model.xAxes).toHaveLength(1); // navigator axis excluded
      expect(model.yAxes).toHaveLength(1);
      const s = model.series[0]!;
      expect(s.dataSemantics).toMatchObject({ mode: 'rendered', grouped: true, sourcePointCount: 400, renderedPointCount: 14 });
      expect(s.points).toHaveLength(14);
      expect(codes(model)).toContain('DATA_GROUPED');
      const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
      expect(raw.series[0]!.points).toHaveLength(400);
      expect(codes(raw)).not.toContain('DATA_GROUPED');
    });

    it('extractChartModelFromOptions matches the rendered extraction', () => {
      for (const fixture of [F.simpleLine, F.columnChart, F.pieChart, F.multiLine, F.scatterChart]) {
        const rendered = extract(fixture).model;
        const fromOptions = extractChartModelFromOptions(fixture, opts());
        expect(fromOptions.meta.extraction).toBe('headless');
        expect(fromOptions.series.map((s) => [s.name, s.kind, s.points.length])).toEqual(rendered.series.map((s) => [s.name, s.kind, s.points.length]));
        expect(fromOptions.series.map((s) => s.points.map((p) => [p.x, p.y, p.name]))).toEqual(rendered.series.map((s) => s.points.map((p) => [p.x, p.y, p.name])));
        expect(fromOptions.series.map((s) => hex(s.color))).toEqual(rendered.series.map((s) => hex(s.color)));
        expect(fromOptions.title?.text).toBe(rendered.title?.text);
        expect(fromOptions.xAxes.map((a) => [a.kind, a.categories])).toEqual(rendered.xAxes.map((a) => [a.kind, a.categories]));
        expect(fromOptions.legend.enabled).toBe(rendered.legend.enabled);
        expect(fromOptions.plotArea.box).toBeNull();
      }
      const pie = extractChartModelFromOptions(F.pieChart, opts()).series[0]!;
      expect(pie.points.map((p) => hex(p.color))).toEqual(['#2caffe', '#ff0000', '#00e272', '#fe6a35']);
      const sized = extractChartModelFromOptions(F.customStyling, opts());
      expect([sized.width, sized.height]).toEqual([900, 500]);
      expect([extractChartModelFromOptions(F.simpleLine, opts()).width, extractChartModelFromOptions(F.simpleLine, opts()).height]).toEqual([600, 400]);
    });

    it('unsupported type, empty chart, polar and overrides', () => {
      const empty = extract(F.emptyChart).model;
      expect(empty.series).toEqual([]);
      const d = empty.warnings.find((w) => w.code === 'EMPTY_CHART');
      expect(d).toMatchObject({ outcome: 'blocking', severity: 'error' });
      expect(extractChartModelFromOptions(F.emptyChart, opts()).warnings.some((w) => w.code === 'EMPTY_CHART')).toBe(true);
      expect(extract(F.polarChart).model.polar).toBe(true);
      const sized = extract(F.simpleLine, { chartWidth: 1200, chartHeight: 700 }).model;
      expect([sized.width, sized.height]).toEqual([1200, 700]);
      expect(extract(F.barChart).model.inverted).toBe(true);
    });

    it('large charts stay in rendered form without grouping', () => {
      const { model } = extract(F.largeChart);
      expect(model.series.map((s) => s.points.length)).toEqual([1000, 1000, 1000]);
      expect(model.series[0]!.marker?.enabled).toBe(false);
      expect(extract(F.simpleLine).model.series[0]!.marker?.enabled).toBe(true);
    });

    it('reports annotations, plot bands, tooltip formatters and 3D without executing callbacks', () => {
      let called = false;
      const fixture = {
        ...F.simpleLine,
        chart: { options3d: { enabled: true } },
        tooltip: {
          formatter(): string {
            called = true;
            return '';
          },
        },
        yAxis: { plotBands: [{ from: 1, to: 2 }] },
        annotations: [{ labels: [{ point: { x: 0, y: 0 }, text: 'x' }] }],
      };
      const { model } = extract(fixture as unknown as Parameters<typeof renderChart>[1]);
      expect(called).toBe(false);
      const props = model.warnings.map((w) => `${w.code}@${w.property}`);
      expect(props).toEqual(
        expect.arrayContaining(['UNSUPPORTED_TOOLTIP@tooltip.formatter', 'UNSUPPORTED_PLOT_BAND@yAxis[0].plotBands', 'UNSUPPORTED_ANNOTATION@annotations', 'UNSUPPORTED_3D@chart.options3d']),
      );
    });
  });
}
