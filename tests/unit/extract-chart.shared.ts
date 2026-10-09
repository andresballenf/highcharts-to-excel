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

// biome-ignore lint/suspicious/noExplicitAny: tests read live Highcharts internals that the public types omit
type AnyChart = any;

function opts(over: Partial<ExtractOptions> = {}): ExtractOptions {
  return { dataMode: 'rendered', seriesVisibility: 'visible', diagnostics: new DiagnosticCollector(), ...over };
}

const hex = (c: { r: number; g: number; b: number; a: number } | null | undefined): string | null =>
  c ? `#${colorToHex(c).toLowerCase()}` : null;
const codes = (m: ChartModel): string[] => m.warnings.map((w) => w.code);

export function runExtractorSuite(Highcharts: HighchartsLike, label: string): void {
  const H = Highcharts;
  const extract = (
    fixture: Parameters<typeof renderChart>[1],
    over: Partial<ExtractOptions> = {},
  ): { chart: AnyChart; model: ChartModel } => {
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
      expect(model.series[0]!.sourceType).toBe('boxplot');
      expect(kinds(F.columnRangeChart)).toEqual(['columnrange']);
      expect(kinds(F.errorBarChart)).toEqual(['column', 'errorbar']);
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
      const proxyChart = new Proxy(chart, {
        get: (t, k) => (k === 'series' ? [proxySeries] : (Reflect.get(t, k) as unknown)),
      });
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
      expect(model.series[0]!.points.map((p) => p.x)).toEqual([
        Date.UTC(2024, 0, 1),
        Date.UTC(2024, 0, 2),
        Date.UTC(2024, 0, 3),
        Date.UTC(2024, 0, 4),
      ]);
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
      const fixture = {
        series: [
          {
            type: 'line' as const,
            name: 'Daily',
            data,
            dataGrouping: { forced: true, enabled: true, units: [['month', [1]]] as Array<[string, number[]]> },
          },
        ],
      };
      const chart = renderChart(H, fixture, 'stockChart');
      const model = extractChartModel(chart, opts());
      expect(model.series).toHaveLength(1); // navigator series excluded
      expect(model.xAxes).toHaveLength(1); // navigator axis excluded
      expect(model.yAxes).toHaveLength(1);
      const s = model.series[0]!;
      expect(s.dataSemantics).toMatchObject({
        mode: 'rendered',
        grouped: true,
        sourcePointCount: 400,
        renderedPointCount: 14,
      });
      expect(s.points).toHaveLength(14);
      expect(codes(model)).toContain('DATA_GROUPED');
      const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
      expect(raw.series[0]!.points).toHaveLength(400);
      expect(codes(raw)).not.toContain('DATA_GROUPED');
    });

    it('a stock chart zoomed below cropThreshold reports DATA_CROPPED and exports only the visible points', () => {
      if (!H.stockChart) throw new Error('stock module not loaded');
      const day = 86_400_000;
      const start = Date.UTC(2024, 0, 1);
      const data: number[][] = [];
      for (let i = 0; i < 200; i++) data.push([start + i * day, i]);
      const fixture = {
        navigator: { enabled: false },
        scrollbar: { enabled: false },
        rangeSelector: { enabled: false },
        series: [{ type: 'line' as const, name: 'Daily', data, cropThreshold: 10, dataGrouping: { enabled: false } }],
      };
      const chart = renderChart(H, fixture, 'stockChart') as AnyChart;
      chart.xAxis[0].setExtremes(start + 50 * day, start + 59 * day);
      const model = extractChartModel(chart, opts());
      const s = model.series[0]!;
      expect(s.dataSemantics.cropped).toBe(true);
      expect(s.dataSemantics.sourcePointCount).toBe(200);
      expect(s.points.length).toBeLessThan(200);
      expect(model.warnings.find((w) => w.code === 'DATA_CROPPED')).toMatchObject({
        code: 'DATA_CROPPED',
        property: 'series[0].data',
        seriesIndex: 0,
      });
      const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
      expect(raw.series[0]!.points).toHaveLength(200);
      expect(codes(raw)).not.toContain('DATA_CROPPED');
    });

    it('reports EMPTY_SERIES for a series without points', () => {
      const { model } = extract({
        title: { text: 'Empty series' },
        series: [
          { type: 'line', name: 'Filled', data: [1, 2, 3] },
          { type: 'line', name: 'Nothing', data: [] },
        ],
      });
      expect(model.series.map((s) => s.points.length)).toEqual([3, 0]);
      const empty = model.warnings.filter((w) => w.code === 'EMPTY_SERIES');
      expect(empty).toHaveLength(1);
      expect(empty[0]).toMatchObject({ code: 'EMPTY_SERIES', severity: 'info', seriesIndex: 1 });
      expect(codes(model)).not.toContain('EMPTY_CHART');
    });

    it('raw mode skips a string data value with NON_NUMERIC_VALUE (nothing is fabricated)', () => {
      const fixture = {
        series: [{ type: 'line' as const, name: 'Mixed', data: [1, 'oops', 3] as unknown as number[] }],
      };
      const { model } = extract(fixture, { dataMode: 'raw' });
      expect(model.series[0]!.points.map((p) => p.y)).toEqual([1, 3]);
      const d = model.warnings.find((w) => w.code === 'NON_NUMERIC_VALUE');
      expect(d).toMatchObject({
        code: 'NON_NUMERIC_VALUE',
        outcome: 'approximated',
        property: 'series[0].data[1]',
        seriesIndex: 0,
      });
      expect(d?.details).toMatchObject({ valueType: 'string', count: 1, firstIndices: [1] });
    });

    it('extractChartModelFromOptions matches the rendered extraction', () => {
      for (const fixture of [F.simpleLine, F.columnChart, F.pieChart, F.multiLine, F.scatterChart]) {
        const rendered = extract(fixture).model;
        const fromOptions = extractChartModelFromOptions(fixture, opts());
        expect(fromOptions.meta.extraction).toBe('headless');
        expect(fromOptions.series.map((s) => [s.name, s.kind, s.points.length])).toEqual(
          rendered.series.map((s) => [s.name, s.kind, s.points.length]),
        );
        expect(fromOptions.series.map((s) => s.points.map((p) => [p.x, p.y, p.name]))).toEqual(
          rendered.series.map((s) => s.points.map((p) => [p.x, p.y, p.name])),
        );
        expect(fromOptions.series.map((s) => hex(s.color))).toEqual(rendered.series.map((s) => hex(s.color)));
        expect(fromOptions.title?.text).toBe(rendered.title?.text);
        expect(fromOptions.xAxes.map((a) => [a.kind, a.categories])).toEqual(
          rendered.xAxes.map((a) => [a.kind, a.categories]),
        );
        expect(fromOptions.legend.enabled).toBe(rendered.legend.enabled);
        expect(fromOptions.plotArea.box).toBeNull();
      }
      const pie = extractChartModelFromOptions(F.pieChart, opts()).series[0]!;
      expect(pie.points.map((p) => hex(p.color))).toEqual(['#2caffe', '#ff0000', '#00e272', '#fe6a35']);
      const sized = extractChartModelFromOptions(F.customStyling, opts());
      expect([sized.width, sized.height]).toEqual([900, 500]);
      expect([
        extractChartModelFromOptions(F.simpleLine, opts()).width,
        extractChartModelFromOptions(F.simpleLine, opts()).height,
      ]).toEqual([600, 400]);
    });

    it('unsupported type, empty chart, polar and overrides', () => {
      const empty = extract(F.emptyChart).model;
      expect(empty.series).toEqual([]);
      const d = empty.warnings.find((w) => w.code === 'EMPTY_CHART');
      expect(d).toMatchObject({ outcome: 'blocking', severity: 'error' });
      expect(extractChartModelFromOptions(F.emptyChart, opts()).warnings.some((w) => w.code === 'EMPTY_CHART')).toBe(
        true,
      );
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
        expect.arrayContaining([
          'UNSUPPORTED_TOOLTIP@tooltip.formatter',
          'UNSUPPORTED_PLOT_BAND@yAxis[0].plotBands',
          'UNSUPPORTED_ANNOTATION@annotations',
          'UNSUPPORTED_3D@chart.options3d',
        ]),
      );
    });
  });

  describe(`extractor audit fixes (${label})`, () => {
    afterEach(() => destroyAll());
    const major = Number.parseInt(String(H.version ?? '0'), 10);
    const day = 86_400_000;
    const xs = (m: ChartModel, i = 0): Array<number | null> => m.series[i]!.points.map((p) => p.x);

    it('E3 shifts datetime x values to the wall-clock time the chart shows (live chart)', () => {
      const tokyoMidnight = Date.UTC(2024, 0, 1) - 9 * 3_600_000;
      const fixture = {
        time: { timezone: 'Asia/Tokyo' },
        xAxis: { type: 'datetime' as const },
        series: [
          {
            type: 'line' as const,
            data: [
              [tokyoMidnight, 1],
              [tokyoMidnight + day, 2],
            ],
          },
        ],
      };
      const { chart, model } = extract(fixture);
      // getTimezoneOffset is in ms, positive west (as Date#getTimezoneOffset): Tokyo → -9 h.
      expect(chart.time.getTimezoneOffset(tokyoMidnight)).toBe(-9 * 3_600_000);
      expect(chart.time.dateFormat('%Y-%m-%d %H:%M', chart.series[0].points[0].x)).toBe('2024-01-01 00:00');
      expect(xs(model)).toEqual([Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2)]);
      expect(model.xAxes[0]!.dataMin).toBe(Date.UTC(2024, 0, 1));
      expect(model.xAxes[0]!.dataMax).toBe(Date.UTC(2024, 0, 2));
      expect(model.meta.datetimeOffsetMinutes).toBe(540);
      const tz = model.warnings.filter((w) => w.code === 'APPROXIMATED_DATETIME');
      expect(tz).toHaveLength(1);
      expect(tz[0]).toMatchObject({ property: 'time.timezone', outcome: 'approximated', severity: 'info' });
      const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
      expect(xs(raw)).toEqual([Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2)]);
      // UTC charts are untouched.
      const utc = extract(F.datetimeChart).model;
      expect(utc.meta.datetimeOffsetMinutes).toBe(0);
      expect(codes(utc)).not.toContain('APPROXIMATED_DATETIME');
    });

    it('E3 options path applies time.timezone / timezoneOffset through Intl', () => {
      const tokyoMidnight = Date.UTC(2024, 0, 1) - 9 * 3_600_000;
      const base = {
        xAxis: { type: 'datetime' },
        series: [
          {
            data: [
              [tokyoMidnight, 1],
              [tokyoMidnight + day, 2],
            ],
          },
        ],
      };
      const m = extractChartModelFromOptions({ ...base, time: { timezone: 'Asia/Tokyo' } }, opts());
      expect(xs(m)).toEqual([Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2)]);
      expect(m.meta.datetimeOffsetMinutes).toBe(540);
      expect(m.warnings.filter((w) => w.code === 'APPROXIMATED_DATETIME')).toHaveLength(1);
      const fixed = extractChartModelFromOptions({ ...base, time: { timezoneOffset: -540 } }, opts());
      expect(xs(fixed)).toEqual([Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2)]);
      // New York in summer: UTC-4.
      const ny = extractChartModelFromOptions(
        {
          xAxis: { type: 'datetime' },
          time: { timezone: 'America/New_York' },
          series: [{ data: [[Date.UTC(2024, 6, 1, 16), 1]] }],
        },
        opts(),
      );
      expect(xs(ny)).toEqual([Date.UTC(2024, 6, 1, 12)]);
      expect(ny.meta.datetimeOffsetMinutes).toBe(-240);
      const plain = extractChartModelFromOptions(base, opts());
      expect(xs(plain)).toEqual([tokyoMidnight, tokyoMidnight + day]);
      expect(plain.meta.datetimeOffsetMinutes).toBe(0);
      const bad = extractChartModelFromOptions({ ...base, time: { timezone: 'Not/AZone' } }, opts());
      expect(bad.meta.datetimeOffsetMinutes).toBeNull();
    });

    it('E4 parses string x values and string pointStart on datetime axes', () => {
      const fixture = {
        xAxis: { type: 'datetime' as const },
        series: [
          {
            type: 'line' as const,
            data: [
              ['2024-01-01', 1],
              ['2024-01-02', 2],
            ],
          },
          { type: 'line' as const, pointStart: '2024-01-01', pointIntervalUnit: 'day', data: [3, 4] },
          {
            type: 'line' as const,
            data: [
              { x: '2024-01-03', y: 5 },
              { x: 'not a date', y: 6 },
            ],
          },
        ],
      } as unknown as Parameters<typeof renderChart>[1];
      const expected = [Date.UTC(2024, 0, 1), Date.UTC(2024, 0, 2)];
      const m = extractChartModelFromOptions(fixture, opts());
      expect(xs(m, 0)).toEqual(expected);
      expect(m.series[0]!.points.map((p) => p.name)).toEqual([null, null]);
      expect(xs(m, 1)).toEqual(expected);
      expect(xs(m, 2)).toEqual([Date.UTC(2024, 0, 3)]);
      expect(m.warnings.find((w) => w.code === 'NON_NUMERIC_VALUE')).toMatchObject({ seriesIndex: 2 });
      if (major >= 12) {
        // Highcharts 12+ parses date strings itself: raw mode must agree with what is rendered.
        const { chart } = extract(fixture);
        const rendered = extractChartModel(chart, opts());
        const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
        expect(xs(rendered, 0)).toEqual(expected);
        expect(xs(raw, 0)).toEqual(xs(rendered, 0));
        expect(xs(raw, 1)).toEqual(xs(rendered, 1));
      }
    });

    it('E5 honours series.keys and relativeXValue in raw data', () => {
      const fixture = {
        series: [
          {
            type: 'line' as const,
            keys: ['y', 'color'],
            data: [
              [5, '#ff0000'],
              [6, '#00ff00'],
            ],
          },
          {
            type: 'line' as const,
            keys: ['name', 'x', 'y'],
            data: [
              ['a', 10, 1],
              ['b', 20, 2],
            ],
          },
          {
            type: 'line' as const,
            pointStart: 100,
            relativeXValue: true,
            data: [
              [0, 1],
              [2, 3],
            ],
          },
        ],
      } as unknown as Parameters<typeof renderChart>[1];
      const pts = (m: ChartModel, i: number) => m.series[i]!.points.map((p) => [p.x, p.name, p.y]);
      const m = extractChartModelFromOptions(fixture, opts());
      expect(pts(m, 0)).toEqual([
        [0, null, 5],
        [1, null, 6],
      ]);
      expect(hex(m.series[0]!.points[0]!.color)).toBe('#ff0000');
      expect(pts(m, 1)).toEqual([
        [10, 'a', 1],
        [20, 'b', 2],
      ]);
      expect(pts(m, 2)).toEqual([
        [100, null, 1],
        [102, null, 3],
      ]);
      expect(codes(m)).not.toContain('NON_NUMERIC_VALUE');
      const { chart, model: rendered } = extract(fixture);
      const raw = extractChartModel(chart, opts({ dataMode: 'raw' }));
      for (const i of [0, 1, 2]) expect(pts(raw, i)).toEqual(pts(rendered, i));
    });

    it('E6 keeps literal < and > in titles and maps <br> to a line break', () => {
      const fixture = {
        title: { text: 'Growth < 5% vs > 3% target' },
        subtitle: { text: 'Line1<br>Line2<br/>Line3' },
        xAxis: { categories: ['a<b>c</b>', 'R&amp;D'] },
        series: [{ type: 'line' as const, data: [1, 2] }],
      };
      for (const m of [extract(fixture).model, extractChartModelFromOptions(fixture, opts())]) {
        expect(m.title?.text).toBe('Growth < 5% vs > 3% target');
        expect(m.subtitle?.text).toBe('Line1\nLine2\nLine3');
        expect(m.xAxes[0]!.categories).toEqual(['ac', 'R&D']);
      }
    });

    it('E7 legend visibility counts exported series only and showInLegend is carried', () => {
      const fixture = {
        series: [
          { type: 'line' as const, data: [1, 2], showInLegend: false },
          { type: 'line' as const, data: [2, 3], visible: false },
        ],
      };
      const { model } = extract(fixture);
      expect(model.series).toHaveLength(1);
      expect(model.series[0]!.showInLegend).toBe(false);
      expect(model.legend.enabled).toBe(false);
      const all = extract(fixture, { seriesVisibility: 'all' }).model;
      expect(all.series.map((s) => s.showInLegend)).toEqual([false, true]);
      expect(all.legend.enabled).toBe(true);
      expect(extractChartModelFromOptions(fixture, opts()).legend.enabled).toBe(false);
      const pie = extract(F.pieChart).model;
      expect(pie.series[0]!.showInLegend).toBe(false);
      expect(extract(F.simpleLine).model.series[0]!.showInLegend).toBe(true);
    });

    it('E9 accepts typed arrays as y-only data without a runtime chart', () => {
      const m = extractChartModelFromOptions({ series: [{ data: new Float64Array([1, 2, 3]) }] }, opts());
      expect(m.series[0]!.points.map((p) => [p.x, p.y])).toEqual([
        [0, 1],
        [1, 2],
        [2, 3],
      ]);
      expect(codes(m)).not.toContain('DATA_MODE_FALLBACK');
      expect(codes(m)).not.toContain('EMPTY_CHART');
      const odd = extractChartModelFromOptions({ series: [{ data: [1, 2] }, { data: 'nope' }] }, opts());
      expect(odd.warnings.some((w) => /rendered points/.test(w.message))).toBe(false);
      expect(odd.series[1]!.dataSemantics.mode).toBe('raw');
    });

    it('E10 aggregates per-point diagnostics (one per series)', () => {
      const data = Array.from({ length: 20_000 }, (_, i) => (i % 2 ? 'x' : i));
      const m = extractChartModelFromOptions({ series: [{ data }] }, opts());
      expect(m.series[0]!.points).toHaveLength(10_000);
      expect(m.warnings.length).toBeLessThanOrEqual(2);
      const d = m.warnings.find((w) => w.code === 'NON_NUMERIC_VALUE');
      expect(d?.details).toMatchObject({ count: 10_000, firstIndices: [1, 3, 5, 7, 9] });
      const labelled = extractChartModelFromOptions(
        {
          series: [
            {
              data: Array.from({ length: 50 }, (_, i) => ({ y: i, dataLabels: { formatter: () => String(i) } })),
            },
          ],
        },
        opts(),
      );
      const labelCodes = labelled.warnings.filter((w) => w.property.includes('dataLabels'));
      expect(labelCodes.length).toBeGreaterThan(0);
      expect(new Set(labelCodes.map((w) => w.code)).size).toBe(labelCodes.length);
      expect(labelCodes[0]!.details).toMatchObject({ count: 50, firstIndices: [0, 1, 2, 3, 4] });
    });

    it('E11 a live series exported from its source data reports raw mode', () => {
      const chart: AnyChart = renderChart(H, { series: [{ type: 'line' as const, data: [1, 2, 3] }] });
      // A series with no generated points (never rendered): only options.data is left.
      const proxySeries = new Proxy(chart.series[0], {
        get: (target, key) => (key === 'points' || key === 'data' ? [] : (Reflect.get(target, key) as unknown)),
      });
      const proxyChart = new Proxy(chart, {
        get: (t, k) => (k === 'series' ? [proxySeries] : (Reflect.get(t, k) as unknown)),
      });
      const model = extractChartModel(proxyChart, opts());
      const s = model.series[0]!;
      expect(s.points.map((p) => p.y)).toEqual([1, 2, 3]);
      expect(s.dataSemantics.mode).toBe('raw');
      expect(codes(model)).toContain('DATA_MODE_FALLBACK');
    });

    // Last on purpose: the boost module stays loaded for the rest of the file. Highcharts 13 only
    // (the module attaches to the core loaded in this file; one version per file).
    it.runIf(major === 13)('E1 boosted series export their values, not the pixel pseudo-points', async () => {
      const proto = HTMLCanvasElement.prototype as unknown as { getContext: unknown };
      const original = proto.getContext;
      const fakeGL = (): unknown =>
        new Proxy(
          {},
          {
            set: () => true,
            get: (_t, k) =>
              k === 'canvas'
                ? document.createElement('canvas')
                : typeof k === 'string' && /^[A-Z_0-9]+$/.test(k)
                  ? 1
                  : () => ({}),
          },
        );
      proto.getContext = () => fakeGL();
      (window as unknown as { WebGLRenderingContext: unknown }).WebGLRenderingContext =
        function WebGLRenderingContext() {};
      try {
        await import('highcharts/modules/boost');
        const data = Array.from({ length: 6000 }, (_, i) => [i, i % 100]);
        const chart = renderChart(H, {
          boost: { enabled: true, useGPUTranslations: false, seriesThreshold: 1 },
          series: [{ type: 'line', boostThreshold: 1, data }],
        } as unknown as Parameters<typeof renderChart>[1]) as AnyChart;
        expect(chart.series[0].boosted).toBe(true);
        const model = extractChartModel(chart, opts());
        const s = model.series[0]!;
        expect(s.points).toHaveLength(6000);
        expect(s.points.filter((p) => typeof p.y === 'number')).toHaveLength(6000);
        expect(s.points[5]).toMatchObject({ x: 5, y: 5 });
        const d = model.warnings.find((w) => w.code === 'DATA_MODE_FALLBACK');
        expect(d).toMatchObject({ outcome: 'approximated', severity: 'info' });
        expect(d?.message).toMatch(/boost module active; values read from series data/);
        expect(codes(model)).not.toContain('DATA_CROPPED');
      } finally {
        proto.getContext = original;
      }
    });
  });

  describe(`fidelity backlog extraction (${label})`, () => {
    afterEach(() => destroyAll());

    it('errorbar series carry low/high, their parent id and the black errorbar color', () => {
      const live = extract(F.errorBarChart).model;
      const plain = extractChartModelFromOptions(F.errorBarChart, opts());
      for (const m of [live, plain]) {
        const [rain, err] = m.series;
        expect(rain!.id).toBe('rain');
        expect(err!.kind).toBe('errorbar');
        expect(err!.linkedTo).toBe('rain');
        expect(err!.points.map((p) => [p.low, p.high])).toEqual([
          [48, 51],
          [68, 73],
          [92, 110],
          [128, 136],
        ]);
        expect(hex(err!.color)).toBe('#000000');
        expect(err!.line?.width).toBeGreaterThan(0);
        // The errorbar does not consume a palette slot: the column keeps color 0.
        expect(hex(rain!.color)).toBe('#2caffe');
        expect(m.warnings.filter((w) => w.code === 'NON_NUMERIC_VALUE')).toEqual([]);
      }
    });

    it("an errorbar without linkedTo links to the previous series (':previous' default)", () => {
      const fixture = {
        ...F.errorBarChart,
        series: [
          { type: 'column' as const, name: 'Rainfall', data: [1, 2] },
          {
            type: 'errorbar' as const,
            name: 'Err',
            data: [
              [0.5, 1.5],
              [1, 3],
            ],
          },
        ],
      };
      for (const m of [extract(fixture).model, extractChartModelFromOptions(fixture, opts())]) {
        expect(m.series[1]!.linkedTo).toBe(m.series[0]!.id);
      }
    });

    it('column ranges carry low/high and no y', () => {
      const m = extract(F.columnRangeChart).model;
      expect(m.series[0]!.kind).toBe('columnrange');
      expect(m.series[0]!.points.map((p) => [p.low, p.high, p.y])).toEqual([
        [-9.5, 8, null],
        [-7.8, 8.3, null],
        [-13.1, 9.2, null],
      ]);
      expect(m.series[0]!.bars).not.toBeNull();
      expect(codes(m)).not.toContain('NON_NUMERIC_VALUE');
    });

    it('lang separators that differ from Excel raise APPROXIMATED_NUMBER_FORMAT once per chart', () => {
      const fixture = {
        ...F.multiLine,
        lang: { thousandsSep: '.', decimalPoint: ',' },
        plotOptions: { series: { dataLabels: { enabled: true, format: '{y:,.1f}' } } },
      };
      for (const m of [extract(fixture).model, extractChartModelFromOptions(fixture, opts())]) {
        const approx = m.warnings.filter((w) => w.code === 'APPROXIMATED_NUMBER_FORMAT');
        expect(approx.map((w) => w.property).sort()).toEqual(['lang.decimalPoint', 'lang.thousandsSep']);
      }
    });

    it('a space thousands separator only matters for formats that group digits', () => {
      const grouped = {
        ...F.simpleLine,
        lang: { thousandsSep: ' ' },
        plotOptions: { series: { dataLabels: { enabled: true, format: '{y:,.0f}' } } },
      };
      const plain = {
        ...grouped,
        plotOptions: { series: { dataLabels: { enabled: true, format: '{y:.1f}' } } },
      };
      const props = (m: ChartModel): string[] =>
        m.warnings.filter((w) => w.code === 'APPROXIMATED_NUMBER_FORMAT').map((w) => w.property);
      expect(props(extractChartModelFromOptions(grouped, opts()))).toEqual(['lang.thousandsSep']);
      expect(props(extractChartModelFromOptions(plain, opts()))).toEqual([]);
    });

    it('cssVariables overrides resolve styled-mode colors headless without STYLED_MODE_FALLBACK', () => {
      const cssVariables = {
        '--highcharts-color-0': '#8e44ad',
        '--highcharts-color-1': '#16a085',
        '--highcharts-background-color': '#fafafa',
      };
      for (const m of [
        extract(F.styledMode, { cssVariables }).model,
        extractChartModelFromOptions(F.styledMode, opts({ cssVariables })),
      ]) {
        expect(m.series.map((s) => hex(s.color))).toEqual(['#8e44ad', '#16a085']);
        expect(codes(m)).not.toContain('STYLED_MODE_FALLBACK');
        expect(m.background).toMatchObject({ type: 'solid', color: { r: 0xfa, g: 0xfa, b: 0xfa } });
        expect(hex(m.colors[0])).toBe('#8e44ad');
      }
      // A variable missing from the overrides keeps the fallback (and its diagnostic).
      const partial = extract(F.styledMode, { cssVariables: { '--highcharts-color-0': '#8e44ad' } }).model;
      expect(partial.series.map((s) => hex(s.color))).toEqual(['#8e44ad', '#544fc5']);
      expect(codes(partial)).toContain('STYLED_MODE_FALLBACK');
    });
  });
}
