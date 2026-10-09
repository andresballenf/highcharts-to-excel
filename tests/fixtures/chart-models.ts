/**
 * Hand-built ChartModels for the core translation tests.
 */

import {
  createEmptyChartModel,
  type AxisModel,
  type ChartModel,
  type Color,
  type DataLabelStyle,
  type Font,
  type PointModel,
  type SeriesModel,
  type TextBlock,
} from '../../src/types/chart-model';

export function rgb(hex: string, a = 1): Color {
  const h = hex.replace('#', '');
  return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16), a };
}

export const PALETTE = ['#2caffe', '#544fc5', '#00e272', '#fe6a35', '#6b8abc'].map((c) => rgb(c));

export function font(partial: Partial<Font> = {}): Font {
  return { family: 'Helvetica', size: 12, bold: false, italic: false, color: rgb('#333333'), ...partial };
}

export function text(t: string, partial: Partial<Font> = {}): TextBlock {
  return { text: t, font: font(partial), align: 'center', verticalAlign: 'top' };
}

export function axis(partial: Partial<AxisModel> & Pick<AxisModel, 'index' | 'kind'>): AxisModel {
  return {
    id: null,
    categories: null,
    title: null,
    labels: { enabled: true, font: null, format: null, rotation: 0 },
    min: null,
    max: null,
    dataMin: null,
    dataMax: null,
    tickInterval: null,
    minorTickInterval: null,
    reversed: false,
    opposite: false,
    visible: true,
    gridLines: null,
    minorGridLines: null,
    axisLine: null,
    tickMarks: null,
    crossing: null,
    logBase: null,
    dateFormat: null,
    ...partial,
  };
}

export function point(partial: Partial<PointModel> = {}): PointModel {
  return {
    x: null,
    name: null,
    y: null,
    z: null,
    isNull: false,
    color: null,
    border: null,
    marker: null,
    sliced: null,
    dataLabels: null,
    visible: true,
    ...partial,
  };
}

/** Points with x = index. */
export function pts(ys: Array<number | null>): PointModel[] {
  return ys.map((y, i) => point({ x: i, y, isNull: y === null }));
}

export function series(partial: Partial<SeriesModel> & Pick<SeriesModel, 'kind'>): SeriesModel {
  const index = partial.index ?? 0;
  const color = partial.color === undefined ? (PALETTE[index % PALETTE.length] ?? null) : partial.color;
  const points = partial.points ?? [];
  return {
    id: `s${index}`,
    index,
    name: `Series ${index + 1}`,
    sourceType: partial.kind,
    visible: true,
    showInLegend: true,
    xAxisIndex: 0,
    yAxisIndex: 0,
    color,
    stacking: null,
    stackGroup: null,
    line: null,
    fill: null,
    fillOpacity: 1,
    border: null,
    marker: null,
    dataLabels: null,
    smooth: false,
    bars: null,
    pie: null,
    yFormat: null,
    dataSemantics: {
      mode: 'rendered',
      grouped: false,
      cropped: false,
      sourcePointCount: points.length,
      renderedPointCount: points.length,
    },
    ...partial,
    points,
  };
}

export function labels(partial: Partial<DataLabelStyle> = {}): DataLabelStyle {
  return {
    enabled: true,
    font: null,
    format: null,
    position: 'auto',
    showValue: true,
    showCategoryName: false,
    showSeriesName: false,
    showPercentage: false,
    background: null,
    border: null,
    ...partial,
  };
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr'];

export function baseModel(partial: Partial<ChartModel> = {}): ChartModel {
  const m = createEmptyChartModel({
    sourceLibrary: 'highcharts',
    sourceVersion: '13.1.1',
    datetimeOffsetMinutes: 0,
    sourceChartType: 'line',
    extraction: 'headless',
    styledMode: false,
    chartId: null,
  });
  m.xAxes = [
    axis({
      index: 0,
      kind: 'category',
      categories: [...MONTHS],
      tickMarks: { color: rgb('#ccd6eb'), width: 1, dash: 'solid' },
    }),
  ];
  m.yAxes = [
    axis({
      index: 0,
      kind: 'linear',
      title: text('Values'),
      gridLines: { color: rgb('#e6e6e6'), width: 1, dash: 'solid' },
    }),
  ];
  m.colors = [...PALETTE];
  return { ...m, ...partial };
}

export function lineModel(): ChartModel {
  return baseModel({
    title: text('Monthly sales', { size: 18 }),
    series: [
      series({
        kind: 'line',
        index: 0,
        name: 'Sales',
        points: pts([1, 2, 3, 4]),
        line: { color: null, width: 2, dash: 'solid' },
        marker: { enabled: true, symbol: 'circle', radius: 4, fill: null, stroke: null, strokeWidth: 0 },
      }),
    ],
  });
}

/** Linear x axis; series do not share x values. */
export function multiSeriesLineModel(): ChartModel {
  const m = baseModel({
    series: [
      series({ kind: 'line', index: 0, name: 'A', points: [1, 2, 4].map((x, i) => point({ x, y: 10 + i })) }),
      series({
        kind: 'spline',
        index: 1,
        name: 'B',
        smooth: true,
        points: [2, 3, 4].map((x, i) => point({ x, y: 20 + i })),
      }),
    ],
  });
  m.xAxes = [axis({ index: 0, kind: 'linear' })];
  return m;
}

export function columnStackedModel(): ChartModel {
  const bars = { pointPadding: 0.1, groupPadding: 0.2, borderRadius: 3 };
  return baseModel({
    meta: { ...baseModel().meta, sourceChartType: 'column' },
    series: [
      series({
        kind: 'column',
        index: 0,
        name: 'North',
        stacking: 'normal',
        bars,
        points: pts([1, 2, 3, 4]),
        dataLabels: labels({ position: 'outsideEnd' }),
      }),
      series({ kind: 'column', index: 1, name: 'South', stacking: 'normal', bars, points: pts([4, 3, 2, 1]) }),
    ],
  });
}

/** Column series on an inverted chart → horizontal bars (the extractor reports the x axis reversed). */
export function barModel(): ChartModel {
  const m = baseModel({
    inverted: true,
    series: [
      series({
        kind: 'column',
        index: 0,
        name: 'Bars',
        bars: { pointPadding: 0, groupPadding: 0.1, borderRadius: 0 },
        border: { color: rgb('#ffffff'), width: 0, dash: 'solid' },
        points: pts([5, 6, 7, 8]),
      }),
    ],
  });
  m.xAxes = [{ ...m.xAxes[0]!, reversed: true }];
  return m;
}

export function areaPercentModel(): ChartModel {
  return baseModel({
    series: [
      series({ kind: 'area', index: 0, name: 'A', stacking: 'percent', fillOpacity: 0.75, points: pts([1, 2, 3, 4]) }),
      series({ kind: 'area', index: 1, name: 'B', stacking: 'percent', fillOpacity: 0.75, points: pts([3, 2, 1, 0]) }),
    ],
  });
}

export function pieModel(): ChartModel {
  const names = ['Chrome', 'Edge', 'Firefox', 'Safari'];
  const m = baseModel({
    title: text('Browsers'),
    legend: { ...baseModel().legend, enabled: false },
    series: [
      series({
        kind: 'pie',
        index: 0,
        name: 'Share',
        pie: { innerSize: 0, startAngle: -90, endAngle: null },
        border: { color: rgb('#ffffff'), width: 1, dash: 'solid' },
        dataLabels: labels({ showValue: false, showCategoryName: true, position: 'outsideEnd' }),
        points: names.map((name, i) =>
          point({ name, y: [60, 15, 15, 10][i]!, color: PALETTE[i]!, sliced: i === 1 ? 10 : null }),
        ),
      }),
    ],
  });
  m.xAxes = [];
  m.yAxes = [];
  return m;
}

export function doughnutModel(): ChartModel {
  const m = baseModel({
    series: [
      series({
        kind: 'doughnut',
        index: 0,
        name: 'Inner',
        pie: { innerSize: 0.4, startAngle: 0, endAngle: null },
        points: ['A', 'B'].map((name, i) => point({ name, y: i + 1, color: PALETTE[i]! })),
      }),
      series({
        kind: 'doughnut',
        index: 1,
        name: 'Outer',
        pie: { innerSize: 0.6, startAngle: 0, endAngle: null },
        points: ['A', 'C'].map((name, i) => point({ name, y: i + 5, color: PALETTE[i + 2]! })),
      }),
    ],
  });
  m.xAxes = [];
  m.yAxes = [];
  return m;
}

export function scatterModel(): ChartModel {
  const m = baseModel({
    series: [
      series({
        kind: 'scatter',
        index: 0,
        name: 'Men',
        points: [
          [1, 2],
          [3, 4],
          [5, 6],
        ].map(([x, y]) => point({ x: x!, y: y! })),
      }),
      series({
        kind: 'scatter',
        index: 1,
        name: 'Women',
        points: [[1.5, 2.5]].map(([x, y]) => point({ x: x!, y: y! })),
      }),
    ],
  });
  m.xAxes = [axis({ index: 0, kind: 'linear', title: text('Height') })];
  return m;
}

export function bubbleModel(): ChartModel {
  const m = baseModel({
    series: [
      series({
        kind: 'bubble',
        index: 0,
        name: 'Countries',
        fillOpacity: 0.5,
        points: [point({ x: 1, y: 2, z: 10 }), point({ x: 2, y: 3, z: 20 })],
      }),
    ],
  });
  m.xAxes = [axis({ index: 0, kind: 'linear' })];
  return m;
}

export const DAY = 86_400_000;

export function datetimeModel(): ChartModel {
  const t0 = Date.UTC(2024, 0, 1);
  const m = baseModel({
    series: [
      series({
        kind: 'line',
        index: 0,
        name: 'Temp',
        points: [0, 1, 2, 3].map((d) => point({ x: t0 + d * DAY, y: 10 + d })),
      }),
    ],
  });
  m.xAxes = [axis({ index: 0, kind: 'datetime', tickInterval: 2 * DAY })];
  return m;
}

export function comboModel(): ChartModel {
  return baseModel({
    series: [
      series({ kind: 'column', index: 0, name: 'Rainfall', points: pts([49.9, 71.5, 106.4, 129.2]) }),
      series({
        kind: 'line',
        index: 1,
        name: 'Average',
        points: pts([60, 70, 80, 90]),
        line: { color: null, width: 3, dash: 'dash' },
      }),
    ],
  });
}

export function secondaryAxisModel(): ChartModel {
  const m = baseModel({
    series: [
      series({ kind: 'column', index: 0, name: 'Rainfall', points: pts([49.9, 71.5, 106.4, 129.2]) }),
      series({
        kind: 'spline',
        index: 1,
        name: 'Temperature',
        yAxisIndex: 1,
        smooth: true,
        points: pts([7, 6.9, 9.5, 14.5]),
      }),
    ],
  });
  m.yAxes = [
    axis({
      index: 0,
      kind: 'linear',
      title: text('Rainfall (mm)'),
      labels: { enabled: true, font: null, format: { kind: 'excel', code: '0" mm"' }, rotation: 0 },
    }),
    axis({ index: 1, kind: 'linear', title: text('Temperature (°C)'), opposite: true }),
  ];
  return m;
}

export function styledModel(): ChartModel {
  const m = baseModel({
    width: 800,
    height: 500,
    title: text('Styled', { family: '"Open Sans", sans-serif', size: 20, bold: true, color: rgb('#112233') }),
    subtitle: text('Second line', { size: 12 }),
    background: {
      type: 'gradient',
      angle: 90,
      stops: [
        { offset: 0, color: rgb('#ffffff') },
        { offset: 1, color: rgb('#dddddd') },
      ],
    },
    border: { color: rgb('#333333'), width: 2, dash: 'solid' },
    plotArea: {
      background: { type: 'solid', color: rgb('#fafafa') },
      border: null,
      box: { left: 80, top: 60, width: 600, height: 350 },
    },
    legend: {
      enabled: true,
      position: 'right',
      layout: 'vertical',
      font: font({ size: 11, bold: true }),
      background: { type: 'solid', color: rgb('#eeeeee') },
      border: { color: rgb('#999999'), width: 1, dash: 'solid' },
      overlay: false,
      reversed: false,
    },
    series: [
      series({
        kind: 'column',
        index: 0,
        name: 'Styled',
        color: rgb('#ff0000'),
        border: { color: rgb('#000000'), width: 1, dash: 'solid' },
        dataLabels: labels({
          position: 'outsideEnd',
          font: font({ size: 10, bold: true }),
          format: { kind: 'excel', code: '0.0' },
        }),
        points: pts([1, 2, 3, 4]).map((p, i) => (i === 2 ? { ...p, color: rgb('#00ff00') } : p)),
      }),
    ],
  });
  m.yAxes = [
    axis({
      index: 0,
      kind: 'linear',
      title: text('Y title', { size: 14 }),
      gridLines: { color: rgb('#cccccc'), width: 2, dash: 'dash' },
      min: 0,
      max: 10,
      tickInterval: 2,
      labels: { enabled: true, font: font({ size: 11 }), format: null, rotation: 0 },
    }),
  ];
  return m;
}

export function nullNegativeModel(): ChartModel {
  return baseModel({
    series: [
      series({
        kind: 'column',
        index: 0,
        name: 'Delta',
        points: [
          point({ x: 0, y: 5 }),
          point({ x: 1, y: null, isNull: true }),
          point({ x: 2, y: -3 }),
          point({ x: 3, y: 7 }),
        ],
      }),
    ],
  });
}

export function hiddenSeriesModel(): ChartModel {
  return baseModel({
    series: [
      series({ kind: 'line', index: 0, name: 'Shown', points: pts([1, 2, 3, 4]) }),
      series({ kind: 'line', index: 1, name: 'Hidden', visible: false, points: pts([4, 3, 2, 1]) }),
    ],
  });
}

export function unknownTypeModel(): ChartModel {
  return baseModel({
    series: [series({ kind: 'unknown', sourceType: 'treemap', index: 0, name: 'Tree', points: pts([1, 2]) })],
  });
}

export function polarModel(): ChartModel {
  return { ...lineModel(), polar: true };
}

export function emptyModel(): ChartModel {
  return baseModel();
}

export function formulaLikeModel(): ChartModel {
  const m = baseModel({
    series: [series({ kind: 'column', index: 0, name: '+foo', points: pts([1, 2]) })],
  });
  m.xAxes = [axis({ index: 0, kind: 'category', categories: ['=SUM(A1)', 'plain'] })];
  return m;
}

/** 40,000 points: above Excel's 32,000-points-per-series guidance, far below the row limit. */
export function hugeModel(): ChartModel {
  const points: PointModel[] = new Array(40_000);
  for (let i = 0; i < points.length; i++) points[i] = point({ x: i, y: i % 100 });
  const m = baseModel({ series: [series({ kind: 'line', index: 0, name: 'Big', points })] });
  m.xAxes = [axis({ index: 0, kind: 'linear' })];
  return m;
}
