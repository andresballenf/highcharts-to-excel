/**
 * Hand-built WorkbookSpec / ExcelChartSpec fixtures for the OOXML writer tests.
 */
import type {
  CellSpec,
  ExcelAxisSpec,
  ExcelChartSpec,
  ExcelDataLabelsSpec,
  ExcelFontSpec,
  ExcelLineSpec,
  ExcelSeriesSpec,
  PlotGroupSpec,
  RowSpec,
  WorkbookSpec,
} from '../../src/excel/writer-interface';

export const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

export const CAT_F = "'Data'!$A$2:$A$5";
export const VAL1_F = "'Data'!$B$2:$B$5";
export const VAL2_F = "'Data'!$C$2:$C$5";
export const NAME1_F = "'Data'!$B$1";
export const NAME2_F = "'Data'!$C$1";

export const font = (o: Partial<ExcelFontSpec> = {}): ExcelFontSpec => ({
  typeface: 'Arial',
  sizeHundredthsPt: 1100,
  bold: false,
  italic: false,
  colorHex: '333333',
  ...o,
});

export const line = (o: Partial<ExcelLineSpec> = {}): ExcelLineSpec => ({
  widthPx: 2,
  hex: '4472C4',
  alpha: 1,
  dash: 'solid',
  noFill: false,
  ...o,
});

export const labels = (o: Partial<ExcelDataLabelsSpec> = {}): ExcelDataLabelsSpec => ({
  showValue: true,
  showCategoryName: false,
  showSeriesName: false,
  showPercent: false,
  position: null,
  numberFormat: null,
  font: null,
  fill: null,
  line: null,
  ...o,
});

export function series(i: number, o: Partial<ExcelSeriesSpec> = {}): ExcelSeriesSpec {
  const second = i % 2 === 1;
  return {
    idx: i,
    order: i,
    name: { kind: 'ref', formula: second ? NAME2_F : NAME1_F, cache: second ? 'Costs' : 'Sales' },
    categories: { kind: 'str', formula: CAT_F, cache: ['Q1', 'Q2', null, 'Q4'] },
    values: { formula: second ? VAL2_F : VAL1_F, cache: second ? [4, 3, 2, 1] : [1, null, 3, Number.NaN] },
    bubbleSizes: null,
    shape: { fill: { type: 'solid', hex: second ? 'ED7D31' : '4472C4', alpha: 0.8 }, line: null },
    marker: null,
    smooth: null,
    dataPoints: [],
    dataLabels: null,
    invertIfNegative: false,
    ...o,
  };
}

export function axis(
  id: number,
  kind: ExcelAxisSpec['kind'],
  crossAxisId: number,
  o: Partial<ExcelAxisSpec> = {},
): ExcelAxisSpec {
  return {
    id,
    kind,
    position: kind === 'val' ? 'l' : 'b',
    crossAxisId,
    deleted: false,
    title: null,
    numberFormat: null,
    majorGridlines: kind === 'val' ? line({ widthPx: 1, hex: 'D9D9D9' }) : null,
    minorGridlines: null,
    axisLine: null,
    labels: { position: 'nextTo', font: null, rotation: null },
    scaling: { min: null, max: null, orientation: 'minMax', logBase: null },
    majorUnit: null,
    minorUnit: null,
    crosses: 'autoZero',
    majorTickMark: 'out',
    dateAxis: null,
    ...o,
  };
}

export function chart(
  plotGroups: PlotGroupSpec[],
  axes: ExcelAxisSpec[],
  o: Partial<ExcelChartSpec> = {},
): ExcelChartSpec {
  return {
    title: { lines: ['Quarterly results'], font: font({ bold: true, sizeHundredthsPt: 1400 }), overlay: false },
    textDefaults: font(),
    chartArea: {
      fill: { type: 'solid', hex: 'FFFFFF', alpha: 1 },
      line: { widthPx: 0, hex: null, alpha: 1, dash: 'solid', noFill: true },
    },
    plotArea: { fill: null, line: null, manualLayout: null },
    plotGroups,
    axes,
    legend: {
      position: 'b',
      overlay: false,
      font: font({ sizeHundredthsPt: 900 }),
      fill: null,
      line: null,
      deletedEntries: [],
    },
    dispBlanksAs: 'gap',
    style: null,
    ...o,
  };
}

const catVal = (): ExcelAxisSpec[] => [axis(100, 'cat', 200), axis(200, 'val', 100)];

export const chartSpecs = {
  barClustered: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bar',
          barDir: 'col',
          grouping: 'clustered',
          gapWidth: 150,
          overlap: -10,
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: null,
          series: [
            series(0, {
              dataLabels: labels({ position: 'outEnd', numberFormat: '0.0', font: font({ bold: true }) }),
              dataPoints: [
                {
                  idx: 2,
                  shape: { fill: { type: 'solid', hex: 'FF0000', alpha: 1 }, line: null },
                  marker: null,
                  explosion: null,
                  dataLabels: null,
                },
                { idx: 0, shape: null, marker: null, explosion: null, dataLabels: labels({ showValue: false }) },
              ],
            }),
            series(1),
          ],
        },
      ],
      catVal(),
    ),
  barStacked: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bar',
          barDir: 'col',
          grouping: 'stacked',
          gapWidth: 80,
          overlap: null,
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: labels({ position: 'outEnd' }),
          series: [series(0, { dataLabels: labels({ position: 'outEnd' }) }), series(1)],
        },
      ],
      catVal(),
    ),
  barHorizontal: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bar',
          barDir: 'bar',
          grouping: 'clustered',
          gapWidth: 999,
          overlap: 0,
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: null,
          series: [series(0)],
        },
      ],
      [axis(100, 'cat', 200, { position: 'l' }), axis(200, 'val', 100, { position: 'b' })],
    ),
  line: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'line',
          grouping: 'standard',
          varyColors: false,
          showMarkers: true,
          axisIds: [100, 200],
          dataLabels: null,
          series: [
            series(0, {
              shape: { fill: null, line: line({ widthPx: 2, hex: '4472C4', dash: 'dash' }) },
              marker: { symbol: 'circle', size: 7, fill: { type: 'solid', hex: '4472C4', alpha: 1 }, line: null },
              smooth: true,
              dataLabels: labels({ position: 't' }),
              dataPoints: [
                {
                  idx: 1,
                  shape: null,
                  marker: { symbol: 'diamond', size: 100, fill: null, line: null },
                  explosion: null,
                  dataLabels: null,
                },
              ],
            }),
            series(1, { marker: { symbol: 'none', size: 5, fill: null, line: null } }),
          ],
        },
      ],
      catVal(),
    ),
  areaStacked: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'area',
          grouping: 'stacked',
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: labels({ position: 'ctr' }),
          series: [series(0, { dataLabels: labels({ position: 'outEnd' }) }), series(1)],
        },
      ],
      catVal(),
    ),
  scatter: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'scatter',
          scatterStyle: 'lineMarker',
          varyColors: false,
          axisIds: [300, 400],
          dataLabels: null,
          series: [
            series(0, {
              categories: { kind: 'num', formula: "'Data'!$D$2:$D$5", cache: [1.5, 2.5, null, 4.5] },
              marker: { symbol: 'square', size: 6, fill: null, line: line({ widthPx: 1 }) },
            }),
          ],
        },
      ],
      [axis(300, 'val', 400, { position: 'b' }), axis(400, 'val', 300)],
    ),
  bubble: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bubble',
          varyColors: false,
          bubbleScale: 400,
          axisIds: [300, 400],
          dataLabels: null,
          series: [
            series(0, {
              categories: { kind: 'num', formula: "'Data'!$D$2:$D$5", cache: [1, 2, 3, 4] },
              bubbleSizes: { formula: "'Data'!$E$2:$E$5", cache: [10, 20, null, 40] },
              invertIfNegative: false,
            }),
          ],
        },
      ],
      [axis(300, 'val', 400, { position: 'b' }), axis(400, 'val', 300)],
    ),
  pie: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'pie',
          varyColors: true,
          firstSliceAngle: 90,
          dataLabels: null,
          series: [
            series(0, {
              values: { formula: VAL2_F, cache: [4, 3, 2, 1] },
              dataLabels: labels({ showPercent: true, showCategoryName: true, position: 'bestFit' }),
              dataPoints: [
                {
                  idx: 0,
                  shape: { fill: { type: 'solid', hex: '70AD47', alpha: 1 }, line: null },
                  marker: null,
                  explosion: 15,
                  dataLabels: null,
                },
              ],
            }),
          ],
        },
      ],
      [],
    ),
  doughnut: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'doughnut',
          varyColors: true,
          firstSliceAngle: 400,
          holeSize: 5,
          dataLabels: null,
          series: [
            series(0, { values: { formula: VAL2_F, cache: [4, 3, 2, 1] }, dataLabels: labels({ position: 'outEnd' }) }),
          ],
        },
      ],
      [],
    ),
  combo: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bar',
          barDir: 'col',
          grouping: 'clustered',
          gapWidth: 150,
          overlap: null,
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: null,
          series: [series(0)],
        },
        {
          kind: 'line',
          grouping: 'standard',
          varyColors: false,
          showMarkers: false,
          axisIds: [100, 200],
          dataLabels: null,
          series: [series(1)],
        },
      ],
      catVal(),
    ),
  secondary: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bar',
          barDir: 'col',
          grouping: 'clustered',
          gapWidth: 150,
          overlap: null,
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: null,
          series: [series(0)],
        },
        {
          kind: 'line',
          grouping: 'standard',
          varyColors: false,
          showMarkers: true,
          axisIds: [101, 201],
          dataLabels: null,
          series: [series(1)],
        },
      ],
      [
        axis(100, 'cat', 200),
        axis(200, 'val', 100, { title: { lines: ['Sales', '(USD)'], font: null, overlay: false } }),
        axis(101, 'cat', 201, { deleted: true }),
        axis(201, 'val', 101, { position: 'r', crosses: 'max', majorGridlines: null }),
      ],
    ),
  dateAxis: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'line',
          grouping: 'standard',
          varyColors: false,
          showMarkers: true,
          axisIds: [100, 200],
          dataLabels: null,
          series: [
            series(0, {
              categories: {
                kind: 'num',
                formula: "'Data'!$F$2:$F$5",
                cache: [45292, 45323, 45352, 45383],
                formatCode: 'yyyy-mm-dd',
              },
            }),
          ],
        },
      ],
      [
        axis(100, 'date', 200, {
          numberFormat: { code: 'mmm yy', sourceLinked: false },
          dateAxis: { baseTimeUnit: 'months' },
          majorUnit: 1,
          labels: { position: 'low', font: font({ colorHex: '666666' }), rotation: -45 },
        }),
        axis(200, 'val', 100),
      ],
    ),
  logAxis: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'line',
          grouping: 'standard',
          varyColors: false,
          showMarkers: true,
          axisIds: [100, 200],
          dataLabels: null,
          series: [series(0, { values: { formula: VAL2_F, cache: [1, 10, 100, 1000] } })],
        },
      ],
      [
        axis(100, 'cat', 200, { majorUnit: 2 }),
        axis(200, 'val', 100, {
          scaling: { min: 1, max: 10000, orientation: 'maxMin', logBase: 10 },
          majorUnit: 1,
          crosses: { at: 1 },
          numberFormat: { code: '#,##0', sourceLinked: false },
          axisLine: line({ widthPx: 1, hex: '000000' }),
        }),
      ],
    ),
  gradientLayout: (): ExcelChartSpec =>
    chart(
      [
        {
          kind: 'bar',
          barDir: 'col',
          grouping: 'percentStacked',
          gapWidth: 50,
          overlap: 20,
          varyColors: false,
          axisIds: [100, 200],
          dataLabels: null,
          series: [
            series(0, {
              shape: {
                fill: {
                  type: 'gradient',
                  angle: 45,
                  stops: [
                    { pos: 1, hex: '1F4E79', alpha: 1 },
                    { pos: 0, hex: '9DC3E6', alpha: 0.5 },
                  ],
                },
                line: line({ widthPx: 2, hex: '1F4E79', alpha: 0.8, dash: 'sysDot' }),
              },
            }),
          ],
        },
      ],
      catVal(),
      {
        plotArea: {
          fill: { type: 'solid', hex: 'F2F2F2', alpha: 1 },
          line: null,
          manualLayout: { x: 0.1, y: 0.15, w: 0.8, h: 1.5 },
        },
        style: 2,
        title: null,
        legend: null,
        dispBlanksAs: 'span',
      },
    ),
} satisfies Record<string, () => ExcelChartSpec>;

export type ChartSpecName = keyof typeof chartSpecs;

const str = (col0: number, row0: number, value: string, style?: CellSpec['style']): CellSpec => ({
  col0,
  row0,
  value: { type: 'string', value },
  ...(style ? { style } : {}),
});
const num = (col0: number, row0: number, value: number, style?: CellSpec['style']): CellSpec => ({
  col0,
  row0,
  value: { type: 'number', value },
  ...(style ? { style } : {}),
});

/** A Data sheet with header row, number formats, risky strings; a Chart sheet; a hidden sheet. */
export function fullWorkbook(o: { charts?: ExcelChartSpec[]; image?: boolean } = {}): WorkbookSpec {
  const header = { bold: true, fillHex: 'DDEBF7', align: 'center' as const };
  const pct = { numberFormat: '0.0%' };
  const date = { numberFormat: 'yyyy-mm-dd' };
  const rows: RowSpec[] = [
    {
      row0: 0,
      cells: [
        str(0, 0, 'Quarter', header),
        str(1, 0, 'Sales', header),
        str(2, 0, 'Costs', header),
        str(3, 0, 'X', header),
        str(4, 0, 'Size', header),
        str(5, 0, 'Date', header),
        str(6, 0, 'Share', header),
      ],
    },
    {
      row0: 1,
      cells: [
        str(0, 1, 'Q1'),
        num(1, 1, 1),
        num(2, 1, 4),
        num(3, 1, 1.5),
        num(4, 1, 10),
        num(5, 1, 45292, date),
        num(6, 1, 0.125, pct),
      ],
    },
    // Sales Q2 is null: written as a blank (omitted) cell.
    {
      row0: 2,
      cells: [
        str(0, 2, 'Q2'),
        { col0: 1, row0: 2, value: { type: 'blank' } },
        num(2, 2, 3),
        num(3, 2, 2.5),
        num(4, 2, 20),
        num(5, 2, 45323, date),
        num(6, 2, 0.5, pct),
      ],
    },
    {
      row0: 3,
      cells: [
        { col0: 0, row0: 3, value: { type: 'blank' }, style: { italic: true } },
        num(1, 3, 3),
        num(2, 3, 2),
        num(3, 3, 3.5),
        num(4, 3, 30),
        num(5, 3, 45352, date),
        num(6, 3, 1, pct),
      ],
    },
    {
      row0: 4,
      cells: [
        str(0, 4, 'Q4'),
        num(1, 4, Number.NaN),
        num(2, 4, 1),
        num(3, 4, 4.5),
        num(4, 4, 40),
        num(5, 4, 45383, date),
        num(6, 4, -0, pct),
      ],
    },
    // Rows given out of order on purpose; the writer sorts them.
    {
      row0: 7,
      cells: [
        str(0, 7, '=1+1'),
        str(1, 7, '+foo'),
        str(2, 7, '-5'),
        str(3, 7, '@x'),
        str(4, 7, 'Ünïcødé & <tags> "q" \u0001end'),
        { col0: 5, row0: 7, value: { type: 'boolean', value: true } },
        { col0: 6, row0: 7, value: { type: 'error', value: '#N/A' } },
      ],
    },
    {
      row0: 6,
      cells: [str(0, 6, 'Notes', { bold: true, italic: true, fontColorHex: 'C00000', numberFormat: '0.00' })],
    },
  ];
  const charts = o.charts ?? [chartSpecs.combo()];
  return {
    properties: { title: 'Writer test & co', creator: 'vitest', created: new Date('2024-05-06T07:08:09.123Z') },
    sheets: [
      {
        name: 'Data',
        hidden: false,
        columns: [
          { col0: 0, widthChars: 14 },
          { col0: 4, widthChars: 30.5 },
        ],
        rows,
        freezeHeaderRow: true,
        drawings: [],
      },
      {
        name: 'Chart',
        hidden: false,
        columns: [],
        rows: [],
        freezeHeaderRow: false,
        drawings: [
          ...charts.map((c, i) => ({
            kind: 'chart' as const,
            chart: c,
            name: `Chart ${i + 1}`,
            anchor: { col0: 1, row0: 1 + i * 22, colOffsetPx: 4, rowOffsetPx: 2, widthPx: 600, heightPx: 400 },
          })),
          ...(o.image === false
            ? []
            : [
                {
                  kind: 'image' as const,
                  png: PNG_1X1,
                  name: 'Logo',
                  anchor: { col0: 12, row0: 1, colOffsetPx: 0, rowOffsetPx: 0, widthPx: 32, heightPx: 32 },
                },
              ]),
        ],
      },
      {
        name: 'Meta',
        hidden: true,
        columns: [],
        rows: [{ row0: 0, cells: [str(0, 0, 'hidden metadata')] }],
        freezeHeaderRow: false,
        drawings: [],
      },
    ],
  };
}
