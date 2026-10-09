/**
 * Writer-agnostic description of the workbook to produce, expressed in Excel/OOXML vocabulary.
 *
 * The translation layer (src/core) turns a ChartModel into a WorkbookSpec; an ExcelWriter turns the
 * WorkbookSpec into bytes. Nothing in here references Highcharts.
 */

// ---------------------------------------------------------------------------
// Workbook / worksheet
// ---------------------------------------------------------------------------

export interface WorkbookSpec {
  properties: { title?: string; creator?: string; created?: Date };
  /** Tab order. At least one sheet must be visible. */
  sheets: SheetSpec[];
}

export interface SheetSpec {
  /** Already sanitized and unique (see utils/filenames.sanitizeSheetName). */
  name: string;
  hidden: boolean;
  columns: Array<{ col0: number; widthChars: number }>;
  rows: RowSpec[];
  freezeHeaderRow: boolean;
  drawings: DrawingSpec[];
}

export interface RowSpec {
  row0: number;
  cells: CellSpec[];
}

export type CellValue =
  | { type: 'string'; value: string }
  | { type: 'number'; value: number }
  | { type: 'boolean'; value: boolean }
  | { type: 'blank' };

export interface CellStyleSpec {
  bold?: boolean;
  italic?: boolean;
  /** Excel number format code, e.g. "0.00", "yyyy-mm-dd". */
  numberFormat?: string;
  /** RRGGBB */
  fillHex?: string;
  /** RRGGBB */
  fontColorHex?: string;
  align?: 'left' | 'center' | 'right';
}

export interface CellSpec {
  col0: number;
  row0: number;
  value: CellValue;
  style?: CellStyleSpec;
}

export interface AnchorSpec {
  col0: number;
  row0: number;
  colOffsetPx: number;
  rowOffsetPx: number;
  widthPx: number;
  heightPx: number;
}

export type DrawingSpec =
  | { kind: 'chart'; chart: ExcelChartSpec; anchor: AnchorSpec; name: string }
  | { kind: 'image'; png: Uint8Array; anchor: AnchorSpec; name: string };

// ---------------------------------------------------------------------------
// Drawing primitives (DrawingML)
// ---------------------------------------------------------------------------

export interface ExcelFontSpec {
  typeface: string | null;
  /** Hundredths of a point (a:rPr sz). */
  sizeHundredthsPt: number | null;
  bold: boolean;
  italic: boolean;
  /** RRGGBB */
  colorHex: string | null;
}

export type ExcelFillSpec =
  | { type: 'none' }
  | { type: 'solid'; hex: string; alpha: number }
  | { type: 'gradient'; stops: Array<{ pos: number; hex: string; alpha: number }>; angle: number };

/** OOXML a:prstDash values. */
export type OoxmlDash =
  | 'solid'
  | 'dash'
  | 'dot'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot'
  | 'sysDashDot'
  | 'sysDashDotDot';

export interface ExcelLineSpec {
  /** Width in CSS px; converted to EMU by the writer. */
  widthPx: number;
  /** RRGGBB; null with noFill=false means "inherit". */
  hex: string | null;
  alpha: number;
  dash: OoxmlDash;
  /** Explicit "no line". */
  noFill: boolean;
}

export interface ExcelTextSpec {
  /** Lines are joined with line breaks in a single paragraph each. */
  lines: string[];
  font: ExcelFontSpec | null;
  overlay: boolean;
}

export interface ExcelShapeStyle {
  /** null = writer default (inherit from theme / Excel automatic). */
  fill: ExcelFillSpec | null;
  line: ExcelLineSpec | null;
}

// ---------------------------------------------------------------------------
// Chart
// ---------------------------------------------------------------------------

export type ExcelDataLabelPosition = 'ctr' | 'inEnd' | 'inBase' | 'outEnd' | 't' | 'b' | 'l' | 'r' | 'bestFit';

export interface ExcelDataLabelsSpec {
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  showPercent: boolean;
  /** null = Excel default for the chart type. Only positions valid for the chart type may be used. */
  position: ExcelDataLabelPosition | null;
  numberFormat: string | null;
  font: ExcelFontSpec | null;
  fill: ExcelFillSpec | null;
  line: ExcelLineSpec | null;
}

export interface ExcelMarkerSpec {
  symbol: 'circle' | 'square' | 'diamond' | 'triangle' | 'none' | 'auto';
  /** Points, 2-72. */
  size: number;
  fill: ExcelFillSpec | null;
  line: ExcelLineSpec | null;
}

export interface ExcelSeriesRef<T> {
  /** Sheet-qualified absolute formula, e.g. 'Data'!$B$2:$B$5 */
  formula: string;
  /** Values to write as cache so the chart renders before recalculation. */
  cache: T[];
  /** Number format code for numCache; defaults to General. */
  formatCode?: string;
}

export interface ExcelSeriesSpec {
  idx: number;
  order: number;
  name: { kind: 'ref'; formula: string; cache: string } | { kind: 'literal'; text: string };
  /** Category labels (bar/line/area/pie) or X values (scatter/bubble). Null for an index-based axis. */
  categories: (ExcelSeriesRef<string | number | null> & { kind: 'str' | 'num' }) | null;
  values: ExcelSeriesRef<number | null>;
  /** Bubble sizes. */
  bubbleSizes: ExcelSeriesRef<number | null> | null;
  shape: ExcelShapeStyle;
  marker: ExcelMarkerSpec | null;
  smooth: boolean | null;
  dataPoints: Array<{
    idx: number;
    shape: ExcelShapeStyle | null;
    marker: ExcelMarkerSpec | null;
    /** Pie slice explosion percent. */
    explosion: number | null;
    /** Hide this point's data label / set per-point labels. */
    dataLabels: ExcelDataLabelsSpec | null;
  }>;
  dataLabels: ExcelDataLabelsSpec | null;
  invertIfNegative: boolean;
}

export type ExcelGrouping = 'clustered' | 'stacked' | 'percentStacked';
export type ExcelLineGrouping = 'standard' | 'stacked' | 'percentStacked';
export type ExcelScatterStyle = 'lineMarker' | 'marker' | 'smoothMarker' | 'line' | 'smooth' | 'none';

export type PlotGroupSpec =
  | {
      kind: 'bar';
      barDir: 'col' | 'bar';
      grouping: ExcelGrouping;
      /** Percent, 0-500. */
      gapWidth: number;
      /** Percent, -100..100. Stacked charts need 100. */
      overlap: number | null;
      varyColors: boolean;
      series: ExcelSeriesSpec[];
      axisIds: [number, number];
      dataLabels: ExcelDataLabelsSpec | null;
    }
  | {
      kind: 'line';
      grouping: ExcelLineGrouping;
      varyColors: boolean;
      showMarkers: boolean;
      series: ExcelSeriesSpec[];
      axisIds: [number, number];
      dataLabels: ExcelDataLabelsSpec | null;
    }
  | {
      kind: 'area';
      grouping: ExcelLineGrouping;
      varyColors: boolean;
      series: ExcelSeriesSpec[];
      axisIds: [number, number];
      dataLabels: ExcelDataLabelsSpec | null;
    }
  | {
      kind: 'scatter';
      scatterStyle: ExcelScatterStyle;
      varyColors: boolean;
      series: ExcelSeriesSpec[];
      axisIds: [number, number];
      dataLabels: ExcelDataLabelsSpec | null;
    }
  | {
      kind: 'bubble';
      varyColors: boolean;
      /** Percent, 0-300. */
      bubbleScale: number;
      series: ExcelSeriesSpec[];
      axisIds: [number, number];
      dataLabels: ExcelDataLabelsSpec | null;
    }
  | {
      kind: 'pie';
      varyColors: boolean;
      /** Degrees, 0-360. */
      firstSliceAngle: number;
      series: ExcelSeriesSpec[];
      dataLabels: ExcelDataLabelsSpec | null;
    }
  | {
      kind: 'doughnut';
      varyColors: boolean;
      firstSliceAngle: number;
      /** Percent, 10-90. */
      holeSize: number;
      series: ExcelSeriesSpec[];
      dataLabels: ExcelDataLabelsSpec | null;
    };

export type ExcelAxisKind = 'cat' | 'val' | 'date';
export type ExcelAxisPosition = 'b' | 'l' | 'r' | 't';

export interface ExcelAxisSpec {
  id: number;
  kind: ExcelAxisKind;
  position: ExcelAxisPosition;
  crossAxisId: number;
  /** Hidden axis (c:delete). */
  deleted: boolean;
  title: ExcelTextSpec | null;
  numberFormat: { code: string; sourceLinked: boolean } | null;
  /** null = no gridlines. */
  majorGridlines: ExcelLineSpec | null;
  minorGridlines: ExcelLineSpec | null;
  /** null = default axis line; use noFill for none. */
  axisLine: ExcelLineSpec | null;
  labels: {
    /** 'none' hides labels. */
    position: 'nextTo' | 'low' | 'high' | 'none';
    font: ExcelFontSpec | null;
    /** Degrees, -90..90; null = auto. */
    rotation: number | null;
  };
  scaling: {
    min: number | null;
    max: number | null;
    orientation: 'minMax' | 'maxMin';
    logBase: number | null;
  };
  majorUnit: number | null;
  minorUnit: number | null;
  crosses: 'autoZero' | 'min' | 'max' | { at: number };
  majorTickMark: 'none' | 'out' | 'in' | 'cross';
  /** Date axes only. */
  dateAxis: { baseTimeUnit: 'days' | 'months' | 'years' | null } | null;
}

export interface ExcelChartSpec {
  title: ExcelTextSpec | null;
  /** Default text style for the whole chart (c:chartSpace/c:txPr). */
  textDefaults: ExcelFontSpec | null;
  chartArea: ExcelShapeStyle;
  plotArea: ExcelShapeStyle & {
    /** Fractions of the chart area (0-1). null = automatic layout. */
    manualLayout: { x: number; y: number; w: number; h: number } | null;
  };
  plotGroups: PlotGroupSpec[];
  /** Empty for pie/doughnut. */
  axes: ExcelAxisSpec[];
  legend: {
    position: 'b' | 't' | 'l' | 'r' | 'tr';
    overlay: boolean;
    font: ExcelFontSpec | null;
    fill: ExcelFillSpec | null;
    line: ExcelLineSpec | null;
  } | null;
  dispBlanksAs: 'gap' | 'zero' | 'span';
  /** Excel chart style number (1-48). Null = omit. */
  style: number | null;
}

// ---------------------------------------------------------------------------
// Writer
// ---------------------------------------------------------------------------

export interface ExcelWriter {
  readonly name: string;
  write(workbook: WorkbookSpec): Promise<Uint8Array>;
}

/** Excel limits (XLSX, Excel 2007+). */
export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLUMNS = 16_384;
/** Excel 2007 guidance for points per series; later versions are memory-bound. Exceeding it raises a warning. */
export const EXCEL_MAX_POINTS_PER_SERIES = 32_000;
/** Maximum series per chart. */
export const EXCEL_MAX_SERIES_PER_CHART = 255;
