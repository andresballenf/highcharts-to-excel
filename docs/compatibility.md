# Compatibility reference

This page details the [README compatibility matrix](../README.md#compatibility-matrix). The tables come from reading the extractors (`src/highcharts/extract-*.ts`), the translator (`src/core/translate-chart.ts`, `src/core/style-mapping.ts`, `src/core/data-layout.ts`), the chart-type registry (`src/core/chart-type-registry.ts`) and `src/translators/*`.

Outcomes:

- **native**: represented as Excel allows; listed in `report.supported`, usually with no diagnostic.
- **approximated**: represented with a visible difference.
- **unsupported**: dropped.
- **blocking**: the export throws `CHART_NOT_EDITABLE`.

Severity defaults: `blocking` is `error`, `approximated` and `unsupported` are `warning`, and `translated` is `info`, unless the table says *(info)*.

Versioning: diagnostic codes are append-only (a minor version can add codes; renaming or removing one is a major change). The tables below describe the current output (sheet names, cell layout, chart XML), which is documented but not covered by semver. See [Stability and versioning](../README.md#stability-and-versioning).

Tested with Highcharts 12.6.2 and 13.1.1 in Node 22 + jsdom, Chromium (Playwright) and LibreOffice (smoke render). **Not verified in Microsoft Excel.** See [manual-qa.md](manual-qa.md).

## Chart types

| Highcharts type | Excel chart | Support | Notes |
| --- | --- | --- | --- |
| line | line | native | Line chart; stacking maps to stacked/percent-stacked line. |
| spline | line (smoothed) | native | Excel smoothing differs slightly from Highcharts splines. |
| area | area | native | Stacking and percent stacking supported; fill opacity kept. |
| areaspline | area | approximated | Excel area charts cannot be smoothed; drawn with straight segments. |
| column | column (clustered/stacked/100%) | native | groupPadding/pointPadding map to gap width/overlap. |
| bar | bar (clustered/stacked/100%) | native | Horizontal bars. |
| pie | pie | native | One series per pie; slice colors and sliced points (explosion) kept. |
| pie + innerSize | doughnut | native | Hole size from innerSize (10-90%); multiple rings share one hole size. |
| scatter | scatter | native | X/Y columns per series; marker-only unless lineWidth > 0. |
| bubble | bubble | native | X/Y/Size columns per series; cannot be combined with other types. |
| columnrange | stacked column (hidden base) | approximated | Floating bars: a hidden base series stacked under a Range column (=High-Low formula cells, so editing Low/High updates the chart). Ranges crossing zero use Base/Up/Down helper columns. One range series per axis. |
| arearange | stacked area (hidden base) | approximated | A hidden Low area with a Range area (=High-Low formula cells) stacked on top. One range series per axis. |
| boxplot | - | unsupported | Excel box & whisker is a chartex type that cannot reference this layout. |
| heatmap | - | unsupported | Excel has no heatmap chart (only conditional formatting). |
| treemap | - | unsupported | Excel treemap is a chartex type not produced by this library. |
| waterfall | - | unsupported | Excel waterfall is a chartex type not produced by this library. |
| funnel | - | unsupported | Excel funnel is a chartex type not produced by this library. |
| gauge | - | unsupported | Excel has no gauge chart. |
| polar (line/spline/area) | radar (marker/standard/filled) | approximated | Categories are spaced evenly around the circle and lines are straight. Polar columns/bars and other types stay blocking (Excel has no polar columns). |
| variablepie | - | unsupported | Excel pies cannot vary slice radius. |
| sankey | - | unsupported | Excel has no flow diagrams. |
| networkgraph | - | unsupported | Excel has no network/graph layout chart. |
| timeline | - | unsupported | Excel has no timeline chart type. |
| histogram | - | unsupported | Excel histogram is a chartex type computed from raw data, not from this series. |
| bellcurve | - | unsupported | Derived series (computed in the browser) with no Excel equivalent. |
| errorbar | error bars (custom) | approximated | Drawn as custom Excel error bars on the linked parent series (bar, line, area, scatter, bubble); +err/-err columns on the data sheet. Unlinked error bars are not exported. |
| lollipop | - | unsupported | No Excel equivalent without helper series and error bars. |
| dumbbell | - | unsupported | No Excel equivalent without helper series and high-low lines. |

`report.excelChartType` names the result:

- One plot group gives `line`, `column`, `bar`, `area`, `scatter`, `bubble`, `pie`, `doughnut`, `radar` or `filledRadar` (stacked types are prefixed, for example `stackedColumn` for a column range).
- Stacking adds a prefix: `stacked*` or `percentStacked*`, for example `stackedColumn` or `percentStackedArea`.
- Several groups give `combo:<a>+<b>`, for example `combo:column+line`.

## Chart level

| Highcharts option path | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| `chart.type`, `series[i].type` | native / blocking | `UNSUPPORTED_CHART_TYPE` (blocking) | Blocks when no series maps to Excel, or when pie or bubble is mixed with other types. |
| mixed series types | native | `MIXED_SERIES_TYPES` *(info)* | Excel combo chart. |
| `series[i].type` (unknown type) | unsupported | `UNSUPPORTED_SERIES_TYPE` | The series is dropped from chart and data sheet; the others export. |
| second and later plain `pie` series | approximated | `UNSUPPORTED_SERIES_TYPE` | Only the first pie is drawn. The others stay on the data sheet. |
| `pie` + doughnut ring | approximated | `APPROXIMATED_CHART_TYPE` | The pie becomes the inner doughnut ring. |
| doughnut ring `size` | approximated | `APPROXIMATED_LAYOUT` *(info)* | All rings have the same thickness. |
| doughnut rings with different slice names or order | native | none | Each ring gets its own category column (`Category 1 \| Ring 1 \| Category 2 \| Ring 2 …`) and keeps its slices in its own order. Rings listing the same names in the same order share one `Category` column. LibreOffice labels every ring (and the legend) with the first ring's category names; Excel's legend also lists the first ring's categories (per-ring labels in Excel: unverified, see `docs/manual-qa.md`). |
| `series[i].endAngle` (semi-circle) | approximated | `APPROXIMATED_CHART_TYPE` | Drawn as a full circle. |
| `series[i].startAngle` | native | none | `firstSliceAngle`. |
| `series[i].innerSize` | native | none | Hole size clamped to 10-90%. |
| `data[j].sliced` | native | none | Explosion from `slicedOffset` relative to the chart size. |
| `chart.polar` (line, spline, area, areaspline) | approximated | `APPROXIMATED_CHART_TYPE` *(info)* | Excel radar chart on a category/value axis pair: `marker` style (markers on), `standard` (markers off) or `filled` (areas). Categories are spaced evenly around the circle. One radar style per chart: the first series decides (`APPROXIMATED_CHART_TYPE` for the others); stacking, smoothing and extra y axes are approximated. |
| `chart.polar` (column, bar, other types) | blocking | `UNSUPPORTED_POLAR` | Excel has no polar columns. |
| `series[i].type: 'errorbar'` + `linkedTo` | approximated | `UNSUPPORTED_SERIES_TYPE` when unlinked | Custom Excel error bars (`errBarType both`, `errValType cust`) on the parent series (bar, line, area, scatter, bubble): `<parent> +err` = high − y and `<parent> -err` = y − low columns next to the parent. `linkedTo: ':previous'` (the errorbar default) and ids are resolved. Error bars on pie, radar or range series, a second errorbar on one parent, or an unlinked errorbar are not exported. Error bar points without a parent point: `UNALIGNED_X_VALUES`. |
| `series[i].type: 'columnrange'` / `'arearange'` | approximated | `APPROXIMATED_CHART_TYPE` *(info)* | Stacked column/bar/area with a hidden base series (no fill, no line, no legend entry). Data sheet: `<name> Low`, `<name> High`, `<name> Range` (formula cells `=High-Low` with cached values; editing Low/High updates the chart). A column range with a negative low uses `<name> Base` (`=MAX(0,Low)+MIN(0,High)`), `Up` (`=MAX(0,High)-MAX(0,Low)`) and `Down` (`=MIN(0,Low)-MIN(0,High)`) because Excel stacks negative values below the axis. Points without both ends get no range (empty cells). Data labels are not exported (`APPROXIMATED_DATA_LABELS`). A second range series, or plain series of the same type on the same axis: the range series is dropped (`UNSUPPORTED_SERIES_TYPE`). |
| `chart.options3d.enabled` | approximated | `UNSUPPORTED_3D` | Exported flat. |
| `chart.inverted` (line/area/scatter/bubble) | approximated | `APPROXIMATED_CHART_TYPE` | Excel cannot invert these. Inverted columns become horizontal bars (native). |
| horizontal bars + line/area | approximated | `APPROXIMATED_CHART_TYPE` *(info)* | Lines follow the bar orientation. |
| `series[i].stack` (several stack groups) | approximated | `APPROXIMATED_CHART_TYPE` | Merged into one stack per type and axis. |
| `series[i].stacking` | native | none | `normal` → stacked, `percent` → 100% stacked. |
| mixed stacking within one type on one axis | approximated | `APPROXIMATED_CHART_TYPE` | One group with the majority's stacking (ties → stacked); reported on each other series' `stacking`. |
| `series[i].groupPadding`, `pointPadding` | native | none | Mapped to gap width and overlap so bars keep the Highcharts width (stacked: overlap 100; gap width capped at 500). |
| `series[i].borderRadius` (explicit) | approximated | `UNSUPPORTED_STYLE` *(info)* | Square corners. |
| `chart.width`, `chart.height` / rendered size | native | none | Excel chart object size in pixels. Override with `chartWidth`/`chartHeight`. |
| plot area box (`plotLeft/Top/Width/Height`) | approximated | `APPROXIMATED_LAYOUT` *(info)* | Pinned as a manual layout, clipped to the chart area; not for pies; not in `fidelity: 'minimal'`. |
| `chart.backgroundColor`, `borderColor`/`borderWidth` | native | `UNRESOLVED_COLOR` if unparseable | |
| `chart.plotBackgroundColor`, `plotBorderColor`/`plotBorderWidth` | native | `UNRESOLVED_COLOR` if unparseable | |
| `chart.styledMode` | approximated | `STYLED_MODE_FALLBACK` | `themeOverrides.cssVariables` first (`--highcharts-color-<colorIndex % chart.colorCount>`, every slice variable for pies, `--highcharts-background-color`); no diagnostic when a series' variables all come from there. Otherwise real browser: computed SVG colors; headless: palette by `colorIndex`, background white. |
| `title.text`, `title.style` | native | none | Also written to cell A1 of the chart sheet. |
| `subtitle.text`, `subtitle.style` | approximated | `APPROXIMATED_LAYOUT` *(info)* | Second title paragraph in its own font (size, weight and color from the subtitle style). Used as the title when there is none. LibreOffice draws the whole title in the first paragraph's style; Excel honours per-paragraph fonts. |
| `annotations` | unsupported | `UNSUPPORTED_ANNOTATION` | |
| `tooltip.formatter`, `tooltip.pointFormatter`, `series[i].tooltip.*Formatter` | unsupported | `UNSUPPORTED_TOOLTIP` *(info)* | Tooltips are never exported. |
| `tooltip.valueDecimals`, `valuePrefix`, `valueSuffix` | native | `UNSUPPORTED_NUMBER_FORMAT` / `APPROXIMATED_NUMBER_FORMAT` when not representable | Become the number format of the series' data cells. |

## Series styling

| Highcharts option path | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| `series[i].color`, `colors` palette, `colorIndex` | native | `UNRESOLVED_COLOR` if unparseable | CSS variables (Highcharts 13 `var(--highcharts-color-N)`) resolve to the default palette when headless. |
| `series[i].lineWidth` | native | none | Line, spline, area outline, and scatter (`0` = markers only). |
| `series[i].dashStyle` | native | none | 11 dash styles → OOXML presets. Unknown names → solid. |
| `series[i].fillColor`, `fillOpacity` (area) | native | none | `fillOpacity` (default 0.75) applies only to the derived fill, as in Highcharts. |
| `series[i].borderColor`, `borderWidth` (column/bar/pie) | native | none | |
| `series[i].marker.enabled/symbol/radius/fillColor/lineColor/lineWidth` | native | `APPROXIMATED_MARKER` | Line, spline and scatter only. `triangle-down` becomes triangle, custom or `url()` becomes circle. `enabled` undefined follows Highcharts' spacing threshold (estimated when headless). |
| `data[j].color`, `colorByPoint`, `series[i].colors` | native | none | Fills for column, bar, bubble and pie. Marker fill for line and scatter. |
| `data[j].marker`, `data[j].borderColor/Width` | native | none | Per-point overrides. |
| `series[i].negativeColor` | approximated | `UNSUPPORTED_STYLE` | Series color used. |
| `series[i].zones` | approximated | `UNSUPPORTED_STYLE` | Base color used. |
| gradient colors (`linearGradient`) | native | none | Angle from the gradient vector. |
| gradient colors (`radialGradient`) | approximated | `UNSUPPORTED_GRADIENT` | Top-to-bottom linear gradient. |
| unresolvable gradient stops | approximated | `UNRESOLVED_COLOR` | Stops dropped. |
| pattern fills / other color objects | unsupported | `UNRESOLVED_COLOR` | Excel's automatic color. |
| `areaspline` smoothing | approximated | `APPROXIMATED_CHART_TYPE` | Straight segments. |

## Data labels

| Highcharts option path | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| `series[i].dataLabels.enabled` | native | none | Per-point `data[j].dataLabels.enabled: false` hides single labels. |
| `dataLabels.format` (`{y}`, `{point.name}`, `{series.name}`, `{percentage}`, `{y:.1f}`…) | native / approximated | `APPROXIMATED_DATA_LABELS`, `UNSUPPORTED_NUMBER_FORMAT` | Mapped to show value, category, series name and percent plus an Excel number format. Literal text between parts becomes Excel's separator. Percentages show only on pie/doughnut; elsewhere a percentage-only label shows the value instead. |
| `dataLabels.formatter` | unsupported | `UNSUPPORTED_FORMATTER` | Excel shows the raw value. |
| `dataLabels.align/verticalAlign/inside/distance` | native / approximated | `APPROXIMATED_DATA_LABELS` | Mapped to an Excel position. Positions that the Excel chart type does not allow use the default. Outside end on stacked bars becomes inside end. Area, doughnut and radar cannot be positioned *(info)*. |
| `dataLabels.style`, `backgroundColor`, `borderColor/Width` | native | `APPROXIMATED_FONT` for generic families | |

## Axes

| Highcharts option path | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| `xAxis.categories`, `type: 'category'` | native | none | Shared category column on the data sheet. |
| `xAxis.type: 'datetime'` (whole days) | native | none | Excel date axis; base unit days, months (all month starts) or years (all January 1). Values are exported in the chart's displayed timezone (`time.timezone` / `useUTC`), so Excel shows the same wall-clock dates and times. |
| `xAxis.type: 'datetime'` before 1899-12-31 | approximated | `APPROXIMATED_DATETIME` | Excel has no such dates: ISO date text on a category axis (scatter: plain day numbers). Dates before 1900-03-01 account for Excel's 1900 leap-year bug. |
| `xAxis.type: 'datetime'` (intraday) | approximated | `APPROXIMATED_DATETIME` | Evenly spaced category axis with date-time labels. |
| `series[i].pointIntervalUnit` | approximated | `APPROXIMATED_DATETIME` *(info)* | Calendar steps from `pointStart` on the displayed dates; month-end dates may differ from Highcharts' step-from-previous-point. |
| `time.timezone`, `time.useUTC: false`, `time.timezoneOffset` (non-UTC) | approximated | `APPROXIMATED_DATETIME` *(info)*, property `time.timezone` | Each datetime x value (and datetime x-axis bound) moves to the wall-clock time the chart displays; the first point's offset is `meta.datetimeOffsetMinutes`. An unknown zone name is exported as UTC (warning, `datetimeOffsetMinutes: null`). |
| date strings on a datetime axis (`['2024-01-01', 1]`, `x: '2024-01-01'`, `pointStart: '2024-01-01'`) | translated | `NON_NUMERIC_VALUE` when not a date | Parsed as the chart's wall-clock time (with `Z`/`±hh:mm`: as that instant). |
| `xAxis.type: 'linear'`, uneven x (line/column) | approximated | `APPROXIMATED_AXIS_SCALE` | Categories are evenly spaced. |
| `xAxis.type: 'logarithmic'` (line/column) | approximated | `APPROXIMATED_AXIS_SCALE` | Evenly spaced categories. |
| `yAxis.type: 'logarithmic'`, scatter `xAxis` log | native | none | `logBase` 10. |
| `yAxis[n]` second axis | native | `SECONDARY_AXIS` *(info)* | Excel secondary value axis. |
| `yAxis[n]` third axis and beyond | approximated | `UNSUPPORTED_AXIS_FEATURE` | Plotted against the secondary axis. |
| `xAxis[n]`, n ≥ 1 | approximated | `MULTIPLE_X_AXES` | The first x axis is used for all series. |
| scatter in a category combo | approximated | `APPROXIMATED_AXIS_SCALE` | Own hidden X/Y axes. |
| `min`, `max` (value axes) | native | none | |
| `min`, `max` (category axes) | approximated | `APPROXIMATED_AXIS_SCALE` *(info)* | All categories shown. |
| zoom (`userMin`/`userMax`) | approximated | `APPROXIMATED_AXIS_SCALE` *(info)* | The zoomed window becomes fixed bounds. |
| `tickInterval`, `minorTickInterval` | native | none | Value axes (not log). Date axes with base unit days. |
| `reversed`, `opposite`, `crossing` | native | none | Excel orientation, crossing and label positions. |
| `visible: false`, `labels.enabled: false` | native | none | Axis deleted, or labels hidden. |
| `labels.rotation` | native | none | Clamped to ±90°. |
| `labels.format` | native / approximated / unsupported | `APPROXIMATED_NUMBER_FORMAT`, `UNSUPPORTED_NUMBER_FORMAT` | Excel number or date format. Literal text around category and datetime labels is dropped. |
| `labels.formatter` | unsupported | `UNSUPPORTED_FORMATTER` | Excel default format. |
| `title.text`, `title.style`, `labels.style` | native | `APPROXIMATED_FONT` | |
| `gridLineWidth/Color/DashStyle`, `minorGridLine*` | native | none | |
| `lineWidth/lineColor`, `tickWidth` | native | none | |
| `plotBands`, `plotLines` | unsupported | `UNSUPPORTED_PLOT_BAND` | |

## Legend

| Highcharts option path | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| `legend.enabled` | native | none | Off when no **exported** series shows in the legend (pies default to `showInLegend: false`; hidden series left out with `seriesVisibility: 'visible'` do not count). |
| `legend.align`/`verticalAlign`: bottom, top, left, right, top-right | native | none | |
| top-left, bottom-left, bottom-right | approximated | `APPROXIMATED_LEGEND_POSITION` | Centered at top or bottom. |
| centered over the plot (`verticalAlign: 'middle'`, `align: 'center'`) | approximated | `APPROXIMATED_LEGEND_POSITION` *(info)* | Placed at the right. |
| `legend.floating` | native | none | Overlay. |
| `legend.reversed` | unsupported | `UNSUPPORTED_STYLE` *(info)* | |
| `series[i].showInLegend: false` | native | none | Legend entry deleted (`c:legendEntry`); not for pie slices. |
| `legend.itemStyle`, `backgroundColor`, `borderColor/Width` | native | `APPROXIMATED_FONT` | |

## Fonts and numbers

| Highcharts option path | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| `*.style.fontFamily` | native / approximated | `APPROXIMATED_FONT` | First usable family. System aliases are skipped. Generic families map to Office fonts (`sans-serif` → Arial and so on). |
| `*.style.fontSize`, `fontWeight`, `fontStyle`, `color` | native | none | px × 0.75 = pt. `em` sizes are resolved against the chart font (default 16 px). |
| `lang.thousandsSep`, `lang.decimalPoint` | approximated | `APPROXIMATED_NUMBER_FORMAT` (once per chart, property `lang.thousandsSep` / `lang.decimalPoint`) | Excel always uses the viewer's locale separators. Reported only when a translated format actually groups digits (`{y:,.0f}`, `valueDecimals`) with a separator other than `,`, or shows decimals with a point other than `.`. Highcharts 11's default thousands separator is a space; Highcharts 12+ leaves it unset (locale), which is not reported. |

## Data

| Situation | Outcome | Diagnostic code | Notes |
| --- | --- | --- | --- |
| grouped data (Stock `dataGrouping`) in `rendered` mode | approximated | `DATA_GROUPED` *(info)* | Grouped points exported; counts in `details`. |
| cropped points (zoom, navigator) in `rendered` mode | approximated | `DATA_CROPPED` *(info)* | |
| `raw` mode without source data | approximated | `DATA_MODE_FALLBACK` | Rendered points exported. |
| `rendered` mode, series with no rendered points (never drawn) | approximated | `DATA_MODE_FALLBACK` *(info)* | Source data exported; `dataSemantics.mode` is `'raw'`. |
| boosted series (`modules/boost`) | approximated | `DATA_MODE_FALLBACK` *(info)*, property `series[i].boostThreshold` | `series.points` hold pixels only; values come from the processed data columns (or `options.data`). |
| `series[i].keys`, `relativeXValue`, typed-array `data` (raw / options path) | translated | none | Mapped as Highcharts' `Point.optionsToObject` does. |
| hidden series, `seriesVisibility: 'visible'` | translated | `HIDDEN_SERIES_EXCLUDED` *(info)* | |
| hidden series, `seriesVisibility: 'all'` | approximated | `HIDDEN_SERIES_INCLUDED` *(info)* | Visible in Excel. |
| hidden pie slice | approximated | `HIDDEN_POINT` *(info)* | Shown in Excel. |
| `null` y values | translated | `NULL_VALUES` *(info)* | Empty cells, drawn as gaps. |
| series missing a shared x value | approximated | `UNALIGNED_X_VALUES` *(info)* | `#N/A` cell (null in the chart cache), so lines and areas stay connected as in Highcharts; real `null` points stay empty cells (gaps). |
| x between or beyond the categories of a category axis (e.g. `x: 0.5`, `x: -1`) | approximated | `UNALIGNED_X_VALUES` | Each such x becomes its own category row; never mapped onto a category by position. |
| two points of one series at the same x | approximated | `UNALIGNED_X_VALUES` | The first point is kept; the repeats stay out of the chart (`details.duplicates`). |
| non-numeric values / unknown point shapes | approximated | `NON_NUMERIC_VALUE` | Skipped or left empty, never guessed. Reported once per series (`details.count`, `details.firstIndices`); per-point data-label diagnostics are aggregated the same way. |
| empty series | translated | `EMPTY_SERIES` *(info)* | |
| no series or no data | blocking | `EMPTY_CHART` | |
| more than 1,048,576 rows / 16,384 columns | blocking | `ROW_LIMIT_EXCEEDED` / `COLUMN_LIMIT_EXCEEDED` | |
| more than 32,000 points in one series | approximated (non-blocking) | `ROW_LIMIT_EXCEEDED` | Exported anyway. 32,000 was the Excel 2007 per-series cap; Excel 2010+ is memory-bound. Older Excel versions may truncate the series or render it slowly. |
| text starting with `= + - @`, tab or CR | translated | `FORMULA_LIKE_TEXT_ESCAPED` *(info)* | Written as inline string text; reported once. |
| invalid or duplicate sheet names | approximated | `SHEET_NAME_ADJUSTED` *(info)* | |
| `includeReferenceImage` unavailable | unsupported | `WRITER_LIMITATION` | Export continues without the image (info on the options-only path). |
