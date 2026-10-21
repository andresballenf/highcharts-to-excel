---
'highcharts-editable-excel': minor
---

Closer to what the chart shows, out of the box:

- **Number formats from formatters.** On a live chart, value-axis `labels.formatter`, `dataLabels.formatter` and tooltip formatters now produce Excel number formats: the rendered labels and the tooltip text are matched to their values and the shared pattern (currency prefix or suffix, `k`/`M`/`B` scaling, percent, decimals, grouping) becomes the format code, e.g. `$5K` → `$#,##0,"K"` and `$12,500.50` → `$#,##0.00`. Highcharts' default `10k` axis labels become `#,##0,"k"`. Reported as `APPROXIMATED_NUMBER_FORMAT` (info). New option `inferFormats` (default `true`; `false` never calls a formatter).
- **Data cell formats** also come from `tooltip.pointFormat`, then the data label format, then a non-scaling y axis format, when the tooltip gives none.
- **Series names** follow the legend: the rendered legend label (`legend.labelFormatter` / `labelFormat`) instead of the internal `series.name`.
- **Spacer series** with no values that are unnamed or have `showInLegend: false` are no longer exported as "Series N" columns full of `#N/A` (`EMPTY_SERIES`, info, `details.excluded: true`). Empty series that are kept get blank cells instead of `#N/A`.
