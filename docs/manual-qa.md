# Manual QA in Microsoft Excel

The automated checks prove that the generated `.xlsx` packages are well formed. They also prove that the chart series reference the right cells, and that the files open and render in LibreOffice. They do **not** prove that Microsoft Excel opens the files without a repair prompt or draws them the way Highcharts does. **No workbook was opened in Excel in the development environment.** This checklist is that missing step. Run it on Excel for Windows or Mac (Microsoft 365, or 2019 and later) before a release, and whenever `src/excel/` or `src/core/` changes.

## 1. Get the workbooks

Option A: from the test suite.

```bash
pnpm install
pnpm test
```

Open the files in `tests/output/export/v13/`. These are workbooks exported from real Highcharts 13 charts by `tests/integration/export-fixtures.v13.test.ts`:

| File | What it exercises |
| --- | --- |
| `line.xlsx`, `multi-line.xlsx`, `spline.xlsx` | Line charts, markers, smoothing |
| `column.xlsx`, `stacked-column.xlsx`, `percent.xlsx`, `bar.xlsx` | Clustered, stacked and 100% bars, horizontal bars, gap/overlap |
| `area.xlsx` | Area fill opacity |
| `pie.xlsx`, `doughnut.xlsx` | Slice colors, explosion, hole size |
| `scatter.xlsx` | X/Y blocks, value axes |
| `datetime.xlsx` | Excel date axis and date formats |
| `combo.xlsx`, `secondary.xlsx` | Combo chart, secondary value axis |
| `custom-colors.xlsx`, `styled-mode.xlsx`, `custom-axes.xlsx` | Colors, styled-mode fallback, reversed, opposite and log axes, min/max |
| `null-negative.xlsx`, `hidden-series.xlsx`, `hidden-series-all.xlsx`, `runtime-updated.xlsx` | Gaps, negatives, hidden series, runtime updates |
| `hidden-data-sheet.xlsx` | `includeSourceData: false` (hidden Data sheet) |

The same cases are written for Highcharts 12 in `tests/output/export/v12/`. Spot-check a few of them too. `tests/output/writer/*.xlsx` holds the writer's own fixtures, one chart type per file plus `all-charts.xlsx`.

Option B: from the demo, to compare against the live Highcharts rendering.

```bash
pnpm demo   # http://127.0.0.1:4173
```

Open a chart's context menu (☰), choose **Download editable Excel chart**, and keep the browser window open next to Excel.

## 2. Checklist per workbook

1. **Opens cleanly.** No "We found a problem with some content… Do you want us to try to recover?" prompt, and no repair log.
2. **Native chart object.** Click the chart. The *Chart Design* and *Format* ribbon tabs appear. Right-click → **Select Data…** (Windows) or **Edit Data** shows series ranges on the `Data` sheet, for example `=Data!$B$2:$B$7`.
3. **Live link.** Change a value on the `Data` sheet. The chart updates immediately. Change a header cell. The legend and series name update (scatter and bubble names are literal and do not follow the header).
4. **Change Chart Type** (Chart Design → Change Chart Type) works and produces a sensible chart.
5. **Copy into PowerPoint.** Copy the chart and paste it with *Keep Source Formatting & Embed Workbook* or *Use Destination Theme & Link Data*. In PowerPoint, *Edit Data* opens the data and edits update the chart.
6. **Visual comparison with the Highcharts rendering** (demo, or the reference image when exported with `includeReferenceImage: true`):
   - proportions and plot-area size
   - series colors, line widths, dash styles, markers
   - series order and names, and stacking
   - axes: type (category/date/value/log), min/max, reversed/opposite, number formats, label rotation
   - titles: chart title, with the subtitle as a second paragraph in its own (smaller, lighter) font, and axis titles
   - fonts: family, size, bold
   - legend position and entries
   - gridlines (major/minor, color)
   - data labels: content and position
7. **Warnings match reality.** Every visible difference should correspond to a diagnostic in `result.warnings` (the demo shows them). A difference without a diagnostic is a bug.

## 3. Results template

Copy this table into the release notes or the pull request.

| Workbook | Excel version / OS | Opens cleanly | Native chart / Select Data | Cell edit updates chart | Change Chart Type | PowerPoint keeps editability | Visual match (notes) | Tester / date |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| line.xlsx | | | | | | | | |
| column.xlsx | | | | | | | | |
| stacked-column.xlsx | | | | | | | | |
| bar.xlsx | | | | | | | | |
| area.xlsx | | | | | | | | |
| pie.xlsx | | | | | | | | |
| doughnut.xlsx | | | | | | | | |
| scatter.xlsx | | | | | | | | |
| datetime.xlsx | | | | | | | | |
| combo.xlsx | | | | | | | | |
| secondary.xlsx | | | | | | | | |
| custom-axes.xlsx | | | | | | | | |

Use ✅ / ❌ and put details in the notes column. File an issue for every ❌ with the workbook attached.

## 4. OOXML schema validation (optional)

`pnpm validate:xsd` (after `pnpm test`) runs `scripts/validate-ooxml.py` on `tests/output/export/v13/*.xlsx` and `tests/output/writer/*.xlsx`. It needs Python with `lxml` and the ECMA-376 schemas, which this repository does not ship: download ECMA-376 Part 1/4 schemas from ecma-international.org and set `OOXML_SCHEMA_DIR` to a folder containing `ISO-IEC29500-4_2016/*.xsd` and `ecma/fouth-edition/opc-*.xsd`. It catches element-order and type errors that make Excel repair or drop a chart, but a schema-valid file is still not proof that Excel opens it correctly: the checklist above is.

## 5. LibreOffice renders are only a smoke check

`tests/integration/render-libreoffice.test.ts` converts workbooks to PDF and PNG in `tests/output/render/` when `soffice`, `pdftoppm` and `pdftotext` are on the PATH. `tests/integration/writer-render.test.ts` does the same in `tests/output/writer/all-charts.pdf`. These renders show that the chart parts parse and draw. LibreOffice's chart engine is not Excel's, though: fonts, label placement, axis crossing and repair behaviour differ. For example, LibreOffice draws a rich chart title entirely in its first paragraph's style, so the subtitle line appears in the title font there. Excel honours the per-paragraph fonts, and that is what the checklist must confirm. A good LibreOffice render is **not** evidence that Excel accepts the file. Only the checklist above is.

## 6. Automated validation layers, and what still needs a human in Excel

Two layers check the package structure without Excel, plus one that watches the LibreOffice renders for unreviewed changes:

| Layer | Command | Where it runs | What it catches |
| --- | --- | --- | --- |
| Schema validation (XSD) | `pnpm validate:xsd` | Locally, when `OOXML_SCHEMA_DIR` and `lxml` are available (section 4) | Every XML part against the ECMA-376 schemas: element order, unknown elements, value types. |
| Open XML SDK validation | `pnpm validate:openxml` (after `pnpm test`) | CI job `openxml-validate`; locally with the .NET 8 SDK | `tools/ooxml-validator` opens each workbook with `SpreadsheetDocument.Open` from Microsoft's `DocumentFormat.OpenXml` and runs `OpenXmlValidator` for Office 2016: schema plus the SDK's semantic constraints (attribute ranges, relationship targets, part-to-part references, elements Office 2016 does not know). It prints `file: part: path: description [id]` per error and exits 1 on any. See `tools/ooxml-validator/README.md`. |
| LibreOffice visual regression | `pnpm test:visual` | CI `verify` job, after coverage | Eight fixtures rendered by LibreOffice, cropped to the chart and compared with `tests/baselines/render/*.png` (pixelmatch threshold 0.1, at most 1.5% differing pixels). A change in what LibreOffice draws fails until someone reviews it and runs `UPDATE_BASELINES=1 pnpm test:visual`. It says nothing about Excel. |

The Open XML SDK is the closest automated check to Excel's own parser, but it is not Excel: Excel's loader applies extra rules the SDK does not model, and it decides on its own when to show the repair prompt. A file that passes both layers can still be repaired or drawn differently by Excel. These still need a person with Excel (the checklist in section 2):

- The file opens with no "We found a problem with some content" prompt, and no chart is dropped.
- The chart looks right: colors, fonts and per-paragraph title fonts, label placement, axis crossing, log and date axes, gap/overlap, hole size.
- The chart is editable: Select Data shows the Data sheet ranges, editing a cell redraws the chart, the series names come from cells.
- Number formats, hidden-sheet behaviour and the Meta sheet read as intended.
