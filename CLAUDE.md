# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`highcharts-editable-excel`: an ESM TypeScript library (pnpm, Node 22, TS strict) that turns a live Highcharts chart, or a plain options object, into an `.xlsx` workbook. The workbook holds a **native Excel chart** whose series reference cells on a data sheet. Public API: `src/index.ts`. User docs: `README.md`, `docs/compatibility.md`, `docs/manual-qa.md`.

## Commands

```bash
pnpm typecheck                         # tsc over src, tests, e2e, demo/src, demo/vite.config.ts, scripts
pnpm lint                              # biome check . (lint + format check); pnpm lint:fix / pnpm format write
pnpm build                             # tsup → dist/index.js (ESM) + dist/index.cjs (CJS), ~285 KB each, Highcharts external, fflate bundled
pnpm test                              # vitest run (jsdom); writes tests/output/** (gitignored)
npx vitest run tests/unit/colors.test.ts   # single test file (add -t "name" for one test)
pnpm test:e2e                          # Playwright vs the Vite demo on :4173
PW_CHROMIUM_EXECUTABLE=/path/to/chrome pnpm test:e2e   # sandbox: use the preinstalled Chromium
pnpm demo                              # Vite demo; aliases the package to src/ (no build needed)
pnpm bench                             # tests/integration/bench.test.ts with BENCH=1 → tests/output/bench.json
pnpm coverage                          # vitest --coverage (v8) → tests/output/coverage; thresholds lines 90 / functions 92 / branches 78 / statements 87
REQUIRE_RENDER=1 pnpm test             # LibreOffice/poppler render tests FAIL instead of skipping when the tools are missing (CI)
pnpm pack-check [--no-build]           # pack, publint + attw, install in a temp dir, every src/index.ts export present, Highcharts not bundled
pnpm examples:typecheck                # tsc -p examples/tsconfig.json (shims in examples/shims.d.ts)
```

CI (`.github/workflows/ci.yml`) runs typecheck, build, examples:typecheck, demo:build, coverage (with `REQUIRE_RENDER=1` after installing LibreOffice + poppler), test:e2e and pack-check on Node 22 and 24, and uploads workbooks, renders and e2e output. `prepublishOnly` runs `pnpm check && pnpm pack-check`; a `v*` tag runs `.github/workflows/release.yml` (check, pack-check, `pnpm publish --provenance` with `NPM_TOKEN`). Record user-facing changes in `CHANGELOG.md`.

## Pipeline and folder → layer map

```
chart | options → extract (src/highcharts) → ChartModel IR (src/types/chart-model.ts)
  → hooks.transformModel → applyThemeOverrides (src/core/theme-overrides.ts)
  → translateChartModel (src/core: chart-type-registry → data-layout → style-mapping)
  → WorkbookSpec (src/excel/writer-interface.ts) → OoxmlExcelWriter (src/excel) → Uint8Array
  → triggerDownload (src/browser)
```

- `src/api/export.ts`: orchestration, strict mode, `ExportError` wrapping, multi-chart workbooks.
- `src/highcharts/`: the only code that knows Highcharts. Reads live chart internals (`series.points`, `getColumn`, `userMin`, `plotLeft`…) and options. `install-export-menu.ts` takes the Highcharts namespace as an argument and never imports `highcharts`.
- `src/translators/`: pure functions (colors, fonts, dash and marker styles, number formats), each returning value + optional diagnostic.
- `src/core/`: Highcharts-free and OOXML-free decisions. `chart-type-registry.ts` holds `CHART_TYPE_MATRIX` (the README table) and the combo, pie and axis rules.
- `src/excel/`: knows only `WorkbookSpec`. Chart XML must follow the ECMA-376 element order (`chart-xml.ts`), or Excel drops the chart.
- **The IR in `src/types/chart-model.ts` is the contract between layers.** Keep it JSON-serializable. Change it deliberately and update both sides.

## Diagnostics

Codes, outcomes and severities live in `src/types/diagnostics.ts` (`DiagnosticCode`, `buildCompatibilityReport`, the `dataConcerns` set). Raise them with `createDiagnostic`/`DiagnosticCollector.report` (both public exports, together with `buildCompatibilityReport`). `APPROXIMATED_COLOR`, `APPROXIMATED_FONT_SIZE` and `APPROXIMATED_DASH_STYLE` are reserved (declared, not raised). The collector dedupes on code + property + series. Blocking diagnostics always become `CHART_NOT_EDITABLE`. Error codes are in `src/types/public-api.ts`. New codes need a README and `docs/compatibility.md` entry.

## Highcharts versions in tests

- Three versions are installed: `highcharts` (13.1.1), `highcharts12` (npm alias of 12.6.2) and `highcharts11` (npm alias of 11.4.9). Tested against 11.4.9, 12.6.2 and 13.1.1.
- **One version per test file. Never import two of `highcharts`, `highcharts12`, `highcharts11` in the same file.** Shared suites live in `*.shared.ts` and are run from `*.v11.test.ts` / `*.v12.test.ts` / `*.v13.test.ts`.
- v12/v13 modules are side-effect imports placed right after their core: `import 'highcharts/modules/exporting'` (or `'highcharts12/modules/exporting'`). They are not factories. **v11 modules are factories**: `import Exporting from 'highcharts11/modules/exporting'; Exporting(Highcharts);` (see the `*.v11.test.ts` files).
- `tests/helpers/render-chart.ts` renders into jsdom with animation and accessibility off.
- jsdom cannot compute CSS, so styled-mode and font-measurement paths behave as "headless" (`isRealBrowser()` is false under jsdom).

## Writer decision

`@office-kit/xlsx` 0.24.1 was evaluated and rejected as the writer: no combo charts or secondary axes, a broken date axis, nulls cached as the text `"null"`, and out-of-order `logBase`. The library ships its own narrowly scoped OOXML writer on `fflate` behind `ExcelWriter` (users can pass their own via `ExportOptions.writer`). office-kit is a **dev-only** round-trip validator (`tests/integration/writer-render.test.ts`). Do not add runtime dependencies.

## Validation caveat

Automated checks:

- structural XLSX inspection (`tests/helpers/inspect-xlsx.ts`)
- the office-kit round-trip
- LibreOffice renders (`tests/output/writer/`, `tests/output/render/`, skipped when `soffice` is missing)
- Playwright downloads

**None of these is Microsoft Excel, and the files were never opened in Excel here.** Do not claim Excel compatibility beyond that. Point to `docs/manual-qa.md`.

## Conventions

- Never mutate the user's chart or options (tests assert this). `transformModel` returns a new model.
- Never fabricate data: missing values stay empty cells and are reported.
- All cell text is written as inline strings (formula-injection safe). Keep it that way.
- Do not edit `package.json` scripts without updating the README Development section.
