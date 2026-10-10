# Contributing

Thanks for helping make Highcharts charts editable in Excel. This page covers the workflow; the
technical details (architecture, how to add a chart type or a translator, the writer decision) are in
the [Development](README.md#development) section of the README and in `CLAUDE.md`.

## Ways to contribute

- **Report a bug** with the smallest chart configuration that reproduces it, the versions involved and
  the `warnings` array of the export result. Use the bug report template.
- **Test in Microsoft Excel.** The test suite validates the output structurally, round-trips it through
  `@office-kit/xlsx` and renders it in LibreOffice, but it cannot run Excel. Opening the workbooks from
  `docs/manual-qa.md` in Excel for Windows, Excel for Mac or Excel Online and reporting what you see is
  one of the most valuable contributions.
- **Improve fidelity.** Every `APPROXIMATED_*` or `UNSUPPORTED_*` diagnostic is a candidate: if Excel
  can express the feature, it can be translated.
- **Sponsor the project** if your company depends on it; see [Support the project](README.md#support-the-project).

## Development setup

```bash
git clone https://github.com/andresballenf/highcharts-to-excel.git
cd highcharts-to-excel
pnpm install          # Node >= 22, pnpm 10 (corepack enable)
pnpm demo             # Vite demo on http://127.0.0.1:4173, serves the library from src/
pnpm check            # typecheck + build + tests
```

Optional tools: LibreOffice and poppler (`brew install --cask libreoffice && brew install poppler`)
enable the render tests, which skip when the tools are missing. The visual-regression comparison
(`tests/integration/visual-regression.test.ts`) stays skipped outside CI even with the tools
installed, because its baselines are rendered on the CI image and other fonts fail the pixel limit;
run it on purpose with `VISUAL_REGRESSION=1 pnpm test:visual`. Playwright's Chromium
(`pnpm exec playwright install chromium`) enables `pnpm test:e2e`.

## Pull requests

1. Open an issue first for anything larger than a bug fix, so the approach can be agreed on.
2. Keep the change focused. Refactors go in their own pull request.
3. Add or update tests. Chart-type and translator changes need a test per supported Highcharts
   version (11, 12 and 13 are installed as `highcharts11`, `highcharts12` and `highcharts`).
4. Run `pnpm lint` (Biome) and `pnpm check` before pushing. CI runs the same plus the e2e tests,
   coverage thresholds and `pnpm pack-check`.
5. Add a changeset with `pnpm changeset`. Pick `patch` for fixes, `minor` for new features or
   changes to experimental surfaces, and `major` only for breaking changes to the stable surface
   (see "Stability and versioning" in the README). Do not edit `CHANGELOG.md` by hand.
6. Document user-visible changes: README, `docs/compatibility.md` for chart types and diagnostic
   codes, `docs/chartjs.md` for the Chart.js adapter.

Be precise about Excel. Unless you opened the file in Microsoft Excel, say "validates" or "renders
in LibreOffice", not "works in Excel".

## Releases

Releases are automated with [Changesets](https://github.com/changesets/changesets). Merging pull
requests with changesets into `main` opens or refreshes a `chore: release` pull request; merging that
pull request publishes to npm through trusted publishing (OIDC, no stored token), with a provenance
attestation, and creates the GitHub release and tag.

## Code of conduct

This project follows the [Contributor Covenant](CODE_OF_CONDUCT.md). By participating you agree to
uphold it.

## License

By contributing you agree that your contributions are licensed under the project's [MIT License](LICENSE).
