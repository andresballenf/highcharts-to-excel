# Browser tests (Playwright)

`e2e/export.spec.ts` drives the demo (`demo/`) in Chromium: it opens the export menus, clicks the
real buttons, captures the real downloads and validates every `.xlsx` with
`tests/helpers/inspect-xlsx.ts` (well-formed XML, chart groups, `<c:f>` formulas into the data
sheet, styling, row counts). Every test also fails on any `console.error` or uncaught page error.

```sh
pnpm test:e2e      # starts `pnpm demo` on http://127.0.0.1:4173 automatically
```

The first run needs a browser: `pnpm exec playwright install chromium` (CI does this).
In the Claude Code cloud sandbox the bundled download is unavailable; point Playwright at the
preinstalled Chromium instead:

```sh
PW_CHROMIUM_EXECUTABLE=/opt/pw-browsers/chromium pnpm test:e2e
```

Outputs go to `e2e/output/` (gitignored): a copy of each downloaded workbook, `menu-items.json`,
`styled-mode-colors.json`, `screenshots/` (desktop gallery, open menu, warnings panel, 375 px
gallery) and `test-results/` traces for failures.
