---
"highcharts-editable-excel": minor
---

Customize the export button and menu: `button` (symbol, custom `svgPath`, fill/stroke, theme, className, title), `menuIcon` (`excel`, custom svg/html), `menuStyle`/`menuItemStyle`/`menuItemHoverStyle`, and i18n via `langKey` (`lang.downloadEditableXLSX`). Per-chart `exporting.editableExcel.menuIcon`. Uninstall restores everything it changed.

Behavior change: without `menuText`, the global menu item definition has no `text`; its text comes from `lang.downloadEditableXLSX` (registered with the default text at install when absent). Text order: per-chart `menuText` → install `menuText` → `lang[langKey]` → `DEFAULT_MENU_TEXT`.
