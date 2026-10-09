/**
 * Selectors for the Highcharts DOM the e2e tests drive. Highcharts' class names live here only, so
 * a rename in a future Highcharts version is a one-file change.
 */

import type { Locator, Page } from '@playwright/test';

export const HC = {
  root: '.highcharts-root',
  contextButton: '.highcharts-contextbutton',
  contextMenu: '.highcharts-contextmenu',
  menuItem: '.highcharts-menu-item',
} as const;

/** The demo card's chart container (`#chart-<name>`). */
export function chartContainer(page: Page, name: string): Locator {
  return page.locator(`#chart-${name}`);
}

/** The rendered SVG root of a chart. */
export function chartRoot(chart: Locator): Locator {
  return chart.locator(HC.root);
}

/** The ☰ export menu button of a chart. */
export function contextButton(chart: Locator): Locator {
  return chart.locator(HC.contextButton);
}

/** The open export menu of a chart. */
export function contextMenu(chart: Locator): Locator {
  return chart.locator(HC.contextMenu);
}

/** The items of a chart's open export menu. */
export function menuItems(chart: Locator): Locator {
  return chart.locator(`${HC.contextMenu} ${HC.menuItem}`);
}

/** Every export-menu item on the page, across charts. */
export function allMenuItems(page: Page): Locator {
  return page.locator(HC.menuItem);
}
