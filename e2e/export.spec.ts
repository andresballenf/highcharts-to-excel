/**
 * Browser tests for the demo (demo/): real Highcharts in Chromium, real downloads, and the
 * downloaded .xlsx files validated with tests/helpers/inspect-xlsx.ts.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test as base, expect, type Locator, type Page } from '@playwright/test';
import { inspectXlsx, type XlsxInspection } from '../tests/helpers/inspect-xlsx';
import {
  EXCEL_MENU_ICON,
  allMenuItems,
  buttonSymbol,
  chartContainer,
  chartRoot,
  contextButton,
  contextMenu,
  menuItems as chartMenuItems,
  menuList,
} from './helpers';

const OUTPUT_DIR = fileURLToPath(new URL('./output/', import.meta.url));
const SCREENSHOT_DIR = path.join(OUTPUT_DIR, 'screenshots');
fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });

const EXCEL_MENU_TEXT = 'Download editable Excel chart';
/** demo/src/charts.ts GERMAN_MENU_TEXT, set by the branded card's "Deutsch" button. */
const GERMAN_MENU_TEXT = 'Als Excel-Diagramm herunterladen';
/** demo/src/charts.ts BRAND_COLOR (#1d6f42) as computed CSS. */
const BRAND_RGB = 'rgb(29, 111, 66)';

/** Every test collects console errors and page errors; any of them fails the test at teardown. */
const test = base.extend<{ pageErrors: string[] }>({
  pageErrors: [
    async ({ page }, use) => {
      const errors: string[] = [];
      page.on('console', (msg) => {
        if (msg.type() === 'error') errors.push(`console.error: ${msg.text()}`);
      });
      page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
      await use(errors);
      expect(errors, 'console errors and uncaught page errors').toEqual([]);
    },
    { auto: true },
  ],
});

async function openDemo(page: Page): Promise<void> {
  await page.goto('/');
  await expect(page).toHaveTitle('highcharts-editable-excel demo');
  await page.waitForFunction(() => window.__demo?.ready === true);
  await chartRoot(chartContainer(page, 'stacked')).waitFor();
}

/** Runs `action`, waits for the download it triggers, saves a copy and returns the parsed workbook. */
async function captureDownload(
  page: Page,
  action: () => Promise<void>,
  saveAs?: string,
): Promise<{ filename: string; x: XlsxInspection }> {
  const [download] = await Promise.all([page.waitForEvent('download'), action()]);
  const filename = download.suggestedFilename();
  expect(filename).toMatch(/\.xlsx$/);
  const file = await download.path();
  const bytes = new Uint8Array(fs.readFileSync(file));
  fs.writeFileSync(path.join(OUTPUT_DIR, saveAs ?? filename), bytes);
  const x = await inspectXlsx(bytes);
  x.assertWellFormed();
  return { filename, x };
}

function menuItems(page: Page, name: string): Locator {
  return chartMenuItems(chartContainer(page, name));
}

async function openMenu(page: Page, name: string): Promise<string[]> {
  const chart = chartContainer(page, name);
  await contextButton(chart).click();
  await expect(contextMenu(chart)).toBeVisible();
  return (await chartMenuItems(chart).allInnerTexts()).map((t) => t.trim());
}

function rowCount(x: XlsxInspection, sheet = 'Data'): number {
  return (x.text(x.sheetPath(sheet)).match(/<row\b/g) ?? []).length;
}

/** First srgbClr in each `<c:ser>`'s shape properties = the series fill (columns) or line color. */
function seriesColors(chartXml: string): string[] {
  return [...chartXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map((m) => {
    const spPr = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(m[1]!)?.[1] ?? '';
    return /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(spPr)?.[1]?.toUpperCase() ?? '';
  });
}

function drawingXml(x: XlsxInspection): string {
  const drawing = x.parts.find((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p));
  expect(drawing, 'drawing part').toBeDefined();
  return x.text(drawing!);
}

test('the page loads every card without errors', async ({ page }) => {
  await openDemo(page);
  const names = await page.evaluate(() => Object.keys(window.__demo.charts));
  expect(names.sort()).toEqual(
    [
      'area',
      'bar',
      'branded',
      'column',
      'combo',
      'custom',
      'datetime',
      'doughnut',
      'dynamic',
      'line',
      'multi-line',
      'pie',
      'scatter',
      'secondary',
      'stacked',
      'styled',
    ].sort(),
  );
  for (const name of names) {
    await expect(chartRoot(chartContainer(page, name))).toBeVisible();
    await expect(page.getByTestId(`export-${name}`)).toHaveText('Export to Excel');
    await expect(page.getByTestId(`analyze-${name}`)).toHaveText('Analyze');
    await expect(page.getByTestId(`warnings-${name}`)).toBeAttached();
  }
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, 'gallery-desktop.png'), fullPage: true });
});

test('the export menu gains exactly one Excel item and keeps the original items', async ({ page }) => {
  await openDemo(page);

  const line = await openMenu(page, 'line');
  expect(line.filter((t) => t === EXCEL_MENU_TEXT)).toHaveLength(1);
  expect(line).toContain('Download PNG image');
  expect(line).toContain('View in full screen');
  expect(line).toContain('Download SVG vector image');
  await page.locator('#chart-line').screenshot({ path: path.join(SCREENSHOT_DIR, 'menu-chart-line.png') });

  const pie = await openMenu(page, 'pie');
  expect(pie.filter((t) => t === EXCEL_MENU_TEXT)).toHaveLength(1);
  expect(pie).toContain('Show total (custom item)');
  expect(pie).toContain('Download PNG image');
  expect(pie).toContain('View in full screen');

  const bar = await openMenu(page, 'bar');
  expect(bar).toContain('Save as Excel chart');
  expect(bar).not.toContain(EXCEL_MENU_TEXT);
  expect(bar.filter((t) => t === 'Save as Excel chart')).toHaveLength(1);
  expect(bar).toContain('Download PNG image');

  fs.writeFileSync(path.join(OUTPUT_DIR, 'menu-items.json'), JSON.stringify({ line, pie, bar }, null, 2));
});

test('clicking the menu item downloads a native, linked line chart', async ({ page }) => {
  await openDemo(page);
  await openMenu(page, 'line');
  const { filename, x } = await captureDownload(
    page,
    () => menuItems(page, 'line').filter({ hasText: EXCEL_MENU_TEXT }).click(),
    'line.xlsx',
  );
  expect(filename).toBe('Monthly sales.xlsx');
  expect(x.chartPaths()).toHaveLength(1);
  const chartXml = x.chartXml(0);
  expect(chartXml).toContain('<c:lineChart>');
  expect(chartXml).toMatch(/<c:f>Data!\$B\$2:\$B\$7<\/c:f>/);
  expect(x.seriesFormulas(x.chartPaths()[0]!)).toEqual([
    { name: 'Data!$B$1', cat: 'Data!$A$2:$A$7', val: 'Data!$B$2:$B$7' },
  ]);
  expect(x.hasImages()).toBe(false);
  expect(x.parts.some((p) => p.startsWith('xl/media/'))).toBe(false);
  // The menu path reports through onExport into the same panel.
  await expect(page.getByTestId('warnings-line').locator('pre')).toContainText('"editable": true');
});

test('the programmatic Export button downloads a column chart and fills the warnings panel', async ({ page }) => {
  await openDemo(page);
  const { filename, x } = await captureDownload(page, () => page.getByTestId('export-column').click());
  expect(filename).toBe('column.xlsx');
  expect(x.chartXml(0)).toContain('<c:barChart>');
  expect(x.chartXml(0)).toContain('<c:barDir val="col"/>');
  const panel = page.getByTestId('warnings-column');
  await expect(panel).toHaveAttribute('open', '');
  await expect(panel.locator('pre')).toContainText('"editable": true');
  await expect(panel.locator('pre')).toContainText('"excelChartType"');
  // Every warning line is rendered as "code — property — message".
  const warnings = await page.evaluate(
    () =>
      (window.__demo.lastResults.column as { warnings: Array<{ code: string; property: string; message: string }> })
        .warnings,
  );
  const lines = await panel.locator('li').allInnerTexts();
  if (warnings.length > 0) expect(lines).toEqual(warnings.map((w) => `${w.code} — ${w.property} — ${w.message}`));
  await page
    .locator('article[data-chart="column"]')
    .screenshot({ path: path.join(SCREENSHOT_DIR, 'warnings-panel-column.png') });

  // Analyze renders a report without downloading.
  await page.getByTestId('analyze-bar').click();
  await expect(page.getByTestId('warnings-bar').locator('pre')).toContainText('"kind": "analysis"');
  await expect(page.getByTestId('warnings-bar').locator('pre')).toContainText('"editable": true');
});

test('includeReferenceImage adds a PNG next to the native, still-linked chart', async ({ page }) => {
  await openDemo(page);
  const { bytes } = await page.evaluate(() => window.__demo.exportWithImage('column'));
  const data = new Uint8Array(bytes);
  fs.writeFileSync(path.join(OUTPUT_DIR, 'column-with-image.xlsx'), data);
  const x = await inspectXlsx(data);
  x.assertWellFormed();
  expect(x.hasImages()).toBe(true);
  expect(x.parts).toContain('xl/media/image1.png');
  // PNG signature: 89 50 4E 47 0D 0A 1A 0A.
  expect(Array.from(x.bytes('xl/media/image1.png').subarray(0, 8))).toEqual([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
  ]);
  expect(x.chartPaths()).toHaveLength(1);
  expect(x.chartXml(0)).toContain('<c:barChart>');
  expect(x.chartXml(0)).toMatch(/<c:f>Data!\$B\$2:\$B\$\d+<\/c:f>/);
});

test('several charts on one page export independently through their menus', async ({ page }) => {
  await openDemo(page);
  const expected: Record<string, string> = { line: 'lineChart', column: 'barChart', pie: 'pieChart' };
  for (const [name, group] of Object.entries(expected)) {
    await openMenu(page, name);
    const { x } = await captureDownload(
      page,
      () => menuItems(page, name).filter({ hasText: EXCEL_MENU_TEXT }).click(),
      `multi-${name}.xlsx`,
    );
    expect(x.plotGroupKinds(x.chartPaths()[0]!)).toEqual([group]);
  }
  // Two menus open in turn: each lists the Excel item exactly once.
  for (const name of ['line', 'column']) {
    await openMenu(page, name);
    await expect(menuItems(page, name).filter({ hasText: EXCEL_MENU_TEXT })).toHaveCount(1);
  }
  await expect(allMenuItems(page).filter({ hasText: EXCEL_MENU_TEXT })).toHaveCount(3);
});

test('dynamic data: added points and setData are what gets exported', async ({ page }) => {
  await openDemo(page);
  const before = await captureDownload(page, () => page.getByTestId('export-dynamic').click(), 'dynamic-before.xlsx');
  const originalRows = rowCount(before.x);
  expect(originalRows).toBe(1 + 6);

  await page.getByTestId('add-point-chart-dynamic').click();
  await page.getByTestId('add-point-chart-dynamic').click();
  const added = await captureDownload(page, () => page.getByTestId('export-dynamic').click(), 'dynamic-added.xlsx');
  expect(rowCount(added.x)).toBe(originalRows + 2);
  expect(added.x.seriesFormulas(added.x.chartPaths()[0]!)[0]?.val).toBe('Data!$B$2:$B$9');

  await page.getByTestId('randomize-chart-dynamic').click();
  const ys = await page.evaluate(() => window.__demo.charts.dynamic!.series[0]!.points.map((p) => p.y));
  const randomized = await captureDownload(
    page,
    () => page.getByTestId('export-dynamic').click(),
    'dynamic-randomized.xlsx',
  );
  const data = randomized.x.sheetPath('Data');
  const exported = ys.map((_, i) => randomized.x.cellValue(data, `B${i + 2}`));
  expect(exported).toEqual(ys);
  expect(randomized.x.cellValue(data, `B${ys.length + 2}`)).toBeNull();
  // Deterministic sequence: the randomized values differ from the initial data.
  expect(ys.slice(0, 6)).not.toEqual([12, 15, 11, 18, 16, 21]);
});

test('styled mode: series colors come from the CSS variables computed by the browser', async ({ page }) => {
  await openDemo(page);
  // The values as written in the demo's <style id="styled-mode-theme"> block.
  const css = await page.locator('#styled-mode-theme').innerHTML();
  const defined = [0, 1, 2].map((i) =>
    new RegExp(`--highcharts-color-${i}:\\s*#([0-9a-fA-F]{6})`).exec(css)?.[1]?.toUpperCase(),
  );
  expect(defined.every(Boolean)).toBe(true);

  const { x } = await captureDownload(page, () => page.getByTestId('export-styled').click());
  expect(seriesColors(x.chartXml(0))).toEqual(defined);

  const codes = await page.evaluate(() =>
    (window.__demo.lastResults.styled as { warnings: Array<{ code: string }> }).warnings.map((w) => w.code),
  );
  expect(codes).not.toContain('STYLED_MODE_FALLBACK');
  fs.writeFileSync(path.join(OUTPUT_DIR, 'styled-mode-colors.json'), JSON.stringify(defined));
});

test('custom styling survives: font, dashed gridlines, data labels, legend, size', async ({ page }) => {
  await openDemo(page);
  const { x } = await captureDownload(page, () => page.getByTestId('export-custom').click());
  const chartXml = x.chartXml(0);
  expect(chartXml).toContain('<a:latin typeface="Georgia"/>');
  expect(chartXml).toMatch(/<a:prstDash val="(dash|sysDash|lgDash)"\/>/);
  expect(chartXml).toContain('<c:showVal val="1"/>');
  expect(chartXml).toContain('<c:legendPos val="r"/>');
  expect(seriesColors(chartXml)).toEqual(['C0392B', '2C3E50']);
  // 800 CSS px × 9525 EMU/px.
  expect(drawingXml(x)).toContain('<xdr:ext cx="7620000"');
});

test('export all: one workbook with a chart and a data sheet per chart', async ({ page }) => {
  await openDemo(page);
  const n = await page.evaluate(() => Object.keys(window.__demo.charts).length);
  const { filename, x } = await captureDownload(page, () => page.getByTestId('export-all').click());
  expect(filename).toBe('all-charts.xlsx');
  expect(x.sheetNames()).toHaveLength(2 * n);
  expect(x.chartPaths()).toHaveLength(n);
  expect(x.hasImages()).toBe(false);
  await expect(page.getByTestId('export-all-status')).toContainText(`${n} charts`);
});

/** The `d` of the context-button symbol of a throwaway chart that keeps Highcharts' default 'menu'. */
async function defaultButtonPath(page: Page): Promise<string> {
  return page.evaluate(() => {
    const H = window.__demo.highcharts;
    const div = document.body.appendChild(document.createElement('div'));
    div.style.width = '400px';
    const chart = H.chart(div, {
      series: [{ type: 'line', data: [1, 2] }],
      exporting: { buttons: { contextButton: { symbol: 'menu' } } },
    });
    const d = chart.container.querySelector('.highcharts-contextbutton .highcharts-button-symbol')?.getAttribute('d');
    chart.destroy();
    div.remove();
    return d ?? '';
  });
}

test('branded export button and menu: custom symbol, Excel icon, menu styles, translated text', async ({ page }) => {
  await openDemo(page);
  const chart = chartContainer(page, 'branded');
  await chart.scrollIntoViewIfNeeded();

  // The context button draws the registered download-arrow symbol in the brand color.
  const symbol = buttonSymbol(chart);
  await expect(symbol).toHaveAttribute('stroke', '#1d6f42');
  await expect(symbol).toHaveAttribute('fill', '#1d6f42');
  const brandedD = await symbol.getAttribute('d');
  const defaultD = await defaultButtonPath(page);
  expect(defaultD).not.toBe('');
  expect(brandedD).toBeTruthy();
  expect(brandedD).not.toBe(defaultD);

  // The open menu: one Excel item with the built-in SVG icon, styled by menuStyle / menuItemStyle.
  const items = await openMenu(page, 'branded');
  expect(items.filter((t) => t === EXCEL_MENU_TEXT)).toHaveLength(1);
  const excelItem = menuItems(page, 'branded').filter({ hasText: EXCEL_MENU_TEXT });
  await expect(excelItem.locator(EXCEL_MENU_ICON)).toHaveCount(1);
  await expect(excelItem.locator(EXCEL_MENU_ICON)).toBeVisible();
  expect(await excelItem.locator(`${EXCEL_MENU_ICON} path`).count()).toBeGreaterThan(0);
  expect(await excelItem.locator('.highcharts-editable-excel-label').textContent()).toBe(EXCEL_MENU_TEXT);
  const itemCss = await excelItem.evaluate((li) => {
    const cs = getComputedStyle(li);
    return { fontSize: cs.fontSize, paddingLeft: cs.paddingLeft };
  });
  expect(itemCss).toEqual({ fontSize: '13px', paddingLeft: '14px' });
  const listCss = await menuList(chart).evaluate((ul) => {
    const cs = getComputedStyle(ul);
    return { borderColor: cs.borderTopColor, borderWidth: cs.borderTopWidth, radius: cs.borderTopLeftRadius };
  });
  expect(listCss).toEqual({ borderColor: BRAND_RGB, borderWidth: '1px', radius: '6px' });
  await excelItem.hover();
  await expect.poll(() => excelItem.evaluate((li) => getComputedStyle(li).color)).toBe(BRAND_RGB);
  await page
    .locator('article[data-chart="branded"]')
    .screenshot({ path: path.join(SCREENSHOT_DIR, 'branded-menu-open.png') });

  // Exports still download a valid, native workbook.
  const english = await captureDownload(page, () => excelItem.click(), 'branded.xlsx');
  expect(english.x.chartXml(0)).toContain('<c:barChart>');
  expect(english.x.hasImages()).toBe(false);

  // langKey: the "Deutsch" button sets lang.downloadEditableXLSX and re-creates the chart.
  await page.getByTestId('lang-de').click();
  const german = await openMenu(page, 'branded');
  expect(german).toContain(GERMAN_MENU_TEXT);
  expect(german).not.toContain(EXCEL_MENU_TEXT);
  const germanItem = menuItems(page, 'branded').filter({ hasText: GERMAN_MENU_TEXT });
  await expect(germanItem.locator(EXCEL_MENU_ICON)).toHaveCount(1);
  const translated = await captureDownload(page, () => germanItem.click(), 'branded-de.xlsx');
  expect(translated.x.chartXml(0)).toContain('<c:barChart>');
  // Charts not re-created keep the text they were created with.
  expect(await openMenu(page, 'line')).toContain(EXCEL_MENU_TEXT);
});

test.describe('mobile viewport', () => {
  test.use({ viewport: { width: 375, height: 812 }, isMobile: true, hasTouch: true });

  test('renders without horizontal scrolling and still exports', async ({ page }) => {
    await openDemo(page);
    await expect(page.locator('article[data-chart="line"]')).toBeVisible();
    await expect(chartRoot(chartContainer(page, 'line'))).toBeVisible();
    const { scrollWidth, innerWidth } = await page.evaluate(() => ({
      scrollWidth: document.documentElement.scrollWidth,
      innerWidth: window.innerWidth,
    }));
    expect(scrollWidth).toBeLessThanOrEqual(innerWidth);
    await page.screenshot({ path: path.join(SCREENSHOT_DIR, 'gallery-mobile.png') });

    const { x } = await captureDownload(page, () => page.getByTestId('export-line').click(), 'line-mobile.xlsx');
    expect(x.chartXml(0)).toContain('<c:lineChart>');
  });

  test('the branded menu opens and fits at 375px', async ({ page }) => {
    await openDemo(page);
    const chart = chartContainer(page, 'branded');
    await chart.scrollIntoViewIfNeeded();
    await contextButton(chart).click();
    await expect(contextMenu(chart)).toBeVisible();
    const excelItem = menuItems(page, 'branded').filter({ hasText: EXCEL_MENU_TEXT });
    await expect(excelItem.locator(EXCEL_MENU_ICON)).toBeVisible();
    const box = await menuList(chart).boundingBox();
    expect(box).not.toBeNull();
    expect(box!.x).toBeGreaterThanOrEqual(0);
    expect(box!.x + box!.width).toBeLessThanOrEqual(375);
    const scrollWidth = await page.evaluate(() => document.documentElement.scrollWidth);
    expect(scrollWidth).toBeLessThanOrEqual(375);
    await page.screenshot({ path: path.join(SCREENSHOT_DIR, 'branded-mobile.png') });
    const { x } = await captureDownload(page, () => excelItem.click(), 'branded-mobile.xlsx');
    expect(x.chartXml(0)).toContain('<c:barChart>');
  });
});
