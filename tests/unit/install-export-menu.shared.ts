/**
 * Export-menu integration assertions shared by the per-version test files. Each version file
 * imports its own Highcharts build (+ exporting and highcharts-more) and calls `runInstallSuite`.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { installHighchartsExcelExport, addEditableExcelMenuItem } from '../../src/highcharts/install-export-menu';
import {
  DEFAULT_BUTTON_SYMBOL,
  DEFAULT_LANG_KEY,
  DEFAULT_MENU_ITEM_KEY,
  DEFAULT_MENU_TEXT,
  ExportError,
  MENU_ICON_EXCEL_SVG,
  type ExportResult,
  type Installation,
} from '../../src/types/public-api';
import * as F from '../fixtures/highcharts-options';
import { destroyAll, renderChart, type HighchartsLike } from '../helpers/render-chart';

// biome-ignore lint/suspicious/noExplicitAny: tests read live Highcharts internals that the public types omit
type AnyRec = any;

export interface HighchartsNamespace extends HighchartsLike {
  getOptions(): AnyRec;
  setOptions(o: AnyRec): AnyRec;
  addEvent: (...args: AnyRec[]) => () => void;
  Chart: AnyRec;
}

const KEY = DEFAULT_MENU_ITEM_KEY;

function globalItems(H: HighchartsNamespace): unknown[] {
  return H.getOptions().exporting.buttons.contextButton.menuItems as unknown[];
}

function chartItems(chart: AnyRec): unknown[] {
  return chart.options.exporting.buttons.contextButton.menuItems as unknown[];
}

function count(items: unknown[], key: string): number {
  return items.filter((i) => i === key).length;
}

function definition(H: HighchartsNamespace): AnyRec {
  return H.getOptions().exporting.menuItemDefinitions[KEY];
}

/** The markup Highcharts renders for a menu item: `item.text || chart.options.lang[item.textKey]`. */
function itemMarkup(chart: AnyRec, key = KEY): string {
  const def = chart.options.exporting.menuItemDefinitions[key];
  return def.text || chart.options.lang[def.textKey];
}

/** Renders a menu item's markup the way Highcharts does (AST-sanitized) into a detached <li>. */
function renderItem(H: HighchartsNamespace, chart: AnyRec, key = KEY): HTMLLIElement {
  const li = document.createElement('li');
  (H as AnyRec).AST.setElementHTML(li, itemMarkup(chart, key));
  return li;
}

/** The plain-text label a user sees (textContent of the rendered <li>). */
function itemLabel(H: HighchartsNamespace, chart: AnyRec, key = KEY): string {
  return renderItem(H, chart, key).textContent ?? '';
}

/** Deep clone of plain option data (functions kept by reference). */
function snap<T>(value: T): T {
  if (Array.isArray(value)) return value.map((v: unknown) => snap(v)) as T;
  if (value !== null && typeof value === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) out[k] = snap(v);
    return out as T;
  }
  return value;
}

/** The global option subtrees the installer touches. */
function globalState(H: HighchartsNamespace): AnyRec {
  const o = H.getOptions();
  return snap({
    contextButton: o.exporting.buttons.contextButton,
    definitionKeys: Object.keys(o.exporting.menuItemDefinitions),
    navigation: o.navigation,
    lang: o.lang,
  });
}

function symbols(H: HighchartsNamespace): AnyRec {
  const h = H as AnyRec;
  return (h.SVGRenderer ?? h.Renderer).prototype.symbols;
}

/** Waits for the next onExport/onError callback. */
function settle() {
  let resolve!: () => void;
  const done = new Promise<void>((r) => (resolve = r));
  const onExport = vi.fn((_result: ExportResult, _chart: unknown) => resolve());
  const onError = vi.fn((_error: unknown, _chart: unknown) => resolve());
  return { onExport, onError, done };
}

export function runInstallSuite(H: HighchartsNamespace, label: string): void {
  describe(`installHighchartsExcelExport (${label})`, () => {
    let installation: Installation | null = null;
    let originalItems: unknown[];
    let originalState: AnyRec;
    const install = (
      ...args: Parameters<typeof installHighchartsExcelExport> extends [unknown, ...infer R] ? R : never
    ): Installation => (installation = installHighchartsExcelExport(H, ...args));

    beforeEach(() => {
      originalItems = globalItems(H);
      originalState = globalState(H);
      vi.stubGlobal(
        'URL',
        Object.assign(Object.create(URL) as object, {
          createObjectURL: vi.fn(() => 'blob:x'),
          revokeObjectURL: vi.fn(),
        }),
      );
      vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
    });

    afterEach(() => {
      installation?.uninstall();
      installation = null;
      destroyAll();
      vi.restoreAllMocks();
      vi.unstubAllGlobals();
      // Leave the shared Highcharts defaults exactly as found.
      H.setOptions({ exporting: { buttons: { contextButton: { menuItems: originalItems } } } });
      delete H.getOptions().exporting.menuItemDefinitions.myItem;
      // Lang keys a test set itself (translations registered before install are the app's).
      const o = H.getOptions();
      for (const k of Object.keys(o.lang)) if (!(k in originalState.lang)) delete o.lang[k];
      Object.assign(o.lang, snap(originalState.lang));
      o.navigation = snap(originalState.navigation);
      o.exporting.buttons.contextButton = snap(originalState.contextButton);
    });

    it('adds the item once and returns the same installation on a second call', () => {
      const a = install();
      const b = installHighchartsExcelExport(H, { menuText: 'ignored' });
      expect(b).toBe(a);
      expect(a.options.menuText).toBeUndefined();
      expect(count(globalItems(H), KEY)).toBe(1);
      // No menuText: the text comes from lang[langKey] through Highcharts' own textKey lookup.
      expect(definition(H).text).toBeUndefined();
      expect(definition(H).textKey).toBe(DEFAULT_LANG_KEY);
      expect(H.getOptions().lang[DEFAULT_LANG_KEY]).toBe(DEFAULT_MENU_TEXT);
    });

    it('preserves the default items in order', () => {
      const before = [...globalItems(H)];
      install();
      const after = globalItems(H);
      expect(after.filter((i) => i !== KEY)).toEqual(before);
      // Inserted after the last of downloadXLS / downloadCSV / downloadSVG.
      const anchor = ['downloadXLS', 'downloadCSV', 'downloadSVG']
        .map((k) => before.indexOf(k))
        .reduce((m, i) => Math.max(m, i), -1);
      expect(after.indexOf(KEY)).toBe(anchor >= 0 ? anchor + 1 : after.length - 1);
    });

    it('keeps an app-registered custom item and honours insertAfter', () => {
      const defaults = [...globalItems(H)];
      H.setOptions({
        exporting: {
          menuItemDefinitions: { myItem: { text: 'Mine', onclick() {} } },
          buttons: { contextButton: { menuItems: [...defaults, 'myItem'] } },
        },
      });
      install({ insertAfter: 'myItem' });
      const items = globalItems(H);
      expect(items.slice(0, defaults.length)).toEqual(defaults);
      expect(items.indexOf(KEY)).toBe(items.indexOf('myItem') + 1);
      expect(H.getOptions().exporting.menuItemDefinitions.myItem.text).toBe('Mine');
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      expect(chartItems(chart)).toContain('myItem');
      expect(count(chartItems(chart), KEY)).toBe(1);
    });

    it('keeps a custom item without insertAfter (key placed after the default download items)', () => {
      const defaults = [...globalItems(H)];
      H.setOptions({
        exporting: {
          menuItemDefinitions: { myItem: { text: 'Mine', onclick() {} } },
          buttons: { contextButton: { menuItems: [...defaults, 'myItem'] } },
        },
      });
      install();
      const items = globalItems(H);
      expect(items).toContain('myItem');
      expect(count(items, KEY)).toBe(1);
      expect(items.filter((i) => i !== KEY)).toEqual([...defaults, 'myItem']);
    });

    it('lists the key exactly once on charts created after install', () => {
      install();
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      const items = chartItems(chart);
      expect(count(items, KEY)).toBe(1);
      expect(itemLabel(H, chart)).toBe(DEFAULT_MENU_TEXT);
    });

    it('per-chart enabled:false removes the key for that chart only', () => {
      install();
      const disabled = renderChart(H, {
        ...F.simpleLine,
        exporting: { editableExcel: { enabled: false } } as AnyRec,
      }) as AnyRec;
      expect(chartItems(disabled)).not.toContain(KEY);
      expect(chartItems(disabled)).toEqual(globalItems(H).filter((i) => i !== KEY));
      const next = renderChart(H, F.simpleLine) as AnyRec;
      expect(count(chartItems(next), KEY)).toBe(1);
      expect(count(globalItems(H), KEY)).toBe(1);
    });

    it('per-chart menuText is applied without changing the global text', () => {
      install();
      const chart = renderChart(H, {
        ...F.simpleLine,
        exporting: { editableExcel: { menuText: 'Excel (editable)' } } as AnyRec,
      }) as AnyRec;
      expect(chart.options.exporting.menuItemDefinitions[KEY].text).toBe('Excel (editable)');
      expect(typeof chart.options.exporting.menuItemDefinitions[KEY].onclick).toBe('function');
      expect(definition(H).text).toBeUndefined();
      expect(itemLabel(H, renderChart(H, F.simpleLine))).toBe(DEFAULT_MENU_TEXT);
    });

    it('re-applies per-chart settings after chart.update({ exporting }) (afterUpdate)', () => {
      install();
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      expect(count(chartItems(chart), KEY)).toBe(1);
      const captured = (): AnyRec => chart.exporting?.options ?? chart.options.exporting;

      chart.update({ exporting: { editableExcel: { enabled: false } } });
      expect(chartItems(chart)).not.toContain(KEY);
      expect(captured().buttons.contextButton.menuItems).not.toContain(KEY);
      expect(count(globalItems(H), KEY)).toBe(1);

      chart.update({ exporting: { editableExcel: { enabled: true } } });
      expect(count(chartItems(chart), KEY)).toBe(1);

      const own = ['downloadPNG', 'printChart'];
      chart.update({ exporting: { buttons: { contextButton: { menuItems: own } } } });
      expect(chartItems(chart)).toEqual(['downloadPNG', 'printChart', KEY]);
      expect(captured().buttons.contextButton.menuItems).toEqual(['downloadPNG', 'printChart', KEY]);
      expect(own).toEqual(['downloadPNG', 'printChart']);

      chart.update({ exporting: { editableExcel: { menuText: 'Excel (updated)' } } });
      expect(chart.options.exporting.menuItemDefinitions[KEY].text).toBe('Excel (updated)');

      // Updates that do not touch exporting leave the menu alone.
      chart.update({ title: { text: 'Retitled' } });
      expect(chartItems(chart)).toEqual(['downloadPNG', 'printChart', KEY]);
    });

    it('the rendered context menu follows chart.update (the button is redrawn)', () => {
      install();
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      const menuTexts = (): string[] => {
        const button = chart.container.querySelector('.highcharts-contextbutton');
        button?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        return [...chart.container.parentNode.querySelectorAll('.highcharts-menu-item')].map(
          (li: Element) => li.textContent ?? '',
        );
      };
      chart.update({ exporting: { buttons: { contextButton: { menuItems: ['downloadPNG'] } } } });
      const shown = menuTexts();
      expect(shown).toEqual(['Download PNG image', DEFAULT_MENU_TEXT]);
    });

    it('unbinds the afterUpdate handler on uninstall', () => {
      const inst = install();
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      inst.uninstall();
      installation = null;
      chart.update({ exporting: { buttons: { contextButton: { menuItems: ['downloadPNG'] } } } });
      expect(chartItems(chart)).toEqual(['downloadPNG']);
    });

    it('appends the key to explicit per-chart menuItems without mutating the caller array', () => {
      install();
      const own = ['downloadPNG', 'printChart'];
      const userOptions = {
        ...F.simpleLine,
        accessibility: { enabled: false },
        exporting: { buttons: { contextButton: { menuItems: own } } },
      };
      const chart = (H as AnyRec).chart(document.body.appendChild(document.createElement('div')), userOptions);
      expect(chartItems(chart)).toEqual(['downloadPNG', 'printChart', KEY]);
      expect(own).toEqual(['downloadPNG', 'printChart']);
      // The exporting module renders the context button from the options object it captured at init
      // (`chart.exporting.options` in v12+); that object must carry the item too, or the browser menu
      // never shows it even though `chart.options` does (regression caught by the Playwright suite).
      const captured = chart.exporting?.options ?? chart.options.exporting;
      expect(captured.buttons.contextButton.menuItems).toContain(KEY);
      expect(captured).toBe(chart.options.exporting);
      chart.destroy();
    });

    it('menu onclick exports and calls onExport with bytes', async () => {
      const s = settle();
      install({ onExport: s.onExport, onError: s.onError, exportOptions: { filename: 'from-install' } });
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      chart.options.exporting.menuItemDefinitions[KEY].onclick.call(chart);
      await s.done;
      expect(s.onError).not.toHaveBeenCalled();
      const [result, receivedChart] = s.onExport.mock.calls[0] as [ExportResult, unknown];
      expect(receivedChart).toBe(chart);
      expect(result.bytes.length).toBeGreaterThan(0);
      expect(result.filename).toBe('from-install.xlsx');
      expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);
    });

    it('per-chart editableExcel export options override install defaults', async () => {
      const s = settle();
      install({ onExport: s.onExport, onError: s.onError, exportOptions: { filename: 'from-install' } });
      const chart = renderChart(H, {
        ...F.simpleLine,
        exporting: { editableExcel: { filename: 'per-chart' } } as AnyRec,
      }) as AnyRec;
      chart.options.exporting.menuItemDefinitions[KEY].onclick.call(chart);
      await s.done;
      expect((s.onExport.mock.calls[0]![0] as ExportResult).filename).toBe('per-chart.xlsx');
    });

    it('routes a polar-column failure to onError as an ExportError', async () => {
      const s = settle();
      install({ onExport: s.onExport, onError: s.onError });
      const chart = renderChart(H, F.polarColumnChart) as AnyRec;
      definition(H).onclick.call(chart);
      await s.done;
      expect(s.onExport).not.toHaveBeenCalled();
      const [error, receivedChart] = s.onError.mock.calls[0] as [unknown, unknown];
      expect(error).toBeInstanceOf(ExportError);
      expect((error as ExportError).code).toBe('CHART_NOT_EDITABLE');
      expect(receivedChart).toBe(chart);
    });

    it('default onError logs to console.error', async () => {
      const spy = vi.spyOn(console, 'error').mockImplementation(() => undefined);
      install();
      const chart = renderChart(H, F.polarColumnChart) as AnyRec;
      definition(H).onclick.call(chart);
      await vi.waitFor(() => expect(spy).toHaveBeenCalled());
      expect(spy.mock.calls[0]?.[1]).toBeInstanceOf(ExportError);
    });

    it('uninstall restores the menu, removes the definition and allows a new install', () => {
      const before = globalItems(H);
      const first = install();
      first.uninstall();
      installation = null;
      expect(globalItems(H)).toBe(before);
      expect(H.getOptions().exporting.menuItemDefinitions[KEY]).toBeUndefined();
      const plain = renderChart(H, {
        ...F.simpleLine,
        exporting: { editableExcel: { menuText: 'x' } } as AnyRec,
      }) as AnyRec;
      expect(chartItems(plain)).not.toContain(KEY);
      expect(plain.options.exporting.menuItemDefinitions[KEY]).toBeUndefined();

      const second = install({ menuText: 'Again' });
      expect(second).not.toBe(first);
      expect(definition(H).text).toBe('Again');
      expect(count(globalItems(H), KEY)).toBe(1);
    });

    it('supports a custom menuItemKey', () => {
      install({ menuItemKey: 'myExcel', menuText: 'XLSX' });
      expect(globalItems(H)).toContain('myExcel');
      expect(H.getOptions().exporting.menuItemDefinitions.myExcel.text).toBe('XLSX');
      installation?.uninstall();
      installation = null;
      expect(H.getOptions().exporting.menuItemDefinitions.myExcel).toBeUndefined();
    });

    it('addEditableExcelMenuItem patches an already-rendered chart', () => {
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      expect(chartItems(chart)).not.toContain(KEY);
      addEditableExcelMenuItem(chart, { menuText: 'Late' });
      expect(count(chartItems(chart), KEY)).toBe(1);
      expect(chart.options.exporting.menuItemDefinitions[KEY].text).toBe('Late');
      addEditableExcelMenuItem(chart);
      expect(count(chartItems(chart), KEY)).toBe(1);
      expect(globalItems(H)).not.toContain(KEY);
    });

    it('throws EXPORTING_MODULE_MISSING for a Highcharts object without exporting', () => {
      const fake = {
        getOptions() {
          return {};
        },
        setOptions() {},
        addEvent() {
          return () => {};
        },
        Chart: function Chart() {},
      };
      expect(() => installHighchartsExcelExport(fake)).toThrow(
        expect.objectContaining({ code: 'EXPORTING_MODULE_MISSING' }),
      );
    });

    it('throws INVALID_OPTIONS for a non-Highcharts argument', () => {
      for (const bad of [undefined, null, 42, {}, { getOptions() {} }]) {
        expect(() => installHighchartsExcelExport(bad)).toThrow(expect.objectContaining({ code: 'INVALID_OPTIONS' }));
      }
    });

    const directChart = (options: AnyRec): AnyRec => {
      const container = document.body.appendChild(document.createElement('div'));
      return (H as AnyRec).chart(container, options);
    };

    it('E2 frozen options still render after install and keep their per-chart settings', () => {
      install();
      const exporting = Object.freeze({ editableExcel: Object.freeze({ menuText: 'Frozen text' }) });
      const frozen = Object.freeze({
        accessibility: Object.freeze({ enabled: false }),
        series: [{ type: 'line', data: [1, 2] }],
        exporting,
      });
      const chart = directChart(frozen);
      expect(frozen.exporting).toBe(exporting);
      expect(count(chartItems(chart), KEY)).toBe(1);
      expect(chart.options.exporting.menuItemDefinitions[KEY].text).toBe('Frozen text');
      chart.destroy();
    });

    it('E2 never assigns into the caller options (reused objects stay identical)', () => {
      install();
      const own = ['downloadPNG'];
      const exporting = { editableExcel: { menuText: 'Shared' }, buttons: { contextButton: { menuItems: own } } };
      const shared = { accessibility: { enabled: false }, series: [{ type: 'line', data: [1, 2] }], exporting };
      const snapshot = JSON.stringify(shared);
      const a = directChart(shared);
      const b = directChart(shared);
      expect(shared.exporting).toBe(exporting);
      expect(JSON.stringify(shared)).toBe(snapshot);
      expect(Object.keys(shared)).toEqual(['accessibility', 'series', 'exporting']);
      for (const c of [a, b]) {
        expect(chartItems(c)).toEqual(['downloadPNG', KEY]);
        expect(c.options.exporting.menuItemDefinitions[KEY].text).toBe('Shared');
        expect(c.userOptions.exporting.buttons.contextButton.menuItems).toEqual(['downloadPNG', KEY]);
      }
      const disabled = directChart({ ...shared, exporting: { editableExcel: { enabled: false } } });
      expect(chartItems(disabled)).not.toContain(KEY);
      for (const c of [a, b, disabled]) c.destroy();
    });

    it('A1 a throwing onExport reaches onError', async () => {
      const errors: unknown[] = [];
      let resolve!: () => void;
      const done = new Promise<void>((r) => (resolve = r));
      install({
        onExport: () => {
          throw new Error('onExport failed');
        },
        onError: (e) => {
          errors.push(e);
          resolve();
        },
      });
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      definition(H).onclick.call(chart);
      await done;
      expect((errors[0] as Error).message).toBe('onExport failed');
    });

    it('A1 a throwing onError is logged, never an unhandled rejection', async () => {
      const spy = vi.spyOn(console, 'error').mockImplementation(() => undefined);
      const unhandled: unknown[] = [];
      const onUnhandled = (reason: unknown): void => {
        unhandled.push(reason);
      };
      process.on('unhandledRejection', onUnhandled);
      try {
        install({
          onError: () => {
            throw new Error('onError failed');
          },
        });
        const chart = renderChart(H, F.polarColumnChart) as AnyRec;
        definition(H).onclick.call(chart);
        await vi.waitFor(() =>
          expect(spy.mock.calls.flat().some((a) => a instanceof Error && a.message === 'onError failed')).toBe(true),
        );
        await new Promise((r) => setTimeout(r, 20));
        expect(unhandled).toEqual([]);
      } finally {
        process.off('unhandledRejection', onUnhandled);
      }
    });

    it('A1 ignores a second click while an export of the same chart is running', async () => {
      const s = settle();
      install({ onExport: s.onExport, onError: s.onError });
      const chart = renderChart(H, F.simpleLine) as AnyRec;
      const click = (): void => definition(H).onclick.call(chart);
      click();
      click();
      await s.done;
      await new Promise((r) => setTimeout(r, 50));
      expect(s.onExport).toHaveBeenCalledTimes(1);
      expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);
      click();
      await vi.waitFor(() => expect(s.onExport).toHaveBeenCalledTimes(2));
    });

    describe('menu text, langKey and i18n', () => {
      const GERMAN = 'Excel herunterladen';

      it('lang set BEFORE install wins over the default and is left alone by uninstall', () => {
        H.setOptions({ lang: { [DEFAULT_LANG_KEY]: GERMAN } });
        const inst = install();
        expect(inst.langKey).toBe(DEFAULT_LANG_KEY);
        expect(inst.menuItemKey).toBe(KEY);
        expect(H.getOptions().lang[DEFAULT_LANG_KEY]).toBe(GERMAN);
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        expect(itemLabel(H, chart)).toBe(GERMAN);
        inst.uninstall();
        installation = null;
        expect(H.getOptions().lang[DEFAULT_LANG_KEY]).toBe(GERMAN);
      });

      it('lang changed AFTER install is picked up by charts created afterwards', () => {
        install();
        const before = renderChart(H, F.simpleLine) as AnyRec;
        H.setOptions({ lang: { [DEFAULT_LANG_KEY]: GERMAN } });
        const after = renderChart(H, F.simpleLine) as AnyRec;
        expect(after.options.lang[DEFAULT_LANG_KEY]).toBe(GERMAN);
        expect(itemLabel(H, after)).toBe(GERMAN);
        expect(itemLabel(H, before)).toBe(DEFAULT_MENU_TEXT);
      });

      it('install menuText wins over lang', () => {
        H.setOptions({ lang: { [DEFAULT_LANG_KEY]: GERMAN } });
        install({ menuText: 'From install' });
        expect(itemLabel(H, renderChart(H, F.simpleLine))).toBe('From install');
      });

      it('per-chart menuText wins over install menuText and lang', () => {
        H.setOptions({ lang: { [DEFAULT_LANG_KEY]: GERMAN } });
        install({ menuText: 'From install' });
        const chart = renderChart(H, {
          ...F.simpleLine,
          exporting: { editableExcel: { menuText: 'From chart' } } as AnyRec,
        });
        expect(itemLabel(H, chart)).toBe('From chart');
      });

      it('falls back to DEFAULT_MENU_TEXT, registered under lang only while installed', () => {
        expect(H.getOptions().lang[DEFAULT_LANG_KEY]).toBeUndefined();
        const inst = install();
        expect(H.getOptions().lang[DEFAULT_LANG_KEY]).toBe(DEFAULT_MENU_TEXT);
        inst.uninstall();
        installation = null;
        expect(H.getOptions().lang[DEFAULT_LANG_KEY]).toBeUndefined();
      });

      it('a custom langKey drives the textKey and is translatable', () => {
        const inst = install({ langKey: 'myExcelText' });
        expect(inst.langKey).toBe('myExcelText');
        expect(definition(H).textKey).toBe('myExcelText');
        expect(H.getOptions().lang.myExcelText).toBe(DEFAULT_MENU_TEXT);
        H.setOptions({ lang: { myExcelText: 'Exportar a Excel' } });
        expect(itemLabel(H, renderChart(H, F.simpleLine))).toBe('Exportar a Excel');
        inst.uninstall();
        installation = null;
        // The app changed the value after install: it is the app's now and is kept.
        expect(H.getOptions().lang.myExcelText).toBe('Exportar a Excel');
      });
    });

    describe('menuIcon', () => {
      it("'excel' renders the built-in SVG before a plain-text label", () => {
        install({ menuIcon: 'excel' });
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        const markup = itemMarkup(chart);
        expect(markup).toContain('<svg');
        expect(markup).toContain(MENU_ICON_EXCEL_SVG);
        expect(markup).toContain('highcharts-editable-excel-label');
        const li = renderItem(H, chart);
        // Highcharts' AST allow-list keeps the icon.
        expect(li.querySelector('svg.hc-excel-menu-icon')).not.toBeNull();
        expect(li.querySelector('svg.hc-excel-menu-icon path')).not.toBeNull();
        expect(li.querySelector('.highcharts-editable-excel-item')).not.toBeNull();
        expect(itemLabel(H, chart)).toBe(DEFAULT_MENU_TEXT);
      });

      it('the icon label still follows lang (resolved per chart from chart.options.lang)', () => {
        install({ menuIcon: 'excel' });
        H.setOptions({ lang: { [DEFAULT_LANG_KEY]: 'Excel herunterladen' } });
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        expect(itemLabel(H, chart)).toBe('Excel herunterladen');
        expect(renderItem(H, chart).querySelector('svg')).not.toBeNull();
      });

      it('accepts { svg } and { html }, and a per-chart null removes the icon', () => {
        install({ menuIcon: { svg: '<svg class="mine" width="10" height="10"><path d="M0 0L10 10"/></svg>' } });
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        expect(renderItem(H, chart).querySelector('svg.mine path')).not.toBeNull();
        expect(itemLabel(H, chart)).toBe(DEFAULT_MENU_TEXT);
        const plain = renderChart(H, {
          ...F.simpleLine,
          exporting: { editableExcel: { menuIcon: null } } as AnyRec,
        }) as AnyRec;
        expect(itemMarkup(plain)).toBe(DEFAULT_MENU_TEXT);
        const html = renderChart(H, {
          ...F.simpleLine,
          exporting: { editableExcel: { menuIcon: { html: '<b class="badge">XL</b>' }, menuText: 'Go' } } as AnyRec,
        }) as AnyRec;
        const li = renderItem(H, html);
        expect(li.querySelector('b.badge')?.textContent).toBe('XL');
        expect(li.querySelector('.highcharts-editable-excel-label')?.textContent).toBe('Go');
      });

      it('rejects an unknown menuIcon with INVALID_OPTIONS', () => {
        for (const bad of ['word', { svg: '<div></div>' }, { other: 1 }, 42]) {
          expect(() => installHighchartsExcelExport(H, { menuIcon: bad as AnyRec })).toThrow(
            expect.objectContaining({ code: 'INVALID_OPTIONS' }),
          );
        }
        expect(definition(H)).toBeUndefined();
      });
    });

    describe('button and menu styling', () => {
      const ARROW = 'M0.5 0.1 L0.5 0.7 M0.25 0.45 L0.5 0.7 L0.75 0.45 M0.15 0.9 H0.85';

      it('button.svgPath registers a symbol and sets it on the global context button', () => {
        install({ button: { svgPath: ARROW, symbolFill: '#1d6f42', symbolStroke: '#1d6f42', symbolStrokeWidth: 2 } });
        const cb = H.getOptions().exporting.buttons.contextButton;
        expect(cb.symbol).toBe(DEFAULT_BUTTON_SYMBOL);
        expect(cb.symbolFill).toBe('#1d6f42');
        expect(cb.symbolStroke).toBe('#1d6f42');
        expect(cb.symbolStrokeWidth).toBe(2);
        const fn = symbols(H)[DEFAULT_BUTTON_SYMBOL];
        expect(typeof fn).toBe('function');
        const path = fn(0, 0, 20, 20) as unknown[][];
        expect(path.length).toBeGreaterThan(0);
        expect(path[0]).toEqual(['M', 10, 2]);
        expect(path[path.length - 1]).toEqual(['L', 17, 18]);
        // The rendered chart draws the custom symbol.
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        const d = chart.container
          .querySelector('.highcharts-contextbutton path.highcharts-button-symbol')
          ?.getAttribute('d');
        expect(d).toBeTruthy();
        expect(d).not.toMatch(/^M\s*7\.5\s+9/); // not the default "menu" (three bars)
      });

      it('svgPath under a given symbol name, with viewBox inference and relative commands', () => {
        install({ button: { symbol: 'myIcon', svgPath: 'M2 2 l20 20 h-20 Z' } });
        expect(H.getOptions().exporting.buttons.contextButton.symbol).toBe('myIcon');
        expect(symbols(H).myIcon(0, 0, 10, 10)).toEqual([['M', 0, 0], ['L', 10, 10], ['L', 0, 10], ['Z']]);
        expect(symbols(H)[DEFAULT_BUTTON_SYMBOL]).toBeUndefined();
      });

      it('accepts a flat path array and curve/arc commands', () => {
        install({
          button: { svgPath: ['M', 0, 0, 'C', 0, 1, 1, 1, 1, 0, 'Q', 0.5, 0.5, 0, 0, 'A', 0.5, 0.5, 0, 0, 1, 1, 1] },
        });
        const p = symbols(H)[DEFAULT_BUTTON_SYMBOL](0, 0, 2, 2);
        expect(p).toEqual([
          ['M', 0, 0],
          ['C', 0, 2, 2, 2, 2, 0],
          ['Q', 1, 1, 0, 0],
          ['A', 1, 1, 0, 0, 1, 2, 2],
        ]);
      });

      it('rejects an unparsable svgPath with INVALID_OPTIONS', () => {
        for (const bad of ['M0 0 S1 1 2 2', 'hello', 'M0 0 L1', '']) {
          expect(() => installHighchartsExcelExport(H, { button: { svgPath: bad } })).toThrow(
            expect.objectContaining({ code: 'INVALID_OPTIONS' }),
          );
        }
        expect(definition(H)).toBeUndefined();
      });

      it('button.title, className, text and theme land in the global options', () => {
        install({
          button: {
            title: 'Export menu',
            className: 'my-btn',
            text: 'Export',
            theme: { fill: '#eef' },
            symbolSize: 12,
          },
        });
        const o = H.getOptions();
        expect(o.lang.contextButtonTitle).toBe('Export menu');
        const cb = o.exporting.buttons.contextButton;
        expect(cb.className).toBe('highcharts-contextbutton my-btn');
        expect(cb.text).toBe('Export');
        expect(cb.theme.fill).toBe('#eef');
        expect(cb.symbolSize).toBe(12);
        expect(cb.symbol).toBe('menu');
      });

      it('menuStyle / menuItemStyle / menuItemHoverStyle merge into navigation and are restored', () => {
        const before = snap(H.getOptions().navigation);
        const inst = install({
          menuStyle: { border: '1px solid #1d6f42', borderRadius: '6px' },
          menuItemStyle: { fontSize: '13px', padding: '6px 14px' },
          menuItemHoverStyle: { background: '#e8f3ec', color: '#1d6f42' },
        });
        const nav = H.getOptions().navigation;
        expect(nav.menuStyle).toMatchObject({ border: '1px solid #1d6f42', borderRadius: '6px' });
        expect(nav.menuStyle.background).toBe(before.menuStyle.background);
        expect(nav.menuItemStyle).toMatchObject({ fontSize: '13px', padding: '6px 14px' });
        expect(nav.menuItemHoverStyle).toMatchObject({ background: '#e8f3ec', color: '#1d6f42' });
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        expect(chart.options.navigation.menuItemStyle.fontSize).toBe('13px');
        inst.uninstall();
        installation = null;
        expect(H.getOptions().navigation).toEqual(before);
      });

      it('uninstall restores every global option it touched', () => {
        const before = globalState(H);
        const previousMenuSymbol = symbols(H).menu;
        const inst = install({
          langKey: 'brandedExcel',
          menuIcon: 'excel',
          button: {
            svgPath: ARROW,
            symbolFill: '#1d6f42',
            symbolStroke: '#1d6f42',
            title: 'Menu',
            theme: { fill: '#fff' },
          },
          menuStyle: { border: '1px solid #1d6f42' },
          menuItemStyle: { fontSize: '13px' },
          menuItemHoverStyle: { color: '#1d6f42' },
        });
        expect(globalState(H)).not.toEqual(before);
        inst.uninstall();
        installation = null;
        expect(globalState(H)).toEqual(before);
        expect(symbols(H)[DEFAULT_BUTTON_SYMBOL]).toBeUndefined();
        expect(symbols(H).menu).toBe(previousMenuSymbol);
        const chart = renderChart(H, F.simpleLine) as AnyRec;
        expect(chartItems(chart)).not.toContain(KEY);
      });
    });
  });
}
