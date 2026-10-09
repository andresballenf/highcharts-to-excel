/**
 * Export-menu integration assertions shared by the per-version test files. Each version file
 * imports its own Highcharts build (+ exporting and highcharts-more) and calls `runInstallSuite`.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { installHighchartsExcelExport, addEditableExcelMenuItem } from '../../src/highcharts/install-export-menu';
import {
  DEFAULT_MENU_ITEM_KEY,
  DEFAULT_MENU_TEXT,
  ExportError,
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
    const install = (
      ...args: Parameters<typeof installHighchartsExcelExport> extends [unknown, ...infer R] ? R : never
    ): Installation => (installation = installHighchartsExcelExport(H, ...args));

    beforeEach(() => {
      originalItems = globalItems(H);
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
    });

    it('adds the item once and returns the same installation on a second call', () => {
      const a = install();
      const b = installHighchartsExcelExport(H, { menuText: 'ignored' });
      expect(b).toBe(a);
      expect(a.options.menuText).toBeUndefined();
      expect(count(globalItems(H), KEY)).toBe(1);
      expect(definition(H).text).toBe(DEFAULT_MENU_TEXT);
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
      expect(chart.options.exporting.menuItemDefinitions[KEY].text).toBe(DEFAULT_MENU_TEXT);
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
      expect(definition(H).text).toBe(DEFAULT_MENU_TEXT);
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

    it('routes a polar-chart failure to onError as an ExportError', async () => {
      const s = settle();
      install({ onExport: s.onExport, onError: s.onError });
      const chart = renderChart(H, F.polarChart) as AnyRec;
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
      const chart = renderChart(H, F.polarChart) as AnyRec;
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
        const chart = renderChart(H, F.polarChart) as AnyRec;
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
  });
}
