/**
 * Highcharts exporting-menu integration: adds a "Download editable Excel chart" item to the
 * context menu of every chart (and per-chart control through `exporting.editableExcel`).
 *
 * No Highcharts import: the namespace is passed in by the application.
 */

import { downloadHighchartsAsXlsx } from '../browser/download';
import {
  DEFAULT_MENU_ITEM_KEY,
  DEFAULT_MENU_TEXT,
  ExportError,
  type ExportOptions,
  type InstallOptions,
  type Installation,
  type PerChartExportConfig,
} from '../types/public-api';

type Rec = Record<string, unknown>;

interface HighchartsNamespaceLike {
  getOptions(): Rec;
  setOptions(options: Rec): unknown;
  addEvent(target: unknown, type: string, fn: (this: unknown, e: unknown) => void): (() => void) | undefined;
  Chart: { prototype?: Rec };
  Exporting?: unknown;
}

/** Highcharts' default context-menu items (v12/v13 without export-data). */
const FALLBACK_MENU_ITEMS: readonly string[] = Object.freeze([
  'viewFullscreen',
  'printChart',
  'separator',
  'downloadPNG',
  'downloadJPEG',
  'downloadSVG',
]);

/** Anchors the item is inserted after by default (the last one present wins). */
const DEFAULT_ANCHORS: readonly string[] = ['downloadXLS', 'downloadCSV', 'downloadSVG'];

const installations = new WeakMap<object, Installation>();
/** Charts with an export started from the menu and not finished yet (double-click guard). */
const inFlight = new WeakSet<object>();

function logError(message: string, error: unknown): void {
  console.error(`[highcharts-editable-excel] ${message}`, error);
}

function isRec(x: unknown): x is Rec {
  return typeof x === 'object' && x !== null && !Array.isArray(x);
}

function recAt(x: unknown, ...path: string[]): Rec | undefined {
  let cur: unknown = x;
  for (const key of path) {
    if (!isRec(cur)) return undefined;
    cur = cur[key];
  }
  return isRec(cur) ? cur : undefined;
}

function isHighchartsNamespace(x: unknown): x is HighchartsNamespaceLike {
  if ((typeof x !== 'object' && typeof x !== 'function') || x === null) return false;
  const h = x as Rec;
  return (
    typeof h.getOptions === 'function' &&
    typeof h.setOptions === 'function' &&
    typeof h.addEvent === 'function' &&
    typeof h.Chart === 'function'
  );
}

function hasExportingModule(H: HighchartsNamespaceLike): boolean {
  const proto = H.Chart.prototype;
  if (proto && typeof proto.exportChart === 'function') return true;
  if (H.Exporting !== undefined && H.Exporting !== null) return true;
  return recAt(H.getOptions(), 'exporting', 'menuItemDefinitions') !== undefined;
}

function menuItemsOf(options: unknown): unknown[] | undefined {
  const items = recAt(options, 'exporting', 'buttons', 'contextButton')?.menuItems;
  return Array.isArray(items) ? items : undefined;
}

/** Returns a copy of `items` with `key` inserted once (no-op copy when already present). */
export function insertMenuItem(items: readonly unknown[], key: string, insertAfter?: string): unknown[] {
  const out = [...items];
  if (out.includes(key)) return out;
  let at = -1;
  if (insertAfter !== undefined) at = out.indexOf(insertAfter);
  if (at < 0) {
    for (const anchor of DEFAULT_ANCHORS) at = Math.max(at, out.indexOf(anchor));
  }
  if (at < 0) out.push(key);
  else out.splice(at + 1, 0, key);
  return out;
}

function withoutKey(items: readonly unknown[], key: string): unknown[] {
  return items.filter((item) => item !== key);
}

/** Per-chart config with the menu-only fields removed. */
function perChartExportOptions(config: PerChartExportConfig | undefined): ExportOptions {
  if (!config) return {};
  const { enabled: _enabled, menuText: _menuText, ...rest } = config;
  return rest;
}

function readPerChartConfig(chart: unknown): PerChartExportConfig | undefined {
  const config = recAt(chart, 'options', 'exporting', 'editableExcel');
  return config as PerChartExportConfig | undefined;
}

/** Builds the menu item definition; its `onclick` runs with `this` = the chart. */
function createMenuItemDefinition(
  options: InstallOptions,
  text: string,
): { text: string; onclick: (this: unknown) => void } {
  return {
    text,
    onclick(this: unknown): void {
      const key = typeof this === 'object' && this !== null ? this : null;
      // A second click while this chart's export is still running is ignored.
      if (key !== null && inFlight.has(key)) return;
      const merged: ExportOptions = {
        ...(options.exportOptions ?? {}),
        ...perChartExportOptions(readPerChartConfig(this)),
      };
      // Neither callback may produce an unhandled rejection: a throwing onExport goes to onError,
      // a throwing onError is logged.
      const handleError = (error: unknown): void => {
        if (!options.onError) {
          logError('Export failed:', error);
          return;
        }
        try {
          options.onError(error, this);
        } catch (callbackError) {
          logError('onError threw:', callbackError);
        }
      };
      let pending: Promise<Awaited<ReturnType<typeof downloadHighchartsAsXlsx>>>;
      try {
        pending = downloadHighchartsAsXlsx(this, merged);
      } catch (error) {
        handleError(error);
        return;
      }
      if (key !== null) inFlight.add(key);
      pending
        .then((result) => options.onExport?.(result, this))
        .catch(handleError)
        .finally(() => {
          if (key !== null) inFlight.delete(key);
        });
    },
  };
}

/**
 * Applies the per-chart settings (`exporting.editableExcel`, explicit `menuItems`) to a chart that
 * has just merged its options (Chart 'afterInit', before the first render). Only objects owned by
 * the chart are replaced: `chart.userOptions` (Highcharts' shallow copy) and `chart.options`
 * (the merged result). The caller's options object, possibly frozen or reused for several charts,
 * is never assigned into.
 *
 * (The Chart 'init' event cannot do this without mutation: Highcharts' module code is strict, so
 * replacing `e.args[0]` does not change the options the chart merges.)
 */
function applyPerChartOptions(
  H: HighchartsNamespaceLike,
  chart: Rec,
  key: string,
  insertAfter: string | undefined,
): void {
  const userOptions = isRec(chart.userOptions) ? chart.userOptions : undefined;
  const merged = isRec(chart.options) ? chart.options : undefined;
  if (!userOptions || !merged) return;
  const exporting = recAt(userOptions, 'exporting');
  const config = recAt(exporting, 'editableExcel') as PerChartExportConfig | undefined;
  const ownItems = menuItemsOf(userOptions);
  const disabled = config?.enabled === false;
  const menuText = typeof config?.menuText === 'string' && config.menuText !== '' ? config.menuText : undefined;

  let newItems: unknown[] | undefined;
  if (disabled) {
    const base = ownItems ?? menuItemsOf(H.getOptions()) ?? FALLBACK_MENU_ITEMS;
    newItems = withoutKey(base, key);
  } else if (ownItems && !ownItems.includes(key)) {
    newItems = insertMenuItem(ownItems, key, insertAfter);
  }
  if (newItems === undefined && menuText === undefined) return;

  // Shallow copies all the way down: the caller's nested objects and arrays are never edited.
  const withSettings = (base: Rec | undefined, defsFallback: Rec | undefined): Rec => {
    const next: Rec = { ...(base ?? {}) };
    if (newItems !== undefined) {
      const buttons = recAt(base, 'buttons');
      const contextButton = recAt(buttons, 'contextButton');
      next.buttons = { ...(buttons ?? {}), contextButton: { ...(contextButton ?? {}), menuItems: newItems } };
    }
    if (menuText !== undefined && !disabled) {
      const defs = recAt(base, 'menuItemDefinitions');
      const def = recAt(defs, key) ?? recAt(defsFallback, key) ?? {};
      next.menuItemDefinitions = { ...(defs ?? {}), [key]: { ...def, text: menuText } };
    }
    return next;
  };
  const globalDefs = recAt(H.getOptions(), 'exporting', 'menuItemDefinitions');
  chart.userOptions = { ...userOptions, exporting: withSettings(exporting, globalDefs) };

  // The merged `chart.options.exporting` object is chart-owned and already referenced by the
  // exporting module (`chart.exporting.options`, and the context button copies `menuItems` when it
  // renders), so it must be updated IN PLACE: assign new arrays/objects onto it rather than
  // replacing it. The caller's arrays are still never edited (a new `menuItems` array is assigned).
  const mergedExporting = recAt(merged, 'exporting');
  const updated = withSettings(mergedExporting, globalDefs);
  if (mergedExporting) {
    if (newItems !== undefined) {
      let buttons = recAt(mergedExporting, 'buttons');
      if (!buttons) {
        buttons = {};
        mergedExporting.buttons = buttons;
      }
      let contextButton = recAt(buttons, 'contextButton');
      if (!contextButton) {
        contextButton = {};
        buttons.contextButton = contextButton;
      }
      contextButton.menuItems = newItems;
    }
    if (updated.menuItemDefinitions !== undefined) mergedExporting.menuItemDefinitions = updated.menuItemDefinitions;
  } else {
    merged.exporting = updated;
  }
  const exportingInstance = recAt(chart, 'exporting');
  if (exportingInstance && exportingInstance.options !== recAt(merged, 'exporting')) {
    exportingInstance.options = recAt(merged, 'exporting');
  }
}

/**
 * Registers the "Download editable Excel chart" item in Highcharts' exporting menu for every chart
 * created afterwards. Idempotent per Highcharts namespace: a second call returns the same
 * installation (the first call's options win). After `uninstall()` a new install is allowed.
 *
 * Per chart, `exporting.editableExcel` (PerChartExportConfig) can disable the item
 * (`enabled: false`), change its text (`menuText`) or override export options.
 *
 * Per-chart `exporting.editableExcel.enabled` / `menuText` and explicit `menuItems` are applied
 * when the chart is CREATED. After `chart.update({ exporting: … })` that changes them, call
 * `addEditableExcelMenuItem(chart)` again (export options in `editableExcel` are read on every
 * click and need nothing). The caller's options object is never modified, so frozen or shared
 * option objects are fine.
 *
 * @throws ExportError INVALID_OPTIONS when `Highcharts` is not a Highcharts namespace;
 *   EXPORTING_MODULE_MISSING when the exporting module is not loaded.
 */
export function installHighchartsExcelExport(Highcharts: unknown, options: InstallOptions = {}): Installation {
  if (!isHighchartsNamespace(Highcharts)) {
    throw new ExportError(
      'INVALID_OPTIONS',
      'installHighchartsExcelExport expects the Highcharts namespace (with getOptions, setOptions, addEvent and Chart).',
    );
  }
  const H = Highcharts;
  const existing = installations.get(H);
  if (existing) return existing;
  if (!hasExportingModule(H)) {
    throw new ExportError(
      'EXPORTING_MODULE_MISSING',
      'Load highcharts/modules/exporting before installing the editable Excel export menu item.',
    );
  }

  const frozenOptions: Readonly<InstallOptions> = Object.freeze({ ...options });
  const key = options.menuItemKey ?? DEFAULT_MENU_ITEM_KEY;
  const text = options.menuText ?? DEFAULT_MENU_TEXT;

  // 1. Definition (setOptions merges menuItemDefinitions).
  H.setOptions({ exporting: { menuItemDefinitions: { [key]: createMenuItemDefinition(frozenOptions, text) } } });

  // 2. Global default menu (setOptions replaces menuItems arrays).
  const originalItems = menuItemsOf(H.getOptions());
  const baseItems = originalItems ?? FALLBACK_MENU_ITEMS;
  const addedToMenu = !baseItems.includes(key);
  if (addedToMenu) {
    H.setOptions({
      exporting: { buttons: { contextButton: { menuItems: insertMenuItem(baseItems, key, options.insertAfter) } } },
    });
  }

  // 3. Per-chart options. Chart 'afterInit' fires after `chart.options`/`chart.userOptions` are set
  //    and before the first render (which draws the exporting button).
  const unbind = H.addEvent(H.Chart, 'afterInit', function onChartAfterInit(this: unknown): void {
    if (isRec(this)) applyPerChartOptions(H, this, key, options.insertAfter);
  });

  let installed = true;
  const installation: Installation = {
    options: frozenOptions,
    uninstall(): void {
      if (!installed) return;
      installed = false;
      if (typeof unbind === 'function') unbind();
      const defs = recAt(H.getOptions(), 'exporting', 'menuItemDefinitions');
      if (defs) delete defs[key];
      if (addedToMenu) {
        if (originalItems) {
          H.setOptions({ exporting: { buttons: { contextButton: { menuItems: originalItems } } } });
        } else {
          const contextButton = recAt(H.getOptions(), 'exporting', 'buttons', 'contextButton');
          if (contextButton) delete contextButton.menuItems;
        }
      }
      if (installations.get(H) === installation) installations.delete(H);
    },
  };
  installations.set(H, installation);
  return installation;
}

/**
 * Adds the menu item to an ALREADY-RENDERED chart (charts created before
 * `installHighchartsExcelExport`, or charts that override `menuItems`). Uses `chart.update`, which
 * re-renders the exporting button; the item definition is set on the chart itself so it works
 * without a global install. No-op when the chart already lists the item.
 *
 * @throws ExportError INVALID_CHART when `chart` has no `update` method / options.
 */
export function addEditableExcelMenuItem(chart: unknown, options: InstallOptions = {}): void {
  if (!isRec(chart) || typeof chart.update !== 'function' || !isRec(chart.options)) {
    throw new ExportError('INVALID_CHART', 'addEditableExcelMenuItem expects a rendered Highcharts chart.');
  }
  const key = options.menuItemKey ?? DEFAULT_MENU_ITEM_KEY;
  const config = readPerChartConfig(chart);
  if (config?.enabled === false) return;
  const current = menuItemsOf(chart.options) ?? FALLBACK_MENU_ITEMS;
  const hasDefinition = recAt(chart.options, 'exporting', 'menuItemDefinitions', key) !== undefined;
  if (current.includes(key) && hasDefinition) return;
  const text = config?.menuText ?? options.menuText ?? DEFAULT_MENU_TEXT;
  (chart.update as (o: Rec, redraw?: boolean) => void).call(
    chart,
    {
      exporting: {
        menuItemDefinitions: { [key]: createMenuItemDefinition(options, text) },
        buttons: { contextButton: { menuItems: insertMenuItem(current, key, options.insertAfter) } },
      },
    },
    true,
  );
}
