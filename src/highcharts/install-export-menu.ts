/**
 * Highcharts exporting-menu integration: adds a "Download editable Excel chart" item to the
 * context menu of every chart (and per-chart control through `exporting.editableExcel`).
 *
 * No Highcharts import: the namespace is passed in by the application.
 */

import { downloadHighchartsAsXlsx } from '../browser/download';
import {
  type ContextButtonOptions,
  DEFAULT_BUTTON_SYMBOL,
  DEFAULT_LANG_KEY,
  DEFAULT_MENU_ITEM_KEY,
  DEFAULT_MENU_TEXT,
  ExportError,
  type ExportOptions,
  type InstallOptions,
  type Installation,
  MENU_ICON_EXCEL_SVG,
  type MenuIconOption,
  type PerChartExportConfig,
} from '../types/public-api';

type Rec = Record<string, unknown>;

interface HighchartsNamespaceLike {
  getOptions(): Rec;
  setOptions(options: Rec): unknown;
  addEvent(target: unknown, type: string, fn: (this: unknown, e: unknown) => void): (() => void) | undefined;
  Chart: { prototype?: Rec };
  Exporting?: unknown;
  SVGRenderer?: { prototype?: Rec };
  /** v11 alias of the default renderer class. */
  Renderer?: { prototype?: Rec };
}

/** A Highcharts path array (`[['M', x, y], ['L', x, y], ['Z']]`). */
export type PathArray = Array<[string, ...number[]]>;
/** Signature of a Highcharts renderer symbol. */
export type SymbolFunction = (x: number, y: number, w: number, h: number) => PathArray;

/** Highcharts' default context-menu items (v12/v13 without export-data). */
const FALLBACK_MENU_ITEMS: readonly string[] = Object.freeze([
  'viewFullscreen',
  'printChart',
  'separator',
  'downloadPNG',
  'downloadJPEG',
  'downloadSVG',
]);

/** `ContextButtonOptions` keys copied as-is to `exporting.buttons.contextButton`. */
const BUTTON_KEYS = [
  'symbol',
  'symbolFill',
  'symbolStroke',
  'symbolStrokeWidth',
  'symbolSize',
  'theme',
  'text',
] as const;

/** `InstallOptions` keys merged into `navigation`. */
const MENU_STYLE_KEYS = ['menuStyle', 'menuItemStyle', 'menuItemHoverStyle'] as const;

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

function invalidInstallOption(property: string, expected: string): ExportError {
  return new ExportError('INVALID_OPTIONS', `Invalid install option "${property}": expected ${expected}.`, {
    property,
  });
}

/** Markup for a menu icon option (validated); undefined for `null`/`undefined`. */
function iconMarkup(icon: MenuIconOption | undefined, property: string): string | undefined {
  if (icon === undefined || icon === null) return undefined;
  if (icon === 'excel') return MENU_ICON_EXCEL_SVG;
  const value: unknown = icon;
  if (isRec(value)) {
    if (typeof value.svg === 'string' && /^\s*<svg[\s>]/i.test(value.svg)) return value.svg;
    if (typeof value.html === 'string') return value.html;
  }
  throw invalidInstallOption(property, "'excel', { svg: '<svg …>' }, { html: string } or null");
}

/**
 * The item markup with an icon: `<span class="highcharts-editable-excel-item">ICON<span
 * class="highcharts-editable-excel-label">TEXT</span></span>`. The label keeps the text as given
 * (Highcharts treats item text as markup too), so the item's `textContent` is the plain label when
 * the icon has no text nodes.
 */
function withIcon(icon: string | undefined, label: string): string {
  if (icon === undefined) return label;
  return (
    '<span class="highcharts-editable-excel-item">' +
    `${icon}<span class="highcharts-editable-excel-label">${label}</span></span>`
  );
}

function nonEmptyString(x: unknown): string | undefined {
  return typeof x === 'string' && x !== '' ? x : undefined;
}

/** Deep copy of plain option data (records and arrays; functions and other values by reference). */
function clonePlain<T>(value: T): T {
  if (Array.isArray(value)) return value.map((v: unknown) => clonePlain(v)) as T;
  if (isRec(value) && Object.getPrototypeOf(value) === Object.prototype) {
    const out: Rec = {};
    for (const [k, v] of Object.entries(value)) out[k] = clonePlain(v);
    return out as T;
  }
  return value;
}

/**
 * Records `holder()[key]` now and returns a function that puts it back (or deletes it). The holder
 * is looked up again on restore, in case `setOptions` replaced it.
 */
function snapshotKey(holder: () => Rec, key: string): () => void {
  const before = holder();
  const had = Object.hasOwn(before, key);
  const original = clonePlain(before[key]);
  return () => {
    const target = holder();
    if (had) target[key] = original;
    else delete target[key];
  };
}

/** Returns `getOptions()[path…]`, creating empty records along the way. */
function ensureRec(root: Rec, ...path: string[]): Rec {
  let cur = root;
  for (const key of path) {
    let next = cur[key];
    if (!isRec(next)) {
      next = {};
      cur[key] = next;
    }
    cur = next as Rec;
  }
  return cur;
}

// --- SVG path → Highcharts symbol --------------------------------------------------------------

/** Arguments per command; `A` has 7 (rx ry rotation large-arc sweep x y). */
const PATH_ARITY: Readonly<Record<string, number>> = { M: 2, L: 2, H: 1, V: 1, C: 6, Q: 4, A: 7, Z: 0 };

interface PathSegment {
  /** Absolute, uppercase: M L C Q A Z (H/V become L). */
  cmd: 'M' | 'L' | 'C' | 'Q' | 'A' | 'Z';
  args: number[];
}

function pathError(detail: string): ExportError {
  return invalidInstallOption(
    'button.svgPath',
    `an SVG path with M, L, H, V, C, Q, A, Z commands (absolute or relative); ${detail}`,
  );
}

/**
 * Parses an SVG path `d` (or a flat `['M', 0, 0, 'L', …]` array) into absolute segments.
 *
 * @throws ExportError INVALID_OPTIONS (property `button.svgPath`) for unsupported commands, bad
 *   numbers or a wrong number of arguments.
 */
export function parseSvgPath(path: string | ReadonlyArray<string | number>): PathSegment[] {
  const d = typeof path === 'string' ? path : path.join(' ');
  const tokens = d.match(/[a-zA-Z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  const leftover = d.replace(/[a-zA-Z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|[\s,]+/g, '');
  if (leftover !== '') throw pathError(`unexpected "${leftover.slice(0, 20)}"`);
  if (tokens.length === 0 || !/^[Mm]$/.test(tokens[0] ?? '')) throw pathError('it must start with M');

  const out: PathSegment[] = [];
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let i = 0;
  while (i < tokens.length) {
    const letter = tokens[i++] as string;
    const upper = letter.toUpperCase();
    const arity = PATH_ARITY[upper];
    if (arity === undefined) throw pathError(`"${letter}" is not supported`);
    const relative = letter !== upper;
    const nums: number[] = [];
    while (i < tokens.length && !/^[a-zA-Z]$/.test(tokens[i] as string)) nums.push(Number(tokens[i++]));
    if (arity === 0) {
      if (nums.length > 0) throw pathError('Z takes no arguments');
      out.push({ cmd: 'Z', args: [] });
      x = startX;
      y = startY;
      continue;
    }
    if (nums.length === 0 || nums.length % arity !== 0) {
      throw pathError(`"${letter}" needs a multiple of ${arity} numbers`);
    }
    for (let k = 0; k < nums.length; k += arity) {
      const a = nums.slice(k, k + arity);
      const ox = relative ? x : 0;
      const oy = relative ? y : 0;
      // Implicit repeats after a moveto are linetos.
      const cmd = upper === 'M' && k > 0 ? 'L' : upper;
      switch (cmd) {
        case 'M':
        case 'L':
          x = (a[0] as number) + ox;
          y = (a[1] as number) + oy;
          if (cmd === 'M') {
            startX = x;
            startY = y;
          }
          out.push({ cmd, args: [x, y] });
          break;
        case 'H':
          x = (a[0] as number) + ox;
          out.push({ cmd: 'L', args: [x, y] });
          break;
        case 'V':
          y = (a[0] as number) + oy;
          out.push({ cmd: 'L', args: [x, y] });
          break;
        case 'C':
        case 'Q': {
          const abs = a.map((v, n) => v + (n % 2 === 0 ? ox : oy));
          x = abs[abs.length - 2] as number;
          y = abs[abs.length - 1] as number;
          out.push({ cmd, args: abs });
          break;
        }
        default: {
          // A: rx ry rotation large-arc sweep x y
          x = (a[5] as number) + ox;
          y = (a[6] as number) + oy;
          out.push({
            cmd: 'A',
            args: [a[0] as number, a[1] as number, a[2] as number, a[3] as number, a[4] as number, x, y],
          });
        }
      }
    }
  }
  return out;
}

/** Coordinates of a segment as [x, y] pairs (arc radii and flags excluded). */
function segmentPoints(seg: PathSegment): Array<[number, number]> {
  const a = seg.cmd === 'A' ? seg.args.slice(5) : seg.args;
  const pts: Array<[number, number]> = [];
  for (let k = 0; k + 1 < a.length; k += 2) pts.push([a[k] as number, a[k + 1] as number]);
  return pts;
}

const round4 = (v: number): number => Math.round(v * 1e4) / 1e4;

/**
 * Builds a renderer symbol from an SVG path. Coordinates all within 0..1 are a unit box; otherwise
 * the bounding box, made square around its center, is scaled to the symbol box.
 */
export function svgPathToSymbol(path: string | ReadonlyArray<string | number>): SymbolFunction {
  const segments = parseSvgPath(path);
  const pts = segments.flatMap(segmentPoints);
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const [minX, maxX, minY, maxY] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
  if (![minX, maxX, minY, maxY].every(Number.isFinite)) throw pathError('it has no coordinates');
  let box: [number, number, number];
  if (minX >= 0 && minY >= 0 && maxX <= 1 && maxY <= 1) box = [0, 0, 1];
  else {
    const size = Math.max(maxX - minX, maxY - minY) || 1;
    box = [(minX + maxX - size) / 2, (minY + maxY - size) / 2, size];
  }
  const [bx, by, bs] = box;
  return (x, y, w, h) => {
    const sx = w / bs;
    const sy = h / bs;
    const px = (v: number): number => round4(x + (v - bx) * sx);
    const py = (v: number): number => round4(y + (v - by) * sy);
    return segments.map((seg): [string, ...number[]] => {
      if (seg.cmd === 'Z') return ['Z'];
      if (seg.cmd === 'A') {
        const [rx, ry, rot, large, sweep, ex, ey] = seg.args as [
          number,
          number,
          number,
          number,
          number,
          number,
          number,
        ];
        return ['A', round4(rx * sx), round4(ry * sy), rot, large, sweep, px(ex), py(ey)];
      }
      return [seg.cmd, ...seg.args.map((v, n) => (n % 2 === 0 ? px(v) : py(v)))];
    });
  };
}

/** The renderer's symbol table (`SVGRenderer.prototype.symbols`, v11 also `Renderer.prototype`). */
function symbolTable(H: HighchartsNamespaceLike): Rec | undefined {
  return recAt(H.SVGRenderer?.prototype, 'symbols') ?? recAt(H.Renderer?.prototype, 'symbols');
}

/** Per-chart config with the menu-only fields removed. */
function perChartExportOptions(config: PerChartExportConfig | undefined): ExportOptions {
  if (!config) return {};
  const { enabled: _enabled, menuText: _menuText, menuIcon: _menuIcon, ...rest } = config;
  return rest;
}

function readPerChartConfig(chart: unknown): PerChartExportConfig | undefined {
  const config = recAt(chart, 'options', 'exporting', 'editableExcel');
  return config as PerChartExportConfig | undefined;
}

/** Builds the menu item definition; its `onclick` runs with `this` = the chart. */
function createMenuItemDefinition(
  options: InstallOptions,
  text: string | undefined,
  textKey: string,
): { text?: string; textKey: string; onclick: (this: unknown) => void } {
  return {
    // Without `text`, Highcharts renders `chart.options.lang[textKey]` (its own i18n lookup).
    ...(text !== undefined ? { text } : {}),
    textKey,
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
/** Install-wide settings the per-chart hook needs. */
interface MenuContext {
  key: string;
  langKey: string;
  insertAfter: string | undefined;
  /** Install `menuText` (non-empty). */
  menuText: string | undefined;
  /** Install icon markup. */
  icon: string | undefined;
}

/** Per-chart icon markup; an invalid per-chart value is logged and the install icon is kept. */
function chartIcon(config: PerChartExportConfig | undefined, fallback: string | undefined): string | undefined {
  if (config?.menuIcon === undefined) return fallback;
  try {
    return iconMarkup(config.menuIcon, 'exporting.editableExcel.menuIcon');
  } catch (error) {
    logError('Ignoring exporting.editableExcel.menuIcon:', error);
    return fallback;
  }
}

/**
 * The item text one chart needs in its own definition, or undefined when the global definition
 * (text, or `textKey` resolved by Highcharts) already renders the right thing.
 * Label order: per-chart menuText → install menuText → chart.options.lang[langKey] → default.
 */
function perChartItemText(ctx: MenuContext, config: PerChartExportConfig | undefined, lang: Rec | undefined) {
  const chartText = nonEmptyString(config?.menuText);
  if (chartText === undefined && config?.menuIcon === undefined && ctx.icon === undefined) return undefined;
  const label = chartText ?? ctx.menuText ?? nonEmptyString(lang?.[ctx.langKey]) ?? DEFAULT_MENU_TEXT;
  return withIcon(chartIcon(config, ctx.icon), label);
}

/**
 * Returns true when it changed the chart's menu items or item definition (so a rendered context
 * button must be redrawn).
 */
function applyPerChartOptions(H: HighchartsNamespaceLike, chart: Rec, ctx: MenuContext): boolean {
  const { key, insertAfter } = ctx;
  const userOptions = isRec(chart.userOptions) ? chart.userOptions : undefined;
  const merged = isRec(chart.options) ? chart.options : undefined;
  if (!userOptions || !merged) return false;
  const exporting = recAt(userOptions, 'exporting');
  const config = recAt(exporting, 'editableExcel') as PerChartExportConfig | undefined;
  const ownItems = menuItemsOf(userOptions);
  const disabled = config?.enabled === false;
  const menuText = disabled ? undefined : perChartItemText(ctx, config, recAt(merged, 'lang'));

  let newItems: unknown[] | undefined;
  if (disabled) {
    const base = ownItems ?? menuItemsOf(H.getOptions()) ?? FALLBACK_MENU_ITEMS;
    newItems = withoutKey(base, key);
  } else if (ownItems && !ownItems.includes(key)) {
    newItems = insertMenuItem(ownItems, key, insertAfter);
  }
  if (newItems === undefined && menuText === undefined) return false;
  const currentItems = menuItemsOf(merged);
  const sameItems =
    newItems === undefined ||
    (currentItems !== undefined &&
      currentItems.length === newItems.length &&
      currentItems.every((item, i) => item === newItems![i]));
  const currentText = recAt(merged, 'exporting', 'menuItemDefinitions', key)?.text;
  const changed = !sameItems || (menuText !== undefined && currentText !== menuText);

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
  return changed;
}

/**
 * Redraws the exporting buttons of a rendered chart so the context menu is rebuilt from the
 * current options (Highcharts 12/13: `chart.exporting.render()`, Highcharts 11: `renderExporting`).
 */
function rerenderExportingButtons(chart: Rec): void {
  const exporting = recAt(chart, 'exporting');
  try {
    if (exporting && typeof exporting.render === 'function') {
      exporting.isDirty = true;
      (exporting.render as () => void).call(exporting);
    } else if (typeof chart.renderExporting === 'function') {
      chart.isDirtyExporting = true;
      (chart.renderExporting as () => void).call(chart);
    }
  } catch (error) {
    logError('Could not redraw the exporting menu after chart.update:', error);
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
 * when the chart is created and again after every `chart.update({ exporting: … })` (Chart
 * 'afterUpdate'), so `chart.update({ exporting: { editableExcel: { enabled: false } } })` removes the
 * item and a new `buttons.contextButton.menuItems` list gets it re-inserted. Export options in
 * `editableExcel` are read on every click. The caller's options object is never modified, so frozen
 * or shared option objects are fine.
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

  // 0. Validate everything before touching the global options.
  const key = options.menuItemKey ?? DEFAULT_MENU_ITEM_KEY;
  const langKey = nonEmptyString(options.langKey) ?? DEFAULT_LANG_KEY;
  const icon = iconMarkup(options.menuIcon, 'menuIcon');
  const button: ContextButtonOptions | undefined = options.button;
  if (button !== undefined && !isRec(button)) throw invalidInstallOption('button', 'an object');
  const svgPath: unknown = button?.svgPath;
  if (svgPath !== undefined && typeof svgPath !== 'string' && !Array.isArray(svgPath)) {
    throw invalidInstallOption('button.svgPath', 'an SVG path string or a flat path array');
  }
  const symbolFn = svgPath !== undefined ? svgPathToSymbol(svgPath as string | Array<string | number>) : undefined;
  const symbols = symbolFn ? symbolTable(H) : undefined;
  if (symbolFn && !symbols) {
    throw invalidInstallOption('button.svgPath', 'a Highcharts build exposing SVGRenderer.prototype.symbols');
  }
  for (const k of MENU_STYLE_KEYS) {
    if (options[k] !== undefined && !isRec(options[k])) throw invalidInstallOption(k, 'an object of CSS declarations');
  }

  const frozenOptions: Readonly<InstallOptions> = Object.freeze({ ...options });
  const menuText = nonEmptyString(options.menuText);
  const restores: Array<() => void> = [];
  const globals = (): Rec => H.getOptions();

  // 1. lang[langKey]: registered only when absent (an app translation set before install wins).
  const lang = (): Rec => ensureRec(globals(), 'lang');
  const addedLang = typeof lang()[langKey] !== 'string';
  if (addedLang) H.setOptions({ lang: { [langKey]: DEFAULT_MENU_TEXT } });

  // 2. Definition (setOptions merges menuItemDefinitions). Without menuText and icon, `text` stays
  //    undefined and Highcharts renders `chart.options.lang[langKey]` (its own i18n). With an icon,
  //    the per-chart hook writes each chart's text from that chart's lang.
  const text =
    menuText !== undefined
      ? withIcon(icon, menuText)
      : icon !== undefined
        ? withIcon(icon, nonEmptyString(lang()[langKey]) ?? DEFAULT_MENU_TEXT)
        : undefined;
  H.setOptions({
    exporting: { menuItemDefinitions: { [key]: createMenuItemDefinition(frozenOptions, text, langKey) } },
  });

  // 2. Global default menu (setOptions replaces menuItems arrays).
  const originalItems = menuItemsOf(H.getOptions());
  const baseItems = originalItems ?? FALLBACK_MENU_ITEMS;
  const addedToMenu = !baseItems.includes(key);
  if (addedToMenu) {
    H.setOptions({
      exporting: { buttons: { contextButton: { menuItems: insertMenuItem(baseItems, key, options.insertAfter) } } },
    });
  }

  // 4. Context button branding (global; restored on uninstall).
  if (button) {
    const contextButton = (): Rec => ensureRec(globals(), 'exporting', 'buttons', 'contextButton');
    const patch: Rec = {};
    for (const k of BUTTON_KEYS) if (button[k] !== undefined) patch[k] = button[k];
    if (symbolFn && symbols) {
      const name = nonEmptyString(button.symbol) ?? DEFAULT_BUTTON_SYMBOL;
      restores.push(snapshotKey(() => symbols, name));
      symbols[name] = symbolFn;
      patch.symbol = name;
    }
    const extraClass = nonEmptyString(button.className);
    if (extraClass !== undefined) {
      patch.className = `${nonEmptyString(contextButton().className) ?? 'highcharts-contextbutton'} ${extraClass}`;
    }
    for (const k of Object.keys(patch)) restores.push(snapshotKey(contextButton, k));
    H.setOptions({ exporting: { buttons: { contextButton: patch } } });
    if (button.title !== undefined) {
      restores.push(snapshotKey(lang, 'contextButtonTitle'));
      H.setOptions({ lang: { contextButtonTitle: button.title } });
    }
  }

  // 5. Dropdown styling: merged into navigation.menuStyle / menuItemStyle / menuItemHoverStyle.
  for (const k of MENU_STYLE_KEYS) {
    const style = options[k];
    if (style === undefined) continue;
    restores.push(snapshotKey(() => ensureRec(globals(), 'navigation'), k));
    H.setOptions({ navigation: { [k]: style } });
  }

  // 6. Per-chart options. Chart 'afterInit' fires after `chart.options`/`chart.userOptions` are set
  //    and before the first render (which draws the exporting button).
  const ctx: MenuContext = { key, langKey, insertAfter: options.insertAfter, menuText, icon };
  const unbind = H.addEvent(H.Chart, 'afterInit', function onChartAfterInit(this: unknown): void {
    if (isRec(this)) applyPerChartOptions(H, this, ctx);
  });
  // `chart.update({ exporting: … })` merges new per-chart settings: apply them again (in place, as
  // above) and redraw the button, whose menu was built from the previous items.
  const unbindUpdate = H.addEvent(H.Chart, 'afterUpdate', function onChartAfterUpdate(this: unknown, e: unknown): void {
    if (!isRec(this)) return;
    const updated = isRec(e) ? e.options : undefined;
    if (!isRec(updated) || updated.exporting === undefined) return;
    if (applyPerChartOptions(H, this, ctx) && this.renderer !== undefined) rerenderExportingButtons(this);
  });

  let installed = true;
  const installation: Installation = {
    options: frozenOptions,
    langKey,
    menuItemKey: key,
    uninstall(): void {
      if (!installed) return;
      installed = false;
      if (typeof unbind === 'function') unbind();
      if (typeof unbindUpdate === 'function') unbindUpdate();
      for (const restore of restores.reverse()) restore();
      // Our default text only: a translation the app set meanwhile stays.
      if (addedLang && lang()[langKey] === DEFAULT_MENU_TEXT) delete lang()[langKey];
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
  const langKey = nonEmptyString(options.langKey) ?? DEFAULT_LANG_KEY;
  const config = readPerChartConfig(chart);
  if (config?.enabled === false) return;
  const current = menuItemsOf(chart.options) ?? FALLBACK_MENU_ITEMS;
  const hasDefinition = recAt(chart.options, 'exporting', 'menuItemDefinitions', key) !== undefined;
  if (current.includes(key) && hasDefinition) return;
  const icon = chartIcon(config, iconMarkup(options.menuIcon, 'menuIcon'));
  const label =
    nonEmptyString(config?.menuText) ??
    nonEmptyString(options.menuText) ??
    nonEmptyString(recAt(chart.options, 'lang')?.[langKey]) ??
    DEFAULT_MENU_TEXT;
  (chart.update as (o: Rec, redraw?: boolean) => void).call(
    chart,
    {
      exporting: {
        menuItemDefinitions: { [key]: createMenuItemDefinition(options, withIcon(icon, label), langKey) },
        buttons: { contextButton: { menuItems: insertMenuItem(current, key, options.insertAfter) } },
      },
    },
    true,
  );
}
