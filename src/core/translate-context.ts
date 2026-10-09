/**
 * Shared state of one `translateChartModel` run, defined once and passed to every builder
 * (axes, series, plot groups, chart chrome): the model, the fidelity flag, the two diagnostic
 * sinks, the supported-property recorder, and the text/font helpers that report into them.
 */

import type { ChartModel, Font, TextBlock } from '../types/chart-model';
import type { Diagnostic, DiagnosticCode, DiagnosticCollector } from '../types/diagnostics';
import type { FidelityMode } from '../types/public-api';
import type { ExcelFontSpec, ExcelTextSpec } from '../excel/writer-interface';
import { stripControlChars } from '../utils/text';
import { toExcelFont } from './style-mapping';

/** Diagnostic codes that only concern styling (dropped in 'minimal' fidelity). */
const STYLE_CODES: ReadonlySet<DiagnosticCode> = new Set<DiagnosticCode>([
  'APPROXIMATED_FONT',
  'APPROXIMATED_FONT_SIZE',
  'APPROXIMATED_MARKER',
  'APPROXIMATED_COLOR',
  'UNSUPPORTED_GRADIENT',
  'APPROXIMATED_DASH_STYLE',
  'UNSUPPORTED_STYLE',
]);

export interface TranslateContext {
  readonly model: ChartModel;
  /** Fidelity flag: true unless the fidelity is 'minimal' (styles are then left to Excel's defaults). */
  readonly best: boolean;
  /** Diagnostics that are always reported. */
  readonly out: Diagnostic[];
  /** Style-only diagnostics; in 'minimal' fidelity the `STYLE_CODES` among them are dropped on flush. */
  readonly styleSink: Diagnostic[];
  /** Supported source properties, in first-report order (deduplicated in the result). */
  readonly supported: string[];
  /** Records supported source properties. */
  support(...paths: string[]): void;
  /** Records supported style properties (only in 'best' fidelity). */
  supportStyle(...paths: string[]): void;
  /** Font → Excel font (null in 'minimal' fidelity); reports into `styleSink`. */
  fontOf(font: Font | null, property: string): ExcelFontSpec | null;
  /** Text block → Excel rich text, optionally followed by extra lines in their own font. */
  textSpec(
    tb: TextBlock | null,
    property: string,
    extra?: { lines: string[]; font: ExcelFontSpec | null } | null,
  ): ExcelTextSpec | null;
  /** Moves both sinks into the collector (filtering style codes in 'minimal' fidelity) and clears them. */
  flush(diagnostics: DiagnosticCollector): void;
}

/** Non-empty, trimmed lines of a text block (control characters stripped, `<br>` splits lines). */
export function textOf(tb: TextBlock): string[] {
  return stripControlChars(tb.text)
    .split(/\r?\n|<br\s*\/?>/i)
    .map((l) => l.trim())
    .filter((l) => l !== '');
}

export function createTranslateContext(model: ChartModel, fidelity: FidelityMode): TranslateContext {
  const best = fidelity !== 'minimal';
  const out: Diagnostic[] = [];
  const styleSink: Diagnostic[] = [];
  const supported: string[] = [];
  const support = (...paths: string[]): void => {
    for (const p of paths) supported.push(p);
  };
  const fontOf = (font: Font | null, property: string): ExcelFontSpec | null =>
    best ? toExcelFont(font, property, styleSink) : null;
  return {
    model,
    best,
    out,
    styleSink,
    supported,
    support,
    supportStyle: (...paths: string[]): void => {
      if (best) support(...paths);
    },
    fontOf,
    textSpec: (tb, property, extra = null) => {
      if (!tb) return null;
      const own = textOf(tb);
      const lines = [...own, ...(extra?.lines ?? [])];
      if (lines.length === 0) return null;
      const spec: ExcelTextSpec = { lines, font: fontOf(tb.font, `${property}.style`), overlay: false };
      if (extra && extra.lines.length > 0 && extra.font) {
        spec.lineFonts = [...own.map(() => null), ...extra.lines.map(() => extra.font)];
      }
      return spec;
    },
    flush: (diagnostics) => {
      diagnostics.addAll(out);
      if (best) diagnostics.addAll(styleSink.filter(Boolean));
      else diagnostics.addAll(styleSink.filter((d) => !STYLE_CODES.has(d.code)));
      out.length = 0;
      styleSink.length = 0;
    },
  };
}
