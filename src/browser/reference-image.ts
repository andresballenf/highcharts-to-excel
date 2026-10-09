/**
 * Opt-in reference image: a PNG rendering of the source chart placed next to the native Excel chart
 * for side-by-side comparison. Browser only (needs SVG rasterization through <canvas>).
 *
 * Never throws: any failure becomes a WRITER_LIMITATION diagnostic and the export continues
 * without the image.
 */

import type { DiagnosticCollector } from '../types/diagnostics';
import { isRealBrowser } from '../highcharts/guards';

export interface ReferenceImage {
  png: Uint8Array;
  /** Size of the image on the sheet, in CSS pixels (the PNG itself is rendered at 2x). */
  widthPx: number;
  heightPx: number;
}

/** Diagnostic property used for every reference-image diagnostic. */
export const REFERENCE_IMAGE_PROPERTY = 'includeReferenceImage';

const SCALE = 2;
const LOAD_TIMEOUT_MS = 10_000;

type GetSvg = (chartOptions?: unknown) => unknown;

function findGetSvg(chart: unknown): GetSvg | null {
  if (typeof chart !== 'object' || chart === null) return null;
  const c = chart as { getSVG?: unknown; exporting?: { getSVG?: unknown } };
  if (typeof c.getSVG === 'function') return (opts) => (c.getSVG as (o?: unknown) => unknown).call(chart, opts);
  const exp = c.exporting;
  if (exp && typeof exp.getSVG === 'function') return (opts) => (exp.getSVG as (o?: unknown, a?: boolean) => unknown).call(exp, opts, false);
  return null;
}

/** Reports that the reference image could not be produced. */
export function reportReferenceImageUnavailable(diagnostics: DiagnosticCollector, reason: string, severity: 'info' | 'warning' = 'warning'): void {
  diagnostics.report('WRITER_LIMITATION', 'unsupported', REFERENCE_IMAGE_PROPERTY, `Reference image omitted: ${reason}`, { severity });
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const timer = setTimeout(() => reject(new Error('timed out loading the SVG rendering')), LOAD_TIMEOUT_MS);
    img.onload = () => {
      clearTimeout(timer);
      resolve(img);
    };
    img.onerror = () => {
      clearTimeout(timer);
      reject(new Error('the browser could not load the SVG rendering'));
    };
    img.src = url;
  });
}

function canvasToPng(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error('canvas.toBlob returned no data'))), 'image/png');
  });
}

/**
 * Renders the chart to PNG via `chart.getSVG()` → <img> → <canvas> (2x) → PNG bytes.
 * Returns null (after reporting a diagnostic) when not possible.
 */
export async function renderReferenceImage(
  chart: unknown,
  widthPx: number,
  heightPx: number,
  diagnostics: DiagnosticCollector,
): Promise<ReferenceImage | null> {
  if (!isRealBrowser()) {
    reportReferenceImageUnavailable(diagnostics, 'rendering the reference image requires a browser.');
    return null;
  }
  const getSvg = findGetSvg(chart);
  if (!getSvg) {
    reportReferenceImageUnavailable(diagnostics, 'chart.getSVG is not available (load highcharts/modules/exporting).');
    return null;
  }
  const w = Math.max(1, Math.round(widthPx));
  const h = Math.max(1, Math.round(heightPx));
  let url: string | null = null;
  try {
    const svg = await getSvg({ chart: { width: w, height: h } });
    if (typeof svg !== 'string' || svg === '') throw new Error('chart.getSVG returned no SVG markup');
    url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }));
    const img = await loadImage(url);
    const canvas = document.createElement('canvas');
    canvas.width = w * SCALE;
    canvas.height = h * SCALE;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('2D canvas context unavailable');
    ctx.scale(SCALE, SCALE);
    ctx.drawImage(img, 0, 0, w, h);
    const blob = await canvasToPng(canvas);
    const png = new Uint8Array(await blob.arrayBuffer());
    if (png.length === 0) throw new Error('empty PNG');
    return { png, widthPx: w, heightPx: h };
  } catch (error) {
    reportReferenceImageUnavailable(diagnostics, `rendering failed (${error instanceof Error ? error.message : String(error)}).`);
    return null;
  } finally {
    if (url !== null) URL.revokeObjectURL(url);
  }
}
