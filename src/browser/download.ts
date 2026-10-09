/**
 * Browser download helpers.
 */

import { abortedError, exportHighchartsToXlsx } from '../api/export';
import { ExportError, XLSX_MIME_TYPE, type ExportOptions, type ExportResult } from '../types/public-api';

/**
 * How long the object URL stays valid after the click. Revoking on the next tick can cancel the
 * download in some browsers (notably Safari and Firefox with large files), so it is kept for a minute.
 */
const REVOKE_DELAY_MS = 60_000;

/** A Blob part for `bytes`: the view itself when it spans its whole buffer, else a compact copy. */
function blobPart(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
  const whole =
    bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength && bytes.buffer instanceof ArrayBuffer;
  return (whole ? bytes : bytes.slice()) as Uint8Array<ArrayBuffer>;
}

/**
 * Saves `bytes` as a file through a temporary `<a download>` element.
 *
 * @throws ExportError BROWSER_REQUIRED when `document` or `URL.createObjectURL` is unavailable.
 */
export function triggerDownload(bytes: Uint8Array, filename: string, mimeType: string = XLSX_MIME_TYPE): void {
  if (
    typeof document === 'undefined' ||
    !document.body ||
    typeof URL === 'undefined' ||
    typeof URL.createObjectURL !== 'function' ||
    typeof Blob === 'undefined'
  ) {
    throw new ExportError('BROWSER_REQUIRED', 'Downloading requires a browser (document and URL.createObjectURL).');
  }
  const blob = new Blob([blobPart(bytes)], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const revoke = (): void => {
    if (typeof URL.revokeObjectURL === 'function') URL.revokeObjectURL(url);
  };
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.rel = 'noopener';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  let clicked = false;
  try {
    anchor.click();
    clicked = true;
  } finally {
    anchor.remove();
    if (clicked) setTimeout(revoke, REVOKE_DELAY_MS);
    else revoke();
  }
}

/**
 * Exports the chart and immediately downloads the workbook. Library errors propagate as
 * ExportError (see `exportHighchartsToXlsx`); errors thrown by your own callbacks (`onWarning`,
 * `onProgress`, `hooks.transformModel`) propagate unwrapped. Once `options.signal` is aborted it
 * rejects with ExportError ABORTED and never starts the download.
 */
export async function downloadHighchartsAsXlsx(chart: unknown, options?: ExportOptions): Promise<ExportResult> {
  const result = await exportHighchartsToXlsx(chart, options);
  const signal = options?.signal;
  if (signal?.aborted) throw abortedError(signal);
  triggerDownload(result.bytes, result.filename, result.mimeType);
  return result;
}
