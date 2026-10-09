/**
 * Browser download helpers.
 */

import { exportHighchartsToXlsx } from '../api/export';
import { ExportError, XLSX_MIME_TYPE, type ExportOptions, type ExportResult } from '../types/public-api';

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
  // Copy into a fresh ArrayBuffer-backed view so the Blob never sees a shared/oversized buffer.
  const blob = new Blob([bytes.slice()], { type: mimeType });
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
    if (clicked) setTimeout(revoke, 0);
    else revoke();
  }
}

/** Exports the chart and immediately downloads the workbook. Errors propagate as ExportError. */
export async function downloadHighchartsAsXlsx(chart: unknown, options?: ExportOptions): Promise<ExportResult> {
  const result = await exportHighchartsToXlsx(chart, options);
  triggerDownload(result.bytes, result.filename, result.mimeType);
  return result;
}
