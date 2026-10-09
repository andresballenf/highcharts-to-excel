import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ExportError, XLSX_MIME_TYPE } from '../../src/types/public-api';
import { triggerDownload } from '../../src/browser/download';

describe('triggerDownload', () => {
  let blobs: Blob[];
  let clicked: HTMLAnchorElement[];
  let createObjectURL: ReturnType<typeof vi.fn>;
  let revokeObjectURL: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    blobs = [];
    clicked = [];
    createObjectURL = vi.fn((b: Blob) => {
      blobs.push(b);
      return 'blob:mock-url';
    });
    revokeObjectURL = vi.fn();
    // jsdom has no createObjectURL: stub a URL-like global for the duration of each test.
    vi.stubGlobal('URL', { createObjectURL, revokeObjectURL });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      expect(this.isConnected).toBe(true);
      clicked.push(this);
    });
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('downloads through a temporary anchor and revokes the URL afterwards', async () => {
    const bytes = new Uint8Array([0x50, 0x4b, 3, 4]);
    triggerDownload(bytes, 'report.xlsx');
    expect(blobs).toHaveLength(1);
    expect(blobs[0]?.type).toBe(XLSX_MIME_TYPE);
    expect(new Uint8Array(await blobs[0]!.arrayBuffer())).toEqual(bytes);
    expect(clicked).toHaveLength(1);
    const a = clicked[0]!;
    expect(a.getAttribute('download')).toBe('report.xlsx');
    expect(a.getAttribute('href')).toBe('blob:mock-url');
    expect(a.isConnected).toBe(false);
    expect(document.querySelectorAll('a[download]')).toHaveLength(0);
    expect(revokeObjectURL).not.toHaveBeenCalled();
    vi.runAllTimers();
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:mock-url');
  });

  it('A7 keeps the object URL alive for 60 s before revoking it', () => {
    triggerDownload(new Uint8Array([1, 2]), 'slow.xlsx');
    vi.advanceTimersByTime(59_999);
    expect(revokeObjectURL).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:mock-url');
  });

  it('A7 hands whole-buffer bytes to the Blob without copying and copies a sub-view', () => {
    const parts: unknown[][] = [];
    vi.stubGlobal(
      'Blob',
      class {
        readonly type: string;
        constructor(p: unknown[], o: { type: string }) {
          parts.push(p);
          this.type = o.type;
        }
      },
    );
    const whole = new Uint8Array([1, 2, 3, 4]);
    triggerDownload(whole, 'whole.xlsx');
    expect(parts[0]?.[0]).toBe(whole);
    const sub = new Uint8Array(new ArrayBuffer(16), 4, 4);
    sub.set([9, 8, 7, 6]);
    triggerDownload(sub, 'sub.xlsx');
    const copied = parts[1]?.[0] as Uint8Array;
    expect(copied).not.toBe(sub);
    expect(Array.from(copied)).toEqual([9, 8, 7, 6]);
    expect(copied.byteLength).toBe(copied.buffer.byteLength);
  });

  it('uses a custom MIME type', () => {
    triggerDownload(new Uint8Array([1]), 'x.bin', 'application/octet-stream');
    expect(blobs[0]?.type).toBe('application/octet-stream');
  });

  it('revokes synchronously and removes the anchor when click throws', () => {
    vi.mocked(HTMLAnchorElement.prototype.click).mockImplementation(() => {
      throw new Error('blocked');
    });
    expect(() => triggerDownload(new Uint8Array([1]), 'a.xlsx')).toThrow('blocked');
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:mock-url');
    expect(document.querySelectorAll('a[download]')).toHaveLength(0);
  });

  it('throws BROWSER_REQUIRED without document', () => {
    vi.stubGlobal('document', undefined);
    let error: unknown;
    try {
      triggerDownload(new Uint8Array([1]), 'a.xlsx');
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExportError);
    expect((error as ExportError).code).toBe('BROWSER_REQUIRED');
    expect(createObjectURL).not.toHaveBeenCalled();
  });

  it('throws BROWSER_REQUIRED without URL.createObjectURL', () => {
    vi.stubGlobal('URL', { revokeObjectURL });
    expect(() => triggerDownload(new Uint8Array([1]), 'a.xlsx')).toThrow(
      expect.objectContaining({ code: 'BROWSER_REQUIRED' }),
    );
  });
});
