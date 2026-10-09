/**
 * Unit conversions between CSS pixels and OOXML units (EMU, points, hundredths of a point).
 * 1 px = 9525 EMU = 0.75 pt at 96 DPI.
 */

export const EMU_PER_PX = 9525;
const PT_PER_PX = 0.75;

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToPt(px: number): number {
  return px * PT_PER_PX;
}

export function ptToPx(pt: number): number {
  return pt / PT_PER_PX;
}

/** Font size for OOXML `sz` attributes (hundredths of a point). */
export function pxToHundredthsPt(px: number): number {
  return Math.round(px * 75);
}

/** Line width for OOXML `a:ln w` (EMU). */
export function pxToLineWidthEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

export function clamp(n: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, n));
}

export function roundTo(n: number, decimals: number): number {
  const f = 10 ** decimals;
  return Math.round(n * f) / f;
}
