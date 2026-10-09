/**
 * Chart time zone handling for datetime x values.
 *
 * Highcharts stores datetime values as UTC timestamps and DISPLAYS them in the chart's time zone
 * (`time.timezone`, or the legacy `useUTC: false` / `timezoneOffset` / `getTimezoneOffset`).
 * Excel has no time zones: a serial date is shown as stored. So every datetime x value is moved
 * to the wall-clock time the chart displays, written as a UTC timestamp:
 *
 *   wall = ts - time.getTimezoneOffset(ts)
 *
 * `getTimezoneOffset` returns milliseconds, positive WEST of UTC (as `Date#getTimezoneOffset`,
 * but in ms): Asia/Tokyo gives -9 h, so Tokyo midnight (15:00Z the day before) becomes 00:00Z.
 */

import { bool, num, rec, str, type Rec } from './guards';

export interface TimeZoneInfo {
  /** Milliseconds to ADD to a UTC timestamp to get the displayed wall-clock time; null when unknown. */
  shiftAt: (ts: number) => number | null;
  /** Human-readable zone, e.g. "UTC", "Asia/Tokyo", "local", "UTC+09:00". */
  label: string;
  /** False when the zone could not be determined; values are then exported unshifted. */
  known: boolean;
}

const UTC_ZONE: TimeZoneInfo = { shiftAt: () => 0, label: 'UTC', known: true };
/** Offsets only change on 15-minute boundaries, so lookups are cached per 15-minute bucket. */
const BUCKET_MS = 15 * 60_000;

function offsetLabel(shiftMs: number): string {
  const m = Math.round(Math.abs(shiftMs) / 60_000);
  const hh = String(Math.floor(m / 60)).padStart(2, '0');
  const mm = String(m % 60).padStart(2, '0');
  return `UTC${shiftMs < 0 ? '-' : '+'}${hh}:${mm}`;
}

function fixedZone(shiftMs: number): TimeZoneInfo {
  return shiftMs === 0 ? UTC_ZONE : { shiftAt: () => shiftMs, label: offsetLabel(shiftMs), known: true };
}

function unknownZone(label: string): TimeZoneInfo {
  return { shiftAt: () => null, label, known: false };
}

const localZone: TimeZoneInfo = {
  shiftAt: (ts) => {
    const o = new Date(ts).getTimezoneOffset();
    return Number.isFinite(o) ? -o * 60_000 || 0 : null;
  },
  label: 'local',
  known: true,
};

/** IANA zone through Intl; null when the runtime does not know the zone. */
function intlZone(zone: string): TimeZoneInfo | null {
  let fmt: Intl.DateTimeFormat;
  try {
    fmt = new Intl.DateTimeFormat('en-US', {
      timeZone: zone,
      hourCycle: 'h23',
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
    });
  } catch {
    return null;
  }
  return {
    label: zone,
    known: true,
    shiftAt: (ts) => {
      if (!Number.isFinite(ts)) return null;
      const parts: Record<string, string> = {};
      for (const p of fmt.formatToParts(new Date(ts))) parts[p.type] = p.value;
      const wall = Date.UTC(
        Number(parts.year),
        Number(parts.month) - 1,
        Number(parts.day),
        Number(parts.hour) % 24,
        Number(parts.minute),
        Number(parts.second),
      );
      const wholeSeconds = ts - (((ts % 1000) + 1000) % 1000);
      const shift = wall - wholeSeconds;
      return Number.isFinite(shift) ? shift : null;
    },
  };
}

/**
 * Time zone of a plain options object, with Highcharts 12/13 precedence: `useUTC` (legacy) decides
 * between UTC and the local zone and overrides `timezone`; a non-zero `timezoneOffset` (minutes,
 * positive west) overrides both; a `getTimezoneOffset` callback (Highcharts 11) is used when set.
 * Default: UTC.
 */
export function timeZoneFromOptions(options: Rec): TimeZoneInfo {
  const t = rec(options.time) ?? {};
  const useUTC = bool(t.useUTC);
  const tzo = num(t.timezoneOffset);
  if (useUTC === false) return localZone;
  if (tzo !== undefined && tzo !== 0) return fixedZone(-tzo * 60_000);
  const callback = t.getTimezoneOffset;
  if (typeof callback === 'function' && t.timezone === undefined) {
    return {
      label: 'getTimezoneOffset',
      known: true,
      shiftAt: (ts) => {
        try {
          const o = num((callback as (x: number) => unknown)(ts));
          return o === undefined ? null : -o * 60_000 || 0;
        } catch {
          return null;
        }
      },
    };
  }
  const zone = useUTC === true ? 'UTC' : (str(t.timezone) ?? 'UTC');
  if (zone === 'UTC' || zone === 'Etc/UTC' || zone === 'GMT') return UTC_ZONE;
  return intlZone(zone) ?? unknownZone(zone);
}

/** Time zone of a live chart: asks `chart.time.getTimezoneOffset` (Highcharts 11-13). */
export function timeZoneFromChart(chart: { time?: unknown; options?: unknown }): TimeZoneInfo {
  const time = chart.time as { getTimezoneOffset?: unknown; timezone?: unknown; options?: unknown } | undefined;
  const fn = time?.getTimezoneOffset;
  if (typeof fn !== 'function') return timeZoneFromOptions(rec(chart.options) ?? {});
  const tOpts = rec(time?.options) ?? {};
  const label =
    str(time?.timezone) ?? str(tOpts.timezone) ?? (tOpts.useUTC === false ? 'local' : undefined) ?? 'chart time zone';
  return {
    label,
    known: true,
    shiftAt: (ts) => {
      try {
        const o = (fn as (x: number) => unknown).call(time, ts);
        return typeof o === 'number' && Number.isFinite(o) ? -o || 0 : null;
      } catch {
        return null;
      }
    },
  };
}

const ISO_DATE =
  /^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3})\d*)?)?)?\s*(Z|[+-]\d{2}(?::?\d{2})?)?$/i;

/**
 * Converts datetime x values to displayed wall-clock time and records the offset applied.
 * One instance per extraction.
 */
export class DatetimeShifter {
  /** Offset (ms) applied to the first datetime point; null until one is seen. */
  private firstShiftMs: number | null = null;
  private seen = false;
  private nonZero = false;
  private readonly cache = new Map<number, number | null>();

  constructor(readonly zone: TimeZoneInfo) {}

  private shiftOf(ts: number): number | null {
    if (this.zone === UTC_ZONE) return 0;
    const key = Math.floor(ts / BUCKET_MS);
    let v = this.cache.get(key);
    if (v === undefined && !this.cache.has(key)) {
      v = this.zone.shiftAt(ts);
      this.cache.set(key, v);
    }
    return v ?? null;
  }

  private record(shift: number | null): void {
    if (!this.seen) {
      this.seen = true;
      this.firstShiftMs = shift;
    }
    if (shift !== null && shift !== 0) this.nonZero = true;
  }

  /** Wall-clock value (UTC ms) of a point's timestamp. */
  point(ts: number): number {
    const shift = this.shiftOf(ts);
    this.record(shift);
    return shift === null ? ts : ts + shift;
  }

  /** A point x that is already a wall-clock time (date string, calendar step): recorded, unchanged. */
  wallPoint(wall: number): number {
    this.record(this.shiftOf(wall));
    return wall;
  }

  /** Wall-clock value of an axis bound (not counted as a point). */
  bound(ts: number): number {
    const shift = this.shiftOf(ts);
    return shift === null ? ts : ts + shift;
  }

  /**
   * Parses a date string the way Highcharts 12+ does on datetime axes: without an explicit zone
   * it is a wall-clock time in the chart's zone (returned as is, in UTC ms); with `Z` or `±hh:mm`
   * it is an instant, moved to wall-clock time. Undefined when the string is not a date.
   */
  parse(s: string, asPoint = true): number | undefined {
    const text = s.trim();
    const m = ISO_DATE.exec(text);
    let wall: number | undefined;
    if (m) {
      const [, y, mo, d, h, mi, sec, ms, zone] = m;
      const utc = Date.UTC(
        Number(y),
        Number(mo) - 1,
        Number(d ?? 1),
        Number(h ?? 0),
        Number(mi ?? 0),
        Number(sec ?? 0),
        Number((ms ?? '0').padEnd(3, '0')),
      );
      if (!Number.isFinite(utc)) return undefined;
      if (zone !== undefined) {
        let offsetMin = 0;
        if (zone.toUpperCase() !== 'Z') {
          const sign = zone.startsWith('-') ? -1 : 1;
          const digits = zone.slice(1).replace(':', '');
          offsetMin = sign * (Number(digits.slice(0, 2)) * 60 + Number(digits.slice(2) || 0));
        }
        const instant = utc - offsetMin * 60_000;
        return asPoint ? this.point(instant) : this.bound(instant);
      }
      wall = utc;
    } else {
      const parsed = Date.parse(`${text} UTC`);
      if (!Number.isFinite(parsed)) return undefined;
      wall = parsed;
    }
    return asPoint ? this.wallPoint(wall) : wall;
  }

  /**
   * Offset in minutes added to the first datetime point: 0 when there was none or the zone is UTC,
   * null when the zone is unknown.
   */
  offsetMinutes(): number | null {
    if (!this.zone.known) return null;
    if (!this.seen) return 0;
    return this.firstShiftMs === null ? null : Math.round(this.firstShiftMs / 60_000) || 0;
  }

  /** True when some datetime value was moved (the chart is not in UTC). */
  get shifted(): boolean {
    return this.nonZero;
  }

  /** True when datetime values were seen but the zone was unknown. */
  get unresolved(): boolean {
    return this.seen && (!this.zone.known || this.firstShiftMs === null);
  }
}
