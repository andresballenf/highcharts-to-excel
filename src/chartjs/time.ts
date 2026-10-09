/**
 * Time values on a Chart.js time scale → the wall-clock milliseconds Excel should show.
 *
 * Chart.js date adapters (date-fns, luxon, moment, dayjs) display time in the browser's local
 * zone by default. Excel has no time zones, so absolute instants (numbers, Dates, ISO strings with
 * an offset) are moved to the local wall-clock time; zone-less date strings already are wall-clock
 * times and are kept as written.
 */

const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;
const LOCAL_DATE_TIME = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d{1,3})?)?$/;

export class WallClock {
  private firstShiftMinutes: number | null = null;
  private nonZero = false;
  private seen = false;

  /** Absolute instant → local wall-clock ms. */
  fromInstant(ms: number): number {
    const offsetMinutes = -new Date(ms).getTimezoneOffset();
    this.seen = true;
    this.firstShiftMinutes ??= offsetMinutes;
    if (offsetMinutes !== 0) this.nonZero = true;
    return ms + offsetMinutes * 60_000;
  }

  /**
   * Parses a Chart.js time value (ms number, Date, ISO/`Date.parse`-able string) into wall-clock
   * ms. Returns null when it cannot be parsed without the chart's date adapter.
   */
  parse(value: unknown): number | null {
    if (typeof value === 'number') return Number.isFinite(value) ? this.fromInstant(value) : null;
    if (value instanceof Date) {
      const t = value.getTime();
      return Number.isFinite(t) ? this.fromInstant(t) : null;
    }
    if (typeof value !== 'string') return null;
    const s = value.trim();
    if (s === '') return null;
    if (DATE_ONLY.test(s)) {
      const t = Date.parse(`${s}T00:00:00Z`);
      return Number.isFinite(t) ? t : null;
    }
    if (LOCAL_DATE_TIME.test(s)) {
      const t = Date.parse(`${s.replace(' ', 'T')}Z`);
      return Number.isFinite(t) ? t : null;
    }
    if (/^-?\d+(\.\d+)?$/.test(s)) return this.fromInstant(Number(s));
    const t = Date.parse(s);
    return Number.isFinite(t) ? this.fromInstant(t) : null;
  }

  /** Minutes added to the first absolute instant; 0 when nothing moved or no instant was seen. */
  offsetMinutes(): number {
    return this.seen ? (this.firstShiftMinutes ?? 0) : 0;
  }

  /** True when some value was moved (the browser is not on UTC). */
  get shifted(): boolean {
    return this.nonZero;
  }
}
