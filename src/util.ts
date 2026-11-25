// Time helpers.
//
// The coach's sheet stores every date and time as a Google Sheets serial number
// (days since 1899-12-30) in the SPREADSHEET's timezone, i.e. wall-clock time.
// We keep that convention: date-range filters and day buckets compare serials
// directly (no timezone math needed), and only durations convert a serial back
// to a real instant so a session that spans a DST change is still measured right.

const DAY_MS = 86_400_000;
const SERIAL_UNIX_EPOCH = 25569; // serial number of 1970-01-01

export const DEFAULT_TIMEZONE = process.env.TIMEZONE || 'America/New_York';

/** Offset (ms) of `tz` from UTC at the given instant. Positive east of UTC. */
function tzOffsetMs(instantMs: number, tz: string): number {
    const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: tz,
        hourCycle: 'h23',
        year: 'numeric', month: 'numeric', day: 'numeric',
        hour: 'numeric', minute: 'numeric', second: 'numeric',
    }).formatToParts(new Date(instantMs));
    const get = (type: string): number => Number(parts.find((p) => p.type === type)?.value ?? 0);
    const wallAsUtc = Date.UTC(get('year'), get('month') - 1, get('day'), get('hour'), get('minute'), get('second'));
    return wallAsUtc - Math.floor(instantMs / 1000) * 1000;
}

/** Sheet serial (wall clock in `tz`) for a real instant. */
export function instantToSerial(date: Date, tz: string): number {
    return (date.getTime() + tzOffsetMs(date.getTime(), tz)) / DAY_MS + SERIAL_UNIX_EPOCH;
}

/** Real instant for a sheet serial (wall clock in `tz`). */
export function serialToInstant(serial: number, tz: string): Date {
    const wall = Math.round((serial - SERIAL_UNIX_EPOCH) * DAY_MS);
    const first = wall - tzOffsetMs(wall, tz);
    const second = wall - tzOffsetMs(first, tz);
    return new Date(second);
}

/** The wall-clock moment of a serial as a Date whose UTC fields hold the wall time. */
export function serialToWallDate(serial: number): Date {
    return new Date(Math.round((serial - SERIAL_UNIX_EPOCH) * DAY_MS));
}

/** yyyy-MM-dd for the day a serial falls on. */
export function serialToYmd(serial: number): string {
    return serialToWallDate(Math.floor(serial)).toISOString().slice(0, 10);
}

/** Serial of midnight on a yyyy-MM-dd date, or NaN. */
export function ymdToSerial(ymd: string): number {
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(ymd).trim());
    if (!m) return NaN;
    return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / DAY_MS + SERIAL_UNIX_EPOCH;
}

/**
 * Serial for a cell that should hold a date or date-time. Date cells arrive as
 * numbers; a cell someone typed as text is parsed as wall-clock time. NaN if neither.
 */
export function cellToSerial(value: unknown): number {
    if (typeof value === 'number') return value;
    const text = String(value ?? '').trim();
    if (!text) return NaN;
    const d = new Date(text);
    if (isNaN(d.getTime())) return NaN;
    return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds())
        / DAY_MS + SERIAL_UNIX_EPOCH;
}

/** "2h 5m" / "45 min" / "3h", as the coach's app prints a session length. */
export function formatDuration(hours: number): string {
    let h = Math.floor(hours);
    let m = Math.round((hours - h) * 60);
    if (m === 60) { h += 1; m = 0; }
    if (h === 0) return m + ' min';
    if (m === 0) return h + 'h';
    return h + 'h ' + m + 'm';
}
