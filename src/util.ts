// Timezone-aware time helpers.
//
// The original beartime hard-subtracted 5 hours to fake "local" time, which is
// wrong half the year (DST) and only affected which date-column a punch lands
// in. Hours math (out - in) is unaffected by any offset since it cancels, so we
// keep timestamps in real UTC and only localize the *date bucket*.

const TIMEZONE = process.env.TIMEZONE || 'America/New_York';

/** Real current time (UTC-based Date). */
export function getNow(): Date {
    return new Date();
}

/** YYYY-MM-DD for `date` in the configured timezone (used for the Log columns). */
export function localDateString(date: Date = new Date()): string {
    // en-CA gives ISO-style YYYY-MM-DD.
    return new Intl.DateTimeFormat('en-CA', {
        timeZone: TIMEZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).format(date);
}

/**
 * Google Sheets serial date number for a YYYY-MM-DD string (days since the
 * 1899-12-30 epoch). Written into the Log header cells so the sheet renders a
 * real date, matching the original behavior.
 */
export function sheetSerialForDate(dateStr: string): number {
    const [y, m, d] = dateStr.split('-').map(Number);
    const utcMs = Date.UTC(y, m - 1, d);
    return Math.floor(utcMs / 8.64e7) + 25569;
}
