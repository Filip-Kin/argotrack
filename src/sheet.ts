import 'dotenv/config';
import { GoogleSpreadsheet, GoogleSpreadsheetWorksheet, GoogleSpreadsheetRow } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import { getNow, localDateString } from './util';

// ─────────────────────────────────────────────────────────────
// Auth
// ─────────────────────────────────────────────────────────────
const serviceAccountAuth = new JWT({
    email: process.env.CLIENT_EMAIL,
    // Coolify/env may store the PEM with any depth of escaped newlines (\n, \\n,
    // even \\\\n after a copy through the API); collapse any run of backslashes
    // before an n into a real newline.
    key: (process.env.API_KEY || '').replace(/\\+n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID || '', serviceAccountAuth);

const USERS_HEADERS = [
    'pin', 'fname', 'lname', 'email', 'type', 'gender',
    'login', 'logout', 'hours', 'total', 'loggedin',
    'sessionType', 'sessionName',
    // Roster metadata copied from the Argonauts Apps Script app (portal/report use these).
    'studentType', 'department', 'driveTeam',
];
const SESSIONS_HEADERS = [
    'timestamp', 'pin', 'name', 'type', 'event', 'sessionType', 'eventName', 'hours',
];
// Coach-app data tabs (season minimums, requirement checklists, per-season meeting-day
// counts). Seasons/Helpers are simple header-row tabs; Checklist is a transposed grid
// (see readChecklist) so it only gets a title on self-provision, not a header row.
const SEASONS_HEADERS = ['name', 'start', 'end', 'minNew', 'minVeteran', 'minCaptain'];
const HELPERS_HEADERS = ['season', 'totalDays'];

let usersSheet: GoogleSpreadsheetWorksheet | undefined;
let logSheet: GoogleSpreadsheetWorksheet | undefined;
let sessionsSheet: GoogleSpreadsheetWorksheet | undefined;
let seasonsSheet: GoogleSpreadsheetWorksheet | undefined;
let checklistSheet: GoogleSpreadsheetWorksheet | undefined;
let helpersSheet: GoogleSpreadsheetWorksheet | undefined;
let connectPromise: Promise<boolean> | undefined;

type UsersRowData = {
    pin: string;
    fname: string;
    lname: string;
    email: string;
    type: string;
    gender: string;
    login: string;
    logout: string;
    hours: number;
    total: number;
    loggedin: string;
    sessionType: string;
    sessionName: string;
    studentType: string;
    department: string;
    driveTeam: string;
};

// ─────────────────────────────────────────────────────────────
// Connection (cached, self-healing)
// ─────────────────────────────────────────────────────────────
async function doConnect(): Promise<boolean> {
    await doc.loadInfo();

    // Self-provision: attaching a brand-new sheet (only "Sheet1") just works.
    // Reuse the first existing tab as Users; create Log/Sessions if missing.
    usersSheet = doc.sheetsByTitle['Users'];
    if (!usersSheet) {
        usersSheet = doc.sheetsByIndex[0];
        if (!usersSheet) throw new Error('Spreadsheet has no sheets');
        await usersSheet.updateProperties({ title: 'Users' });
    }
    logSheet = doc.sheetsByTitle['Log']
        || (await doc.addSheet({ title: 'Log', headerValues: ['pin', 'fname', 'lname'] }));
    sessionsSheet = doc.sheetsByTitle['Sessions']
        || (await doc.addSheet({ title: 'Sessions', headerValues: SESSIONS_HEADERS }));
    seasonsSheet = doc.sheetsByTitle['Seasons']
        || (await doc.addSheet({ title: 'Seasons', headerValues: SEASONS_HEADERS }));
    // Checklist is a transposed grid (deadline/req/detail/consequence header rows, then
    // one row per student). Just ensure the tab exists; readChecklist parses the layout.
    checklistSheet = doc.sheetsByTitle['Checklist']
        || (await doc.addSheet({ title: 'Checklist' }));
    helpersSheet = doc.sheetsByTitle['Helpers']
        || (await doc.addSheet({ title: 'Helpers', headerValues: HELPERS_HEADERS }));

    // Keep tab order Users(0), Log(1), Sessions(2), Seasons(3), Checklist(4), Helpers(5).
    await usersSheet.updateProperties({ index: 0 });
    await logSheet.updateProperties({ index: 1 });
    await sessionsSheet.updateProperties({ index: 2 });
    await seasonsSheet.updateProperties({ index: 3 });
    await checklistSheet.updateProperties({ index: 4 });
    await helpersSheet.updateProperties({ index: 5 });

    // Ensure header rows exist / are correct. Cheap and idempotent.
    await usersSheet.setHeaderRow(USERS_HEADERS);
    await logSheet.loadHeaderRow().catch(async () => {
        await logSheet!.setHeaderRow(['pin', 'fname', 'lname']);
    });
    await sessionsSheet.loadHeaderRow().catch(async () => {
        await sessionsSheet!.setHeaderRow(SESSIONS_HEADERS);
    });
    await seasonsSheet.loadHeaderRow().catch(async () => {
        await seasonsSheet!.setHeaderRow(SEASONS_HEADERS);
    });
    await helpersSheet.loadHeaderRow().catch(async () => {
        await helpersSheet!.setHeaderRow(HELPERS_HEADERS);
    });

    // Make sure every roster pin has a Log row (so per-day hours have a home).
    await ensureLogRows();

    // Lock the Log identity columns / header so they can't drift from Users.
    await ensureProtections().catch((e) =>
        console.error('[protections] could not apply (non-fatal):', e?.message || e)
    );

    return true;
}

/**
 * Connect once and cache the promise. If connecting fails, the cached promise
 * is cleared so the next call retries — this is what lets the app recover from
 * a transient Google outage instead of wedging until a manual restart.
 */
export async function connect(): Promise<boolean> {
    if (!connectPromise) {
        connectPromise = doConnect().catch((e) => {
            connectPromise = undefined; // allow retry on next call
            throw e;
        });
    }
    return connectPromise;
}

/** Force a reconnect on the next `connect()` (called after a failed API op). */
function invalidateConnection() {
    connectPromise = undefined;
}

/**
 * Ensure we're connected, reconnecting once if the cached connection is stale.
 * Retries only the CONNECT, never the caller's operation — safe for mutations.
 */
async function ensureConnected(): Promise<void> {
    try {
        await connect();
    } catch {
        invalidateConnection();
        await connect();
    }
}

/**
 * Run a READ-ONLY sheet operation; on failure, drop the connection and retry
 * once. Never use this for mutations — retrying a half-applied write (e.g. a
 * sign-out that already flipped `loggedin`) would double-apply it.
 */
async function withReconnect<T>(fn: () => Promise<T>): Promise<T> {
    try {
        await connect();
        return await fn();
    } catch (e: any) {
        console.error('[withReconnect] first attempt failed, reconnecting:', e?.message || e);
        invalidateConnection();
        await connect();
        return await fn();
    }
}

// ─────────────────────────────────────────────────────────────
// Roster / Log helpers
// ─────────────────────────────────────────────────────────────
/**
 * Keep the (protected, machine-owned) Log sheet's identity columns in sync with
 * the Users roster: add a Log row for any new pin, and propagate name changes.
 * Users is the single source of truth humans edit; Log is never edited by hand.
 */
async function ensureLogRows() {
    const userRows = await usersSheet!.getRows<UsersRowData>();
    const logRows = await logSheet!.getRows();

    const logByPin = new Map<string, (typeof logRows)[number]>();
    for (const r of logRows) {
        const pin = r.get('pin') ?? r['_rawData']?.[0];
        if (pin) logByPin.set(String(pin), r);
    }

    const toAdd: string[][] = [];
    for (const u of userRows) {
        const pin = u.get('pin');
        if (!pin) continue;
        const existing = logByPin.get(String(pin));
        if (!existing) {
            toAdd.push([pin, u.get('fname'), u.get('lname')]);
        } else {
            // Propagate name edits from the roster.
            const fname = u.get('fname') ?? '';
            const lname = u.get('lname') ?? '';
            if (existing.get('fname') !== fname || existing.get('lname') !== lname) {
                existing.set('fname', fname);
                existing.set('lname', lname);
                await existing.save();
            }
        }
    }
    if (toAdd.length) await logSheet!.addRows(toAdd);
}

/**
 * Lock the parts of the Log sheet that cause the "two pin columns drift apart"
 * breakage, while leaving the day cells open so mentors can mark attendance.
 *
 *  - Identity columns A:C (pin, fname, lname) → service-account only. These are
 *    maintained by the backend from the Users roster; humans never edit them.
 *  - Header row (the date labels) → service-account only, so date columns can't
 *    be renamed/reordered.
 *  - Everything else (the day cells, D2 onward) stays editable so mentors can set
 *    "E" for excused absences and the backend can write hours.
 *
 * Idempotent — each range is tagged by description.
 */
async function ensureProtections() {
    const ID = process.env.SPREADSHEET_ID;
    const editor = process.env.CLIENT_EMAIL;
    if (!ID || !editor || !logSheet) return;
    const sheetId = logSheet.sheetId;

    const ranges: { tag: string; range: any }[] = [
        {
            tag: 'argotrack:log-identity',
            range: { sheetId, startColumnIndex: 0, endColumnIndex: 3 }, // A:C, all rows
        },
        {
            tag: 'argotrack:log-header',
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 }, // row 1, all columns
        },
    ];

    const info: any = await serviceAccountAuth.request({
        url: `https://sheets.googleapis.com/v4/spreadsheets/${ID}` +
            `?fields=sheets(properties(sheetId),protectedRanges(description))`,
    });
    const logInfo = (info.data.sheets || []).find(
        (s: any) => s.properties.sheetId === sheetId
    );
    const present = new Set(
        (logInfo?.protectedRanges || []).map((p: any) => p.description)
    );

    const requests = ranges
        .filter((r) => !present.has(r.tag))
        .map((r) => ({
            addProtectedRange: {
                protectedRange: {
                    range: r.range,
                    description: r.tag,
                    warningOnly: false,
                    editors: { users: [editor] },
                },
            },
        }));

    if (requests.length) {
        await serviceAccountAuth.request({
            url: `https://sheets.googleapis.com/v4/spreadsheets/${ID}:batchUpdate`,
            method: 'POST',
            data: { requests },
        });
    }
}

/** Column index of today's date in the Log sheet, creating it if needed. */
async function getDateColumn(): Promise<number> {
    await logSheet!.loadHeaderRow();
    const dateStr = localDateString();

    const headers = logSheet!.headerValues;
    for (let i = 0; i < headers.length; i++) {
        if (headers[i] === dateStr) return i;
    }

    // Append a new date column at the end.
    const newCol = headers.length;
    if (newCol >= logSheet!.columnCount) {
        await logSheet!.resize({
            columnCount: newCol + 1,
            rowCount: logSheet!.rowCount,
        });
    }
    await logSheet!.loadCells();
    const cell = logSheet!.getCell(0, newCol);
    cell.value = dateStr; // plain ISO string keeps stats month-parsing simple
    await logSheet!.saveUpdatedCells();
    await logSheet!.loadHeaderRow();
    return newCol;
}

async function addHoursForDay(pin: string, hours: number, fname = '', lname = '') {
    // Make sure this pin has a Log row (a member added after startup won't yet).
    const logRows = await logSheet!.getRows();
    const hasRow = logRows.some(
        (r) => String(r.get('pin') ?? r['_rawData']?.[0]).trim() === String(pin).trim()
    );
    if (!hasRow) await logSheet!.addRow({ pin, fname, lname });

    const col = await getDateColumn();
    await logSheet!.loadCells();

    let row = -1;
    for (let i = 1; i < logSheet!.rowCount; i++) {
        if (String(logSheet!.getCell(i, 0).value).trim() === String(pin).trim()) {
            row = i;
            break;
        }
    }
    if (row === -1) throw new Error('Log row not found for pin ' + pin);

    const cell = logSheet!.getCell(row, col);
    const existing = typeof cell.numberValue === 'number' ? cell.numberValue : 0;
    cell.value = existing + hours;
    await logSheet!.saveUpdatedCells();
}

// ─────────────────────────────────────────────────────────────
// Public operations
// ─────────────────────────────────────────────────────────────
export async function getUserFromPin(pin: string) {
    return withReconnect(async () => {
        const rows = await usersSheet!.getRows<UsersRowData>();
        return rows.find((r) => String(r.get('pin')).trim() === String(pin).trim());
    });
}

export type PunchResult =
    | { success: false; message: string }
    | {
          success: true;
          event: 'IN' | 'OUT';
          name: string;
          message: string;
          sessionType?: string;
          eventName?: string;
          duration?: number;
          total?: number;
      };

// Old Apps Script version treated any open IN older than this as stale rather
// than a real session, so a forgotten sign-out doesn't bill someone for days.
const STALE_SESSION_MS = 24 * 60 * 60 * 1000;

// GAS ran each request to completion before the next one started, so two taps
// of the same PIN could never interleave. This async server can await mid-punch,
// so a double-tap (or a client retry) could otherwise read the same starting
// state twice and both write, dropping one of the two updates. Serialize
// punches per-pin to close that window; different pins still run concurrently.
const pinLocks = new Map<string, Promise<unknown>>();
function withPinLock<T>(pin: string, fn: () => Promise<T>): Promise<T> {
    const prior = pinLocks.get(pin) ?? Promise.resolve();
    const run = prior.then(fn, fn);
    pinLocks.set(pin, run.catch(() => {}));
    return run;
}

/** Toggle a user in/out. Records a raw Sessions row and, on OUT, per-day hours. */
export async function punch(pin: string, sessionType?: string, eventName?: string): Promise<PunchResult> {
    // Connection may reconnect once, but the toggle itself runs exactly once —
    // retrying a mutation would double-flip loggedin.
    await ensureConnected();

    return withPinLock(pin, () => punchLocked(pin, sessionType, eventName));
}

async function punchLocked(pin: string, sessionType?: string, eventName?: string): Promise<PunchResult> {
    {
        const rows = await usersSheet!.getRows<UsersRowData>();
        const user = rows.find((r) => String(r.get('pin')).trim() === String(pin).trim());
        if (!user) return { success: false, message: 'PIN not found' };

        const now = getNow();
        const name = `${user.get('fname')} ${user.get('lname')}`.trim();
        const type = user.get('type') || 'STUDENT';
        const loggedInFlag = String(user.get('loggedin')).toUpperCase() === 'TRUE';
        const loginAt = loggedInFlag ? new Date(user.get('login')) : null;
        const isStale = loggedInFlag && (!loginAt || isNaN(loginAt.getTime()) || now.getTime() - loginAt.getTime() >= STALE_SESSION_MS);
        const isIn = loggedInFlag && !isStale;

        if (!isIn) {
            // Sign IN
            user.set('login', now.toISOString());
            user.set('loggedin', 'TRUE');
            user.set('sessionType', sessionType || 'Meeting');
            user.set('sessionName', eventName || '');
            await user.save();

            await sessionsSheet!.addRow({
                timestamp: now.toISOString(),
                pin,
                name,
                type,
                event: 'IN',
                sessionType: sessionType || 'Meeting',
                eventName: eventName || '',
                hours: '',
            });

            return {
                success: true,
                event: 'IN',
                name,
                message: 'Signed in',
                sessionType: sessionType || 'Meeting',
                eventName: eventName || '',
            };
        } else {
            // Sign OUT
            const r = await finalizeSignOut(user, pin, sessionType, eventName);
            return {
                success: true,
                event: 'OUT',
                name: r.name,
                message: 'Signed out',
                sessionType: r.sType,
                eventName: r.sName,
                duration: r.hours,
                total: r.total,
            };
        }
    }
}

/** Perform the sign-out writes (Users row, Log day-hours, Sessions row) shared
 *  by punch()'s OUT branch, signOutAll(), and signOutOne(). */
async function finalizeSignOut(
    user: GoogleSpreadsheetRow<UsersRowData>,
    pin: string,
    sessionTypeOverride?: string,
    eventNameOverride?: string
): Promise<{ name: string; type: string; hours: number; total: number; sType: string; sName: string }> {
    const now = getNow();
    const loginAt = new Date(user.get('login'));
    const hours = Math.max(0, (now.getTime() - loginAt.getTime()) / 36e5);
    const total = parseFloat(String(user.get('total')) || '0') + hours;
    // Session type/name were captured at sign-in; fall back to any passed value.
    const sType = user.get('sessionType') || sessionTypeOverride || 'Meeting';
    const sName = user.get('sessionName') || eventNameOverride || '';
    const name = `${user.get('fname')} ${user.get('lname')}`.trim();
    const type = user.get('type') || 'STUDENT';

    user.set('logout', now.toISOString());
    user.set('hours', hours);
    user.set('total', total);
    user.set('loggedin', 'FALSE');
    await user.save();

    await addHoursForDay(pin, hours, user.get('fname'), user.get('lname'));

    await sessionsSheet!.addRow({
        timestamp: now.toISOString(),
        pin,
        name,
        type,
        event: 'OUT',
        sessionType: sType,
        eventName: sName,
        hours: hours.toFixed(4),
    });

    return { name, type, hours, total, sType, sName };
}

/** Names currently signed in, split into mentors and students. */
export async function getSignedIn(): Promise<{ mentors: string[]; students: string[] }> {
    return withReconnect(async () => {
        const rows = await usersSheet!.getRows<UsersRowData>();
        const mentors: string[] = [];
        const students: string[] = [];
        for (const r of rows) {
            if (String(r.get('loggedin')).toUpperCase() !== 'TRUE') continue;
            const name = displayName(r.get('fname'), r.get('lname'));
            if ((r.get('type') || '').toUpperCase() === 'MENTOR') mentors.push(name);
            else students.push(name);
        }
        mentors.sort();
        students.sort();
        return { mentors, students };
    });
}

/** "Firstname L." for the signed-in panel. */
function displayName(fname: string, lname: string): string {
    const f = (fname || '').trim();
    const l = (lname || '').trim();
    return l ? `${f} ${l.charAt(0)}.` : f;
}

/** Lightweight connectivity probe used by /health. Throws if the sheet is unreachable. */
export async function healthPing(): Promise<void> {
    await connect();
    await usersSheet!.loadCells('A1:A1');
}

// ─────────────────────────────────────────────────────────────
// Stats (mentor-only) — reads the date-pivot Log sheet
// ─────────────────────────────────────────────────────────────
interface MonthStat {
    hours: number;
    attended: number;
    excused: number;
    absent: number;
    attendance: number;
    attendanceWithExcused: number;
}
interface UserStats extends MonthStat {
    pin: string;
    name: string;
    months: { [key: number]: MonthStat };
    byType?: Record<string, number>;
}

function emptyMonth(): MonthStat {
    return { hours: 0, attended: 0, excused: 0, absent: 0, attendance: 0, attendanceWithExcused: 0 };
}
function pct(part: number, whole: number): number {
    return whole > 0 ? (part / whole) * 100 : 0;
}

export async function getStats() {
    return withReconnect(async () => {
        const stats: UserStats[] = [];
        const overall: Omit<UserStats, 'pin' | 'name'> = { ...emptyMonth(), months: {} };

        await logSheet!.loadHeaderRow();
        const headers = logSheet!.headerValues;
        const rows = await logSheet!.getRows();

        for (const row of rows) {
            const raw = row['_rawData'];
            const user: UserStats = {
                pin: raw?.[0],
                name: `${raw?.[1] ?? ''} ${raw?.[2] ?? ''}`.trim(),
                ...emptyMonth(),
                months: {},
            };

            for (let i = 3; i < headers.length; i++) {
                const header = headers[i];
                if (!header) continue;
                const monthNum = parseInt(String(header).split('-')[1]);
                if (isNaN(monthNum)) continue;

                if (!user.months[monthNum]) user.months[monthNum] = emptyMonth();
                if (!overall.months[monthNum]) overall.months[monthNum] = emptyMonth();

                const day = raw?.[i];
                const num = parseFloat(day);
                if (!isNaN(num) && num > 0) {
                    user.attended++; user.hours += num;
                    user.months[monthNum].attended++; user.months[monthNum].hours += num;
                    overall.attended++; overall.hours += num;
                    overall.months[monthNum].attended++; overall.months[monthNum].hours += num;
                } else if (day === 'E') {
                    user.excused++; user.months[monthNum].excused++;
                    overall.excused++; overall.months[monthNum].excused++;
                } else {
                    user.absent++; user.months[monthNum].absent++;
                    overall.absent++; overall.months[monthNum].absent++;
                }
            }

            user.attendance = pct(user.attended, user.attended + user.excused + user.absent);
            user.attendanceWithExcused = pct(user.attended + user.excused, user.attended + user.excused + user.absent);
            for (const m in user.months) {
                const mm = user.months[m];
                mm.attendance = pct(mm.attended, mm.attended + mm.excused + mm.absent);
                mm.attendanceWithExcused = pct(mm.attended + mm.excused, mm.attended + mm.excused + mm.absent);
            }
            stats.push(user);
        }

        overall.attendance = pct(overall.attended, overall.attended + overall.excused + overall.absent);
        overall.attendanceWithExcused = pct(overall.attended + overall.excused, overall.attended + overall.excused + overall.absent);
        for (const m in overall.months) {
            const mm = overall.months[m];
            mm.attendance = pct(mm.attended, mm.attended + mm.excused + mm.absent);
            mm.attendanceWithExcused = pct(mm.attended + mm.excused, mm.attended + mm.excused + mm.absent);
        }

        // Hours by session type (Meeting/Outreach/Competition), from the raw
        // Sessions log — surfaced on the stats page.
        const overallByType: Record<string, number> = {};
        const byPinType = new Map<string, Record<string, number>>();
        const typeSet = new Set<string>();
        const sessionRows = await sessionsSheet!.getRows();
        for (const r of sessionRows) {
            if (String(r.get('event')).toUpperCase() !== 'OUT') continue;
            const p = String(r.get('pin'));
            const st = (r.get('sessionType') || 'Meeting').trim() || 'Meeting';
            const h = parseFloat(r.get('hours')) || 0;
            typeSet.add(st);
            if (!byPinType.has(p)) byPinType.set(p, {});
            byPinType.get(p)![st] = (byPinType.get(p)![st] || 0) + h;
            overallByType[st] = (overallByType[st] || 0) + h;
        }
        const sessionTypes = Array.from(typeSet).sort();
        for (const u of stats) u.byType = byPinType.get(String(u.pin)) || {};
        (overall as any).byType = overallByType;

        return { overall, stats, sessionTypes };
    });
}

// ═════════════════════════════════════════════════════════════
// Argonauts coach-app parity: seasons, checklist, reports, portal.
// Everything here computes from the append-only Sessions log + the
// Seasons/Checklist/Helpers config tabs — no change to the punch/Log path.
// Ported from the coach's Google Apps Script (argo-attendance-code-gs.txt).
// ═════════════════════════════════════════════════════════════

/** A completed (OUT) attendance event, normalized for reporting. */
type OutEvent = { pin: string; date: Date; hours: number; sessionType: string };

async function loadOutEvents(): Promise<OutEvent[]> {
    const rows = await sessionsSheet!.getRows();
    const out: OutEvent[] = [];
    for (const r of rows) {
        if (String(r.get('event')).toUpperCase() !== 'OUT') continue;
        const ts = new Date(r.get('timestamp'));
        if (isNaN(ts.getTime())) continue;
        out.push({
            pin: String(r.get('pin')).trim(),
            date: ts,
            hours: parseFloat(r.get('hours')) || 0,
            sessionType: (r.get('sessionType') || 'Meeting').trim() || 'Meeting',
        });
    }
    return out;
}

type Season = {
    name: string; start: Date; end: Date;
    minNew: number; minVeteran: number; minCaptain: number;
};

async function loadSeasons(): Promise<Season[]> {
    const rows = await seasonsSheet!.getRows();
    const seasons: Season[] = [];
    for (const r of rows) {
        const name = String(r.get('name') || '').trim();
        if (!name) continue;
        const start = new Date(r.get('start'));
        const end = new Date(r.get('end'));
        end.setHours(23, 59, 59, 999);
        seasons.push({
            name, start, end,
            minNew: parseFloat(r.get('minNew')) || 0,
            minVeteran: parseFloat(r.get('minVeteran')) || 0,
            minCaptain: parseFloat(r.get('minCaptain')) || 0,
        });
    }
    return seasons;
}

async function loadHelperDays(): Promise<Record<string, number>> {
    const rows = await helpersSheet!.getRows();
    const map: Record<string, number> = {};
    for (const r of rows) {
        const name = String(r.get('season') || '').trim();
        const days = parseFloat(r.get('totalDays')) || 0;
        if (name) map[name] = days;
    }
    return map;
}

/** yyyy-MM-dd from a date-only value (UTC parts avoid tz off-by-one). */
function ymd(d: Date): string {
    if (isNaN(d.getTime())) return '';
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

// Thresholds for the checklist's numeric requirements, from the Checklist tab's
// own "Details" column text (which varies by student type, so can't be read as
// a single sheet cell): Volunteer Hours "15 hrs rookies / 30 hrs veterans / 40
// hrs captain eligibility", Fall Meeting Attendance "New: 80% | Returning: 60%
// | Captains: 80%", Build Season Attendance "New: 50% | Returning: 60% |
// Captains: 75%".
function checklistThreshold(itemName: string, studentType: string): number | null {
    const n = itemName.toLowerCase();
    if (n.includes('volunteer') && n.includes('hour')) {
        return studentType === 'New' ? 15 : studentType === 'Captain' ? 40 : 30;
    }
    if (n.includes('fall') && (n.includes('attendance') || n.includes('meeting'))) {
        return studentType === 'Captain' ? 80 : studentType === 'New' ? 80 : 60;
    }
    if (n.includes('build') && n.includes('attendance')) {
        return studentType === 'Captain' ? 75 : studentType === 'New' ? 50 : 60;
    }
    return null;
}

/** Map a raw checklist cell to a status + display value (ported verbatim). */
function parseStatusValue(raw: string, threshold: number | null): { status: string; displayValue: string } {
    const lower = String(raw || '').toLowerCase().trim();
    if (lower === 'n/a' || lower === 'na') return { status: 'na', displayValue: 'N/A' };
    if (lower === 'yes' || lower === 'y' || lower === 'complete') return { status: 'complete', displayValue: 'Complete' };
    if (lower === 'no' || lower === 'n' || lower === '') return { status: 'incomplete', displayValue: '' };
    if (lower === 'partial' || lower === 'in progress') return { status: 'partial', displayValue: raw };
    const num = parseFloat(raw);
    if (!isNaN(num)) {
        if (threshold !== null && !isNaN(threshold)) {
            if (num >= threshold) return { status: 'complete', displayValue: String(raw) };
            return { status: num > 0 ? 'partial' : 'incomplete', displayValue: String(raw) };
        }
        return { status: num > 0 ? 'partial' : 'incomplete', displayValue: String(raw) };
    }
    return { status: 'incomplete', displayValue: raw };
}

type ChecklistItem = { col: number; name: string; detail: string; deadline: string; consequence: string };
type ChecklistData = {
    items: ChecklistItem[];
    deadlines: { name: string; consequence: string; items: ChecklistItem[] }[];
    grid: any[][];
    pinRow: Map<string, number>;
};

/**
 * Read the transposed Checklist grid:
 *   row1 deadline group (merged) · row2 requirement · row3 detail · row4 consequence
 *   row5+ one row per student: col A=PIN, col B=Name, col C+ = status per requirement.
 */
async function readChecklist(): Promise<ChecklistData> {
    await checklistSheet!.loadCells();
    const rowCount = checklistSheet!.rowCount;
    const colCount = checklistSheet!.columnCount;
    const grid: any[][] = [];
    for (let r = 0; r < rowCount; r++) {
        const rowArr: any[] = [];
        for (let c = 0; c < colCount; c++) {
            const v = checklistSheet!.getCell(r, c).value;
            rowArr.push(v === null || v === undefined ? '' : v);
        }
        grid.push(rowArr);
    }

    const items: ChecklistItem[] = [];
    let lastDl = '';
    for (let c = 2; c < colCount; c++) {
        const reqName = String(grid[1]?.[c] ?? '').trim();
        if (!reqName) continue;
        const dl = String(grid[0]?.[c] ?? '').trim() || lastDl;
        if (dl) lastDl = dl;
        items.push({
            col: c,
            name: reqName,
            detail: String(grid[2]?.[c] ?? '').trim(),
            deadline: dl,
            consequence: String(grid[3]?.[c] ?? '').trim(),
        });
    }

    const deadlines: ChecklistData['deadlines'] = [];
    let cur: ChecklistData['deadlines'][number] | null = null;
    for (const it of items) {
        if (!cur || cur.name !== it.deadline) {
            cur = { name: it.deadline, consequence: it.consequence, items: [] };
            deadlines.push(cur);
        }
        cur.items.push(it);
    }

    const pinRow = new Map<string, number>();
    for (let r = 4; r < rowCount; r++) {
        const p = String(grid[r]?.[0] ?? '').trim();
        if (p) pinRow.set(p, r);
    }

    return { items, deadlines, grid, pinRow };
}

/** Seasons list for the coach dashboard dropdown. */
export async function getSeasons() {
    return withReconnect(async () => {
        const seasons = (await loadSeasons()).map((s) => ({
            name: s.name,
            startDate: ymd(s.start),
            endDate: ymd(s.end),
            minNew: s.minNew,
            minVeteran: s.minVeteran,
            minCaptain: s.minCaptain,
        }));
        return { success: true, seasons };
    });
}

/** Roster with each member's current IN/OUT status. */
export async function getRoster() {
    return withReconnect(async () => {
        const rows = await usersSheet!.getRows<UsersRowData>();
        const students = rows
            .filter((r) => String(r.get('pin')).trim())
            .map((r) => ({
                name: `${r.get('fname')} ${r.get('lname')}`.trim(),
                pin: String(r.get('pin')).trim(),
                role: (r.get('type') || '').toUpperCase() === 'MENTOR' ? 'Mentor' : 'Student',
                status: String(r.get('loggedin')).toUpperCase() === 'TRUE' ? 'IN' : 'OUT',
                sessionType: r.get('sessionType') || '',
                login: r.get('login') || '',
            }));
        return { success: true, students };
    });
}

/** Date-range attendance/hours report for the coach dashboard. */
export async function getReport(startDate: string, endDate: string) {
    return withReconnect(async () => {
        const start = new Date(startDate);
        const end = new Date(endDate);
        end.setHours(23, 59, 59, 999);

        const userRows = await usersSheet!.getRows<UsersRowData>();
        const outs = await loadOutEvents();

        const map = new Map<string, {
            name: string; pin: string; role: string; totalMins: number;
            sessions: number; days: Set<string>; byType: Record<string, number>;
        }>();
        for (const u of userRows) {
            const pin = String(u.get('pin')).trim();
            if (!pin) continue;
            map.set(pin, {
                name: `${u.get('fname')} ${u.get('lname')}`.trim(),
                pin,
                role: (u.get('type') || '').toUpperCase() === 'MENTOR' ? 'Mentor' : 'Student',
                totalMins: 0, sessions: 0, days: new Set<string>(),
                byType: { Meeting: 0, Outreach: 0, Competition: 0 },
            });
        }

        const allDays = new Set<string>();
        for (const e of outs) {
            if (e.date < start || e.date > end) continue;
            allDays.add(localDateString(e.date));
            const s = map.get(e.pin);
            if (!s) continue;
            const mins = e.hours * 60;
            s.totalMins += mins;
            s.sessions += 1;
            s.days.add(localDateString(e.date));
            if (s.byType[e.sessionType] === undefined) s.byType[e.sessionType] = 0;
            s.byType[e.sessionType] += mins;
        }
        const totalMeetingDays = allDays.size;

        const students = Array.from(map.values()).map((s) => ({
            name: s.name,
            pin: s.pin,
            role: s.role,
            totalMins: Math.round(s.totalMins),
            sessions: s.sessions,
            daysPresent: s.days.size,
            attendancePct: totalMeetingDays > 0 ? Math.round((s.days.size / totalMeetingDays) * 100) : 0,
            avgSession: s.sessions > 0 ? Math.round(s.totalMins / s.sessions) : 0,
            meetingMins: Math.round(s.byType.Meeting || 0),
            outreachMins: Math.round(s.byType.Outreach || 0),
            competitionMins: Math.round(s.byType.Competition || 0),
        }));
        students.sort((a, b) => b.totalMins - a.totalMins);

        return { success: true, students, totalMeetingDays, startDate, endDate };
    });
}

/** Per-student requirement completion grid for the coach dashboard. */
export async function getChecklistOverview() {
    return withReconnect(async () => {
        const cl = await readChecklist();
        const userRows = await usersSheet!.getRows<UsersRowData>();
        const rosterMap = new Map<string, { name: string; role: string; studentType: string }>();
        for (const u of userRows) {
            const pin = String(u.get('pin')).trim();
            if (!pin) continue;
            rosterMap.set(pin, {
                name: `${u.get('fname')} ${u.get('lname')}`.trim(),
                role: (u.get('type') || '').toUpperCase() === 'MENTOR' ? 'Mentor' : 'Student',
                studentType: u.get('studentType') || 'Veteran',
            });
        }

        const requirements = cl.items.map((it) => ({ col: it.col, name: it.name, deadline: it.deadline }));
        const students: any[] = [];
        for (const [pin, rowIdx] of cl.pinRow) {
            const info = rosterMap.get(pin) || { name: pin, role: 'Student', studentType: 'Veteran' };
            let complete = 0, incomplete = 0, partial = 0, na = 0;
            const statuses: string[] = [];
            for (const it of cl.items) {
                const raw = String(cl.grid[rowIdx]?.[it.col] ?? '').trim();
                const st = parseStatusValue(raw, checklistThreshold(it.name, info.studentType)).status;
                if (st === 'complete') complete++;
                else if (st === 'na') na++;
                else if (st === 'partial') partial++;
                else incomplete++;
                statuses.push(st);
            }
            const total = complete + incomplete + partial; // exclude n/a
            students.push({
                pin, name: info.name, role: info.role, studentType: info.studentType,
                complete, incomplete, partial, na, total,
                pct: total > 0 ? Math.round((complete / total) * 100) : 0,
                statuses,
            });
        }
        students.sort((a, b) => a.pct - b.pct); // least complete first
        return { success: true, students, requirements };
    });
}

/** Captain-eligibility policy, ported from the coach app (team-specific thresholds). */
function computeCaptainEligibility(
    pin: string, cl: ChecklistData, rowIdx: number | undefined,
    mine: OutEvent[], seasons: Season[], helperDays: Record<string, number>, outs: OutEvent[]
) {
    let volHours = 0, fallAttPct = 0, buildAttPct = 0;
    let appStatus = 'incomplete', interviewStatus = 'incomplete';

    if (rowIdx !== undefined) {
        for (const it of cl.items) {
            const rn = it.name.toLowerCase();
            const rv = String(cl.grid[rowIdx]?.[it.col] ?? '').trim().toLowerCase();
            const isMet = rv === 'yes' || rv === 'y' || rv === 'complete';
            const numVal = parseFloat(rv) || 0;
            if (rn.includes('volunteer') && rn.includes('hour')) volHours = numVal;
            if (rn.includes('fall') && (rn.includes('attendance') || rn.includes('meeting'))) fallAttPct = numVal;
            if (rn.includes('build') && rn.includes('attendance')) buildAttPct = numVal;
            if (rn.includes('application') && rn.includes('captain')) appStatus = isMet ? 'complete' : 'incomplete';
            if (rn.includes('interview')) interviewStatus = isMet ? 'complete' : 'incomplete';
        }
    }

    // Fall back to the Sessions log when the checklist doesn't carry the numbers.
    if (volHours === 0) {
        let volMins = 0;
        for (const e of mine) volMins += e.hours * 60;
        volHours = Math.round((volMins / 60) * 10) / 10;
    }
    if (fallAttPct === 0 || buildAttPct === 0) {
        for (const s of seasons) {
            const nm = s.name.toLowerCase();
            let totalDays = helperDays[s.name] || 0;
            if (totalDays === 0) {
                const allDays = new Set<string>();
                for (const e of outs) if (e.date >= s.start && e.date <= s.end) allDays.add(localDateString(e.date));
                totalDays = allDays.size;
            }
            if (totalDays === 0) continue;
            const myDays = new Set<string>();
            for (const e of mine) if (e.date >= s.start && e.date <= s.end) myDays.add(localDateString(e.date));
            const pctC = Math.round((myDays.size / totalDays) * 100);
            if (nm.includes('fall') && fallAttPct === 0) fallAttPct = pctC;
            if (nm.includes('build') && buildAttPct === 0) buildAttPct = pctC;
        }
    }

    const volMet = volHours >= 40;
    let attSum = 0, attCount = 0;
    if (fallAttPct > 0) { attSum += fallAttPct; attCount++; }
    if (buildAttPct > 0) { attSum += buildAttPct; attCount++; }
    const avgAtt = attCount > 0 ? Math.round(attSum / attCount) : 0;
    const attMet = avgAtt >= 75;
    const allMet = volMet && attMet && appStatus === 'complete' && interviewStatus === 'complete';

    return {
        deadline: 'May 2027',
        overall: allMet,
        criteria: [
            { name: 'Volunteer Hours', detail: 'Must log 40+ volunteer hours from March 2026 - December 2026', status: volMet ? 'complete' : (volHours > 0 ? 'partial' : 'incomplete'), displayValue: volHours + ' hrs', target: '40 hrs' },
            { name: 'Attendance Average', detail: 'Average of Fall Training + Build Season attendance must be 75% or higher (Sept 2026 - Feb 2027)', status: attMet ? 'complete' : (avgAtt > 0 ? 'partial' : 'incomplete'), displayValue: avgAtt + '%', target: '75%' },
            { name: 'Captain Application', detail: 'Google Form application must be submitted by the date provided (typically April/May)', status: appStatus, displayValue: appStatus === 'complete' ? 'Submitted' : '', target: 'Submitted' },
            { name: 'Coach Interview', detail: 'Sit through an interview with coaches about the position(s) you are applying for (typically April/May)', status: interviewStatus, displayValue: interviewStatus === 'complete' ? 'Complete' : '', target: 'Complete' },
        ],
    };
}

/** Student self-service portal: season hours, attendance, checklists, captain eligibility. */
export async function getStudentPortal(pin: string) {
    return withReconnect(async () => {
        pin = String(pin).trim();
        const userRows = await usersSheet!.getRows<UsersRowData>();
        const u = userRows.find((r) => String(r.get('pin')).trim() === pin);
        if (!u) return { success: false, message: 'PIN not recognized.' };

        const role = (u.get('type') || '').toUpperCase() === 'MENTOR' ? 'Mentor' : 'Student';
        const studentType = u.get('studentType') || 'Veteran';
        const name = `${u.get('fname')} ${u.get('lname')}`.trim();

        const seasons = await loadSeasons();
        const helperDays = await loadHelperDays();
        const outs = await loadOutEvents();
        const mine = outs.filter((e) => e.pin === pin);

        // Hours per season vs the required minimum for this student type.
        const seasonBlocks = seasons.map((s) => {
            let minHours = 0;
            if (role !== 'Mentor') {
                minHours = studentType === 'New' ? s.minNew : studentType === 'Captain' ? s.minCaptain : s.minVeteran;
            }
            let actualMins = 0;
            for (const e of mine) if (e.date >= s.start && e.date <= s.end) actualMins += e.hours * 60;
            return { name: s.name, actualMins: Math.round(actualMins), minMins: Math.round(minHours * 60) };
        });

        // Attendance per season (students only).
        const attendance: any[] = [];
        for (const s of seasons) {
            if (role === 'Mentor') continue;
            const minPct = studentType === 'New' ? 80 : studentType === 'Captain' ? 80 : 60;
            const myDays = new Set<string>();
            const allDays = new Set<string>();
            for (const e of outs) {
                if (e.date < s.start || e.date > s.end) continue;
                allDays.add(localDateString(e.date));
                if (e.pin === pin) myDays.add(localDateString(e.date));
            }
            let totalDays = helperDays[s.name] || 0;
            if (totalDays === 0) totalDays = allDays.size;
            const attPct = totalDays > 0 ? Math.round((myDays.size / totalDays) * 100) : 0;
            attendance.push({ seasonName: s.name, attendancePct: attPct, minPct, presentDays: myDays.size, totalDays });
        }

        // Requirement checklists grouped by deadline.
        const cl = await readChecklist();
        const rowIdx = cl.pinRow.get(pin);
        const deadlines = cl.deadlines.map((dl) => ({
            name: dl.name,
            consequence: dl.consequence,
            items: dl.items.map((it) => {
                const raw = rowIdx !== undefined ? String(cl.grid[rowIdx]?.[it.col] ?? '').trim() : '';
                const threshold = checklistThreshold(it.name, studentType);
                const parsed = parseStatusValue(raw, threshold);
                return { name: it.name, detail: it.detail, status: parsed.status, displayValue: parsed.displayValue, threshold };
            }),
        }));

        const captainEligibility = role !== 'Mentor'
            ? computeCaptainEligibility(pin, cl, rowIdx, mine, seasons, helperDays, outs)
            : null;

        return {
            success: true,
            name, role, studentType,
            department: u.get('department') || '',
            driveTeam: u.get('driveTeam') || '',
            seasons: seasonBlocks,
            attendance,
            deadlines,
            captainEligibility,
        };
    });
}

/** Sign out everyone currently in (coach dashboard button + kiosk PIN 9999). */
export async function signOutAll(): Promise<{ count: number; names: string[] }> {
    await ensureConnected();
    const rows = await usersSheet!.getRows<UsersRowData>();
    const names: string[] = [];
    for (const user of rows) {
        if (String(user.get('loggedin')).toUpperCase() !== 'TRUE') continue;
        const pin = String(user.get('pin')).trim();
        const r = await finalizeSignOut(user, pin);
        names.push(r.name);
    }
    return { count: names.length, names };
}

/** Sign out one specific person by PIN (mentor portal's per-row Sign Out button,
 *  password-gated at the route level). No-ops if they're not currently signed in. */
export async function signOutOne(pin: string): Promise<{ success: boolean; name?: string; message?: string }> {
    await ensureConnected();
    return withPinLock(pin, async () => {
        const rows = await usersSheet!.getRows<UsersRowData>();
        const user = rows.find((r) => String(r.get('pin')).trim() === String(pin).trim());
        if (!user) return { success: false, message: 'PIN not found.' };
        if (String(user.get('loggedin')).toUpperCase() !== 'TRUE') {
            return { success: false, message: 'Not currently signed in.' };
        }
        const r = await finalizeSignOut(user, pin);
        return { success: true, name: r.name };
    });
}
