import 'dotenv/config';
import { GoogleSpreadsheet, GoogleSpreadsheetWorksheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import { getNow, localDateString } from './util';

// ─────────────────────────────────────────────────────────────
// Auth
// ─────────────────────────────────────────────────────────────
const serviceAccountAuth = new JWT({
    email: process.env.CLIENT_EMAIL,
    // Coolify/env may store the PEM with \n or \\n escapes; normalize to real newlines.
    key: (process.env.API_KEY || '').replace(/\\{1,2}n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID || '', serviceAccountAuth);

const USERS_HEADERS = [
    'pin', 'fname', 'lname', 'email', 'type', 'gender',
    'login', 'logout', 'hours', 'total', 'loggedin',
    'sessionType', 'sessionName',
];
const SESSIONS_HEADERS = [
    'timestamp', 'pin', 'name', 'type', 'event', 'sessionType', 'eventName', 'hours',
];

let usersSheet: GoogleSpreadsheetWorksheet | undefined;
let logSheet: GoogleSpreadsheetWorksheet | undefined;
let sessionsSheet: GoogleSpreadsheetWorksheet | undefined;
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

    // Keep tab order Users(0), Log(1), Sessions(2) — ops read/write by index.
    await usersSheet.updateProperties({ index: 0 });
    await logSheet.updateProperties({ index: 1 });
    await sessionsSheet.updateProperties({ index: 2 });

    // Ensure header rows exist / are correct. Cheap and idempotent.
    await usersSheet.setHeaderRow(USERS_HEADERS);
    await logSheet.loadHeaderRow().catch(async () => {
        await logSheet!.setHeaderRow(['pin', 'fname', 'lname']);
    });
    await sessionsSheet.loadHeaderRow().catch(async () => {
        await sessionsSheet!.setHeaderRow(SESSIONS_HEADERS);
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
        (r) => String(r.get('pin') ?? r['_rawData']?.[0]) === String(pin)
    );
    if (!hasRow) await logSheet!.addRow({ pin, fname, lname });

    const col = await getDateColumn();
    await logSheet!.loadCells();

    let row = -1;
    for (let i = 1; i < logSheet!.rowCount; i++) {
        if (String(logSheet!.getCell(i, 0).value) === String(pin)) {
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
        return rows.find((r) => String(r.get('pin')) === String(pin));
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

/** Toggle a user in/out. Records a raw Sessions row and, on OUT, per-day hours. */
export async function punch(pin: string, sessionType?: string, eventName?: string): Promise<PunchResult> {
    // Connection may reconnect once, but the toggle itself runs exactly once —
    // retrying a mutation would double-flip loggedin.
    await ensureConnected();

    {
        const rows = await usersSheet!.getRows<UsersRowData>();
        const user = rows.find((r) => String(r.get('pin')) === String(pin));
        if (!user) return { success: false, message: 'PIN not found' };

        const now = getNow();
        const name = `${user.get('fname')} ${user.get('lname')}`.trim();
        const type = user.get('type') || 'STUDENT';
        const isIn = String(user.get('loggedin')).toUpperCase() === 'TRUE';

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
            const loginAt = new Date(user.get('login'));
            const hours = Math.max(0, (now.getTime() - loginAt.getTime()) / 36e5);
            const total = parseFloat(String(user.get('total')) || '0') + hours;
            // Session type/name were captured at sign-in; fall back to any passed value.
            const sType = user.get('sessionType') || sessionType || 'Meeting';
            const sName = user.get('sessionName') || eventName || '';

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

            return {
                success: true,
                event: 'OUT',
                name,
                message: 'Signed out',
                sessionType: sType,
                eventName: sName,
                duration: hours,
                total,
            };
        }
    }
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

        return { overall, stats };
    });
}
