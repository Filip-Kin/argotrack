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

    usersSheet = doc.sheetsByIndex[0];
    logSheet = doc.sheetsByIndex[1];
    sessionsSheet = doc.sheetsByIndex[2];
    if (!usersSheet || !logSheet || !sessionsSheet) {
        throw new Error('Required sheets not found (need Users, Log, Sessions tabs)');
    }

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

/** Run a sheet operation; on failure, drop the connection and retry once. */
async function withReconnect<T>(fn: () => Promise<T>): Promise<T> {
    try {
        await connect();
        return await fn();
    } catch (e) {
        invalidateConnection();
        await connect();
        return await fn();
    }
}

// ─────────────────────────────────────────────────────────────
// Roster / Log helpers
// ─────────────────────────────────────────────────────────────
async function ensureLogRows() {
    const userRows = await usersSheet!.getRows<UsersRowData>();
    const logRows = await logSheet!.getRows();

    const logPins = new Set<string>();
    for (const r of logRows) {
        const pin = r.get('pin') ?? r['_rawData']?.[0];
        if (pin) logPins.add(String(pin));
    }

    const toAdd: string[][] = [];
    for (const u of userRows) {
        const pin = u.get('pin');
        if (pin && !logPins.has(String(pin))) {
            toAdd.push([pin, u.get('fname'), u.get('lname')]);
        }
    }
    if (toAdd.length) await logSheet!.addRows(toAdd);
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

async function addHoursForDay(pin: string, hours: number) {
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
    return withReconnect(async () => {
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

            await addHoursForDay(pin, hours);

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
    });
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
