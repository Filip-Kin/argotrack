// Attendance logic on the coach's spreadsheet.
//
// The spreadsheet layout is the coach's and stays exactly as she built it:
//   Roster    Name | PIN | Role | Student Type | Department | ... (her own columns)
//   Log       Timestamp | Student Name | PIN | Event | Session Duration (minutes) |
//             Session Type | Event Name | IP Address      (one row per IN / OUT)
//   Summary   formulas over Log and Roster. Never written here.
//   Seasons   Season Name | Start Date | End Date | Min Hours New / Veteran / Captain
//   Helpers   Season | Total Meetings | Total Available Minutes
//   Checklist row1 deadline group, row2 requirement, row3 detail, row4 consequence,
//             row5+ one row per student (col A PIN, col B name, col C+ status)
//
// Who is signed in is derived from each PIN's last Log row, as her Apps Script
// does, so this app and her script can run on the same sheet at the same time.
//
// The one tab this app adds is `Attendance` (see the Attendance region): a
// PIN-by-day grid where mentors mark excused absences with "E".

import {
    Cell, Grid, Row, SheetMeta, TabInfo,
    appendRows, at, batchUpdate, blank, clearValues, columnLetter, dateTimeCell, loadMeta,
    num, numberCell, readValues, tabRange, text, textCell, writeValues,
} from './gs';
import {
    DEFAULT_TIMEZONE, cellToSerial, formatDuration, instantToSerial, serialToInstant, serialToYmd, ymdToSerial,
} from './util';

const TAB = {
    roster: 'Roster',
    log: 'Log',
    summary: 'Summary',
    seasons: 'Seasons',
    helpers: 'Helpers',
    checklist: 'Checklist',
    attendance: 'Attendance',
} as const;

// #region Connection

let metaPromise: Promise<SheetMeta> | undefined;

async function fetchMeta(): Promise<SheetMeta> {
    const meta = await loadMeta();
    for (const required of [TAB.roster, TAB.log]) {
        if (!meta.tabs.some((t) => t.title === required)) {
            throw new Error(`Spreadsheet has no "${required}" tab`);
        }
    }
    return meta;
}

/** Spreadsheet metadata (timezone, tab ids), cached. A failed load is not cached. */
export function connect(): Promise<SheetMeta> {
    if (!metaPromise) {
        metaPromise = fetchMeta().catch((e) => {
            metaPromise = undefined;
            throw e;
        });
    }
    return metaPromise;
}

/** Drop the cached metadata (after this app adds or copies a tab). */
function invalidateMeta(): void {
    metaPromise = undefined;
}

export function findTab(meta: SheetMeta, title: string): TabInfo | undefined {
    return meta.tabs.find((t) => t.title === title);
}

export function timeZoneOf(meta: SheetMeta): string {
    return meta.timeZone || DEFAULT_TIMEZONE;
}

/** Lightweight connectivity probe used by /health. Throws if the sheet is unreachable. */
export async function healthPing(): Promise<void> {
    await connect();
    await readValues([tabRange(TAB.roster, 'A1')]);
}

// #endregion

// #region PINs, roster and log

/**
 * A 4-digit PIN string, or '' when the value is blank or not a valid PIN. The
 * sheet holds PINs both as text ("0716") and as numbers (716); this makes them
 * comparable. A blank must never turn into "0000".
 */
export function normPin(pin: unknown): string {
    const s = String(pin === undefined || pin === null ? '' : pin).trim();
    if (!/^\d{1,4}$/.test(s)) return '';
    return s.padStart(4, '0');
}

export type StudentType = 'New' | 'Veteran' | 'Captain';

/** Accepts "New", "new ", "Rookie", "First Year", "Captain", ...; default Veteran. */
export function normStudentType(raw: unknown): StudentType {
    const t = String(raw === undefined || raw === null ? '' : raw).trim().toLowerCase();
    if (!t) return 'Veteran';
    if (t.includes('captain')) return 'Captain';
    if (t.includes('new') || t.includes('rookie') || t.includes('first') || t.includes('1st')) return 'New';
    return 'Veteran';
}

export type Member = {
    /** 1-based row in the Roster tab. */
    sheetRow: number;
    name: string;
    /** Normalized 4-digit PIN, '' when the roster row has none yet. */
    pin: string;
    role: 'Mentor' | 'Student';
    rawStudentType: string;
    studentType: StudentType;
    department: string;
    driveTeam: string;
};

export function parseRoster(grid: Grid): Member[] {
    // "Drive Team Eligible" is found by its header so the coach can move her own columns.
    const header = grid[0] || [];
    const driveCol = header.findIndex((h) => String(h).toLowerCase().includes('drive team'));
    const members: Member[] = [];
    for (let i = 1; i < grid.length; i++) {
        const name = text(grid[i], 0);
        if (!name) continue;
        members.push({
            sheetRow: i + 1,
            name,
            pin: normPin(at(grid[i], 1)),
            role: text(grid[i], 2).toLowerCase() === 'mentor' ? 'Mentor' : 'Student',
            rawStudentType: text(grid[i], 3),
            studentType: normStudentType(at(grid[i], 3)),
            department: text(grid[i], 4),
            driveTeam: driveCol >= 0 ? text(grid[i], driveCol) : '',
        });
    }
    return members;
}

type LogEntry = {
    /** Wall-clock timestamp as a sheet serial; NaN if the cell is not a date. */
    serial: number;
    name: string;
    pin: string;
    event: string;
    minutes: number;
    sessionType: string;
    eventName: string;
};

function parseLog(grid: Grid): LogEntry[] {
    const entries: LogEntry[] = [];
    for (let i = 1; i < grid.length; i++) {
        const pin = normPin(at(grid[i], 2));
        const name = text(grid[i], 1);
        // A punch always writes a name. A row without one is not a person
        // (the sheet has an old blank-PIN row) and must not show as signed in.
        if (!pin || !name) continue;
        entries.push({
            serial: cellToSerial(at(grid[i], 0)),
            name,
            pin,
            event: text(grid[i], 3),
            minutes: num(grid[i], 4),
            sessionType: text(grid[i], 5),
            eventName: text(grid[i], 6),
        });
    }
    return entries;
}

type Core = { meta: SheetMeta; tz: string; roster: Member[]; log: LogEntry[] };

async function loadCore(): Promise<Core> {
    const meta = await connect();
    const [rosterGrid, logGrid] = await readValues([tabRange(TAB.roster), tabRange(TAB.log, 'A:H')]);
    return { meta, tz: timeZoneOf(meta), roster: parseRoster(rosterGrid), log: parseLog(logGrid) };
}

/** Last Log row per PIN, in sheet order. */
function lastEventByPin(log: LogEntry[]): Map<string, LogEntry> {
    const last = new Map<string, LogEntry>();
    for (const e of log) last.set(e.pin, e);
    return last;
}

/** Roster member for a PIN (undefined if the PIN is not on the roster). */
export async function getMember(pin: string): Promise<Member | undefined> {
    const p = normPin(pin);
    if (!p) return undefined;
    await connect();
    const [rosterGrid] = await readValues([tabRange(TAB.roster)]);
    return parseRoster(rosterGrid).find((m) => m.pin === p);
}

type LogWrite = {
    at: Date; name: string; pin: string; event: 'IN' | 'OUT';
    minutes: number | null; sessionType: string; eventName: string; ip: string;
};

/** Append IN / OUT rows to the Log tab in the coach's column order. */
async function appendLog(core: Core, rows: LogWrite[]): Promise<void> {
    const tab = findTab(core.meta, TAB.log);
    if (!tab) throw new Error('Log tab not found');
    await appendRows(tab.sheetId, rows.map((r) => [
        dateTimeCell(instantToSerial(r.at, core.tz)),
        textCell(r.name),
        // Numeric, as the existing rows are: the Summary and Checklist formulas match on it.
        numberCell(Number(r.pin)),
        textCell(r.event),
        r.minutes === null ? blank : numberCell(r.minutes),
        textCell(r.sessionType),
        textCell(r.eventName),
        textCell(r.ip),
    ]));
}

// #endregion

// #region Punch in / out

export type PunchResult =
    | { success: false; message: string }
    | {
          success: true;
          event: 'IN' | 'OUT';
          name: string;
          message: string;
          sessionType: string;
          eventName: string;
          duration?: number;
      };

// An IN older than this is treated as stale rather than a real session, so a
// forgotten sign-out does not bill someone for days.
const STALE_SESSION_MS = 24 * 60 * 60 * 1000;

// Two taps of the same PIN must not interleave: both would read the same last
// Log row and both would write. Punches are serialized per PIN; different PINs
// still run concurrently.
const pinLocks = new Map<string, Promise<unknown>>();
function withPinLock<T>(pin: string, fn: () => Promise<T>): Promise<T> {
    const prior = pinLocks.get(pin) ?? Promise.resolve();
    const run = prior.then(fn, fn);
    pinLocks.set(pin, run.catch(() => undefined));
    return run;
}

const NOT_RECOGNIZED = 'PIN not recognized. Please try again.';

/** Toggle a person in or out by appending one Log row. */
export async function punch(pinRaw: string, sessionType?: string, eventName?: string, ip?: string): Promise<PunchResult> {
    const pin = normPin(pinRaw);
    if (!pin) return { success: false, message: NOT_RECOGNIZED };
    return withPinLock(pin, async (): Promise<PunchResult> => {
        const core = await loadCore();
        const member = core.roster.find((m) => m.pin === pin);
        if (!member) return { success: false, message: NOT_RECOGNIZED };

        const now = new Date();
        const last = lastEventByPin(core.log).get(pin);
        const signedInAt = last && last.event === 'IN' && !isNaN(last.serial)
            ? serialToInstant(last.serial, core.tz)
            : null;
        const isSignedIn = signedInAt !== null && now.getTime() - signedInAt.getTime() < STALE_SESSION_MS;

        if (isSignedIn && last && signedInAt) {
            const hours = Math.max(0, (now.getTime() - signedInAt.getTime()) / 3_600_000);
            await appendLog(core, [{
                at: now, name: member.name, pin, event: 'OUT', minutes: Math.round(hours * 60),
                sessionType: last.sessionType, eventName: last.eventName, ip: ip || '',
            }]);
            return {
                success: true, event: 'OUT', name: member.name, duration: hours,
                sessionType: last.sessionType, eventName: last.eventName,
                message: 'Signed out successfully. Session: ' + formatDuration(hours),
            };
        }

        const type = sessionType || 'Meeting';
        await appendLog(core, [{
            at: now, name: member.name, pin, event: 'IN', minutes: null,
            sessionType: type, eventName: eventName || '', ip: ip || '',
        }]);
        return {
            success: true, event: 'IN', name: member.name, sessionType: type, eventName: eventName || '',
            message: 'Signed in successfully. Have a great ' + type.toLowerCase() + '!',
        };
    });
}

function signOutRow(core: Core, entry: LogEntry, now: Date): LogWrite {
    const since = isNaN(entry.serial) ? now : serialToInstant(entry.serial, core.tz);
    const hours = Math.max(0, (now.getTime() - since.getTime()) / 3_600_000);
    return {
        at: now, name: entry.name, pin: entry.pin, event: 'OUT', minutes: Math.round(hours * 60),
        sessionType: entry.sessionType, eventName: entry.eventName, ip: '',
    };
}

/** Sign out everyone whose last Log row is IN (coach button and the kiosk end-of-night PIN). */
export async function signOutAll(): Promise<{ count: number; names: string[] }> {
    const core = await loadCore();
    const now = new Date();
    const rows: LogWrite[] = [];
    for (const entry of lastEventByPin(core.log).values()) {
        if (entry.event === 'IN') rows.push(signOutRow(core, entry, now));
    }
    await appendLog(core, rows);
    return { count: rows.length, names: rows.map((r) => r.name) };
}

/** Sign out one person by PIN (mentor portal). No-op if they are not signed in. */
export async function signOutOne(pinRaw: string): Promise<{ success: boolean; name?: string; message?: string }> {
    const pin = normPin(pinRaw);
    if (!pin) return { success: false, message: 'PIN not found.' };
    return withPinLock(pin, async () => {
        const core = await loadCore();
        const member = core.roster.find((m) => m.pin === pin);
        if (!member) return { success: false, message: 'PIN not found.' };
        const last = lastEventByPin(core.log).get(pin);
        if (!last || last.event !== 'IN') return { success: false, message: 'Not currently signed in.' };
        await appendLog(core, [{ ...signOutRow(core, last, new Date()), name: member.name }]);
        return { success: true, name: member.name };
    });
}

// #endregion

// #region Who is signed in, roster, report

/** "Firstname L." for the kiosk's signed-in panel. */
function shortName(full: string): string {
    const parts = full.trim().split(/\s+/);
    const first = parts[0] || '';
    return parts.length > 1 ? `${first} ${parts[parts.length - 1].charAt(0).toUpperCase()}.` : first;
}

/** Names currently signed in, split into mentors and students. */
export async function getSignedIn(): Promise<{ mentors: string[]; students: string[] }> {
    const core = await loadCore();
    const roleByPin = new Map(core.roster.filter((m) => m.pin).map((m) => [m.pin, m.role]));
    const mentors: string[] = [];
    const students: string[] = [];
    for (const entry of lastEventByPin(core.log).values()) {
        if (entry.event !== 'IN') continue;
        if (roleByPin.get(entry.pin) === 'Mentor') mentors.push(shortName(entry.name));
        else students.push(shortName(entry.name));
    }
    mentors.sort();
    students.sort();
    return { mentors, students };
}

/** Roster with each member's current IN/OUT status. */
export async function getRoster() {
    const core = await loadCore();
    const last = lastEventByPin(core.log);
    const students = core.roster.map((m) => {
        const entry = m.pin ? last.get(m.pin) : undefined;
        const isIn = entry?.event === 'IN';
        return {
            name: m.name,
            pin: m.pin,
            role: m.role,
            status: isIn ? 'IN' : 'OUT',
            sessionType: isIn && entry ? entry.sessionType : '',
            login: isIn && entry && !isNaN(entry.serial) ? serialToInstant(entry.serial, core.tz).toISOString() : '',
        };
    });
    return { success: true, students };
}

/** Date-range attendance/hours report for the coach dashboard (both dates inclusive). */
export async function getReport(startDate: string, endDate: string) {
    const core = await loadCore();
    const start = ymdToSerial(startDate);
    const end = ymdToSerial(endDate) + 1; // exclusive: the whole end day counts

    type Tally = {
        name: string; pin: string; role: string; totalMins: number; sessions: number;
        days: Set<number>; byType: Record<string, number>;
    };
    const tallies: Tally[] = core.roster.map((m) => ({
        name: m.name, pin: m.pin, role: m.role, totalMins: 0, sessions: 0,
        days: new Set<number>(), byType: { Meeting: 0, Outreach: 0, Competition: 0 },
    }));
    const byPin = new Map(tallies.filter((t) => t.pin).map((t) => [t.pin, t]));

    const meetingDays = new Set<number>();
    for (const e of core.log) {
        if (isNaN(e.serial) || e.serial < start || e.serial >= end) continue;
        if (e.event !== 'OUT' || e.minutes <= 0) continue;
        const day = Math.floor(e.serial);
        meetingDays.add(day);
        const t = byPin.get(e.pin);
        if (!t) continue;
        t.totalMins += e.minutes;
        t.sessions += 1;
        t.days.add(day);
        const type = e.sessionType || 'Meeting';
        if (t.byType[type] !== undefined) t.byType[type] += e.minutes;
    }
    const totalMeetingDays = meetingDays.size;

    const students = tallies.map((t) => ({
        name: t.name,
        pin: t.pin,
        role: t.role,
        totalMins: Math.round(t.totalMins),
        sessions: t.sessions,
        daysPresent: t.days.size,
        attendancePct: totalMeetingDays > 0 ? Math.round((t.days.size / totalMeetingDays) * 100) : 0,
        avgSession: t.sessions > 0 ? Math.round(t.totalMins / t.sessions) : 0,
        meetingMins: Math.round(t.byType.Meeting || 0),
        outreachMins: Math.round(t.byType.Outreach || 0),
        competitionMins: Math.round(t.byType.Competition || 0),
    }));
    students.sort((a, b) => b.totalMins - a.totalMins);

    return { success: true, students, totalMeetingDays, startDate, endDate };
}

// #endregion

// #region Seasons, helpers, checklist

type Season = {
    name: string;
    /** Serial of the first day (midnight). */
    start: number;
    /** Serial of the day after the last day (exclusive bound). */
    endExclusive: number;
    minNew: number; minVeteran: number; minCaptain: number;
};

function parseSeasons(grid: Grid): Season[] {
    const seasons: Season[] = [];
    for (let i = 1; i < grid.length; i++) {
        const name = text(grid[i], 0);
        if (!name) continue;
        seasons.push({
            name,
            start: Math.floor(cellToSerial(at(grid[i], 1))),
            endExclusive: Math.floor(cellToSerial(at(grid[i], 2))) + 1,
            minNew: num(grid[i], 3),
            minVeteran: num(grid[i], 4),
            minCaptain: num(grid[i], 5),
        });
    }
    return seasons;
}

function inSeason(e: LogEntry, s: Season): boolean {
    return !isNaN(e.serial) && e.serial >= s.start && e.serial < s.endExclusive;
}

/** Season name -> total meeting days, from the Helpers tab. */
function parseHelperDays(grid: Grid): Map<string, number> {
    const map = new Map<string, number>();
    for (let i = 1; i < grid.length; i++) {
        const name = text(grid[i], 0);
        if (name) map.set(name, num(grid[i], 1));
    }
    return map;
}

/**
 * Fall Training and Build Season only count Meeting time (a blank session type
 * is a Meeting). Outreach and Competition time does not count toward those two.
 * Every other season counts every session type.
 */
function countsTowardSeason(seasonName: string, sessionType: string): boolean {
    const n = seasonName.toLowerCase();
    if (!n.includes('fall') && !n.includes('build')) return true;
    const t = sessionType.trim().toLowerCase();
    return t === '' || t === 'meeting';
}

/** Threshold for the checklist's numeric requirements, by student type. */
function thresholdForRequirement(reqName: string, studentType: StudentType): number | null {
    const name = reqName.toLowerCase();
    if (name.includes('volunteer') && name.includes('hour')) {
        return studentType === 'New' ? 15 : studentType === 'Captain' ? 40 : 30;
    }
    if (name.includes('fall') && (name.includes('attendance') || name.includes('meeting'))) {
        return studentType === 'Veteran' ? 60 : 80;
    }
    if (name.includes('build') && name.includes('attendance')) {
        return studentType === 'New' ? 50 : studentType === 'Captain' ? 75 : 60;
    }
    if (name.includes('competition') && name.includes('attendance')) return 0;
    return null;
}

/** Map a raw checklist cell to a status and a display value. */
function parseStatusValue(raw: string, threshold: number | null): { status: string; displayValue: string } {
    const lower = raw.toLowerCase().trim();
    if (lower === 'n/a' || lower === 'na') return { status: 'na', displayValue: 'N/A' };
    if (lower === 'yes' || lower === 'y' || lower === 'complete') return { status: 'complete', displayValue: 'Complete' };
    if (lower === 'no' || lower === 'n' || lower === '') return { status: 'incomplete', displayValue: '' };
    if (lower === 'partial' || lower === 'in progress') return { status: 'partial', displayValue: raw };
    const n = parseFloat(raw);
    if (!isNaN(n)) {
        if (threshold !== null && n >= threshold) return { status: 'complete', displayValue: raw };
        return { status: n > 0 ? 'partial' : 'incomplete', displayValue: raw };
    }
    return { status: 'incomplete', displayValue: raw };
}

type ChecklistItem = { col: number; name: string; detail: string; deadline: string; consequence: string };
type Checklist = {
    items: ChecklistItem[];
    deadlines: { name: string; consequence: string; items: ChecklistItem[] }[];
    grid: Grid;
    /** PIN -> index into grid of that student's row. */
    rowByPin: Map<string, number>;
};

function parseChecklist(grid: Grid): Checklist {
    const deadlineRow = grid[0];
    const requirementRow = grid[1] || [];
    const detailRow = grid[2];
    const consequenceRow = grid[3];

    const items: ChecklistItem[] = [];
    const deadlines: Checklist['deadlines'] = [];
    let lastDeadline = '';
    let current: Checklist['deadlines'][number] | null = null;
    for (let c = 1; c < requirementRow.length; c++) {
        const name = text(requirementRow, c);
        if (!name) continue;
        const deadline = text(deadlineRow, c) || lastDeadline;
        if (!current || deadline !== lastDeadline) {
            current = { name: deadline, consequence: text(consequenceRow, c), items: [] };
            deadlines.push(current);
            lastDeadline = deadline;
        }
        const item = { col: c, name, detail: text(detailRow, c), deadline, consequence: current.consequence };
        items.push(item);
        current.items.push(item);
    }

    const rowByPin = new Map<string, number>();
    for (let r = 4; r < grid.length; r++) {
        const pin = normPin(at(grid[r], 0));
        if (pin && !rowByPin.has(pin)) rowByPin.set(pin, r);
    }
    return { items, deadlines, grid, rowByPin };
}

/** Seasons list for the coach dashboard dropdown. */
export async function getSeasons() {
    await connect();
    const [grid] = await readValues([tabRange(TAB.seasons)]);
    const seasons = parseSeasons(grid).map((s) => ({
        name: s.name,
        startDate: isNaN(s.start) ? '' : serialToYmd(s.start),
        endDate: isNaN(s.endExclusive) ? '' : serialToYmd(s.endExclusive - 1),
        minNew: s.minNew,
        minVeteran: s.minVeteran,
        minCaptain: s.minCaptain,
    }));
    return { success: true, seasons };
}

/** Per-student requirement completion grid for the coach dashboard. */
export async function getChecklistOverview() {
    await connect();
    const [rosterGrid, checklistGrid] = await readValues([tabRange(TAB.roster), tabRange(TAB.checklist)]);
    const roster = parseRoster(rosterGrid);
    const byPin = new Map(roster.filter((m) => m.pin).map((m) => [m.pin, m]));
    const cl = parseChecklist(checklistGrid);

    const requirements = cl.items.map((it) => ({ col: it.col, name: it.name, deadline: it.deadline }));
    const students = [];
    for (const [pin, rowIdx] of cl.rowByPin) {
        const member = byPin.get(pin);
        const studentType = member?.studentType ?? 'Veteran';
        let complete = 0, incomplete = 0, partial = 0, na = 0;
        const statuses: string[] = [];
        for (const it of cl.items) {
            const status = parseStatusValue(text(cl.grid[rowIdx], it.col), thresholdForRequirement(it.name, studentType)).status;
            if (status === 'complete') complete++;
            else if (status === 'na') na++;
            else if (status === 'partial') partial++;
            else incomplete++;
            statuses.push(status);
        }
        const total = complete + incomplete + partial; // n/a does not count
        students.push({
            pin,
            name: member?.name ?? (text(cl.grid[rowIdx], 1) || pin),
            role: member?.role ?? 'Student',
            studentType,
            complete, incomplete, partial, na, total,
            pct: total > 0 ? Math.round((complete / total) * 100) : 0,
            statuses,
        });
    }
    students.sort((a, b) => a.pct - b.pct); // least complete first
    return { success: true, students, requirements };
}

// #endregion

// #region Student portal

/** Minimum attendance % for a season, by student type and season name. */
function minAttendancePct(seasonName: string, studentType: StudentType): number {
    const n = seasonName.toLowerCase();
    if (n.includes('fall')) return studentType === 'Veteran' ? 60 : 80;
    if (n.includes('build')) return studentType === 'New' ? 50 : studentType === 'Captain' ? 75 : 60;
    if (n.includes('competition')) return 0;
    if (n.includes('offseason') || n.includes('off season')) return 0;
    return 60;
}

// Captain volunteer hours count Outreach sessions from March 1 2026 to December 31 2026.
const VOLUNTEER_WINDOW = { start: ymdToSerial('2026-03-01'), endExclusive: ymdToSerial('2027-01-01') };

/** Student self-service portal: season hours, attendance, checklists, captain eligibility. */
export async function getStudentPortal(pinRaw: string) {
    const typed = String(pinRaw ?? '').trim();
    if (!/^\d{4}$/.test(typed)) return { success: false, message: 'Please enter your 4-digit PIN.' };
    const pin = normPin(typed);

    await connect();
    const [rosterGrid, logGrid, seasonsGrid, helpersGrid, checklistGrid] = await readValues([
        tabRange(TAB.roster), tabRange(TAB.log, 'A:H'), tabRange(TAB.seasons), tabRange(TAB.helpers), tabRange(TAB.checklist),
    ]);
    const student = parseRoster(rosterGrid).find((m) => m.pin === pin);
    if (!student) return { success: false, message: 'PIN not recognized.' };

    const log = parseLog(logGrid);
    const mine = log.filter((e) => e.pin === pin && e.event === 'OUT');
    const seasons = parseSeasons(seasonsGrid);
    const helperDays = parseHelperDays(helpersGrid);
    const cl = parseChecklist(checklistGrid);
    const isMentor = student.role === 'Mentor';

    // Hours per season against the minimum for this student type.
    const seasonBlocks = seasons.map((s) => {
        const minHours = isMentor ? 0
            : student.studentType === 'New' ? s.minNew
            : student.studentType === 'Captain' ? s.minCaptain
            : s.minVeteran;
        let actualMins = 0;
        for (const e of mine) {
            if (inSeason(e, s) && countsTowardSeason(s.name, e.sessionType)) actualMins += e.minutes;
        }
        return { name: s.name, actualMins: Math.round(actualMins), minMins: Math.round(minHours * 60) };
    });

    // Requirement checklists grouped by deadline.
    const rowIdx = cl.rowByPin.get(pin);
    const myRow: Row | undefined = rowIdx === undefined ? undefined : cl.grid[rowIdx];
    const deadlines = cl.deadlines.map((dl) => ({
        name: dl.name,
        consequence: dl.consequence,
        items: dl.items.map((it) => {
            const threshold = thresholdForRequirement(it.name, student.studentType);
            const parsed = parseStatusValue(text(myRow, it.col), threshold);
            return { name: it.name, detail: it.detail, status: parsed.status, displayValue: parsed.displayValue, threshold };
        }),
    }));

    // Sessions attended per season against the Helpers tab's total meeting days.
    const sessionsIn = (s: Season): number =>
        mine.filter((e) => inSeason(e, s) && countsTowardSeason(s.name, e.sessionType)).length;
    const attendance: { seasonName: string; attendancePct: number; minPct: number; presentDays: number; totalDays: number }[] = [];
    if (!isMentor) {
        for (const s of seasons) {
            let totalDays = helperDays.get(s.name) || 0;
            if (totalDays === 0) {
                const days = new Set<number>();
                for (const e of log) {
                    if (e.event === 'OUT' && inSeason(e, s) && countsTowardSeason(s.name, e.sessionType)) days.add(Math.floor(e.serial));
                }
                totalDays = days.size;
            }
            const present = sessionsIn(s);
            attendance.push({
                seasonName: s.name,
                attendancePct: totalDays > 0 ? Math.round((present / totalDays) * 100) : 0,
                minPct: minAttendancePct(s.name, student.studentType),
                presentDays: present,
                totalDays,
            });
        }
    }

    let captainEligibility = null;
    if (!isMentor) {
        let volHours = 0, fallAttPct = 0, buildAttPct = 0;
        let appStatus = 'incomplete', interviewStatus = 'incomplete';

        // The Checklist tab is the source of truth for the captain criteria.
        if (myRow) {
            for (const it of cl.items) {
                const reqName = it.name.toLowerCase();
                const value = text(myRow, it.col).toLowerCase();
                const isMet = value === 'yes' || value === 'y' || value === 'complete';
                const numeric = parseFloat(value) || 0;
                if (reqName.includes('volunteer') && reqName.includes('hour')) volHours = numeric;
                if (reqName.includes('fall') && (reqName.includes('attendance') || reqName.includes('meeting'))) fallAttPct = numeric;
                if (reqName.includes('build') && reqName.includes('attendance')) buildAttPct = numeric;
                if (reqName.includes('application') && reqName.includes('captain')) appStatus = isMet ? 'complete' : 'incomplete';
                if (reqName.includes('interview')) interviewStatus = isMet ? 'complete' : 'incomplete';
            }
        }

        // Fall back to the Log when the Checklist has no number.
        if (volHours === 0) {
            let volMins = 0;
            for (const e of mine) {
                if (e.sessionType.toLowerCase() !== 'outreach') continue;
                if (isNaN(e.serial) || e.serial < VOLUNTEER_WINDOW.start || e.serial >= VOLUNTEER_WINDOW.endExclusive) continue;
                volMins += e.minutes;
            }
            volHours = Math.round((volMins / 60) * 10) / 10;
        }
        if (fallAttPct === 0 || buildAttPct === 0) {
            for (const s of seasons) {
                const totalDays = helperDays.get(s.name) || 0;
                if (totalDays === 0) continue;
                const pct = Math.round((sessionsIn(s) / totalDays) * 100);
                const n = s.name.toLowerCase();
                if (n.includes('fall') && fallAttPct === 0) fallAttPct = pct;
                if (n.includes('build') && buildAttPct === 0) buildAttPct = pct;
            }
        }

        // "Attendance Progress" shows the same number as the Checklist.
        for (const a of attendance) {
            const n = a.seasonName.toLowerCase();
            if (n.includes('fall') && fallAttPct > 0) a.attendancePct = fallAttPct;
            if (n.includes('build') && buildAttPct > 0) a.attendancePct = buildAttPct;
        }

        const volMet = volHours >= 40;
        const parts = [fallAttPct, buildAttPct].filter((p) => p > 0);
        const avgAtt = parts.length ? Math.round(parts.reduce((a, b) => a + b, 0) / parts.length) : 0;
        const attMet = avgAtt >= 75;

        captainEligibility = {
            deadline: 'May 2027',
            overall: volMet && attMet && appStatus === 'complete' && interviewStatus === 'complete',
            criteria: [
                {
                    name: 'Volunteer Hours',
                    detail: 'Must log 40+ volunteer hours from March 2026 - December 2026',
                    status: volMet ? 'complete' : (volHours > 0 ? 'partial' : 'incomplete'),
                    displayValue: volHours + ' hrs',
                    target: '40 hrs',
                },
                {
                    name: 'Attendance Average',
                    detail: 'Average of Fall Training + Build Season attendance must be 75% or higher (Sept 2026 - Feb 2027)',
                    status: attMet ? 'complete' : (avgAtt > 0 ? 'partial' : 'incomplete'),
                    displayValue: avgAtt + '%',
                    target: '75%',
                },
                {
                    name: 'Captain Application',
                    detail: 'Google Form application must be submitted by the date provided (typically April/May)',
                    status: appStatus,
                    displayValue: appStatus === 'complete' ? 'Submitted' : '',
                    target: 'Submitted',
                },
                {
                    name: 'Coach Interview',
                    detail: 'Sit through an interview with coaches about the position(s) you are applying for (typically April/May)',
                    status: interviewStatus,
                    displayValue: interviewStatus === 'complete' ? 'Complete' : '',
                    target: 'Complete',
                },
            ],
        };
    }

    return {
        success: true,
        name: student.name,
        role: student.role,
        studentType: student.studentType,
        department: student.department,
        driveTeam: student.driveTeam,
        seasons: seasonBlocks,
        attendance,
        deadlines,
        captainEligibility,
    };
}

// #endregion

// #region Year rollover

/** "2026-27": the school year starts in August. */
function schoolYear(tz: string): string {
    const today = serialToYmd(instantToSerial(new Date(), tz));
    const year = Number(today.slice(0, 4));
    const month = Number(today.slice(5, 7));
    return month >= 8 ? `${year}-${String(year + 1).slice(2)}` : `${year - 1}-${String(year).slice(2)}`;
}

export async function getRolloverStatus() {
    invalidateMeta();
    const meta = await connect();
    const year = schoolYear(timeZoneOf(meta));
    return { success: true, schoolYear: year, alreadyDone: !!findTab(meta, `${TAB.checklist} ${year}`) };
}

function errorMessage(e: unknown): string {
    return e instanceof Error ? e.message : String(e);
}

/**
 * End-of-year rollover, as the coach's script does it: copy Checklist, Summary
 * and Log to "<tab> <school year>" archive tabs, clear the Checklist statuses
 * and the Summary rows, and promote New students to Veteran. The Log itself is
 * the permanent record and is never cleared.
 */
export async function rolloverYear() {
    invalidateMeta();
    const meta = await connect();
    const year = schoolYear(timeZoneOf(meta));
    const archived: string[] = [];
    const errors: string[] = [];
    let tabCount = meta.tabs.length;

    /** Copy a tab to the end of the spreadsheet under its archive name, once. */
    const archive = async (tab: TabInfo): Promise<boolean> => {
        const name = `${tab.title} ${year}`;
        if (findTab(meta, name)) return false;
        await batchUpdate([{ duplicateSheet: { sourceSheetId: tab.sheetId, insertSheetIndex: tabCount, newSheetName: name } }]);
        tabCount++;
        return true;
    };

    try {
        const tab = findTab(meta, TAB.checklist);
        if (tab) {
            if (await archive(tab)) archived.push('Checklist');
            // Keep header rows 1-4 and the PIN / Name columns; clear the statuses.
            if (tab.rowCount > 4 && tab.columnCount > 2) {
                await clearValues([tabRange(TAB.checklist, `C5:${columnLetter(tab.columnCount)}${tab.rowCount}`)]);
            }
        }
    } catch (e) { errors.push('Checklist: ' + errorMessage(e)); }

    try {
        const tab = findTab(meta, TAB.summary);
        if (tab) {
            if (await archive(tab)) archived.push('Summary');
            if (tab.rowCount > 1) {
                await clearValues([tabRange(TAB.summary, `A2:${columnLetter(tab.columnCount)}${tab.rowCount}`)]);
            }
        }
    } catch (e) { errors.push('Summary: ' + errorMessage(e)); }

    try {
        const tab = findTab(meta, TAB.log);
        if (tab && await archive(tab)) archived.push('Log (archived copy - original kept)');
    } catch (e) { errors.push('Log: ' + errorMessage(e)); }

    try {
        const [rosterGrid] = await readValues([tabRange(TAB.roster)]);
        const promote = parseRoster(rosterGrid).filter((m) => m.role === 'Student' && m.rawStudentType === 'New');
        await writeValues(promote.map((m) => ({ range: tabRange(TAB.roster, `D${m.sheetRow}`), values: [['Veteran']] })));
        archived.push(`Roster: ${promote.length} New students promoted to Veteran`);
    } catch (e) { errors.push('Roster: ' + errorMessage(e)); }

    invalidateMeta();
    return {
        success: errors.length === 0,
        schoolYear: year,
        archived,
        errors,
        message: errors.length === 0 ? `Rollover complete for ${year}!` : 'Completed with some errors.',
    };
}

// #endregion

// #region Attendance (the tab this app adds) and mentor stats
//
// `Attendance` is a grid: PIN | Name | one column per day that has a sign-out.
// Each cell holds that person's hours for the day, taken from the Log. A mentor
// types "E" in an empty cell to mark an excused absence. The grid is rebuilt
// from Roster + Log every time the mentor stats are opened, so it also picks up
// sessions recorded by the coach's own script. Only "E" marks are kept from the
// tab itself; everything else in it is derived.

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
    months: Record<number, MonthStat>;
    byType: Record<string, number>;
}

function emptyMonth(): MonthStat {
    return { hours: 0, attended: 0, excused: 0, absent: 0, attendance: 0, attendanceWithExcused: 0 };
}
function pct(part: number, whole: number): number {
    return whole > 0 ? (part / whole) * 100 : 0;
}
function finishStat(s: MonthStat): void {
    const days = s.attended + s.excused + s.absent;
    s.attendance = pct(s.attended, days);
    s.attendanceWithExcused = pct(s.attended + s.excused, days);
}

const ATTENDANCE_HEADERS = ['PIN', 'Name'];

async function ensureAttendanceTab(): Promise<void> {
    const meta = await connect();
    if (findTab(meta, TAB.attendance)) return;
    // Added at the end so none of the coach's tabs move.
    await batchUpdate([{ addSheet: { properties: { title: TAB.attendance, index: meta.tabs.length } } }]);
    invalidateMeta();
}

function sameCell(a: Cell, b: Cell): boolean {
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
    return String(a) === String(b);
}

export async function getStats() {
    await ensureAttendanceTab();
    const core = await loadCore();
    const [existing] = await readValues([tabRange(TAB.attendance)]);

    // Excused marks already in the tab, keyed by "pin|yyyy-mm-dd".
    const excused = new Set<string>();
    const existingHeader = existing[0] || [];
    for (let r = 1; r < existing.length; r++) {
        const pin = normPin(at(existing[r], 0));
        if (!pin) continue;
        for (let c = ATTENDANCE_HEADERS.length; c < existingHeader.length; c++) {
            if (text(existing[r], c).toUpperCase() === 'E') excused.add(`${pin}|${text(existingHeader, c)}`);
        }
    }

    // Hours per person per day, and hours per person per session type, from the Log.
    const hoursByPinDay = new Map<string, number>();
    const byTypeByPin = new Map<string, Record<string, number>>();
    const overallByType: Record<string, number> = {};
    const daySet = new Set<string>();
    for (const e of core.log) {
        if (e.event !== 'OUT' || isNaN(e.serial)) continue;
        const day = serialToYmd(e.serial);
        const hours = e.minutes / 60;
        daySet.add(day);
        hoursByPinDay.set(`${e.pin}|${day}`, (hoursByPinDay.get(`${e.pin}|${day}`) || 0) + hours);
        const type = e.sessionType || 'Meeting';
        const mineByType = byTypeByPin.get(e.pin) || {};
        mineByType[type] = (mineByType[type] || 0) + hours;
        byTypeByPin.set(e.pin, mineByType);
        overallByType[type] = (overallByType[type] || 0) + hours;
    }
    const days = Array.from(daySet).sort();

    const stats: UserStats[] = [];
    const overall: MonthStat & { months: Record<number, MonthStat>; byType: Record<string, number> } =
        { ...emptyMonth(), months: {}, byType: overallByType };
    const grid: Grid = [[...ATTENDANCE_HEADERS, ...days]];

    for (const m of core.roster) {
        if (!m.pin) continue;
        const user: UserStats = { pin: m.pin, name: m.name, ...emptyMonth(), months: {}, byType: byTypeByPin.get(m.pin) || {} };
        const row: Row = [m.pin, m.name];
        for (const day of days) {
            const month = Number(day.slice(5, 7));
            const um = (user.months[month] ??= emptyMonth());
            const om = (overall.months[month] ??= emptyMonth());
            const hours = hoursByPinDay.get(`${m.pin}|${day}`) || 0;
            if (hours > 0) {
                for (const s of [user, um, overall, om]) { s.attended++; s.hours += hours; }
                row.push(Math.round(hours * 100) / 100);
            } else if (excused.has(`${m.pin}|${day}`)) {
                for (const s of [user, um, overall, om]) s.excused++;
                row.push('E');
            } else {
                for (const s of [user, um, overall, om]) s.absent++;
                row.push('');
            }
        }
        finishStat(user);
        Object.values(user.months).forEach(finishStat);
        stats.push(user);
        grid.push(row);
    }
    finishStat(overall);
    Object.values(overall.months).forEach(finishStat);

    await writeAttendance(existing, grid);

    return { overall, stats, sessionTypes: Object.keys(overallByType).sort() };
}

/** Rewrite the Attendance tab when its content differs from the derived grid. */
async function writeAttendance(existing: Grid, grid: Grid): Promise<void> {
    // Cover everything that is in the tab now, so one write both updates the
    // grid and blanks any row or column that is no longer part of it.
    const width = Math.max(grid[0].length, ...existing.map((row) => row.length));
    const height = Math.max(grid.length, existing.length);
    const next: Grid = [];
    let changed = false;
    for (let r = 0; r < height; r++) {
        const row: Row = [];
        for (let c = 0; c < width; c++) {
            const value = at(grid[r], c);
            if (!sameCell(at(existing[r], c), value)) changed = true;
            row.push(value);
        }
        next.push(row);
    }
    if (!changed) return;

    const meta = await connect();
    const tab = findTab(meta, TAB.attendance);
    if (!tab) throw new Error('Attendance tab not found');
    if (tab.rowCount < height || tab.columnCount < width) {
        await batchUpdate([{
            updateSheetProperties: {
                properties: {
                    sheetId: tab.sheetId,
                    gridProperties: {
                        rowCount: Math.max(tab.rowCount, height),
                        columnCount: Math.max(tab.columnCount, width),
                    },
                },
                fields: 'gridProperties(rowCount,columnCount)',
            },
        }]);
        invalidateMeta();
    }
    await writeValues([{ range: tabRange(TAB.attendance, 'A1'), values: next }]);
}

// #endregion
