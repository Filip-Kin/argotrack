// Volunteer sign-ups, on the coach's three tabs:
//   VolEvents   Event ID | Event Name | Date | Who It Is For | Location | Notes for Volunteers |
//               Contact | Max Slots Per Person | Show Names | Sign-up Deadline | Status
//   VolSlots    Slot ID | Event ID | Slot Name | Start Time | End Time | Spots
//   VolSignups  Timestamp | Signup ID | Event ID | Slot ID | Person Name | Type |
//               Student Name | Student PIN | Email | Note | Status
// The coach edits events and slots in the sheet. This app only reads those two
// tabs, appends rows to VolSignups, and sets a sign-up's Status to "Cancelled".

import { randomUUID } from 'crypto';
import {
    Grid, Row, SheetMeta,
    appendRows, at, dateTimeCell, readDisplay, readValues, tabRange, text, textCell, writeValues,
} from './gs';
import { Member, connect, findTab, normPin, parseRoster, timeZoneOf } from './sheet';
import { cellToSerial, instantToSerial, serialToWallDate } from './util';

const VOL = { events: 'VolEvents', slots: 'VolSlots', signups: 'VolSignups', roster: 'Roster' } as const;
const SIGNUP_STATUS_COLUMN = 'K';
const MAX_ACTIVE_PER_STUDENT = 30;

// #region Helpers

/** Trim, drop control characters, cap the length. Values are stored as literal text. */
function clean(value: unknown, max: number): string {
    return String(value === undefined || value === null ? '' : value)
        .replace(/[\u0000-\u001f\u007f]/g, ' ')
        .trim()
        .slice(0, max);
}

function validEmail(s: string): boolean {
    return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s) && s.length <= 120;
}

/** "Firstname L." */
function shortName(full: string): string {
    const parts = full.trim().split(/\s+/);
    if (parts.length < 2) return parts[0] || '';
    return parts[0] + ' ' + parts[parts.length - 1].charAt(0).toUpperCase() + '.';
}

/** yyyyMMdd for a date cell, '' when the cell is blank or not a date. */
function dateKey(value: unknown): string {
    const serial = cellToSerial(value);
    if (isNaN(serial)) return '';
    return serialToWallDate(Math.floor(serial)).toISOString().slice(0, 10).replace(/-/g, '');
}

function todayKey(tz: string): string {
    return dateKey(instantToSerial(new Date(), tz));
}

function formatWall(value: unknown, options: Intl.DateTimeFormatOptions): string {
    const serial = cellToSerial(value);
    if (isNaN(serial)) return '';
    return new Intl.DateTimeFormat('en-US', { ...options, timeZone: 'UTC' }).format(serialToWallDate(serial));
}

/** "Sat, Jan 9, 2027" */
const longDate = (value: unknown): string =>
    formatWall(value, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
/** "Jan 9, 2027" */
const shortDate = (value: unknown): string => formatWall(value, { month: 'short', day: 'numeric', year: 'numeric' });
/** "11/11 9:22 AM" */
function stamp(value: unknown): string {
    const serial = cellToSerial(value);
    if (isNaN(serial)) return '';
    const d = serialToWallDate(serial);
    const h = d.getUTCHours();
    const minutes = String(d.getUTCMinutes()).padStart(2, '0');
    return `${d.getUTCMonth() + 1}/${d.getUTCDate()} ${h % 12 === 0 ? 12 : h % 12}:${minutes} ${h < 12 ? 'AM' : 'PM'}`;
}

/** "3:00 PM" or "15:00" -> minutes after midnight, or -1. */
function parseTime(value: string): number {
    const m = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?$/.exec(value.trim());
    if (!m) return -1;
    let h = parseInt(m[1], 10);
    const min = parseInt(m[2], 10);
    if (m[3]) {
        if (h === 12) h = 0;
        if (m[3].toLowerCase() === 'pm') h += 12;
    }
    if (h > 23 || min > 59) return -1;
    return h * 60 + min;
}

function calendarStamp(day: string, minutes: number): string {
    const h = String(Math.floor(minutes / 60)).padStart(2, '0');
    const m = String(minutes % 60).padStart(2, '0');
    return `${day}T${h}${m}00`;
}

/** 0 means "no limit" for both Spots and Max Slots Per Person. */
function limit(value: unknown): number {
    const n = parseInt(String(value), 10);
    return isNaN(n) || n < 0 ? 0 : n;
}

/** The event is not open for anyone to see or join. */
function isHiddenStatus(status: string): boolean {
    const s = status.toLowerCase();
    return s === 'closed' || s === 'deleted' || s === 'hidden';
}

type Signup = {
    /** 1-based row in the VolSignups tab. */
    sheetRow: number;
    timestamp: unknown;
    signupId: string; eventId: string; slotId: string;
    personName: string; type: string; studentName: string; pin: string; email: string; note: string;
    active: boolean; cancelled: boolean;
};

function parseSignups(grid: Grid): Signup[] {
    // The tab normally has a header row, but the coach's copy has data in row 1.
    // A header's first cell is text ("Timestamp"); a sign-up's is a date.
    const first = grid[0];
    const hasHeader = first !== undefined && isNaN(cellToSerial(at(first, 0)));
    const signups: Signup[] = [];
    for (let i = hasHeader ? 1 : 0; i < grid.length; i++) {
        const row: Row = grid[i];
        const status = text(row, 10).toLowerCase();
        signups.push({
            sheetRow: i + 1,
            timestamp: at(row, 0),
            signupId: text(row, 1), eventId: text(row, 2), slotId: text(row, 3),
            personName: text(row, 4), type: text(row, 5), studentName: text(row, 6),
            pin: normPin(at(row, 7)), email: text(row, 8), note: text(row, 9),
            active: status === 'active', cancelled: status === 'cancelled',
        });
    }
    return signups;
}

type VolData = {
    meta: SheetMeta; tz: string; roster: Member[];
    events: Grid; eventsText: Grid; slots: Grid; slotsText: Grid; signups: Signup[];
};

/** Everything the volunteer screens need. Null when the volunteer tabs do not exist. */
async function loadVolunteerData(): Promise<VolData | null> {
    const meta = await connect();
    if (![VOL.events, VOL.slots, VOL.signups].every((t) => findTab(meta, t))) return null;
    const ranges = [VOL.events, VOL.slots].map((t) => tabRange(t));
    const [[roster, events, slots, signups], [eventsText, slotsText]] = await Promise.all([
        readValues([tabRange(VOL.roster), ...ranges, tabRange(VOL.signups)]),
        readDisplay(ranges),
    ]);
    return {
        meta, tz: timeZoneOf(meta), roster: parseRoster(roster),
        events, eventsText, slots, slotsText, signups: parseSignups(signups),
    };
}

function findStudent(data: VolData, pinRaw: unknown): Member | undefined {
    const typed = String(pinRaw === undefined || pinRaw === null ? '' : pinRaw).trim();
    if (!/^\d{4}$/.test(typed)) return undefined;
    return data.roster.find((m) => m.pin === normPin(typed));
}

// One sign-up or cancellation at a time, so two people cannot take the last spot.
let volunteerQueue: Promise<unknown> = Promise.resolve();
function withVolunteerLock<T>(fn: () => Promise<T>): Promise<T> {
    const run = volunteerQueue.then(fn, fn);
    volunteerQueue = run.catch(() => undefined);
    return run;
}

const NOT_SET_UP = 'Volunteer sign-ups are not set up.';

// #endregion

// #region Student portal: list, sign up, cancel

/** Upcoming events with their slots, and this family's own sign-ups. */
export async function getVolunteerEvents(pinRaw: unknown) {
    const data = await loadVolunteerData();
    if (!data) return { success: false, message: NOT_SET_UP };
    const student = findStudent(data, pinRaw);
    if (!student) return { success: false, message: 'PIN not recognized.' };

    const today = todayKey(data.tz);
    const activeBySlot = new Map<string, Signup[]>();
    for (const s of data.signups) {
        if (!s.active) continue;
        const list = activeBySlot.get(s.slotId) || [];
        list.push(s);
        activeBySlot.set(s.slotId, list);
    }

    const events = [];
    for (let i = 1; i < data.events.length; i++) {
        const row = data.events[i];
        const shown = data.eventsText[i];
        const eventId = text(row, 0);
        if (!eventId) continue;
        if (isHiddenStatus(text(row, 10))) continue;
        const day = dateKey(at(row, 2));
        if (day && day < today) continue; // past events are hidden

        const deadline = dateKey(at(row, 9));
        const showNames = text(shown, 8).toLowerCase() !== 'no';

        let eventStart = -1, eventEnd = -1;
        const slots = [];
        for (let j = 1; j < data.slots.length; j++) {
            if (text(data.slots[j], 1) !== eventId) continue;
            const slotId = text(data.slots[j], 0);
            if (!slotId) continue;
            const start = text(data.slotsText[j], 3);
            const end = text(data.slotsText[j], 4);
            const taken = activeBySlot.get(slotId) || [];
            const startMin = parseTime(start), endMin = parseTime(end);
            if (startMin >= 0 && (eventStart < 0 || startMin < eventStart)) eventStart = startMin;
            if (endMin >= 0 && endMin > eventEnd) eventEnd = endMin;
            slots.push({
                slotId,
                name: text(data.slotsText[j], 2),
                start, end, startMin, endMin,
                spots: limit(at(data.slots[j], 5)),
                taken: taken.length,
                names: showNames
                    ? taken.map((s) => shortName(s.personName) + (s.type === 'Parent' ? ' (parent)' : ''))
                    : [],
                mine: taken
                    .filter((s) => s.pin === student.pin)
                    .map((s) => ({ signupId: s.signupId, personName: s.personName, type: s.type })),
            });
        }

        let calendar = null;
        if (day) {
            if (eventStart >= 0) {
                const end = eventEnd > eventStart ? eventEnd : Math.min(eventStart + 60, 1439);
                calendar = { start: calendarStamp(day, eventStart), end: calendarStamp(day, end), tz: data.tz };
            } else {
                calendar = { allDay: true, start: day, tz: data.tz };
            }
        }

        events.push({
            eventId,
            name: text(shown, 1),
            date: day ? longDate(at(row, 2)) : '',
            dateKey: day,
            audience: text(shown, 3),
            location: text(shown, 4),
            notes: text(shown, 5),
            contact: text(shown, 6),
            maxPerPerson: limit(at(row, 7)),
            deadline: deadline ? shortDate(at(row, 9)) : '',
            signupsOpen: !(deadline && today > deadline),
            calendar,
            slots,
        });
    }
    events.sort((a, b) => (a.dateKey || '99999999') < (b.dateKey || '99999999') ? -1 : 1);

    const mySignups = events.flatMap((ev) => ev.slots.flatMap((slot) => slot.mine.map((mine) => ({
        signupId: mine.signupId,
        personName: mine.personName,
        type: mine.type,
        eventName: ev.name,
        date: ev.date,
        slotName: slot.name,
        start: slot.start,
        end: slot.end,
    }))));

    return { success: true, studentName: student.name, events, mySignups };
}

export type SignUpInput = {
    pin: unknown; slotId: unknown; signStudent: unknown;
    parentName: unknown; parentEmail: unknown; studentEmail: unknown; note: unknown;
};

/** Sign up the student, a parent, or both for one slot. */
export function volunteerSignUp(input: SignUpInput): Promise<{ success: boolean; message: string }> {
    const slotId = clean(input.slotId, 60);
    const parentName = clean(input.parentName, 80);
    const parentEmail = clean(input.parentEmail, 120).toLowerCase();
    const studentEmail = clean(input.studentEmail, 120).toLowerCase();
    const note = clean(input.note, 200);
    const wantStudent = String(input.signStudent).toLowerCase() === 'yes';
    const wantParent = parentName !== '';

    return withVolunteerLock(async () => {
        const data = await loadVolunteerData();
        if (!data) return { success: false, message: NOT_SET_UP };
        const student = findStudent(data, input.pin);
        if (!student) return { success: false, message: 'PIN not recognized.' };

        if (!wantStudent && !wantParent) return { success: false, message: 'Choose to sign up yourself, a parent, or both.' };
        if (wantParent && !validEmail(parentEmail)) return { success: false, message: 'Please enter a valid parent email address.' };
        if (studentEmail && !validEmail(studentEmail)) return { success: false, message: 'Please enter a valid student email address or leave it blank.' };

        const slot = data.slots.slice(1).find((row) => text(row, 0) === slotId);
        if (!slot) return { success: false, message: 'That slot no longer exists.' };
        const eventId = text(slot, 1);
        const spots = limit(at(slot, 5));

        const event = data.events.slice(1).find((row) => text(row, 0) === eventId);
        if (!event) return { success: false, message: 'That event no longer exists.' };
        if (isHiddenStatus(text(event, 10))) return { success: false, message: 'Sign-ups for this event are closed.' };
        const today = todayKey(data.tz);
        const day = dateKey(at(event, 2));
        if (day && day < today) return { success: false, message: 'This event has already happened.' };
        const deadline = dateKey(at(event, 9));
        if (deadline && today > deadline) return { success: false, message: 'The sign-up deadline has passed.' };
        const maxPer = limit(at(event, 7));

        let takenInSlot = 0, studentInEvent = 0, parentInEvent = 0, activeForStudent = 0;
        let dupStudentSlot = false, dupParentSlot = false;
        for (const s of data.signups) {
            if (!s.active) continue;
            if (s.slotId === slotId) takenInSlot++;
            if (s.pin === student.pin) activeForStudent++;
            if (s.eventId !== eventId) continue;
            if (s.type === 'Student' && s.pin === student.pin) {
                studentInEvent++;
                if (s.slotId === slotId) dupStudentSlot = true;
            }
            if (s.type === 'Parent' && wantParent && s.email.toLowerCase() === parentEmail) {
                parentInEvent++;
                if (s.slotId === slotId) dupParentSlot = true;
            }
        }

        const need = (wantStudent ? 1 : 0) + (wantParent ? 1 : 0);
        const plural = (n: number): string => (n === 1 ? '' : 's');
        if (wantStudent && dupStudentSlot) return { success: false, message: 'You are already signed up for this slot.' };
        if (wantParent && dupParentSlot) return { success: false, message: 'That parent is already signed up for this slot.' };
        if (spots > 0 && takenInSlot + need > spots) {
            const left = Math.max(0, spots - takenInSlot);
            return {
                success: false,
                message: left === 0 ? 'Sorry, that slot just filled up.' : `Only ${left} spot${plural(left)} left in that slot.`,
            };
        }
        if (maxPer > 0 && wantStudent && studentInEvent >= maxPer) {
            return { success: false, message: `You can sign up for ${maxPer} slot${plural(maxPer)} per event. Cancel one first to switch.` };
        }
        if (maxPer > 0 && wantParent && parentInEvent >= maxPer) {
            return { success: false, message: `That parent can sign up for ${maxPer} slot${plural(maxPer)} per event.` };
        }
        if (activeForStudent + need > MAX_ACTIVE_PER_STUDENT) {
            return { success: false, message: 'You have reached the limit of active sign-ups. Cancel one first.' };
        }

        const tab = findTab(data.meta, VOL.signups);
        if (!tab) return { success: false, message: NOT_SET_UP };
        const now = dateTimeCell(instantToSerial(new Date(), data.tz));
        const rowFor = (personName: string, type: 'Student' | 'Parent', email: string) => [
            now, textCell(randomUUID().slice(0, 8)), textCell(eventId), textCell(slotId),
            textCell(personName), textCell(type), textCell(student.name), textCell(student.pin),
            textCell(email), textCell(note), textCell('Active'),
        ];
        const rows = [];
        if (wantStudent) rows.push(rowFor(student.name, 'Student', studentEmail));
        if (wantParent) rows.push(rowFor(parentName, 'Parent', parentEmail));
        await appendRows(tab.sheetId, rows);

        return { success: true, message: 'You are signed up! Thank you for volunteering.' };
    });
}

/** Cancel one of this student's own sign-ups. */
export function volunteerCancel(pinRaw: unknown, signupIdRaw: unknown): Promise<{ success: boolean; message: string }> {
    const signupId = clean(signupIdRaw, 40);
    return withVolunteerLock(async () => {
        const data = await loadVolunteerData();
        if (!data) return { success: false, message: NOT_SET_UP };
        const student = findStudent(data, pinRaw);
        if (!student) return { success: false, message: 'PIN not recognized.' };
        if (!signupId) return { success: false, message: 'Missing sign-up.' };

        const signup = data.signups.find((s) => s.signupId === signupId);
        if (!signup) return { success: false, message: 'Sign-up not found.' };
        if (signup.pin !== student.pin) return { success: false, message: 'You can only cancel your own sign-ups.' };
        if (!signup.active) return { success: false, message: 'That sign-up is already cancelled.' };

        await writeValues([{
            range: tabRange(VOL.signups, `${SIGNUP_STATUS_COLUMN}${signup.sheetRow}`),
            values: [['Cancelled']],
        }]);
        return { success: true, message: 'Your spot was cancelled.' };
    });
}

// #endregion

// #region Coach: who signed up

/** Every event with its sign-up count, plus the full roster of one event. */
export async function getVolunteerRoster(eventIdRaw: unknown) {
    const data = await loadVolunteerData();
    if (!data) return { success: false, message: NOT_SET_UP };
    const today = todayKey(data.tz);

    const activeByEvent = new Map<string, number>();
    for (const s of data.signups) {
        if (s.active) activeByEvent.set(s.eventId, (activeByEvent.get(s.eventId) || 0) + 1);
    }

    const events = [];
    for (let i = 1; i < data.events.length; i++) {
        const row = data.events[i];
        const eventId = text(row, 0);
        if (!eventId) continue;
        const day = dateKey(at(row, 2));
        events.push({
            eventId,
            name: text(data.eventsText[i], 1),
            date: day ? longDate(at(row, 2)) : '',
            dateKey: day,
            status: text(data.eventsText[i], 10) || 'Open',
            past: !!(day && day < today),
            count: activeByEvent.get(eventId) || 0,
        });
    }
    // Upcoming first (soonest to latest), then past events (most recent first).
    events.sort((a, b) => {
        if (a.past !== b.past) return a.past ? 1 : -1;
        const ak = a.dateKey || '99999999', bk = b.dateKey || '99999999';
        if (ak === bk) return 0;
        return (ak < bk) !== a.past ? -1 : 1;
    });

    const eventId = String(eventIdRaw ?? '').trim();
    if (!eventId) return { success: true, events, roster: null };
    const event = events.find((e) => e.eventId === eventId);
    if (!event) return { success: false, message: 'Event not found.' };

    type Person = { name: string; type: string; studentName: string; email: string; note: string; signedUp: string };
    const slots: { slotId: string; name: string; time: string; spots: number; people: Person[] }[] = [];
    const slotIndex = new Map<string, number>();
    for (let j = 1; j < data.slots.length; j++) {
        if (text(data.slots[j], 1) !== eventId) continue;
        const slotId = text(data.slots[j], 0);
        if (!slotId) continue;
        const start = text(data.slotsText[j], 3), end = text(data.slotsText[j], 4);
        slotIndex.set(slotId, slots.length);
        slots.push({
            slotId,
            name: text(data.slotsText[j], 2),
            time: start && end ? `${start} - ${end}` : (start ? `Starts ${start}` : ''),
            spots: limit(at(data.slots[j], 5)),
            people: [],
        });
    }

    let students = 0, parents = 0, cancelled = 0;
    for (const s of data.signups) {
        if (s.eventId !== eventId) continue;
        if (!s.active) {
            if (s.cancelled) cancelled++;
            continue;
        }
        const idx = slotIndex.get(s.slotId);
        if (idx === undefined) continue;
        if (s.type === 'Parent') parents++; else students++;
        slots[idx].people.push({
            name: s.personName, type: s.type, studentName: s.studentName,
            email: s.email, note: s.note, signedUp: stamp(s.timestamp),
        });
    }

    return {
        success: true,
        events,
        roster: { event, slots, totals: { students, parents, total: students + parents, cancelled } },
    };
}

// #endregion
