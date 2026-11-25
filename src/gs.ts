// Thin Google Sheets REST client (service-account auth).
//
// The coach's spreadsheet is the only data store and its layout belongs to her.
// Everything here addresses tabs by NAME and cells by POSITION, exactly as her
// Apps Script does, and nothing here renames, reorders or reformats a tab.

import 'dotenv/config';
import { JWT } from 'google-auth-library';

export type Cell = string | number | boolean;
export type Row = Cell[];
export type Grid = Row[];

const auth = new JWT({
    email: process.env.CLIENT_EMAIL,
    // Coolify/env may store the PEM with any depth of escaped newlines (\n, \\n,
    // even \\\\n after a copy through the API); collapse any run of backslashes
    // before an n into a real newline.
    key: (process.env.API_KEY || '').replace(/\\+n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const SPREADSHEET_ID = process.env.SPREADSHEET_ID || '';
const BASE = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}`;

async function call<T>(url: string, method: 'GET' | 'POST' | 'PUT' = 'GET', data?: unknown): Promise<T> {
    const res = await auth.request<T>({ url, method, data });
    return res.data;
}

/** GETs are safe to repeat, so retry once on a transient failure. Writes never retry. */
async function get<T>(url: string): Promise<T> {
    try {
        return await call<T>(url);
    } catch (e) {
        console.error('[sheets] read failed, retrying once:', e instanceof Error ? e.message : e);
        return call<T>(url);
    }
}

// #region Metadata

export type TabInfo = { title: string; sheetId: number; rowCount: number; columnCount: number };
export type SheetMeta = { timeZone: string; tabs: TabInfo[] };

type MetaResponse = {
    properties?: { timeZone?: string };
    sheets?: { properties?: { title?: string; sheetId?: number; gridProperties?: { rowCount?: number; columnCount?: number } } }[];
};

export async function loadMeta(): Promise<SheetMeta> {
    const data = await get<MetaResponse>(
        `${BASE}?fields=properties(timeZone),sheets(properties(title,sheetId,gridProperties(rowCount,columnCount)))`
    );
    return {
        timeZone: data.properties?.timeZone || '',
        tabs: (data.sheets || []).map((s) => ({
            title: s.properties?.title || '',
            sheetId: s.properties?.sheetId ?? 0,
            rowCount: s.properties?.gridProperties?.rowCount ?? 0,
            columnCount: s.properties?.gridProperties?.columnCount ?? 0,
        })),
    };
}

// #endregion

// #region Reads

type BatchGetResponse = { valueRanges?: { values?: Grid }[] };

/** A1 range for a tab name (quoted so names with spaces work). */
export function tabRange(tab: string, a1?: string): string {
    const quoted = `'${tab.replace(/'/g, "''")}'`;
    return a1 ? `${quoted}!${a1}` : quoted;
}

async function batchGet(ranges: string[], render: 'UNFORMATTED_VALUE' | 'FORMATTED_VALUE'): Promise<Grid[]> {
    const qs = ranges.map((r) => 'ranges=' + encodeURIComponent(r)).join('&');
    const data = await get<BatchGetResponse>(
        `${BASE}/values:batchGet?${qs}&valueRenderOption=${render}&dateTimeRenderOption=SERIAL_NUMBER`
    );
    return ranges.map((_, i) => data.valueRanges?.[i]?.values || []);
}

/**
 * Computed values, like Apps Script getValues(): numbers stay numbers and dates
 * come back as serial numbers. Trailing empty rows and cells are dropped.
 */
export function readValues(ranges: string[]): Promise<Grid[]> {
    return batchGet(ranges, 'UNFORMATTED_VALUE');
}

/** What a person sees in each cell, like Apps Script getDisplayValues(). */
export function readDisplay(ranges: string[]): Promise<Grid[]> {
    return batchGet(ranges, 'FORMATTED_VALUE');
}

/** Cell at a position, '' when the row or cell is past the end of the data. */
export function at(row: Row | undefined, index: number): Cell {
    const v = row?.[index];
    return v === undefined || v === null ? '' : v;
}

/** Trimmed text of a cell. */
export function text(row: Row | undefined, index: number): string {
    return String(at(row, index)).trim();
}

/** Numeric value of a cell, 0 when it is blank or not a number. */
export function num(row: Row | undefined, index: number): number {
    const v = at(row, index);
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v));
    return isNaN(n) ? 0 : n;
}

// #endregion

// #region Writes

export type CellWrite =
    | { kind: 'blank' }
    | { kind: 'text'; value: string }
    | { kind: 'number'; value: number }
    | { kind: 'datetime'; serial: number };

export const blank: CellWrite = { kind: 'blank' };
export const textCell = (value: string): CellWrite => (value === '' ? blank : { kind: 'text', value });
export const numberCell = (value: number): CellWrite => ({ kind: 'number', value });
export const dateTimeCell = (serial: number): CellWrite => ({ kind: 'datetime', serial });

function toCellData(c: CellWrite): object {
    switch (c.kind) {
        case 'blank': return {};
        // stringValue is stored as literal text: a value that starts with "=" is
        // never evaluated as a formula.
        case 'text': return { userEnteredValue: { stringValue: c.value } };
        case 'number': return { userEnteredValue: { numberValue: c.value } };
        case 'datetime': return {
            userEnteredValue: { numberValue: c.serial },
            userEnteredFormat: { numberFormat: { type: 'DATE_TIME', pattern: 'M/d/yyyy H:mm:ss' } },
        };
    }
}

export async function batchUpdate(requests: object[]): Promise<void> {
    await call<unknown>(`${BASE}:batchUpdate`, 'POST', { requests });
}

/**
 * Add rows after the last row that has data in the tab. One atomic server-side
 * operation with the same semantics as Apps Script appendRow(), so two writers
 * (this app and the coach's script) can never land on the same row.
 */
export async function appendRows(sheetId: number, rows: CellWrite[][]): Promise<void> {
    if (!rows.length) return;
    await batchUpdate([{
        appendCells: {
            sheetId,
            rows: rows.map((r) => ({ values: r.map(toCellData) })),
            fields: 'userEnteredValue,userEnteredFormat.numberFormat',
        },
    }]);
}

/** Overwrite cells with literal values (no formula parsing). */
export async function writeValues(updates: { range: string; values: Grid }[]): Promise<void> {
    if (!updates.length) return;
    await call<unknown>(`${BASE}/values:batchUpdate`, 'POST', {
        valueInputOption: 'RAW',
        data: updates.map((u) => ({ range: u.range, values: u.values })),
    });
}

/** Clear values (formats stay), like Apps Script clearContent(). */
export async function clearValues(ranges: string[]): Promise<void> {
    if (!ranges.length) return;
    await call<unknown>(`${BASE}/values:batchClear`, 'POST', { ranges });
}

/** Spreadsheet column letters for a 1-based column number (1 -> A, 27 -> AA). */
export function columnLetter(n: number): string {
    let s = '';
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

// #endregion
