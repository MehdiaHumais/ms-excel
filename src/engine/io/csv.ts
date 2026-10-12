/**
 * CSV import/export.
 *
 * Handles quoted fields, embedded commas/quotes/newlines, optional BOM, and
 * delimiter sniffing (comma, tab, semicolon — Excel regional defaults).
 * Exported values use the *displayed* text (numbers & dates formatted), which
 * is what users expect from a spreadsheet "Save as CSV".
 */

import { WorkbookDoc, Worksheet, CellPrimitive } from '../model/types';
import { getSheet, createWorkbook } from '../model/workbook';
import { cellToA1 } from '../model/address';
import { formatValue } from '../format/numberFormat';

export interface CsvOptions {
  delimiter?: string;
  /** Include the first row as-is; used for sniffing. */
  hasHeader?: boolean;
}

const DEFAULT_DELIMITERS = [',', '\t', ';'];

/** Parse a CSV string greedily, honoring RFC 4180 quoting. */
export function parseCsv(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  // Strip BOM.
  if (text.charCodeAt(0) === 0xfeff) i = 1;
  while (i < text.length) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\n') {
      row.push(field);
      field = '';
      rows.push(row);
      row = [];
      i++;
      continue;
    }
    if (ch === '\r') {
      // Handle CRLF and bare CR line endings.
      if (text[i + 1] === '\n') i++;
      row.push(field);
      field = '';
      rows.push(row);
      row = [];
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** Guess the most likely delimiter by counting occurrences in the first line. */
export function sniffCsvDelimiter(text: string): string {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? '';
  let best = ',';
  let bestCount = -1;
  for (const d of DEFAULT_DELIMITERS) {
    const count = firstLine.split(d).length - 1;
    if (count > bestCount) {
      bestCount = count;
      best = d;
    }
  }
  return best;
}

export function toCsvText(rows: string[][], delimiter = ','): string {
  return rows.map((row) => row.map((cell) => quoteCsv(cell, delimiter)).join(delimiter)).join('\r\n');
}

function quoteCsv(value: string, delimiter: string): string {
  if (/["\r\n]/.test(value) || value.includes(delimiter)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

/** Coerce a raw string into a sensible cell primitive (number first, keep strings). */
export function coerceCellValue(raw: string): CellPrimitive {
  const trimmed = raw.trim();
  if (trimmed === '') return null;
  const negative = /^-/.test(trimmed);
  const numberish = trimmed.replace(/,/g, '').replace(/[()]/g, (m) => (m === '(' ? '-' : ''));
  if (/^[-+]?\d+(\.\d+)?$/.test(numberish)) {
    const n = parseFloat(numberish);
    if (!Number.isNaN(n)) return negative && trimmed.startsWith('(') ? -Math.abs(n) : n;
  }
  if (/^(true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === 'true';
  return trimmed;
}

/** Build a WorkbookDoc from an array of parsed CSV rows. */
export function csvRowsToWorkbook(rows: string[][], options: CsvOptions = {}): WorkbookDoc {
  const wb = createWorkbook({ title: 'Imported CSV' });
  const sheet = wb.sheets[0];
  if (!options.hasHeader) {
    // Skip empty trailing/leading padding? No — write everything.
  }
  const cells: Record<string, { value: CellPrimitive }> = {};
  rows.forEach((row, r) => {
    row.forEach((raw, c) => {
      const v = coerceCellValue(raw);
      if (v === null) return;
      cells[cellToA1(c, r + 1)] = { value: v };
    });
  });
  sheet.cells = cells;
  return wb;
}

/**
 * Export the active sheet to CSV using displayed values.
 * `formattedValues` uses number/date formatting so round-trips look right.
 */
export function workbookToCsvData(wb: WorkbookDoc, sheetId?: string, formattedValues = true): string[][] {
  const sheet = sheetId ? getSheet(wb, sheetId) : wb.sheets[wb.sheets.findIndex((s) => s.id === wb.activeSheetId)];
  if (!sheet) return [];
  return sheetToCsvRows(sheet, formattedValues);
}

export function sheetToCsvRows(sheet: Worksheet, formattedValues: boolean): string[][] {
  const maxRow = findMaxRow(sheet);
  const maxCol = findMaxCol(sheet);
  const rows: string[][] = [];
  for (let r = 1; r <= maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const cell = sheet.cells[cellToA1(c, r)];
      if (!cell) {
        row.push('');
        continue;
      }
      if (formattedValues) {
        row.push(formatCellForCsv(cell));
      } else if (cell.value !== undefined) {
        row.push(typeof cell.value === 'boolean' ? (cell.value ? 'TRUE' : 'FALSE') : String(cell.value));
      } else {
        row.push('');
      }
    }
    rows.push(row);
  }
  return rows;
}

function formatCellForCsv(cell: NonNullable<Worksheet['cells'][string]>): string {
  if (cell.formula !== undefined) return `=${cell.formula}`;
  if (cell.value === null || cell.value === undefined) return '';
  if (typeof cell.value === 'boolean') return cell.value ? 'TRUE' : 'FALSE';
  return formatValue(cell.value, cell.style?.numberFormat);
}

function findMaxRow(sheet: Worksheet): number {
  let max = 0;
  for (const ref of Object.keys(sheet.cells)) {
    const r = parseInt(ref.replace(/[A-Z]/gi, ''), 10);
    if (r > max) max = r;
  }
  return max;
}

function findMaxCol(sheet: Worksheet): number {
  let max = 0;
  for (const ref of Object.keys(sheet.cells)) {
    const c = ref.replace(/[0-9]/g, '');
    const col = colFromLabel(c);
    if (col > max) max = col;
  }
  return max;
}

function colFromLabel(label: string): number {
  let n = 0;
  for (let i = 0; i < label.length; i++) n = n * 26 + (label.charCodeAt(i) - 64);
  return n - 1;
}