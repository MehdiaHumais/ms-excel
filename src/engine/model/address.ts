import { SHEET_SIZE } from './types';

/** Convert a 0-based column index to letters, e.g. 0 -> "A", 27 -> "AB". */
export function colToLabel(col: number): string {
  let c = col;
  let out = '';
  while (c >= 0) {
    out = String.fromCharCode(65 + (c % 26)) + out;
    c = Math.floor(c / 26) - 1;
  }
  return out;
}

/** Convert column letters to a 0-based index, e.g. "A" -> 0, "AB" -> 27. Returns -1 if invalid. */
export function labelToCol(label: string): number {
  let col = -1;
  for (const ch of label.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    col = (col + 1) * 26 + (code - 65);
  }
  return col;
}

export function isValidColLabel(label: string): boolean {
  return label.length > 0 && label.length <= 4 && labelToCol(label) >= 0;
}

export interface A1Ref {
  sheet?: string;
  /** 0-based column. */
  col: number;
  /** 1-based row. */
  row: number;
  /** Whether the column part is absolute ($A). */
  absCol?: boolean;
  /** Whether the row part is absolute ($1). */
  absRow?: boolean;
}

/**
 * Parse an A1 style reference. Accepts optional sheet qualifier:
 *   A1, $A$1, Sheet1!A1, 'My Sheet'!A1
 * Does not accept ranges.
 */
export function parseA1(ref: string): A1Ref | null {
  if (!ref) return null;
  let rest = ref.trim();
  let sheet: string | undefined;

  const bang = lastIndexOfBang(rest);
  if (bang >= 1) {
    sheet = rest.slice(0, bang);
    rest = rest.slice(bang + 1);
    if (sheet.startsWith("'") && sheet.endsWith("'")) {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
  }

  const m = /^(\$?)([A-Za-z]{1,4})(\$?)([0-9]{1,7})$/.exec(rest);
  if (!m) return null;
  const col = labelToCol(m[2]);
  const row = parseInt(m[4], 10);
  if (col < 0 || row < 1 || row > SHEET_SIZE.rows) return null;
  return { sheet, col, row, absCol: m[1] === '$', absRow: m[3] === '$' };
}

/** Find the last '!' that is not inside quotes. */
function lastIndexOfBang(s: string): number {
  let inQuote = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'") inQuote = !inQuote;
    else if (ch === '!' && !inQuote) return i;
  }
  return -1;
}

export function rowToRowLabel(row: number): A1Ref {
  return { col: -1, row };
}

/** Format an A1Ref back to a string. */
export function a1ToString(ref: A1Ref): string {
  const sheet = ref.sheet != null ? quoteSheetName(ref.sheet) + '!' : '';
  const colPart = ref.col >= 0 ? (ref.absCol ? '$' : '') + colToLabel(ref.col) : '';
  const rowPart = (ref.absRow ? '$' : '') + ref.row;
  return sheet + colPart + rowPart;
}

/** Format a cell coordinate (excluding sheet) e.g. (0,1) -> "A1". */
export function cellToA1(col: number, row: number, absCol?: boolean, absRow?: boolean): string {
  return (absCol ? '$' : '') + colToLabel(col) + (absRow ? '$' : '') + row;
}

/** Format a range e.g. "A1:B2". If single cell, returns "A1". */
export function rangeToA1(sCol: number, sRow: number, eCol: number, eRow: number): string {
  const hasSheet = false;
  void hasSheet;
  if (sCol === eCol && sRow === eRow) return cellToA1(sCol, sRow);
  return cellToA1(sCol, sRow) + ':' + cellToA1(eCol, eRow);
}

/** Quote a sheet name for use inside formulas. */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";
}

export function unquoteSheetName(name: string): string {
  if (name.startsWith("'") && name.endsWith("'")) return name.slice(1, -1).replace(/''/g, "'");
  return name;
}

/** Walk every cell coordinate in an inclusive range [sCol..eCol] x [sRow..eRow]. */
export function* iterateRange(
  sRow: number,
  sCol: number,
  eRow: number,
  eCol: number
): Generator<{ row: number; col: number }> {
  for (let r = sRow; r <= eRow; r++) {
    for (let c = sCol; c <= eCol; c++) {
      yield { row: r, col: c };
    }
  }
}

export function rangeDimensions(sRow: number, sCol: number, eRow: number, eCol: number): { rows: number; cols: number } {
  return { rows: Math.abs(eRow - sRow) + 1, cols: Math.abs(eCol - sCol) + 1 };
}

export interface ParsedRange {
  sheet?: string;
  sRow: number;
  sCol: number;
  eRow: number;
  eCol: number;
}

/**
 * Parse a range specification. Sheet qualifier allowed on the start only:
 *   A1, A1:B2, Sheet1!A1:B2, $A$1:$B$2
 * Single cell ranges expand to a single coordinate.
 */
export function parseRange(ref: string): ParsedRange | null {
  if (!ref) return null;
  let sheet: string | undefined;
  const bang = lastIndexOfBang(ref);
  if (bang >= 1) {
    sheet = ref.slice(0, bang).trim();
    ref = ref.slice(bang + 1);
    if (sheet.startsWith("'") && sheet.endsWith("'")) {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
  }
  const parts = ref.split(':');
  if (parts.length > 2 || parts.length === 0) return null;
  const a = parseA1(sheet ? sheet + '!' + parts[0] : parts[0]);
  if (!a) return null;
  if (parts.length === 1) {
    return { sheet: a.sheet, sRow: a.row, sCol: a.col, eRow: a.row, eCol: a.col };
  }
  const b = parseA1(sheet ? sheet + '!' + parts[1] : parts[1]);
  if (!b) return null;
  const sRow = Math.min(a.row, b.row);
  const eRow = Math.max(a.row, b.row);
  const sCol = Math.min(a.col, b.col);
  const eCol = Math.max(a.col, b.col);
  return { sheet: a.sheet ?? b.sheet, sRow, sCol, eRow, eCol };
}

export function parsedRangeToA1(r: ParsedRange, includeSheet = false): string {
  const prefix = includeSheet && r.sheet != null ? quoteSheetName(r.sheet) + '!' : '';
  if (r.sCol === r.eCol && r.sRow === r.eRow) return prefix + cellToA1(r.sCol, r.sRow);
  return prefix + cellToA1(r.sCol, r.sRow) + ':' + cellToA1(r.eCol, r.eRow);
}

/** True if coord (col,row) is inside range. */
export function isInRange(row: number, col: number, r: ParsedRange): boolean {
  return row >= r.sRow && row <= r.eRow && col >= r.sCol && col <= r.eCol;
}

/** True if a parsed range represents an entire row or column. */
export function isFullRow(r: ParsedRange): boolean {
  return r.sCol === 0 && r.eCol === SHEET_SIZE.cols - 1;
}
export function isFullCol(r: ParsedRange): boolean {
  return r.sRow === 1 && r.eRow === SHEET_SIZE.rows;
}

/** Normalize a range string so the top-left is always the start. */
export function normalizeRange(range: string): string {
  const p = parseRange(range);
  if (!p) return range;
  return parsedRangeToA1(p);
}