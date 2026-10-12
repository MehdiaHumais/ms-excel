import {
  WorkbookDoc,
  Worksheet,
  Cell,
  CellPrimitive,
  CellStyle,
  RowSettings,
  ColumnSettings,
  MergeDef,
  ChartDef,
  TableDef,
  ConditionalFormatRule,
  DataValidationRule,
  NoteDef,
  NamedRange,
  PageSetup,
  ErrorValue,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT
} from './types';
import {
  cellToA1,
  parseRange,
  parsedRangeToA1,
  ParsedRange,
  rangeToA1,
  normalizeRange
} from './address';

let uidCounter = 0;
export function uid(prefix: string): string {
  uidCounter += 1;
  return `${prefix}_${Date.now().toString(36)}_${uidCounter}_${Math.random().toString(36).slice(2, 8)}`;
}

export function createSheet(name: string, opts: { id?: string } = {}): Worksheet {
  return {
    id: opts.id ?? uid('sheet'),
    name,
    state: 'visible',
    cells: {},
    rows: {},
    columns: {},
    merges: [],
    charts: [],
    tables: [],
    conditionalFormats: [],
    dataValidations: {},
    notes: [],
    view: { zoom: 100 },
    gridlines: true,
    showRowColHeaders: true
  };
}

export function createWorkbook(options: {
  title?: string;
  author?: string;
  sheets?: Array<{ name: string; id?: string }>;
  activeSheetId?: string;
} = {}): WorkbookDoc {
  const time = Date.now();
  const sheets = (options.sheets && options.sheets.length
    ? options.sheets
    : [{ name: 'Sheet1' }]
  ).map((s) => createSheet(s.name, s));
  return {
    version: 1,
    metadata: {
      title: options.title ?? 'Untitled',
      author: options.author ?? '',
      created: time,
      modified: time,
      app: 'BritSheet',
      fileFormatVersion: 1
    },
    sheets,
    activeSheetId: options.activeSheetId ?? sheets[0].id,
    namedRanges: {},
    settings: { calculationMode: 'automatic' }
  };
}

export function getSheet(wb: WorkbookDoc, sheetId: string): Worksheet | undefined {
  return wb.sheets.find((s) => s.id === sheetId);
}

export function getActiveSheet(wb: WorkbookDoc): Worksheet {
  return getSheet(wb, wb.activeSheetId) ?? wb.sheets[0];
}

export function getSheetByName(wb: WorkbookDoc, name: string): Worksheet | undefined {
  return wb.sheets.find((s) => s.name === name);
}

export function getCell(sheet: Worksheet, col: number, row: number): Cell | undefined {
  return sheet.cells[cellToA1(col, row)];
}

export function setCell(sheet: Worksheet, col: number, row: number, cell: Cell | undefined): void {
  const ref = cellToA1(col, row);
  if (!cell || (cell.value === undefined && cell.formula === undefined && cell.style === undefined && cell.note === undefined && cell.hyperlink === undefined)) {
    delete sheet.cells[ref];
  } else {
    sheet.cells[ref] = cell;
  }
}

/** Normalized ref of the top-left of a merge covering (col,row), if any. */
export function mergeAt(sheet: Worksheet, col: number, row: number): MergeDef | null {
  for (const m of sheet.merges) {
    const p = parseRange(m);
    if (!p) continue;
    if (row >= p.sRow && row <= p.eRow && col >= p.sCol && col <= p.eCol) {
      return { startRow: p.sRow, startCol: p.sCol, endRow: p.eRow, endCol: p.eCol };
    }
  }
  return null;
}

export function addMerge(sheet: Worksheet, range: string): void {
  const norm = normalizeRange(range);
  const p = parseRange(norm);
  if (!p || (p.sRow === p.eRow && p.sCol === p.eCol)) return;
  if (!sheet.merges.includes(norm)) sheet.merges.push(norm);
}

export function removeMerge(sheet: Worksheet, range: string): void {
  const norm = normalizeRange(range);
  const idx = sheet.merges.indexOf(norm);
  if (idx >= 0) sheet.merges.splice(idx, 1);
}

/** Row numbers used by any cell in this sheet (for quick iteration). */
export function usedRows(sheet: Worksheet): number[] {
  const set = new Set<number>();
  for (const ref of Object.keys(sheet.cells)) {
    const p = parseRange(ref);
    if (p) set.add(p.sRow);
  }
  for (const r of Object.keys(sheet.rows)) set.add(parseInt(r, 10));
  return Array.from(set).sort((a, b) => a - b);
}

/**
 * The contiguous block of cells around (col,row), matching Excel's "current
 * region" (Ctrl+Shift+Arrow). Returns a single cell when the start cell is
 * empty. Used so inserting a chart from one selected cell picks up the
 * surrounding data instead of rendering an empty "No data in range" frame.
 */
export function currentRegion(sheet: Worksheet, col: number, row: number): { sCol: number; sRow: number; eCol: number; eRow: number } {
  const filled = (c: number, r: number) => {
    const cell = getCell(sheet, c, r);
    if (!cell) return false;
    return cell.value !== undefined && cell.value !== null && cell.value !== '';
  };
  if (!filled(col, row)) return { sCol: col, sRow: row, eCol: col, eRow: row };

  let sCol = col;
  let eCol = col;
  let sRow = row;
  let eRow = row;
  while (sCol > 0 && filled(sCol - 1, row)) sCol--;
  while (eCol < 16383 && filled(eCol + 1, row)) eCol++;
  // The first row of the region defines how far the block extends vertically.
  while (sRow > 1 && filled(col, sRow - 1)) sRow--;
  while (eRow < 1048576 && filled(col, eRow + 1)) eRow++;
  return { sCol, sRow, eCol, eRow };
}

export function usedCols(sheet: Worksheet): number[] {  const set = new Set<number>();
  for (const ref of Object.keys(sheet.cells)) {
    const p = parseRange(ref);
    if (p) set.add(p.sCol);
  }
  for (const c of Object.keys(sheet.columns)) set.add(parseInt(c, 10));
  return Array.from(set).sort((a, b) => a - b);
}

/* ------------------------------------------------------------------ */
/* Reference shifting when rows/columns are inserted or deleted        */
/* ------------------------------------------------------------------ */

export type ShiftOp = {
  axis: 'row' | 'col';
  index: number; // 0-based row index or 0-based col index where insertion begins
  count: number; // positive for insert, negative for delete (magnitude)
};

/** Shift a parsed range by an operation (deterministic, returns null if fully deleted). */
export function shiftParsedRange(r: ParsedRange, op: ShiftOp): ParsedRange | null {
  const start = op.axis === 'row' ? r.sRow - 1 : r.sCol;
  const end = op.axis === 'row' ? r.eRow - 1 : r.eCol;
  const { index, count } = op;
  let s = start;
  let e = end;
  if (count < 0) {
    // Deleting `count` cells (-count on screen) starting at `index`: coords in
    // [remStart, remEnd] disappear, and everything after them slides up by the
    // number of removed cells (previously this slid the wrong way and did not
    // move the start coordinate, so deleted rows/columns corrupted positions).
    const removed = -count;
    const remStart = index;
    const remEnd = index + removed - 1;
    const startDeleted = s >= remStart && s <= remEnd;
    const endDeleted = e >= remStart && e <= remEnd;
    if (startDeleted && endDeleted) return null; // entire range deleted
    const shiftUp = (coord: number) => (coord > remEnd ? coord - removed : coord);
    s = startDeleted ? remStart : shiftUp(s);
    e = endDeleted ? remStart - 1 : shiftUp(e);
    if (e < s) e = s;
  } else {
    if (s >= index) s += count;
    if (e >= index) e += count;
    if (e < s) e = s;
  }
  const out: ParsedRange = { ...r, sCol: r.sCol, eCol: r.eCol, sRow: r.sRow, eRow: r.eRow };
  if (op.axis === 'row') {
    out.sRow = s + 1;
    out.eRow = e + 1;
  } else {
    out.sCol = s;
    out.eCol = e;
  }
  return out;
}

/**
 * Shift an A1 range STRING by an op. Returns null when the range vanishes.
 * Keeps sheet qualifier and $ signs as-is.
 */
export function shiftRangeString(range: string, op: ShiftOp): string | null {
  const p = parseRange(range);
  if (!p) return range;
  const shifted = shiftParsedRange(p, op);
  if (!shifted) return null;
  const prefix = p.sheet != null ? quoteSheetNameSafe(p.sheet) + '!' : '';
  const cell = (col: number, row: number) => colToLabelSafe(col) + row;
  if (shifted.sRow === shifted.eRow && shifted.sCol === shifted.eCol) {
    return prefix + cell(shifted.sCol, shifted.sRow);
  }
  return prefix + cell(shifted.sCol, shifted.sRow) + ':' + cell(shifted.eCol, shifted.eRow);
}

function quoteSheetNameSafe(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";
}
function colToLabelSafe(col: number): string {
  let c = col;
  let out = '';
  while (c >= 0) {
    out = String.fromCharCode(65 + (c % 26)) + out;
    c = Math.floor(c / 26) - 1;
  }
  return out;
}

/** Shift a row range string like "1:5" or "A:C". */
export function shiftHeaderRange(range: string, op: ShiftOp): string | null {
  const m = /^(\$?)([0-9]+)(?::(\$?)([0-9]+))?$/.exec(range.trim());
  if (m) {
    const a = parseInt(m[2], 10);
    const b = m[4] ? parseInt(m[4], 10) : a;
    const row = op.axis === 'row';
    let s = (row ? a - 1 : 0);
    let e = (row ? b - 1 : 0);
    const index = op.index;
    const count = op.count;
    if (count < 0) {
      const rs = index;
      const re = index - count - 1;
      if (s >= rs && s <= re && e <= re) return null;
      if (s >= rs && s <= re) s = rs;
      if (e >= re) e = e - count;
      if (e < s) e = s;
    } else {
      if (s >= index) s += count;
      if (e >= index) e += count;
      if (e < s) e = s;
    }
    return `${s + 1}:${e + 1}`;
  }
  const cm = /^(\$?)([A-Za-z]{1,4})(?::(\$?)([A-Za-z]{1,4}))?$/.exec(range.trim());
  if (cm && op.axis === 'col') {
    const a = labelToColLocal(cm[2]);
    const b = cm[4] ? labelToColLocal(cm[4]) : a;
    let s = a;
    let e = b;
    const index = op.index;
    const count = op.count;
    if (count < 0) {
      const rs = index;
      const re = index - count - 1;
      if (s >= rs && s <= re && e <= re) return null;
      if (s >= rs && s <= re) s = rs;
      if (e >= re) e = e - count;
      if (e < s) e = s;
    } else {
      if (s >= index) s += count;
      if (e >= index) e += count;
      if (e < s) e = s;
    }
    return `${colToLabelLocal(s)}:${colToLabelLocal(e)}`;
  }
  return range;
}

function labelToColLocal(label: string): number {
  let col = -1;
  for (const ch of label.toUpperCase()) col = (col + 1) * 26 + (ch.charCodeAt(0) - 65);
  return col;
}
function colToLabelLocal(col: number): string {
  let c = col;
  let out = '';
  while (c >= 0) {
    out = String.fromCharCode(65 + (c % 26)) + out;
    c = Math.floor(c / 26) - 1;
  }
  return out;
}

/** Shift cell data map keys by an op (cells that move due to insert/delete). */
export function shiftCellMap(
  cells: Record<string, Cell>,
  op: ShiftOp
): Record<string, Cell> {
  const out: Record<string, Cell> = {};
  const keep: Record<string, Cell> = {};
  for (const [ref, cell] of Object.entries(cells)) {
    const p = parseRange(ref);
    if (!p) {
      keep[ref] = cell;
      continue;
    }
    const shifted = shiftParsedRange(p, op);
    if (!shifted) {
      // Cell deleted entirely by a delete op.
      continue;
    }
    const newRef = rangeToA1(shifted.sCol, shifted.sRow, shifted.eCol, shifted.eRow);
    out[newRef] = cell;
  }
  return { ...out, ...keep };
}

/** Shift every formula in a sheet's cells by an op (adjusts row/col references). */
export function shiftFormulasInSheet(sheet: Worksheet, op: ShiftOp): void {
  for (const ref of Object.keys(sheet.cells)) {
    const cell = sheet.cells[ref];
    if (cell?.formula) {
      const shifted = shiftFormula(cell.formula, op);
      if (shifted === null) {
        delete cell.formula;
        cell.value = '#REF!' as ErrorValue;
      } else {
        cell.formula = shifted;
      }
    }
  }
}

/**
 * Shift A1-style references inside a formula string. Ranges and single refs
 * (with sheet qualifiers, $ signs) are shifted; function names are untouched.
 * Returns null when a reference disappears (deletion covering it).
 */
export function shiftFormula(formula: string, op: ShiftOp): string | null {
  let anyDeleted = false;
  const out = formula.replace(
    /(^|[^\w$])((\$?[A-Za-z]{1,4}\$?[0-9]{1,7})(?::(\$?[A-Za-z]{1,4}\$?[0-9]{1,7}))?)/g,
    (_m, pre: string, body: string) => {
      // Sheet-qualified references are handled below via separate pass.
      const colonIdx = body.indexOf(':');
      if (colonIdx < 0) {
        const shifted = shiftSingleRef(body, op);
        if (shifted === null) {
          anyDeleted = true;
          return pre + '#REF!';
        }
        return pre + shifted;
      } else {
        const a = body.slice(0, colonIdx);
        const b = body.slice(colonIdx + 1);
        const sa = shiftSingleRef(a, op);
        const sb = shiftSingleRef(b, op);
        if (sa === null || sb === null) {
          anyDeleted = true;
          return pre + '#REF!';
        }
        return pre + sa + ':' + sb;
      }
    }
  );
  return anyDeleted ? out : out;
}

function shiftSingleRef(ref: string, op: ShiftOp): string | null {
  const m = /^(\$?)([A-Za-z]{1,4})(\$?)([0-9]{1,7})$/.exec(ref);
  if (!m) return ref;
  const col = labelToColLocal(m[2]);
  const row = parseInt(m[4], 10);
  const isRow = op.axis === 'row';
  const coord = isRow ? row - 1 : col;
  const start = op.index;
  const count = op.count;
  let newCoord = coord;
  if (count < 0) {
    const remStart = start;
    const remEnd = start - count - 1;
    if (coord >= remStart && coord <= remEnd) return null;
    if (coord > remEnd) newCoord = coord - Math.abs(count);
  } else {
    if (coord >= start) newCoord = coord + count;
  }
  if (isRow) return (m[1] ? '$' : '') + m[2] + (m[3] ? '$' : '') + (newCoord + 1);
  return (m[1] ? '$' : '') + colToLabelLocal(newCoord) + (m[3] ? '$' : '') + m[4];
}

/** Shift merges, tables, charts, conditional formats, validations, notes, autofilter, freeze. */
export function shiftSheetStructure(sheet: Worksheet, op: ShiftOp, deletion: 'extend' | 'keep' = 'extend'): void {
  const removedMerges: string[] = [];
  sheet.merges = sheet.merges
    .map((m) => shiftRangeString(m, op))
    .filter((m): m is string => {
      if (m === null) {
        removedMerges.push('x');
        return false;
      }
      return true;
    });
  void removedMerges;

  const shiftStr = (r: string | undefined): string | undefined => {
    if (!r) return undefined;
    const s = shiftRangeString(r, op);
    return s ?? undefined;
  };

  sheet.tables = sheet.tables
    .map((t) => {
      const r = shiftRangeString(t.range, op);
      return r ? { ...t, range: r } : null;
    })
    .filter((t): t is TableDef => t !== null);

  sheet.autofilter = shiftStr(sheet.autofilter);

  sheet.conditionalFormats = sheet.conditionalFormats
    .map((cf) => {
      const r = shiftStr(cf.range);
      return r ? { ...cf, range: r } : null;
    })
    .filter((cf): cf is ConditionalFormatRule => cf !== null);

  const newValidations: Record<string, DataValidationRule> = {};
  for (const [ref, rule] of Object.entries(sheet.dataValidations)) {
    const s = shiftRangeString(ref, op);
    if (s) newValidations[s] = rule;
  }
  sheet.dataValidations = newValidations;

  sheet.notes = sheet.notes
    .map((n) => {
      const s = shiftRangeString(n.ref, op);
      return s ? { ...n, ref: s } : null;
    })
    .filter((n): n is NoteDef => n !== null);

  const freezeRow = sheet.freeze?.row;
  const freezeCol = sheet.freeze?.col;
  if (freezeRow && op.axis === 'row') {
    const rowShifted = shiftNumericIndex(op.index, freezeRow - 1, op.count);
    sheet.freeze = {
      ...(sheet.freeze ?? {}),
      row: rowShifted === null ? undefined : rowShifted + 1
    };
  }
  if (freezeCol && op.axis === 'col') {
    const colShifted = shiftNumericIndex(op.index, freezeCol, op.count);
    sheet.freeze = {
      ...(sheet.freeze ?? {}),
      col: colShifted === null ? undefined : colShifted
    };
  }
}

/** Shift a numeric 0-based index: deletion may nullify it. */
function shiftNumericIndex(insertIndex: number, coord: number, count: number): number | null {
  if (count < 0) {
    const remStart = insertIndex;
    const remEnd = insertIndex - count - 1;
    if (coord >= remStart && coord <= remEnd) return null;
    if (coord > remEnd) return coord - Math.abs(count);
    return coord;
  }
  if (coord >= insertIndex) return coord + count;
  return coord;
}

export function clampRangeToSheet(r: ParsedRange): ParsedRange {
  return r;
}

/** Deep clone via JSON (workbook model is JSON-serializable). */
export function cloneWorkbook<T extends WorkbookDoc>(wb: T): T {
  return JSON.parse(JSON.stringify(wb)) as T;
}

export { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT };
export type { PageSetup, Cell, CellPrimitive, CellStyle, RowSettings, ColumnSettings, MergeDef, ChartDef, NamedRange };