import {
  WorkbookDoc,
  Worksheet,
  Cell,
  CellPrimitive,
  CellStyle,
  RowSettings,
  ColumnSettings,
  ChartDef,
  TableDef,
  ConditionalFormatRule,
  DataValidationRule,
  NoteDef,
  NamedRange,
  PageSetup,
  ErrorValue,
  ShapeDef
} from '../model/types';
import {
  getSheet,
  createSheet,
  uid,
  shiftCellMap,
  shiftFormulasInSheet,
  shiftSheetStructure,
  shiftRangeString,
  cloneWorkbook,
  addMerge,
  removeMerge
} from '../model/workbook';
import { parseRange, cellToA1, normalizeRange } from '../model/address';
import { getTheme, ThemeDef } from '../model/theme';
import { Command, CompositeCommand, NoopCommand } from './history';

const MAXCOL = 16383;
const MAXROW = 1048576;

/* ----------------------------- sheet helpers ----------------------------- */

function withSheetTransform(wb: WorkbookDoc, sheetId: string, fn: (s: Worksheet) => Worksheet): WorkbookDoc {
  const idx = wb.sheets.findIndex((s) => s.id === sheetId);
  if (idx < 0) return wb;
  const next = wb.sheets.slice();
  next[idx] = fn(JSON.parse(JSON.stringify(wb.sheets[idx])));
  return { ...wb, sheets: next, metadata: { ...wb.metadata, modified: Date.now() } };
}

function withSheetCells(
  wb: WorkbookDoc,
  sheetId: string,
  cells: Record<string, Cell | undefined>
): WorkbookDoc {
  return withSheetTransform(wb, sheetId, (s) => {
    const nextCells = { ...s.cells };
    for (const [ref, cell] of Object.entries(cells)) {
      if (!cell || (cell.value === undefined && cell.formula === undefined && cell.style === undefined && cell.note === undefined && cell.hyperlink === undefined)) {
        delete nextCells[ref];
      } else {
        const copy: Cell = { ...cell };
        delete (copy as Partial<Cell>).computed;
        nextCells[ref] = copy;
      }
    }
    s.cells = nextCells;
    return s;
  });
}

function captureCells(sheet: Worksheet, refs: string[]): Record<string, Cell | undefined> {
  const out: Record<string, Cell | undefined> = {};
  for (const ref of refs) {
    if (sheet.cells[ref]) out[ref] = JSON.parse(JSON.stringify(sheet.cells[ref]));
  }
  return out;
}

export function rectRefs(sheet: Worksheet, sRow: number, sCol: number, eRow: number, eCol: number): string[] {
  const refs: string[] = [];
  for (let r = sRow; r <= eRow; r++) {
    for (let c = sCol; c <= eCol; c++) {
      refs.push(cellToA1(c, r));
    }
  }
  return refs;
}

/** Shift a numeric index; returns null if deleted, otherwise the new index. */
function shiftIndex(idx: number, insertIndex: number, count: number): number | null {
  if (count < 0) {
    const remStart = insertIndex;
    const remEnd = insertIndex - count - 1;
    if (idx >= remStart && idx <= remEnd) return null;
    if (idx > remEnd) return idx - Math.abs(count);
    return idx;
  }
  if (idx >= insertIndex) return idx + count;
  return idx;
}

/* ---------------------------------- Commands ---------------------------------- */

/** Set cell contents (values/formulas) for explicit refs. */
export class EditCellsCommand implements Command {
  private captured = false;
  constructor(
    public readonly label: string,
    private sheetId: string,
    private edits: Record<string, Cell | undefined>,
    private before: Record<string, Cell | undefined> = {}
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    if (!this.captured) {
      const sheet = getSheet(wb, this.sheetId);
      if (sheet) {
        const snapshot: Record<string, Cell | undefined> = {};
        for (const ref of Object.keys(this.edits)) {
          const cur = sheet.cells[ref];
          snapshot[ref] = cur ? JSON.parse(JSON.stringify(cur)) : undefined;
        }
        if (Object.keys(this.before).length === 0) this.before = snapshot;
        this.captured = true;
      }
    }
    return withSheetCells(wb, this.sheetId, this.edits);
  }

  inverse(): Command {
    return new EditCellsCommand(this.label, this.sheetId, this.before, this.edits);
  }
}

/** Write raw values (and optionally formulas/styles) into a rectangular range. */
export class SetRangeValuesCommand implements Command {
  private edits: Record<string, Cell | undefined>;
  private before: Record<string, Cell | undefined> = {};

  constructor(
    public readonly label: string,
    private sheetId: string,
    private sRow: number,
    private sCol: number,
    private eRow: number,
    private eCol: number,
    private values: CellPrimitive[][],
    private options?: { formulas?: string[][]; styles?: CellStyle[][] }
  ) {
    this.edits = {};
    for (let r = 0; r + sRow <= eRow; r++) {
      for (let c = 0; c + sCol <= eCol; c++) {
        const ref = cellToA1(sCol + c, sRow + r);
        const v = this.values[r]?.[c];
        const f = this.options?.formulas?.[r]?.[c];
        const st = this.options?.styles?.[r]?.[c];
        const cell: Cell = {};
        if (f !== undefined && f !== null && f !== '') cell.formula = f;
        else if (v !== undefined && v !== null) cell.value = v;
        if (st) cell.style = st;
        if (cell.formula === undefined && cell.value === undefined && !st) continue;
        this.edits[ref] = cell;
      }
    }
  }

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    this.before = captureCells(sheet, Object.keys(this.edits));
    return withSheetCells(wb, this.sheetId, this.edits);
  }

  inverse(): Command {
    return new EditCellsCommand(this.label, this.sheetId, this.before, this.edits);
  }
}

/** Apply a style patch to every cell in a range; undo restores prior styles. */
export class FormatCellsCommand implements Command {
  private before: Record<string, Cell | undefined> = {};
  constructor(
    public readonly label: string,
    private sheetId: string,
    private sRow: number,
    private sCol: number,
    private eRow: number,
    private eCol: number,
    private patch: Partial<CellStyle>
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    const refs = rectRefs(sheet, this.sRow, this.sCol, this.eRow, this.eCol);
    this.before = captureCells(sheet, refs);
    const edits: Record<string, Cell | undefined> = {};
    for (const ref of refs) {
      const existing = sheet.cells[ref];
      const cell: Cell = existing ? JSON.parse(JSON.stringify(existing)) : {};
      cell.style = {
        ...(cell.style ?? {}),
        ...this.patch
      };
      for (const k of Object.keys(this.patch)) {
        const key = k as keyof CellStyle;
        if (this.patch[key] === undefined) delete (cell.style as Record<string, unknown>)[key];
      }
      edits[ref] = cell;
    }
    return withSheetCells(wb, this.sheetId, edits);
  }

  inverse(): Command {
    return new EditCellsCommand(this.label, this.sheetId, this.before);
  }
}

/** Clear contents (optionally styles too) of a range. */
export class ClearRangeCommand implements Command {
  private before: Record<string, Cell | undefined> = {};
  constructor(
    public readonly label: string,
    private sheetId: string,
    private sRow: number,
    private sCol: number,
    private eRow: number,
    private eCol: number,
    private clearStyle = false
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    const refs = rectRefs(sheet, this.sRow, this.sCol, this.eRow, this.eCol);
    this.before = captureCells(sheet, refs);
    const edits: Record<string, Cell | undefined> = {};
    for (const ref of refs) {
      const existing = sheet.cells[ref];
      if (!existing) continue;
      if (this.clearStyle) {
        edits[ref] = undefined;
      } else {
        const keep: Cell = {};
        if (existing.style) keep.style = existing.style;
        if (existing.note) keep.note = existing.note;
        if (existing.hyperlink) keep.hyperlink = existing.hyperlink;
        edits[ref] = keep;
      }
    }
    return withSheetCells(wb, this.sheetId, edits);
  }

  inverse(): Command {
    return new EditCellsCommand(this.label, this.sheetId, this.before);
  }
}

/* ------------------------------ Row / column ops ------------------------------ */

class RestoreRowsCommand implements Command {
  label = 'Restore rows';
  constructor(
    private sheetId: string,
    private rows: Record<number, RowSettings | undefined>
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const [key, val] of Object.entries(this.rows)) {
        const r = parseInt(key, 10);
        if (val && Object.keys(val).length) s.rows[r] = val as RowSettings;
        else delete s.rows[r];
      }
      return s;
    });
  }
  inverse(): Command {
    return new NoopCommand('No changes');
  }
}

class RestoreColumnsCommand implements Command {
  label = 'Restore columns';
  constructor(
    private sheetId: string,
    private cols: Record<number, ColumnSettings | undefined>
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const [key, val] of Object.entries(this.cols)) {
        const c = parseInt(key, 10);
        if (val && Object.keys(val).length) s.columns[c] = val as ColumnSettings;
        else delete s.columns[c];
      }
      return s;
    });
  }
  inverse(): Command {
    return new NoopCommand('No changes');
  }
}

class RestoreMergesCommand implements Command {
  label = 'Restore merges';
  constructor(
    private sheetId: string,
    private merges: string[]
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const m of this.merges) addMerge(s, m);
      return s;
    });
  }
  inverse(): Command {
    return new NoopCommand('No changes');
  }
}

export class ResizeRowsCommand implements Command {
  private before: Record<number, RowSettings | undefined> = {};
  constructor(
    public readonly label: string,
    private sheetId: string,
    private rows: number[],
    private height: number
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    for (const r of this.rows) this.before[r] = sheet.rows[r] ? JSON.parse(JSON.stringify(sheet.rows[r])) : undefined;
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const r of this.rows) s.rows[r] = { ...(s.rows[r] ?? {}), height: this.height, hidden: false };
      return s;
    });
  }
  inverse(): Command {
    return new RestoreRowsCommand(this.sheetId, this.before);
  }
}

export class ResizeColumnsCommand implements Command {
  private before: Record<number, ColumnSettings | undefined> = {};
  constructor(
    public readonly label: string,
    private sheetId: string,
    private cols: number[],
    private width: number
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    for (const c of this.cols) this.before[c] = sheet.columns[c] ? JSON.parse(JSON.stringify(sheet.columns[c])) : undefined;
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const c of this.cols) s.columns[c] = { ...(s.columns[c] ?? {}), width: this.width, hidden: false };
      return s;
    });
  }
  inverse(): Command {
    return new RestoreColumnsCommand(this.sheetId, this.before);
  }
}

export class ToggleRowsHiddenCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private rows: number[],
    private hidden: boolean
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const r of this.rows) s.rows[r] = { ...(s.rows[r] ?? {}), hidden: this.hidden };
      return s;
    });
  }
  inverse(): Command {
    return new ToggleRowsHiddenCommand(this.label, this.sheetId, this.rows, !this.hidden);
  }
}

export class ToggleColumnsHiddenCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private cols: number[],
    private hidden: boolean
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (const c of this.cols) s.columns[c] = { ...(s.columns[c] ?? {}), hidden: this.hidden };
      return s;
    });
  }
  inverse(): Command {
    return new ToggleColumnsHiddenCommand(this.label, this.sheetId, this.cols, !this.hidden);
  }
}

/** Insert `count` rows before 0-based `index`. */
export class InsertRowsCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private index: number,
    private count: number
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return wbWithShiftOp(wb, this.sheetId, { axis: 'row', index: this.index, count: this.count });
  }
  inverse(): Command {
    return new DeleteRowsCommand(this.label, this.sheetId, this.index, this.count);
  }
}

export class InsertColumnsCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private index: number,
    private count: number
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return wbWithShiftOp(wb, this.sheetId, { axis: 'col', index: this.index, count: this.count });
  }
  inverse(): Command {
    return new DeleteColumnsCommand(this.label, this.sheetId, this.index, this.count);
  }
}

/** Delete rows and remember everything needed to restore them. */
export class DeleteRowsCommand implements Command {
  private before = {
    cells: {} as Record<string, Cell | undefined>,
    rows: {} as Record<number, RowSettings | undefined>,
    merges: [] as string[]
  };
  constructor(
    public readonly label: string,
    private sheetId: string,
    private index: number,
    private count: number
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    const deletedRowNos = Array.from({ length: this.count }, (_, i) => this.index + i + 1);
    for (let c = 0; c <= MAXCOL; c++) {
      for (const r of deletedRowNos) {
        const ref = cellToA1(c, r);
        if (sheet.cells[ref]) this.before.cells[ref] = JSON.parse(JSON.stringify(sheet.cells[ref]));
      }
    }
    for (const r of deletedRowNos) {
      if (sheet.rows[r]) this.before.rows[r] = JSON.parse(JSON.stringify(sheet.rows[r]));
      else this.before.rows[r] = undefined;
    }
    this.before.merges = sheet.merges.filter((m) => {
      const p = parseRange(m);
      return p && p.sRow <= this.index + this.count && p.eRow >= this.index + 1;
    });
    return wbWithShiftOp(wb, this.sheetId, { axis: 'row', index: this.index, count: -this.count });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const index = this.index;
    const count = this.count;
    const before = this.before;
    return new CompositeCommand('Restore rows', [
      new InsertRowsCommand('Restore rows', sheetId, index, count),
      new EditCellsCommand('Restore rows', sheetId, before.cells),
      new RestoreRowsCommand(sheetId, before.rows),
      new RestoreMergesCommand(sheetId, before.merges)
    ]);
  }
}

export class DeleteColumnsCommand implements Command {
  private before = {
    cells: {} as Record<string, Cell | undefined>,
    cols: {} as Record<number, ColumnSettings | undefined>,
    merges: [] as string[]
  };
  constructor(
    public readonly label: string,
    private sheetId: string,
    private index: number,
    private count: number
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    const deletedColNos = Array.from({ length: this.count }, (_, i) => this.index + i);
    for (const c of deletedColNos) {
      for (let r = 1; r <= MAXROW; r++) {
        const ref = cellToA1(c, r);
        if (sheet.cells[ref]) this.before.cells[ref] = JSON.parse(JSON.stringify(sheet.cells[ref]));
      }
    }
    for (const c of deletedColNos) {
      if (sheet.columns[c]) this.before.cols[c] = JSON.parse(JSON.stringify(sheet.columns[c]));
      else this.before.cols[c] = undefined;
    }
    this.before.merges = sheet.merges.filter((m) => {
      const p = parseRange(m);
      return p && p.sCol <= this.index + this.count - 1 && p.eCol >= this.index;
    });
    return wbWithShiftOp(wb, this.sheetId, { axis: 'col', index: this.index, count: -this.count });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const index = this.index;
    const count = this.count;
    const before = this.before;
    return new CompositeCommand('Restore columns', [
      new InsertColumnsCommand('Restore columns', sheetId, index, count),
      new EditCellsCommand('Restore columns', sheetId, before.cells),
      new RestoreColumnsCommand(sheetId, before.cols),
      new RestoreMergesCommand(sheetId, before.merges)
    ]);
  }
}

/**
 * Apply a row/column structural shift to the whole workbook:
 * all sheets' formulas shift, target sheet's cells / settings / structures
 * shift, and named ranges at workbook level are re-anchored.
 */
function wbWithShiftOp(
  wb: WorkbookDoc,
  targetSheetId: string,
  op: { axis: 'row' | 'col'; index: number; count: number }
): WorkbookDoc {
  const next = cloneWorkbook(wb);
  const target = getSheet(next, targetSheetId);
  if (!target) return wb;

  for (const sheet of next.sheets) {
    shiftFormulasInSheet(sheet, op);
  }

  const sheetClone = JSON.parse(JSON.stringify(target)) as Worksheet;
  sheetClone.cells = shiftCellMap(sheetClone.cells, op);
  shiftSheetStructure(sheetClone, op);

  if (op.axis === 'row') {
    const rows: Record<number, RowSettings> = {};
    for (const [k, v] of Object.entries(sheetClone.rows)) {
      const r = parseInt(k, 10);
      const shifted = shiftIndex(r - 1, op.index, op.count);
      if (shifted != null) rows[shifted + 1] = v;
    }
    sheetClone.rows = rows;
  } else {
    const cols: Record<number, ColumnSettings> = {};
    for (const [k, v] of Object.entries(sheetClone.columns)) {
      const c = parseInt(k, 10);
      const shifted = shiftIndex(c, op.index, op.count);
      if (shifted != null) cols[shifted] = v;
    }
    sheetClone.columns = cols;
  }

  target.cells = sheetClone.cells;
  target.rows = sheetClone.rows;
  target.columns = sheetClone.columns;
  target.merges = sheetClone.merges;
  target.tables = sheetClone.tables;
  target.conditionalFormats = sheetClone.conditionalFormats;
  target.dataValidations = sheetClone.dataValidations;
  target.notes = sheetClone.notes;
  target.autofilter = sheetClone.autofilter;
  target.freeze = sheetClone.freeze;
  target.charts = sheetClone.charts;

  for (const [name, nr] of Object.entries(next.namedRanges)) {
    const s = shiftRangeString(nr.range, op);
    if (s) next.namedRanges[name] = { ...nr, range: s };
  }

  return next;
}

/* ------------------------------- Merge commands ------------------------------- */

export class MergeCellsCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private range: string
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      addMerge(s, this.range);
      return s;
    });
  }
  inverse(): Command {
    return new UnmergeCellsCommand(this.label, this.sheetId, this.range);
  }
}

export class UnmergeCellsCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private range: string
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return withSheetTransform(wb, this.sheetId, (s) => {
      removeMerge(s, this.range);
      return s;
    });
  }
  inverse(): Command {
    return new MergeCellsCommand(this.label, this.sheetId, this.range);
  }
}

export class UnmergeAllInRangeCommand implements Command {
  private removed: string[] = [];
  constructor(
    public readonly label: string,
    private sheetId: string,
    private range: string
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    const p = parseRange(this.range);
    if (!sheet || !p) return wb;
    this.removed = sheet.merges.filter((m) => {
      const mp = parseRange(m);
      return mp && mp.sRow <= p.eRow && mp.eRow >= p.sRow && mp.sCol <= p.eCol && mp.eCol >= p.sCol;
    });
    return withSheetTransform(wb, this.sheetId, (s) => {
      const p2 = parseRange(this.range);
      s.merges = s.merges.filter((m) => {
        const mp = parseRange(m);
        return !(mp && p2 && mp.sRow <= p2.eRow && mp.eRow >= p2.sRow && mp.sCol <= p2.eCol && mp.eCol >= p2.sCol);
      });
      return s;
    });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const removed = this.removed;
    return {
      label: 'Restore merges',
      apply: (wbx) =>
        withSheetTransform(wbx, sheetId, (s) => {
          for (const m of removed) addMerge(s, m);
          return s;
        }),
      inverse: () => this
    } as Command;
  }
}

/* --------------------------------- Sheet ops ---------------------------------- */

export function uniqueSheetName(wb: WorkbookDoc, base: string): string {
  const names = new Set(wb.sheets.map((s) => s.name));
  let name = base;
  let i = 1;
  while (names.has(name)) {
    name = base + ' ' + i++;
  }
  return name;
}

export class AddSheetCommand implements Command {
  private newSheetId = uid('sheet');
  private index: number;
  private name: string;
  readonly label: string;
  constructor(name?: string, index?: number) {
    this.name = name ?? '';
    this.label = 'Add sheet';
    this.index = index ?? 99999;
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const finalName = this.name || uniqueSheetName(wb, 'Sheet' + (wb.sheets.length + 1));
    const sheet = createSheet(finalName, { id: this.newSheetId });
    const sheets = wb.sheets.slice();
    const at = Math.min(this.index, sheets.length);
    sheets.splice(at, 0, sheet);
    return {
      ...wb,
      sheets,
      activeSheetId: this.newSheetId,
      metadata: { ...wb.metadata, modified: Date.now() }
    };
  }
  inverse(): Command {
    const id = this.newSheetId;
    return { label: 'Delete sheet', apply: (wbx) => removeSheetImpl(wbx, id, true), inverse: () => this } as Command;
  }
}

export class RemoveSheetCommand implements Command {
  private captured: Worksheet | null = null;
  private capturedIndex = -1;
  private capturedActive = '';
  private capturedNamedRanges: Record<string, NamedRange> = {};
  constructor(private sheetId: string) {
    this.label = 'Delete sheet';
  }
  label: string;
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.capturedIndex = idx;
    this.captured = JSON.parse(JSON.stringify(wb.sheets[idx]));
    this.capturedActive = wb.activeSheetId;
    this.capturedNamedRanges = JSON.parse(JSON.stringify(wb.namedRanges));
    return removeSheetImpl(wb, this.sheetId, false);
  }
  inverse(): Command {
    const sheet = this.captured;
    const index = this.capturedIndex;
    const activeId = this.capturedActive;
    const namedRanges = this.capturedNamedRanges;
    const sheetId = this.sheetId;
    return {
      label: 'Restore sheet',
      apply: (wbx) => {
        if (!sheet) return wbx;
        const sheets = wbx.sheets.slice();
        sheets.splice(Math.min(index, sheets.length), 0, sheet);
        return {
          ...wbx,
          sheets,
          activeSheetId: activeId,
          namedRanges,
          metadata: { ...wbx.metadata, modified: Date.now() }
        };
      },
      inverse: () => {
        return { label: 'Delete sheet', apply: (wbx) => removeSheetImpl(wbx, sheetId, false), inverse: () => this } as Command;
      }
    } as Command;
  }
}

function removeSheetImpl(wb: WorkbookDoc, sheetId: string, _deleteLast: boolean): WorkbookDoc {
  const idx = wb.sheets.findIndex((s) => s.id === sheetId);
  if (idx < 0) return wb;
  const sheets = wb.sheets.filter((s) => s.id !== sheetId);
  if (sheets.length === 0) return wb;
  const nr: Record<string, NamedRange> = {};
  for (const [k, v] of Object.entries(wb.namedRanges)) {
    if (v.sheet !== sheetId) nr[k] = v;
  }
  const active = wb.activeSheetId === sheetId ? sheets[Math.max(0, idx - 1)]?.id : wb.activeSheetId;
  return {
    ...wb,
    sheets,
    activeSheetId: active ?? sheets[0].id,
    namedRanges: nr,
    metadata: { ...wb.metadata, modified: Date.now() }
  };
}

export class RenameSheetCommand implements Command {
  private old = '';
  constructor(private sheetId: string, private newName: string) {
    this.label = 'Rename sheet';
  }
  label: string;
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.old = wb.sheets[idx].name;
    const sheets = wb.sheets.slice();
    sheets[idx] = { ...sheets[idx], name: this.newName };
    const namedRanges = { ...wb.namedRanges };
    for (const [k, v] of Object.entries(namedRanges)) {
      if (v.sheet === this.old) namedRanges[k] = { ...v, sheet: this.newName };
    }
    return { ...wb, sheets, namedRanges, metadata: { ...wb.metadata, modified: Date.now() } };
  }
  inverse(): Command {
    return new RenameSheetCommand(this.sheetId, this.old);
  }
}

export class DuplicateSheetCommand implements Command {
  private newSheetId = uid('sheet');
  private newName = '';
  readonly label: string;
  constructor(private sheetId: string) {
    this.label = 'Duplicate sheet';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    const src = wb.sheets[idx];
    const copy = JSON.parse(JSON.stringify(src)) as Worksheet;
    this.newName = uniqueSheetName(wb, src.name + ' Copy');
    copy.id = this.newSheetId;
    copy.name = this.newName;
    const sheets = wb.sheets.slice();
    sheets.splice(idx + 1, 0, copy);
    return {
      ...wb,
      sheets,
      activeSheetId: this.newSheetId,
      metadata: { ...wb.metadata, modified: Date.now() }
    };
  }
  inverse(): Command {
    const id = this.newSheetId;
    return { label: 'Delete duplicate', apply: (wbx) => removeSheetImpl(wbx, id, true), inverse: () => this } as Command;
  }
}

export class MoveSheetCommand implements Command {
  private prevIndex = 0;
  readonly label: string;
  constructor(private sheetId: string, private toIndex: number) {
    this.label = 'Move sheet';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prevIndex = idx;
    const sheets = wb.sheets.slice();
    const [sheet] = sheets.splice(idx, 1);
    sheets.splice(Math.max(0, Math.min(this.toIndex, sheets.length)), 0, sheet);
    return { ...wb, sheets, metadata: { ...wb.metadata, modified: Date.now() } };
  }
  inverse(): Command {
    return new MoveSheetCommand(this.sheetId, this.prevIndex);
  }
}

export class SetSheetColorCommand implements Command {
  private prev: string | undefined = undefined;
  readonly label: string;
  constructor(private sheetId: string, private color: string | undefined) {
    this.label = 'Sheet tab color';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prev = wb.sheets[idx].color;
    const sheets = wb.sheets.slice();
    sheets[idx] = { ...sheets[idx], color: this.color };
    return { ...wb, sheets };
  }
  inverse(): Command {
    return new SetSheetColorCommand(this.sheetId, this.prev);
  }
}

export class SetSheetStateCommand implements Command {
  private prev: Worksheet['state'] = 'visible';
  readonly label: string;
  constructor(private sheetId: string, private state: 'visible' | 'hidden' | 'veryHidden') {
    this.label = 'Set sheet state';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prev = wb.sheets[idx].state;
    const S = wb.sheets.map((s) => (s.id === this.sheetId ? { ...s, state: this.state } : s));
    return { ...wb, sheets: S };
  }
  inverse(): Command {
    return new SetSheetStateCommand(this.sheetId, this.prev);
  }
}

/* ------------------------------ Named range ops ------------------------------ */

export class SetNamedRangeCommand implements Command {
  private prev: NamedRange | undefined;
  constructor(
    public readonly label: string,
    private name: string,
    private range: string,
    private sheet: string
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const cleanName = this.name.startsWith('=') ? this.name.slice(1) : this.name;
    const key = cleanName.toUpperCase();
    this.prev = wb.namedRanges[key];
    const nr: NamedRange = { name: cleanName, sheet: this.sheet, range: normalizeRange(this.range) };
    return {
      ...wb,
      namedRanges: { ...wb.namedRanges, [key]: nr },
      metadata: { ...wb.metadata, modified: Date.now() }
    };
  }
  inverse(): Command {
    const key = this.name.toUpperCase();
    if (this.prev) {
      return new SetNamedRangeCommand(this.label, this.prev.name, this.prev.range, this.prev.sheet);
    }
    const name = this.name;
    return {
      label: 'Remove named range',
      apply: (wbx) => {
        const nr = { ...wbx.namedRanges };
        delete nr[name.toUpperCase()];
        return { ...wbx, namedRanges: nr };
      },
      inverse: () => this
    } as Command;
  }
}

export class DeleteNamedRangeCommand implements Command {
  private prev: NamedRange | undefined;
  readonly label: string;
  constructor(private name: string) {
    this.label = 'Delete named range';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    this.prev = wb.namedRanges[this.name.toUpperCase()];
    if (!this.prev) return wb;
    const nr = { ...wb.namedRanges };
    delete nr[this.name.toUpperCase()];
    return { ...wb, namedRanges: nr, metadata: { ...wb.metadata, modified: Date.now() } };
  }
  inverse(): Command {
    if (!this.prev) return new NoopCommand('No changes');
    const key = this.prev.name;
    const sheet = this.prev.sheet;
    const range = this.prev.range;
    return new SetNamedRangeCommand('Restore named range', key, range, sheet);
  }
}

/* ------------------------------ Charts / tables ------------------------------ */

export class AddChartCommand implements Command {
  private id = uid('chart');
  constructor(
    public readonly label: string,
    private sheetId: string,
    private chart: Omit<ChartDef, 'id'>
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const chart = { ...this.chart, id: this.id };
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.charts = [...s.charts, chart];
      return s;
    });
  }
  inverse(): Command {
    return new RemoveChartCommand('Delete chart', this.sheetId, this.id);
  }
}

export class RemoveChartCommand implements Command {
  private prev: ChartDef | null = null;
  constructor(
    public readonly label: string,
    private sheetId: string,
    private chartId: string
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    this.prev = sheet.charts.find((c) => c.id === this.chartId) ?? null;
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.charts = s.charts.filter((c) => c.id !== this.chartId);
      return s;
    });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const prev = this.prev;
    const chartId = this.chartId;
    return {
      label: 'Restore chart',
      apply: (wbx) =>
        prev
          ? withSheetTransform(wbx, sheetId, (s) => {
              s.charts = [...s.charts, { ...prev, id: chartId }];
              return s;
            })
          : wbx,
      inverse: () => this
    } as Command;
  }
}

export class UpdateChartCommand implements Command {
  private prev: ChartDef | null = null;
  constructor(
    public readonly label: string,
    private sheetId: string,
    private chartId: string,
    private patch: Partial<ChartDef>
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    this.prev = sheet.charts.find((c) => c.id === this.chartId) ?? null;
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.charts = s.charts.map((c) => (c.id === this.chartId ? { ...c, ...this.patch } : c));
      return s;
    });
  }
  inverse(): Command {
    if (!this.prev) return new NoopCommand('No changes');
    const p = this.prev;
    return new UpdateChartCommand(this.label, this.sheetId, this.chartId, {
      type: p.type,
      title: p.title,
      dataRange: p.dataRange,
      x: p.x,
      y: p.y,
      width: p.width,
      height: p.height,
      options: p.options
    });
  }
}

export class AddTableCommand implements Command {
  private id = uid('table');
  constructor(
    public readonly label: string,
    private sheetId: string,
    private table: Omit<TableDef, 'id'>
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const table = { ...this.table, id: this.id };
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.tables = [...s.tables, table];
      return s;
    });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const id = this.id;
    return {
      label: 'Delete table',
      apply: (wbx) =>
        withSheetTransform(wbx, sheetId, (s) => {
          s.tables = s.tables.filter((t) => t.id !== id);
          return s;
        }),
      inverse: () => this
    } as Command;
  }
}

export class RemoveTableCommand implements Command {
  private prev: TableDef | null = null;
  readonly label: string;
  constructor(private sheetId: string, private tableId: string) {
    this.label = 'Delete table';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    this.prev = sheet.tables.find((t) => t.id === this.tableId) ?? null;
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.tables = s.tables.filter((t) => t.id !== this.tableId);
      return s;
    });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const prev = this.prev;
    const tableId = this.tableId;
    return {
      label: 'Restore table',
      apply: (wbx) =>
        prev
          ? withSheetTransform(wbx, sheetId, (s) => {
              s.tables = [...s.tables, { ...prev, id: tableId }];
              return s;
            })
          : wbx,
      inverse: () => this
    } as Command;
  }
}

/* ------------------------- Conditional formatting ------------------------- */

export class AddConditionalFormatCommand implements Command {
  constructor(
    public readonly label: string,
    private sheetId: string,
    private rule: ConditionalFormatRule
  ) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const rule = this.rule;
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.conditionalFormats = [...s.conditionalFormats, { ...rule }];
      return s;
    });
  }
  inverse(): Command {
    const rule = this.rule;
    const sheetId = this.sheetId;
    return {
      label: 'Remove conditional format',
      apply: (wbx) =>
        withSheetTransform(wbx, sheetId, (s) => {
          s.conditionalFormats.splice(s.conditionalFormats.indexOf(rule), 1);
          return s;
        }),
      inverse: () => this
    } as Command;
  }
}

export class ClearConditionalFormatsCommand implements Command {
  private prev: ConditionalFormatRule[] = [];
  readonly label: string;
  constructor(private sheetId: string) {
    this.label = 'Clear conditional formats';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    this.prev = sheet.conditionalFormats.slice();
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.conditionalFormats = [];
      return s;
    });
  }
  inverse(): Command {
    const prev = this.prev;
    const sheetId = this.sheetId;
    return {
      label: 'Restore conditional formats',
      apply: (wbx) =>
        withSheetTransform(wbx, sheetId, (s) => {
          s.conditionalFormats = prev.slice();
          return s;
        }),
      inverse: () => this
    } as Command;
  }
}

/* ------------------------------- Validation ------------------------------- */

export class SetValidationCommand implements Command {
  private before: Record<string, DataValidationRule | undefined> = {};
  constructor(
    public readonly label: string,
    private sheetId: string,
    private sRow: number,
    private sCol: number,
    private eRow: number,
    private eCol: number,
    private rule: DataValidationRule | undefined
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    for (let r = this.sRow; r <= this.eRow; r++) {
      for (let c = this.sCol; c <= this.eCol; c++) {
        const ref = cellToA1(c, r);
        if (sheet.dataValidations[ref]) this.before[ref] = JSON.parse(JSON.stringify(sheet.dataValidations[ref]));
        else this.before[ref] = undefined;
      }
    }
    return withSheetTransform(wb, this.sheetId, (s) => {
      for (let r = this.sRow; r <= this.eRow; r++) {
        for (let c = this.sCol; c <= this.eCol; c++) {
          const ref = cellToA1(c, r);
          if (this.rule && this.rule.type !== 'none') s.dataValidations[ref] = JSON.parse(JSON.stringify(this.rule));
          else delete s.dataValidations[ref];
        }
      }
      return s;
    });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const before = this.before;
    const sRow = this.sRow;
    const sCol = this.sCol;
    const eRow = this.eRow;
    const eCol = this.eCol;
    return {
      label: 'Restore validation',
      apply: (wbx) =>
        withSheetTransform(wbx, sheetId, (s) => {
          for (let r = sRow; r <= eRow; r++) {
            for (let c = sCol; c <= eCol; c++) {
              const ref = cellToA1(c, r);
              if (before[ref] !== undefined) s.dataValidations[ref] = JSON.parse(JSON.stringify(before[ref]));
              else delete s.dataValidations[ref];
            }
          }
          return s;
        }),
      inverse: () => this
    } as Command;
  }
}

/* --------------------------------- Sort --------------------------------- */

export class SortRangeCommand implements Command {
  private before: Record<string, Cell | undefined> = {};
  constructor(
    public readonly label: string,
    private sheetId: string,
    private sRow: number,
    private sCol: number,
    private eRow: number,
    private eCol: number,
    private keys: Array<{ col: number; desc?: boolean }>,
    private options?: { header?: boolean }
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    const refs = rectRefs(sheet, this.sRow, this.sCol, this.eRow, this.eCol);
    this.before = captureCells(sheet, refs);
    const startRow = this.options?.header ? this.sRow + 1 : this.sRow;
    const rows = [];
    for (let r = startRow; r <= this.eRow; r++) rows.push(r);

    const getVal = (r: number, col: number): CellPrimitive | ErrorValue | null => {
      const cell = sheet.cells[cellToA1(col, r)];
      if (!cell) return null;
      if (cell.formula != null) return cell.computed ?? cell.value ?? null;
      return cell.value ?? null;
    };
    const cmp = (a: number, b: number): number => {
      for (const key of this.keys) {
        const va = getVal(a, key.col);
        const vb = getVal(b, key.col);
        let r: number;
        if (typeof va === 'number' && typeof vb === 'number') r = va - vb;
        else {
          const sa = va === null || va === undefined ? '' : String(va);
          const sb = vb === null || vb === undefined ? '' : String(vb);
          r = sa.localeCompare(sb, undefined, { numeric: true, sensitivity: 'base' });
        }
        if (r !== 0) return key.desc ? -r : r;
      }
      return 0;
    };
    rows.sort(cmp);

    const newMap: Record<string, Cell | undefined> = {};
    const colCount = this.eCol - this.sCol + 1;
    rows.forEach((origRow, idx) => {
      const targetRow = startRow + idx;
      for (let c = 0; c < colCount; c++) {
        const origCell = sheet.cells[cellToA1(this.sCol + c, origRow)];
        if (origCell) newMap[cellToA1(this.sCol + c, targetRow)] = JSON.parse(JSON.stringify(origCell));
      }
    });
    return withSheetCells(wb, this.sheetId, newMap);
  }
  inverse(): Command {
    return new EditCellsCommand(this.label, this.sheetId, this.before);
  }
}

/**
 * Excel's Data > Remove Duplicates: keeps the first occurrence of each row and
 * compacts the survivors up, clearing whatever rows are left over.
 */
export class RemoveDuplicatesCommand implements Command {
  private before: Record<string, Cell | undefined> = {};
  private removed = 0;
  constructor(
    public readonly label: string,
    private sheetId: string,
    private sRow: number,
    private sCol: number,
    private eRow: number,
    private eCol: number,
    private hasHeader: boolean
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    const refs = rectRefs(sheet, this.sRow, this.sCol, this.eRow, this.eCol);
    this.before = captureCells(sheet, refs);
    const startRow = this.hasHeader ? this.sRow + 1 : this.sRow;
    if (this.eRow < startRow) return wb;

    const colCount = this.eCol - this.sCol + 1;
    const keyOf = (r: number): string => {
      const parts: string[] = [];
      for (let c = 0; c < colCount; c++) {
        const cell = sheet.cells[cellToA1(this.sCol + c, r)];
        if (!cell) {
          parts.push('');
          continue;
        }
        parts.push(String(cell.formula != null ? cell.computed ?? cell.value ?? '' : cell.value ?? ''));
      }
      return JSON.stringify(parts);
    };

    const seen = new Set<string>();
    const keep: number[] = [];
    for (let r = startRow; r <= this.eRow; r++) {
      const key = keyOf(r);
      if (seen.has(key)) continue;
      seen.add(key);
      keep.push(r);
    }
    this.removed = this.eRow - startRow + 1 - keep.length;

    const newMap: Record<string, Cell | undefined> = {};
    for (let r = startRow; r <= this.eRow; r++) {
      const targetIdx = r - startRow;
      for (let c = 0; c < colCount; c++) {
        const targetRef = cellToA1(this.sCol + c, r);
        if (targetIdx >= keep.length) {
          // Leftover tail: clear it, since withSheetCells only overwrites.
          newMap[targetRef] = undefined;
          continue;
        }
        const srcCell = sheet.cells[cellToA1(this.sCol + c, keep[targetIdx])];
        newMap[targetRef] = srcCell ? JSON.parse(JSON.stringify(srcCell)) : undefined;
      }
    }
    return withSheetCells(wb, this.sheetId, newMap);
  }

  /** How many duplicate rows the last apply() discarded. */
  get removedCount(): number {
    return this.removed;
  }

  inverse(): Command {
    return new EditCellsCommand(this.label, this.sheetId, this.before);
  }
}

/* -------------------------------- View state -------------------------------- */

export class SetZoomCommand implements Command {
  private prev = 100;
  readonly label: string;
  constructor(private sheetId: string, private zoom: number) {
    this.label = 'Set zoom';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prev = wb.sheets[idx].view?.zoom ?? 100;
    const sheets = wb.sheets.slice();
    sheets[idx] = {
      ...sheets[idx],
      view: { ...(sheets[idx].view ?? {}), zoom: this.zoom }
    };
    return { ...wb, sheets };
  }
  inverse(): Command {
    return new SetZoomCommand(this.sheetId, this.prev);
  }
}

export class SetFreezePanesCommand implements Command {
  private prev: Worksheet['freeze'] = {};
  readonly label: string;
  constructor(private sheetId: string, private freeze: Worksheet['freeze']) {
    this.label = 'Freeze panes';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prev = JSON.parse(JSON.stringify(wb.sheets[idx].freeze ?? {}));
    const sheets = wb.sheets.slice();
    sheets[idx] = { ...sheets[idx], freeze: this.freeze };
    return { ...wb, sheets };
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev || Object.keys(prev).length === 0) {
      return new SetFreezePanesCommand(this.sheetId, {});
    }
    return new SetFreezePanesCommand(this.sheetId, { row: prev.row, col: prev.col });
  }
}

export class SetThemeCommand implements Command {
  private prev: ThemeDef | undefined;
  readonly label = 'Set theme';
  /** Patch the workbook theme; Theme.gentle is unused here and kept for parity. */
  constructor(private patch: Partial<ThemeDef>) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    this.prev = wb.theme;
    const base = getTheme(wb.theme);
    return { ...wb, theme: { ...base, ...this.patch } };
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev) return new ClearThemeCommand();
    return new SetThemeCommand(prev);
  }
}

export class ClearThemeCommand implements Command {
  readonly label = 'Reset theme';
  private prev: ThemeDef | undefined;
  apply(wb: WorkbookDoc): WorkbookDoc {
    this.prev = wb.theme;
    const { theme, ...rest } = wb;
    return rest as WorkbookDoc;
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev) return new NoopCommand('No changes');
    return new SetThemeCommand(prev);
  }
}

export class SetPageSetupCommand implements Command {
  private prev: PageSetup | undefined;
  readonly label: string;
  /** replace = overwrite the whole pageSetup (undo); default = merge a partial patch. */
  constructor(private sheetId: string, private setup: PageSetup, private replace = false) {
    this.label = 'Page setup';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prev = wb.sheets[idx].pageSetup;
    const sheets = wb.sheets.slice();
    // The ribbon sends partial patches (e.g. only orientation), so merge by default.
    const base = this.replace ? {} : { ...(sheets[idx].pageSetup ?? {}) };
    sheets[idx] = { ...sheets[idx], pageSetup: { ...base, ...this.setup } };
    return { ...wb, sheets };
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev || Object.keys(prev).length === 0) {
      // Restore "no page setup at all" rather than a partial object.
      return new ClearPageSetupCommand(this.sheetId);
    }
    return new SetPageSetupCommand(this.sheetId, prev, true);
  }
}

/** Replaces the whole pageSetup object (used by undo when it was previously empty). */
export class ClearPageSetupCommand implements Command {
  readonly label = 'Reset page setup';
  private had = false;
  private snapshot: PageSetup | undefined;
  constructor(private sheetId: string) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.had = wb.sheets[idx].pageSetup !== undefined;
    this.snapshot = wb.sheets[idx].pageSetup;
    const sheets = wb.sheets.slice();
    const { pageSetup, ...rest } = sheets[idx];
    sheets[idx] = rest as Worksheet;
    return { ...wb, sheets };
  }
  inverse(): Command {
    if (!this.had) return new NoopCommand('No changes');
    const snapshot = this.snapshot;
    return {
      label: 'Restore page setup',
      apply: (wbx) => {
        const i = wbx.sheets.findIndex((s) => s.id === this.sheetId);
        if (i < 0) return wbx;
        const sheets = wbx.sheets.slice();
        sheets[i] = { ...sheets[i], pageSetup: snapshot ? { ...snapshot } : undefined };
        return { ...wbx, sheets };
      },
      inverse: () => this
    } as Command;
  }
}

export class SetPrintAreaCommand implements Command {
  private prev: string | undefined;
  readonly label: string;
  constructor(private sheetId: string, private printArea: string | undefined) {
    this.label = 'Set print area';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    this.prev = wb.sheets[idx].pageSetup?.printArea;
    const sheets = wb.sheets.slice();
    sheets[idx] = {
      ...sheets[idx],
      pageSetup: { ...(sheets[idx].pageSetup ?? {}), printArea: this.printArea }
    };
    return { ...wb, sheets };
  }
  inverse(): Command {
    return new SetPrintAreaCommand(this.sheetId, this.prev);
  }
}

/* --------------------------------- Notes --------------------------------- */

export class SetNoteCommand implements Command {
  private prevNote: NoteDef | undefined;
  readonly label: string;
  constructor(private sheetId: string, private ref: string, private note: string) {
    this.label = note.trim() ? 'Add note' : 'Delete note';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    this.prevNote = sheet.notes.find((n) => n.ref === this.ref);
    const ref = this.ref;
    const text = this.note.trim();
    const author = wb.metadata.author || undefined;
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.notes = s.notes.filter((n) => n.ref !== ref);
      // Empty text removes the note, so no empty marker is left behind.
      if (text) s.notes.push({ ref, text, author });
      return s;
    });
  }
  inverse(): Command {
    const sheetId = this.sheetId;
    const ref = this.ref;
    const prev = this.prevNote;
    return {
      label: 'Restore note',
      apply: (wbx) =>
        withSheetTransform(wbx, sheetId, (s) => {
          s.notes = s.notes.filter((n) => n.ref !== ref);
          if (prev) s.notes.push({ ...prev });
          return s;
        }),
      inverse: () => this
    } as Command;
  }
}

export class AddShapeCommand implements Command {
  readonly label: string;
  constructor(private sheetId: string, private shape: ShapeDef) {
    this.label = 'Insert shape';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    const sheets = wb.sheets.slice();
    const list = sheets[idx].shapes ?? [];
    // Newest shape goes on top.
    const topZ = list.reduce((m, s) => Math.max(m, s.z ?? 0), 0) + 1;
    sheets[idx] = { ...sheets[idx], shapes: [...list, { ...this.shape, z: topZ }] };
    return { ...wb, sheets };
  }
  inverse(): Command {
    const id = this.shape.id;
    const sheetId = this.sheetId;
    return {
      label: 'Remove shape',
      apply: (wbx) => {
        const i = wbx.sheets.findIndex((s) => s.id === sheetId);
        if (i < 0) return wbx;
        const sheets = wbx.sheets.slice();
        sheets[i] = { ...sheets[i], shapes: (sheets[i].shapes ?? []).filter((s) => s.id !== id) };
        return { ...wbx, sheets };
      },
      inverse: () => this
    } as Command;
  }
}

export class UpdateShapeCommand implements Command {
  private prev: ShapeDef | undefined;
  readonly label: string;
  constructor(private sheetId: string, private id: string, private patch: Partial<ShapeDef>) {
    this.label = 'Update shape';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    const list = sheets_at(wb, idx);
    const current = list.find((s) => s.id === this.id);
    if (!current) return wb;
    this.prev = current;
    const sheets = wb.sheets.slice();
    sheets[idx] = {
      ...sheets[idx],
      shapes: list.map((s) => (s.id === this.id ? { ...s, ...this.patch } : s))
    };
    return { ...wb, sheets };
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev) return new NoopCommand('No changes');
    // Replace the whole shape: spreading `prev` would leave keys the patch
    // introduced (e.g. a fill that did not exist before) in place.
    return new ReplaceShapeCommand(this.sheetId, prev);
  }
}

class ReplaceShapeCommand implements Command {
  readonly label: string;
  private prev: ShapeDef | undefined;
  constructor(private sheetId: string, private shape: ShapeDef) {
    this.label = 'Update shape';
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    const list = sheets_at(wb, idx);
    const current = list.find((s) => s.id === this.shape.id);
    if (!current) return wb;
    this.prev = current;
    const sheets = wb.sheets.slice();
    sheets[idx] = {
      ...sheets[idx],
      shapes: list.map((s) => (s.id === this.shape.id ? { ...this.shape } : s))
    };
    return { ...wb, sheets };
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev) return new NoopCommand('No changes');
    return new ReplaceShapeCommand(this.sheetId, prev);
  }
}

function sheets_at(wb: WorkbookDoc, idx: number): ShapeDef[] {
  return wb.sheets[idx].shapes ?? [];
}

export class RemoveShapeCommand implements Command {
  private prev: ShapeDef | undefined;
  readonly label = 'Delete shape';
  constructor(private sheetId: string, private id: string) {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    const idx = wb.sheets.findIndex((s) => s.id === this.sheetId);
    if (idx < 0) return wb;
    const list = sheets_at(wb, idx);
    this.prev = list.find((s) => s.id === this.id);
    if (!this.prev) return wb;
    const sheets = wb.sheets.slice();
    sheets[idx] = { ...sheets[idx], shapes: list.filter((s) => s.id !== this.id) };
    return { ...wb, sheets };
  }
  inverse(): Command {
    const prev = this.prev;
    if (!prev) return new NoopCommand('No changes');
    return new AddShapeCommand(this.sheetId, prev);
  }
}

export class SetSheetProtectedCommand implements Command {
  readonly label: string;
  constructor(private sheetId: string, private protectedFlag: boolean) {
    this.label = protectedFlag ? 'Protect sheet' : 'Unprotect sheet';
  }
  isProtecting(): boolean {
    return this.protectedFlag;
  }
  apply(wb: WorkbookDoc): WorkbookDoc {
    const sheet = getSheet(wb, this.sheetId);
    if (!sheet) return wb;
    return withSheetTransform(wb, this.sheetId, (s) => {
      s.protected = this.protectedFlag;
      return s;
    });
  }
  inverse(): Command {
    return new SetSheetProtectedCommand(this.sheetId, !this.protectedFlag);
  }
}

export { NoopCommand };
