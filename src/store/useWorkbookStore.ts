/**
 * Central workbook store.
 *
 * Owns the immutable WorkbookDoc, the undo/redo history and the formula
 * engine. All mutations go through `applyCommand`, which produces a new
 * workbook, records history, re-attaches the engine and bumps a revision so
 * cached computations are refreshed.
 */

import { create } from 'zustand';
import { WorkbookDoc, Worksheet, CellPrimitive, CellStyle, ErrorValue, Cell, ShapeDef, ShapeKind } from '@/engine/model/types';
import { createWorkbook, getSheet, getCell, cloneWorkbook } from '@/engine/model/workbook';
import { HistoryManager, Command } from '@/engine/commands/history';
import { WorkbookFormulaEngine } from '@/engine/formula/engine';
import { cellToA1, parseA1 } from '@/engine/model/address';
import {
  EditCellsCommand,
  SetRangeValuesCommand,
  FormatCellsCommand,
  ClearRangeCommand,
  SetNoteCommand,
  SetSheetProtectedCommand,
  SetZoomCommand,
  SetFreezePanesCommand,
  SetPageSetupCommand,
  SetPrintAreaCommand,
  SetSheetColorCommand,
  SetSheetStateCommand,
  RenameSheetCommand,
  MoveSheetCommand,
  AddSheetCommand,
  DuplicateSheetCommand,
  RemoveSheetCommand,
  AddShapeCommand,
  UpdateShapeCommand,
  RemoveShapeCommand
} from '@/engine/commands/commands';
import { setClipboard, getClipboard, ClipboardCell } from '@/store/clipboard';
import { useUiStore } from '@/store/useUiStore';

export interface Selection {
  row: number; // 1-based
  col: number; // 0-based
  endRow: number; // 1-based
  endCol: number; // 0-based
}

export const SINGLE = (row: number, col: number): Selection => ({ row, col, endRow: row, endCol: col });

/**
 * Commands that Excel still permits on a protected sheet: view/print settings
 * and workbook-level sheet structure (add, move, rename, hide, colour).
 * Everything else (edits, formatting, sort, rows/cols, charts, comments) is blocked.
 */
function allowedWhileProtected(cmd: Command): boolean {
  return (
    cmd instanceof SetSheetProtectedCommand ||
    cmd instanceof SetZoomCommand ||
    cmd instanceof SetFreezePanesCommand ||
    cmd instanceof SetPageSetupCommand ||
    cmd instanceof SetPrintAreaCommand ||
    cmd instanceof SetSheetColorCommand ||
    cmd instanceof SetSheetStateCommand ||
    cmd instanceof RenameSheetCommand ||
    cmd instanceof MoveSheetCommand ||
    cmd instanceof AddSheetCommand ||
    cmd instanceof DuplicateSheetCommand ||
    cmd instanceof RemoveSheetCommand
  );
}

interface WorkbookState {
  workbook: WorkbookDoc | null;
  activeSheetId: string | null;
  selection: Selection;
  /** Chart currently selected on the active sheet, or null when cells are selected. */
  selectedChartId: string | null;
  /** Shape currently selected on the active sheet, or null. */
  selectedShapeId: string | null;
  /** Show note markers/tooltips (Excel: "Show Comments"). */
  showNotes: boolean;
  editing: { row: number; col: number; text: string; seeded?: boolean } | null;
  history: HistoryManager;
  engine: WorkbookFormulaEngine;
  /** Incremented on every structural change; used to refresh memoized cell values. */
  revision: number;
  dirty: boolean;
  /** Absolute path of the file this workbook was opened from / last saved to. */
  filePath: string | null;

  newWorkbook(title?: string): void;
  loadWorkbook(wb: WorkbookDoc): void;
  /** Save back to `filePath`, or open Save As when there is no path yet. */
  save(): Promise<void>;
  setFilePath(path: string | null): void;
  /** Record a successful save: remember the path (if any) and clear `dirty`. */
  markSaved(path?: string | null): void;
  applyCommand(cmd: Command): void;
  /** True when the active sheet is protected; edits and formatting are blocked. */
  isSheetProtected(): boolean;
  /** Comment on the active cell (or the whole selection's top-left). */
  editNoteOnSelection(text: string): void;
  /** Delete the note on the active cell, if any. */
  deleteNoteAtSelection(): void;
  /** Toggle note markers/tooltips. */
  setShowNotes(on: boolean): void;
  /** Select a shape (null clears). Selecting a shape clears chart/cell selection. */
  setSelectedShapeId(id: string | null): void;
  /** Insert a shape at sheet coordinates and select it. */
  addShape(kind: ShapeKind, x: number, y: number, width: number, height: number): void;
  /** Patch a shape's geometry/style. */
  updateShape(id: string, patch: Partial<ShapeDef>): void;
  /** Delete a shape by id. */
  removeShape(id: string): void;
  /** Change draw order for the selected shape. */
  reorderShape(id: string, dir: 'front' | 'back' | 'forward' | 'backward'): void;
  /** Jump to the previous/next cell that has a note. Returns false if none. */
  gotoNote(dir: -1 | 1): boolean;
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  canRedo(): boolean;
  recalculate(): void;

  // Selection / navigation
  setSelection(sel: Selection): void;
  setSelectedChartId(id: string | null): void;
  selectActive(): void;
  moveActive(dRow: number, dCol: number, extend?: boolean): void;

  // Edits
  commitEdit(text: string): void;
  cancelEdit(): void;
  startEdit(initialText?: string): void;
  setCellRaw(row: number, col: number, text: string): void;
  setCellValue(row: number, col: number, value: CellPrimitive): void;
  clearSelection(): void;

  // Clipboard
  copySelection(): void;
  cutSelection(): void;
  paste(): void | Promise<void>;

  // Find & replace
  findCells(query: string, opts?: { matchCase?: boolean; wholeCell?: boolean }): Array<{ row: number; col: number; hit: string }>;
  replaceCells(query: string, replacement: string, opts?: { matchCase?: boolean; wholeCell?: boolean }): number;

  // Formatting helpers
  formatSelection(patch: Partial<CellStyle>): void;

  // Sheet ops
  setActiveSheet(id: string): void;
}

function makeEngine() {
  return new WorkbookFormulaEngine();
}

/** Best-effort mirror of a copied range onto the OS clipboard (tab-separated). */
async function writeSelectionToOsClipboard(
  getState: () => WorkbookState,
  sheet: Worksheet,
  activeSheetId: string
): Promise<void> {
  try {
    const { selection, engine } = getState();
    if (!navigator.clipboard?.writeText) return;
    const sRow = Math.min(selection.row, selection.endRow);
    const eRow = Math.max(selection.row, selection.endRow);
    const sCol = Math.min(selection.col, selection.endCol);
    const eCol = Math.max(selection.col, selection.endCol);
    const lines: string[] = [];
    for (let r = sRow; r <= eRow; r++) {
      const parts: string[] = [];
      for (let c = sCol; c <= eCol; c++) {
        const cell = getCell(sheet, c, r);
        let value: unknown = cell?.value;
        if (cell?.formula) {
          try {
            value = engine.compute(activeSheetId, r, c);
          } catch {
            value = cell.formula;
          }
        }
        parts.push(value == null || value === '' ? '' : String(value));
      }
      lines.push(parts.join('\t'));
    }
    await navigator.clipboard.writeText(lines.join('\r\n'));
  } catch {
    // Clipboard permission / availability issues are non-blocking: the internal
    // clipboard still holds the full snapshot (values, formulas and styles).
  }
}

async function readOsClipboardText(): Promise<string | null> {
  try {
    if (navigator.clipboard?.readText) return await navigator.clipboard.readText();
  } catch {
    return null;
  }
  return null;
}

function makeHistory() {
  return new HistoryManager();
}

// Boot straight into a blank worksheet instead of the start screen: one store is
// created per app launch, so this simply makes the grid the first thing shown.
const bootWorkbook = createWorkbook({ title: 'Untitled' });
const bootEngine = makeEngine();
bootEngine.attach(bootWorkbook, bootWorkbook.activeSheetId);

export const useWorkbookStore = create<WorkbookState>((set, get) => ({
  workbook: bootWorkbook,
  activeSheetId: bootWorkbook.activeSheetId,
  selection: SINGLE(1, 0),
  selectedChartId: null,
  selectedShapeId: null,
  showNotes: true,
  editing: null,
  history: makeHistory(),
  engine: bootEngine,
  revision: 0,
  dirty: false,
  filePath: null,

  newWorkbook(title = 'Untitled') {
    const wb = createWorkbook({ title });
    const engine = makeEngine();
    const history = makeHistory();
    engine.attach(wb, wb.activeSheetId);
    set({
      workbook: wb,
      activeSheetId: wb.activeSheetId,
      selection: SINGLE(1, 0),
      editing: null,
      history,
      engine,
      revision: 0,
      filePath: null,
      dirty: true
    });
  },

  loadWorkbook(wb) {
    const engine = makeEngine();
    const history = makeHistory();
    engine.attach(wb, wb.activeSheetId);
    set({
      workbook: wb,
      activeSheetId: wb.activeSheetId,
      selection: SINGLE(1, 0),
      editing: null,
      history,
      engine,
      revision: 0,
      filePath: null,
      dirty: false
    });
  },

  /**
   * Ctrl+S / File > Save.
   *
   * Once the workbook has a path on disk this writes straight back to it
   * (atomic write in the desktop build); a brand new workbook is routed
   * through Save As instead, and the browser preview falls back to a
   * download because it has no filesystem access.
   */
  async save() {
    const { workbook, filePath } = get();
    if (!workbook) return;
    if (!filePath) {
      useUiStore.getState().openDialog('saveAs');
      return;
    }
    const name = filePath.split(/[\\/]/).pop() ?? filePath;
    try {
      const bridge = await import('@/platform/bridge');
      const { serializeWorkbook } = await import('@/engine/io/serialize');
      const json = serializeWorkbook(workbook);
      if (bridge.isTauri()) {
        await bridge.saveWorkbookFile(filePath, json);
        get().markSaved(filePath);
        useUiStore
          .getState()
          .addRecent({ path: filePath, name, modified: Date.now() });
        void bridge.addRecent(filePath).catch(() => {});
        useUiStore.getState().toast('success', `Saved ${name}`);
      } else {
        downloadBlob(json, name, 'application/json');
        get().markSaved(filePath);
        useUiStore.getState().toast('success', `Downloaded ${name}`);
      }
    } catch (err) {
      useUiStore.getState().toast('error', `Could not save ${name}: ${(err as Error).message}`);
    }
  },

  setFilePath(path) {
    set({ filePath: path });
  },

  markSaved(path) {
    // Omitting `path` clears `dirty` but keeps the current file association.
    set(path === undefined ? { dirty: false } : { dirty: false, filePath: path });
  },

  applyCommand(cmd) {
    const { workbook, history, engine, activeSheetId } = get();
    if (!workbook) return;
    // Sheet protection: view/print/sheet-structure commands still work.
    const sheet = activeSheetId ? getSheet(workbook, activeSheetId) : null;
    if (sheet?.protected && !allowedWhileProtected(cmd)) {
      useUiStore
        .getState()
        .toast('error', `"${sheet.name}" is protected. Unprotect it on the Review tab to make changes.`);
      return;
    }
    const next = cmd.apply(workbook);
    if (next === workbook) return;
    history.execute(cmd);
    if (cmd instanceof SetSheetProtectedCommand && (cmd as SetSheetProtectedCommand).isProtecting()) {
      // Excel drops the undo stack when a sheet is protected.
      history.clear();
    }
    engine.attach(next, activeSheetId ?? next.activeSheetId);
    set((s) => ({
      workbook: next,
      dirty: true,
      revision: s.revision + 1
    }));
  },

  isSheetProtected() {
    const { workbook, activeSheetId } = get();
    if (!workbook || !activeSheetId) return false;
    return getSheet(workbook, activeSheetId)?.protected === true;
  },

  editNoteOnSelection(text) {
    const { activeSheetId, selection } = get();
    if (!activeSheetId) return;
    const ref = cellToA1(Math.min(selection.col, selection.endCol), Math.min(selection.row, selection.endRow));
    get().applyCommand(new SetNoteCommand(activeSheetId, ref, text));
  },

  deleteNoteAtSelection() {
    const { activeSheetId, selection, workbook } = get();
    if (!activeSheetId || !workbook) return;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return;
    const ref = cellToA1(Math.min(selection.col, selection.endCol), Math.min(selection.row, selection.endRow));
    if (!sheet.notes.some((n) => n.ref === ref)) return;
    get().applyCommand(new SetNoteCommand(activeSheetId, ref, ''));
  },

  gotoNote(dir) {
    const { workbook, activeSheetId, selection } = get();
    if (!workbook || !activeSheetId) return false;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet || sheet.notes.length === 0) return false;
    // Document order: column A first, then by row, then by column.
    const pos = (ref: string) => {
      const p = parseA1(ref);
      return p ? { order: p.col * 1_000_000 + p.row, row: p.row, col: p.col } : null;
    };
    const ordered = sheet.notes
      .map((n) => ({ note: n, at: pos(n.ref) }))
      .filter((x): x is { note: (typeof sheet.notes)[number]; at: NonNullable<ReturnType<typeof pos>> } => !!x.at)
      .sort((a, b) => a.at.order - b.at.order);
    if (ordered.length === 0) return false;

    const cur = pos(cellToA1(Math.min(selection.col, selection.endCol), Math.min(selection.row, selection.endRow)));
    if (!cur) return false;
    const exact = ordered.findIndex((x) => x.at.order === cur.order);
    let target;
    if (exact >= 0) {
      target = dir === 1 ? ordered[(exact + 1) % ordered.length] : ordered[(exact - 1 + ordered.length) % ordered.length];
    } else {
      const after = dir === 1 ? ordered.find((x) => x.at.order > cur.order) : [...ordered].reverse().find((x) => x.at.order < cur.order);
      if (!after) return false;
      target = after;
    }
    get().setSelection(SINGLE(target.at.row, target.at.col));
    return true;
  },

  undo() {
    const { history, workbook, engine, activeSheetId } = get();
    if (!workbook) return;
    const cmd = history.popUndo();
    if (!cmd) return;
    const next = cmd.inverse().apply(workbook);
    history.pushRedo(cmd);
    engine.attach(next, activeSheetId ?? next.activeSheetId);
    set((s) => ({
      workbook: next,
      dirty: true,
      revision: s.revision + 1
    }));
  },

  redo() {
    const { history, workbook, engine, activeSheetId } = get();
    if (!workbook) return;
    const cmd = history.popRedo();
    if (!cmd) return;
    const next = cmd.apply(workbook);
    history.execute(cmd);
    engine.attach(next, activeSheetId ?? next.activeSheetId);
    set((s) => ({
      workbook: next,
      dirty: true,
      revision: s.revision + 1
    }));
  },

  canUndo() {
    return get().history.canUndo();
  },

  canRedo() {
    return get().history.canRedo();
  },

  recalculate() {
    get().engine.invalidateAll();
    set((s) => ({ revision: s.revision + 1 }));
  },

  setSelection(sel) {
    // Clicking a cell always drops any selected chart or shape.
    set({ selection: sel, selectedChartId: null, selectedShapeId: null });
  },

  setSelectedChartId(id) {
    set({ selectedChartId: id, ...(id ? { selectedShapeId: null } : {}) });
  },

  setShowNotes(on) {
    set({ showNotes: on });
  },

  setSelectedShapeId(id) {
    set({ selectedShapeId: id, ...(id ? { selectedChartId: null } : {}) });
  },

  addShape(kind, x, y, width, height) {
    const { activeSheetId } = get();
    if (!activeSheetId) return;
    const id = `shape_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
    const shape: ShapeDef = {
      id,
      kind,
      x: Math.max(0, Math.round(x)),
      y: Math.max(0, Math.round(y)),
      width: Math.max(20, Math.round(width)),
      height: Math.max(16, Math.round(height)),
      ...(kind === 'line' || kind === 'arrow'
        ? { stroke: '#1f2937', strokeWidth: 2, fill: undefined }
        : { fill: '#dbeafe', stroke: '#1f2937', strokeWidth: 1 }),
      ...(kind === 'textbox' ? { text: '', fontSize: 13, fill: undefined, stroke: undefined } : {})
    };
    get().applyCommand(new AddShapeCommand(activeSheetId, shape));
    set({ selectedShapeId: id, selectedChartId: null });
  },

  updateShape(id, patch) {
    const { activeSheetId } = get();
    if (!activeSheetId) return;
    get().applyCommand(new UpdateShapeCommand(activeSheetId, id, patch));
  },

  removeShape(id) {
    const { activeSheetId } = get();
    if (!activeSheetId) return;
    get().applyCommand(new RemoveShapeCommand(activeSheetId, id));
    if (get().selectedShapeId === id) set({ selectedShapeId: null });
  },

  reorderShape(id, dir) {
    const { workbook, activeSheetId } = get();
    if (!workbook || !activeSheetId) return;
    const list = getSheet(workbook, activeSheetId)?.shapes ?? [];
    const zs = list.map((s) => s.z ?? 0);
    const lo = zs.length ? Math.min(...zs) : 0;
    const hi = zs.length ? Math.max(...zs) : 0;
    const z =
      dir === 'front' ? hi + 1 : dir === 'back' ? lo - 1 : dir === 'forward' ? hi + 1 : lo - 1;
    get().updateShape(id, { z });
  },

  selectActive() {
    const { workbook, activeSheetId } = get();
    if (!workbook || !activeSheetId) return;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return;
    const active = sheet.view?.activeCell;
    if (!active) return;
    const m = /^([A-Z]+)(\d+)$/.exec(active);
    if (!m) return;
    const col = colFromLabel(m[1]);
    const row = parseInt(m[2], 10);
    set({ selection: SINGLE(row, col) });
  },

  moveActive(dRow, dCol, extend = false) {
    const { selection, workbook, activeSheetId } = get();
    if (!workbook || !activeSheetId) return;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return;
    const baseRow = selection.row;
    const baseCol = selection.col;
    if (extend) {
      const endRow = Math.max(1, Math.min(sheetRowCount(sheet), selection.endRow + dRow));
      const endCol = Math.max(0, Math.min(sheetColCount(sheet), selection.endCol + dCol));
      set({ selection: { row: baseRow, col: baseCol, endRow, endCol } });
      return;
    }
    const row = Math.max(1, Math.min(sheetRowCount(sheet), baseRow + dRow));
    const col = Math.max(0, Math.min(sheetColCount(sheet), baseCol + dCol));
    set({ selection: SINGLE(row, col) });
  },

  commitEdit(text) {
    const { editing, activeSheetId, workbook } = get();
    if (!editing || !activeSheetId || !workbook) return;
    const { row, col } = editing;
    setCellRawInStore(get, row, col, text);
    set({ editing: null });
  },

  cancelEdit() {
    set({ editing: null });
  },

  startEdit(initialText?: string) {
    const { workbook, activeSheetId, selection, engine } = get();
    if (!workbook || !activeSheetId || !selection) return;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return;
    if (initialText != null) {
      const { row, col } = selection;
      set({ editing: { row, col, text: initialText, seeded: true } });
      return;
    }
    const { row, col } = selection;
    const cell = getCell(sheet, col, row);
    const text = cellText(cell, engine, activeSheetId, row, col, sheet.name);
    set({ editing: { row, col, text, seeded: false } });
  },

  setCellRaw(row, col, text) {
    setCellRawInStore(get, row, col, text);
  },

  setCellValue(row, col, value) {
    const { workbook, activeSheetId, applyCommand } = get();
    if (!workbook || !activeSheetId) return;
    const ref = cellToA1(col, row);
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return;
    applyCommand(new EditCellsCommand('Change cell', activeSheetId, { [ref]: { value } }));
  },

  clearSelection() {
    const { activeSheetId, workbook, applyCommand } = get();
    if (!activeSheetId || !workbook) return;
    const { row, col, endRow, endCol } = get().selection;
    // ClearRangeCommand, not SetRangeValuesCommand: an empty values matrix there
    // produces no edits, so Delete would silently do nothing.
    applyCommand(new ClearRangeCommand('Clear cells', activeSheetId, row, col, endRow, endCol));
  },

  copySelection() {
    const { workbook, activeSheetId, selection } = get();
    if (!workbook || !activeSheetId) return;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return;
    void writeSelectionToOsClipboard(get, sheet, activeSheetId);
    const sRow = Math.min(selection.row, selection.endRow);
    const eRow = Math.max(selection.row, selection.endRow);
    const sCol = Math.min(selection.col, selection.endCol);
    const eCol = Math.max(selection.col, selection.endCol);
    const rows = eRow - sRow + 1;
    const cols = eCol - sCol + 1;
    const cells: ClipboardCell[] = [];
    for (let r = sRow; r <= eRow; r++) {
      for (let c = sCol; c <= eCol; c++) {
        const cell = getCell(sheet, c, r);
        const snapshot: ClipboardCell = {};
        if (cell?.value !== undefined && cell.value !== null) snapshot.value = cell.value;
        if (cell?.formula !== undefined) snapshot.formula = cell.formula;
        if (cell?.style) snapshot.style = JSON.parse(JSON.stringify(cell.style));
        cells.push(snapshot);
      }
    }
    setClipboard({ rows, cols, cells: cells as ClipboardCell[], cut: false });
  },

  cutSelection() {
    get().copySelection();
    const { activeSheetId, workbook, applyCommand, selection } = get();
    if (!activeSheetId || !workbook) return;
    const sRow = Math.min(selection.row, selection.endRow);
    const eRow = Math.max(selection.row, selection.endRow);
    const sCol = Math.min(selection.col, selection.endCol);
    const eCol = Math.max(selection.col, selection.endCol);
    applyCommand(new ClearRangeCommand('Cut cells', activeSheetId, sRow, sCol, eRow, eCol));
    const clip = getClipboard();
    if (clip) setClipboard({ ...clip, cut: true });
  },

  paste: async () => {
    const { workbook, activeSheetId, applyCommand, selection } = get();
    const clip = getClipboard();
    // Excel's paste also accepts external text from the OS clipboard. Fall back
    // to that when the internal clipboard is empty so Copy/Paste round-trips
    // through the system clipboard like a native spreadsheet.
    if (!clip && workbook && activeSheetId) {
      const text = await readOsClipboardText();
      if (text) {
        const lines = text.split(/\r\n|\r|\n/);
        while (lines.length && lines[lines.length - 1] === '') lines.pop();
        if (lines.length === 0) return;
        const values = lines.map((line) => line.split('\t').map((t) => coerceTypedValue(t)));
        const eRow = selection.row + lines.length - 1;
        const eCol = selection.col + Math.max(0, (values[0]?.length ?? 1) - 1);
        applyCommand(new SetRangeValuesCommand('Paste', activeSheetId, selection.row, selection.col, eRow, eCol, values));
        return;
      }
    }
    if (!clip || !workbook || !activeSheetId) return;
    const sRow = Math.max(1, selection.row);
    const sCol = Math.max(0, selection.col);
    const values: CellPrimitive[][] = [];
    const formulas: string[][] = [];
    const styles: CellStyle[][] = [];
    for (let r = 0; r < clip.rows; r++) {
      const vRow: CellPrimitive[] = [];
      const fRow: string[] = [];
      const sRowArr: CellStyle[] = [];
      for (let c = 0; c < clip.cols; c++) {
        const cell = clip.cells[r * clip.cols + c];
        vRow.push(cell?.value ?? null);
        fRow.push(cell?.formula ?? '');
        sRowArr.push(cell?.style ?? {});
      }
      values.push(vRow);
      formulas.push(fRow);
      styles.push(sRowArr);
    }
    applyCommand(
      new SetRangeValuesCommand(
        'Paste',
        activeSheetId,
        sRow,
        sCol,
        sRow + clip.rows - 1,
        sCol + clip.cols - 1,
        values,
        { formulas, styles }
      )
    );
  },

  formatSelection(patch) {
    const { activeSheetId, workbook, applyCommand, selection } = get();
    if (!activeSheetId || !workbook) return;
    const { row, col, endRow, endCol } = selection;
    applyCommand(new FormatCellsCommand('Format cells', activeSheetId, row, col, endRow, endCol, patch));
  },

  setActiveSheet(id: string) {
    const { workbook, engine } = get();
    if (!workbook) return;
    const next = cloneWorkbook(workbook);
    next.activeSheetId = id;
    engine.attach(next, id);
    set({
      workbook: next,
      activeSheetId: id,
      revision: get().revision + 1
    });
  },

  findCells(query, opts = {}) {
    const { workbook, activeSheetId } = get();
    const results: Array<{ row: number; col: number; hit: string }> = [];
    if (!workbook || !activeSheetId || !query) return results;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return results;
    const haystack = (s: unknown): string => {
      if (typeof s === 'string') return s;
      if (typeof s === 'number' || typeof s === 'boolean') return String(s);
      return '';
    };
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!m) continue;
      const col = colFromLabel(m[1]);
      const row = parseInt(m[2], 10);
      const parts = [haystack(cell.value), cell.formula ? '=' + cell.formula : '', haystack(cell.note)];
      for (const part of parts) {
        if (!part) continue;
        let base = part;
        let q = query;
        if (!opts.matchCase) {
          base = base.toLowerCase();
          q = q.toLowerCase();
        }
        const hit = opts.wholeCell ? base === q : base.includes(q);
        if (hit) {
          results.push({ row, col, hit: part });
          break;
        }
      }
    }
    results.sort((a, b) => (a.row - b.row) || (a.col - b.col));
    return results;
  },

  replaceCells(query, replacement, opts = {}) {
    const { workbook, activeSheetId, applyCommand } = get();
    if (!workbook || !activeSheetId || !query) return 0;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return 0;
    const q = opts.matchCase ? query : query.toLowerCase();
    const matches = (s: string): boolean => {
      const base = opts.matchCase ? s : s.toLowerCase();
      return opts.wholeCell ? base === q : base.includes(q);
    };
    const sub = (s: string): string => {
      if (opts.wholeCell) return replacement;
      return opts.matchCase ? s.replaceAll(query, replacement) : replaceCaseInsensitive(s, query, replacement);
    };
    const doesMatch = (cell: Cell | undefined): boolean => {
      if (!cell) return false;
      if (typeof cell.value === 'string' && matches(cell.value)) return true;
      if (typeof cell.note === 'string' && matches(cell.note)) return true;
      return false;
    };
    const edits: Record<string, Cell | undefined> = {};
    let count = 0;
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      if (!doesMatch(cell)) continue;
      const next: Cell = JSON.parse(JSON.stringify(cell));
      let changed = false;
      if (typeof next.value === 'string' && matches(next.value)) {
        next.value = sub(next.value);
        changed = true;
      }
      if (typeof next.note === 'string' && matches(next.note)) {
        next.note = sub(next.note);
        changed = true;
      }
      if (changed) {
        edits[ref] = next;
        count++;
      }
    }
    if (count) applyCommand(new EditCellsCommand('Replace', activeSheetId, edits));
    return count;
  }
}));

/* ------------------------- helpers & wiring ------------------------- */

function sheetRowCount(sheet: Worksheet): number {
  return 1048576;
}
function sheetColCount(sheet: Worksheet): number {
  return 16384;
}
function colFromLabel(label: string): number {
  let col = -1;
  for (const ch of label.toUpperCase()) col = (col + 1) * 26 + (ch.charCodeAt(0) - 65);
  return col;
}

function cellText(
  cell: { value?: CellPrimitive; formula?: string } | undefined,
  engine: WorkbookFormulaEngine,
  sheetId: string,
  row: number,
  col: number,
  sheetName: string
): string {
  if (cell?.formula !== undefined) return '=' + cell.formula;
  if (cell?.value !== undefined && cell.value !== null) return String(cell.value);
  return '';
}

function setCellRawInStore(get: () => WorkbookState, row: number, col: number, text: string): void {
  const { workbook, activeSheetId, applyCommand } = get();
  if (!workbook || !activeSheetId) return;
  const ref = cellToA1(col, row);
  const trimmed = text.trim();
  const sheet = getSheet(workbook, activeSheetId);
  const prev = sheet?.cells[ref];
  // Committing a typed value replaces only the content. Excel keeps the cell's
  // own formatting, notes and links when you type over it, and the ribbon
  // toggles (Bold…) read straight from that style — so preserve them here.
  const keep: Pick<Cell, 'style' | 'note' | 'hyperlink'> = {};
  if (prev) {
    if (prev.style) keep.style = prev.style;
    if (prev.note) keep.note = prev.note;
    if (prev.hyperlink) keep.hyperlink = prev.hyperlink;
  }
  if (trimmed === '') {
    // Content cleared but formatting retained, like the Delete key.
    const edit: Cell | undefined = keep.style || keep.note || keep.hyperlink ? (keep as Cell) : undefined;
    applyCommand(new EditCellsCommand('Clear cell', activeSheetId, { [ref]: edit }));
    return;
  }
  if (trimmed.startsWith('=')) {
    applyCommand(new EditCellsCommand('Change cell', activeSheetId, { [ref]: { ...keep, formula: trimmed.slice(1) } as Cell }));
    return;
  }
  const value = coerceTypedValue(trimmed);
  applyCommand(new EditCellsCommand('Change cell', activeSheetId, { [ref]: { ...keep, value } as Cell }));
}

function coerceTypedValue(text: string): CellPrimitive {
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(text)) {
    const n = parseFloat(text);
    if (!Number.isNaN(n)) return n;
  }
  if (/^(true|false)$/i.test(text)) return text.toLowerCase() === 'true';
  return text;
}

function replaceCaseInsensitive(input: string, find: string, replacement: string): string {
  const lower = input.toLowerCase();
  const q = find.toLowerCase();
  if (!q) return input;
  let out = '';
  let i = 0;
  while (i < input.length) {
    const idx = lower.indexOf(q, i);
    if (idx === -1) {
      out += input.slice(i);
      break;
    }
    out += input.slice(i, idx) + replacement;
    i = idx + find.length;
  }
  return out;
}

export type { Worksheet }; // re-export convenience

/** Browser fallback: hand a file to the user's downloads folder. */
export function downloadBlob(content: string, fileName: string, type: string): void {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  // Revoking immediately can cancel the download in some webviews.
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}