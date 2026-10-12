import { describe, it, expect } from 'vitest';
import { createWorkbook, getSheet, currentRegion } from '@/engine/model/workbook';
import {
  EditCellsCommand,
SetRangeValuesCommand,
  ClearRangeCommand,
  FormatCellsCommand,
  InsertRowsCommand,
  DeleteRowsCommand,
  DeleteColumnsCommand,
  RemoveDuplicatesCommand,
  SetNamedRangeCommand,
  AddChartCommand,
  RemoveChartCommand,
  SetNoteCommand,
  SetSheetProtectedCommand,
  SetZoomCommand,
  SetPageSetupCommand,
  ClearPageSetupCommand,
  SetPrintAreaCommand,
  AddShapeCommand,
  UpdateShapeCommand,
  RemoveShapeCommand,
  AddSheetCommand,
  SetThemeCommand,
  ClearThemeCommand
} from '@/engine/commands/commands';
import { HistoryManager } from '@/engine/commands/history';
import { WorkbookFormulaEngine } from '@/engine/formula/engine';
import { getTheme, accent, THEMES, THEME_COLOR_SETS } from '@/engine/model/theme';
import { serializeWorkbook, deserializeWorkbook } from '@/engine/io/serialize';

describe('workbook model', () => {
  it('creates a workbook with a default sheet', () => {
    const wb = createWorkbook({ title: 'Test' });
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe('Sheet1');
    expect(wb.metadata.title).toBe('Test');
    expect(wb.activeSheetId).toBe(wb.sheets[0].id);
  });

  it('stores a cell value by reference', () => {
    const wb = createWorkbook();
    const sheet = getSheet(wb, wb.activeSheetId)!;
    const cmd = new EditCellsCommand('set', wb.activeSheetId, { A1: { value: 42 } });
    const next = cmd.apply(wb);
    const s2 = getSheet(next, next.activeSheetId)!;
    expect(s2.cells.A1.value).toBe(42);
  });

  it('supports undo/redo through the history manager', () => {
    const wb = createWorkbook();
    const history = new HistoryManager();
    let current = wb;
    const cmd = new EditCellsCommand('set', wb.activeSheetId, { A1: { value: 'hello' } });
    current = cmd.apply(current);
    history.execute(cmd);
    expect((getSheet(current, current.activeSheetId)?.cells.A1 as any)?.value).toBe('hello');

    const undo = history.popUndo()!;
    current = undo.inverse().apply(current);
    history.pushRedo(undo);
    expect(getSheet(current, current.activeSheetId)?.cells.A1).toBeUndefined();

    const redo = history.popRedo()!;
    current = redo.apply(current);
    history.execute(redo);
    expect((getSheet(current, current.activeSheetId)?.cells.A1 as any)?.value).toBe('hello');
  });

  it('immutably returns a fresh workbook on change', () => {
    const wb = createWorkbook();
    const cmd = new EditCellsCommand('set', wb.activeSheetId, { A1: { value: 1 } });
    const next = cmd.apply(wb);
    expect(next).not.toBe(wb);
    expect(getSheet(wb, wb.activeSheetId)?.cells.A1).toBeUndefined();
  });
});

describe('range commands', () => {
  it('writes a rectangular range of values', () => {
    const wb = createWorkbook();
    const cmd = new SetRangeValuesCommand('fill', wb.activeSheetId, 1, 0, 2, 2, [
      [1, 2, 3],
      [4, 5, 6],
    ]);
    const next = cmd.apply(wb);
    const sheet = getSheet(next, next.activeSheetId)!;
    expect(sheet.cells.A1.value).toBe(1);
    expect(sheet.cells.B2.value).toBe(5);
    expect(sheet.cells.C2.value).toBe(6);
  });

  it('clears cell contents and keeps styles', () => {
    const wb = createWorkbook();
    let next = new SetRangeValuesCommand('fill', wb.activeSheetId, 1, 0, 1, 1, [
      ['Widget', '10'],
    ]).apply(wb);
    next = new FormatCellsCommand('bold', wb.activeSheetId, 1, 0, 1, 1, { bold: true }).apply(next);
    let sheet = getSheet(next, next.activeSheetId)!;
    expect(sheet.cells.A1.value).toBe('Widget');

    // Delete clears values but keeps formatting, and undo brings them back.
    const clear = new ClearRangeCommand('Clear cells', wb.activeSheetId, 1, 0, 1, 1);
    next = clear.apply(next);
    sheet = getSheet(next, next.activeSheetId)!;
    expect(sheet.cells.A1.value).toBeUndefined();
    expect(sheet.cells.B1.value).toBeUndefined();
    expect(sheet.cells.A1.style?.bold).toBe(true);

    sheet = getSheet(clear.inverse().apply(next), wb.activeSheetId)!;
    expect(sheet.cells.A1.value).toBe('Widget');
    expect(sheet.cells.B1.value).toBe('10');
  });

  it('applies style patches and clears stale keys', () => {
    const wb = createWorkbook();
    const fmt = new FormatCellsCommand('bold', wb.activeSheetId, 1, 0, 1, 0, { bold: true, italic: true });
    let next = fmt.apply(wb);
    let sheet = getSheet(next, next.activeSheetId)!;
    expect(sheet.cells.A1.style).toMatchObject({ bold: true, italic: true });

    const fmt2 = new FormatCellsCommand('un-italic', wb.activeSheetId, 1, 0, 1, 0, { italic: undefined, bold: true } as never);
    next = fmt2.apply(next);
    sheet = getSheet(next, next.activeSheetId)!;
    expect(sheet.cells.A1.style?.italic).toBeUndefined();
    expect(sheet.cells.A1.style?.bold).toBe(true);
  });

  it('shifts rows on insert', () => {
    const wb = createWorkbook();
    let next = new EditCellsCommand('set', wb.activeSheetId, { A1: { value: 'top' }, A2: { value: 'bottom' } }).apply(wb);
    next = new InsertRowsCommand('insert', next.activeSheetId, 0, 1).apply(next);
    const sheet = getSheet(next, next.activeSheetId)!;
    expect(sheet.cells.A1).toBeUndefined();
    expect((sheet.cells.A2 as any)?.value).toBe('top');
    expect((sheet.cells.A3 as any)?.value).toBe('bottom');
  });

  it('shifts rows up on delete', () => {
    const wb = createWorkbook();
    let next = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { value: 'r1' },
      A2: { value: 'r2' },
      A3: { value: 'r3' }
    }).apply(wb);
    next = new DeleteRowsCommand('delete', next.activeSheetId, 1, 1).apply(next);
    const sheet = getSheet(next, next.activeSheetId)!;
    expect((sheet.cells.A1 as any)?.value).toBe('r1');
    expect((sheet.cells.A2 as any)?.value).toBe('r3');
    expect(sheet.cells.A3).toBeUndefined();
  });

  it('shifts columns left on delete', () => {
    const wb = createWorkbook();
    let next = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { value: 'c1' },
      B1: { value: 'c2' },
      C1: { value: 'c3' }
    }).apply(wb);
    next = new DeleteColumnsCommand('delete', next.activeSheetId, 1, 1).apply(next);
    const sheet = getSheet(next, next.activeSheetId)!;
    expect((sheet.cells.A1 as any)?.value).toBe('c1');
    expect((sheet.cells.B1 as any)?.value).toBe('c3');
    expect(sheet.cells.C1).toBeUndefined();
  });

  it('restores deleted rows on undo', () => {
    const wb = createWorkbook();
    let next = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { value: 'r1' },
      A2: { value: 'r2' },
      A3: { value: 'r3' }
    }).apply(wb);
    const del = new DeleteRowsCommand('delete', next.activeSheetId, 1, 1);
    next = del.apply(next);
    next = del.inverse().apply(next);
    const sheet = getSheet(next, next.activeSheetId)!;
    expect((sheet.cells.A2 as any)?.value).toBe('r2');
    expect((sheet.cells.A3 as any)?.value).toBe('r3');
  });

  it('removes duplicate rows, keeps the first occurrence and clears the tail', () => {
    const wb = createWorkbook();
    let next = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { value: 'Ann' },
      B1: { value: 1 },
      A2: { value: 'Bob' },
      B2: { value: 2 },
      A3: { value: 'Ann' },
      B3: { value: 1 },
      A4: { value: 'Cid' },
      B4: { value: 3 },
      A5: { value: 'Bob' },
      B5: { value: 2 },
    }).apply(wb);

    const cmd = new RemoveDuplicatesCommand('dedupe', next.activeSheetId, 1, 0, 5, 1, false);
    next = cmd.apply(next);
    const sheet = getSheet(next, next.activeSheetId)!;

    expect((sheet.cells.A1 as any)?.value).toBe('Ann');
    expect((sheet.cells.A2 as any)?.value).toBe('Bob');
    expect((sheet.cells.A3 as any)?.value).toBe('Cid');
    // Leftover rows must be cleared, not left behind.
    expect(sheet.cells.A4).toBeUndefined();
    expect(sheet.cells.B4).toBeUndefined();
    expect(cmd.removedCount).toBe(2);

    const restored = cmd.inverse().apply(next);
    const s2 = getSheet(restored, restored.activeSheetId)!;
    expect((s2.cells.A3 as any)?.value).toBe('Ann');
    expect((s2.cells.A5 as any)?.value).toBe('Bob');
  });

  it('keeps the header row when told the data has headers', () => {
    const wb = createWorkbook();
    let next = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { value: 'Name' },
      A2: { value: 'Ann' },
      A3: { value: 'Ann' },
    }).apply(wb);

    next = new RemoveDuplicatesCommand('dedupe', next.activeSheetId, 1, 0, 3, 0, true).apply(next);
    const sheet = getSheet(next, next.activeSheetId)!;
    expect((sheet.cells.A1 as any)?.value).toBe('Name');
    expect((sheet.cells.A2 as any)?.value).toBe('Ann');
    expect(sheet.cells.A3).toBeUndefined();
  });
});

describe('formula engine', () => {
  it('computes arithmetic and references', () => {
    const wb = createWorkbook();
    const engine = new WorkbookFormulaEngine();
    engine.attach(wb, wb.activeSheetId);
    const edits = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { value: 10 },
      A2: { value: 5 },
      B1: { formula: 'A1+A2' },
      B2: { formula: 'SUM(A1:A2)' },
    });
    const next = edits.apply(wb);
    engine.attach(next, next.activeSheetId);
    expect(engine.compute(next.activeSheetId, 1, 1)).toBe(15);
    expect(engine.compute(next.activeSheetId, 2, 1)).toBe(15);
  });

  it('evaluates IF with lazy semantics', () => {
    const wb = createWorkbook();
    const engine = new WorkbookFormulaEngine();
    engine.attach(wb, wb.activeSheetId);
    const edits = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { formula: 'IF(2>1,"yes","no")' },
      A2: { formula: 'IFERROR(1/0,"fallback")' },
    });
    const next = edits.apply(wb);
    engine.attach(next, next.activeSheetId);
    expect(engine.compute(next.activeSheetId, 1, 0)).toBe('yes');
    expect(engine.compute(next.activeSheetId, 2, 0)).toBe('fallback');
  });

  it('surfaces circular references', () => {
    const wb = createWorkbook();
    const engine = new WorkbookFormulaEngine();
    engine.attach(wb, wb.activeSheetId);
    const edits = new EditCellsCommand('set', wb.activeSheetId, {
      A1: { formula: 'B1' },
      B1: { formula: 'A1' },
    });
    const next = edits.apply(wb);
    engine.attach(next, next.activeSheetId);
    expect(engine.compute(next.activeSheetId, 1, 1)).toBe('#CIRC!');
  });

  it('resolves a defined name in a formula regardless of case', () => {    const engine = new WorkbookFormulaEngine();
    const wb0 = createWorkbook();
    const withData = new EditCellsCommand('set', wb0.activeSheetId, {
      A1: { value: 10 },
      A2: { value: 20 },
      A3: { value: 30 },
      A4: { value: 40 }
    }).apply(wb0);
    // SetNamedRangeCommand stores names under an UPPERCASE key.
    const named = new SetNamedRangeCommand('Define name', 'SalesTotal', 'A1:A4', 'Sheet1').apply(withData);
    engine.attach(named, named.activeSheetId);

    const withFormula = new EditCellsCommand('set', named.activeSheetId, {
      C1: { formula: 'SUM(SalesTotal)' },
      C2: { formula: 'AVERAGE(SalesTotal)' },
      C3: { formula: 'MAX(SalesTotal)-MIN(SalesTotal)' }
    }).apply(named);
    engine.attach(withFormula, withFormula.activeSheetId);

    // Columns are 0-based, so C1 is (row 1, col 2).
    expect(engine.compute(withFormula.activeSheetId, 1, 2)).toBe(100);
    expect(engine.compute(withFormula.activeSheetId, 2, 2)).toBe(25);
    expect(engine.compute(withFormula.activeSheetId, 3, 2)).toBe(30);
  });
});

describe('chart region and removal', () => {
  const withData = () => {
    const wb0 = createWorkbook();
    return new EditCellsCommand('set', wb0.activeSheetId, {
      A1: { value: 10 },
      A2: { value: 20 },
      A3: { value: 30 },
      A4: { value: 40 },
      B1: { value: 'x' }
    }).apply(wb0);
  };

  it('expands a single cell to the surrounding data region', () => {
    const wb = withData();
    const sheet = getSheet(wb, wb.activeSheetId)!;
    // Row 3 only has data in column A, so the region is the A1:A4 block.
    expect(currentRegion(sheet, 0, 2)).toEqual({ sCol: 0, sRow: 1, eCol: 0, eRow: 4 });
    // Row 1 has A and B filled, so the block widens there.
    expect(currentRegion(sheet, 1, 1)).toEqual({ sCol: 0, sRow: 1, eCol: 1, eRow: 1 });
    // A cell outside the block stays a single cell.
    expect(currentRegion(sheet, 6, 6)).toEqual({ sCol: 6, sRow: 6, eCol: 6, eRow: 6 });
  });

  it('treats a zero value as data, like Excel', () => {
    const wb0 = createWorkbook();
    const wb = new EditCellsCommand('set', wb0.activeSheetId, {
      A1: { value: 1 },
      A2: { value: 0 },
      A3: { value: 3 }
    }).apply(wb0);
    const sheet = getSheet(wb, wb.activeSheetId)!;
    expect(currentRegion(sheet, 0, 1)).toEqual({ sCol: 0, sRow: 1, eCol: 0, eRow: 3 });
  });

  it('does not bleed across a gap in the data', () => {
    const wb0 = createWorkbook();
    const wb = new EditCellsCommand('set', wb0.activeSheetId, {
      A1: { value: 1 },
      A2: { value: 2 },
      A4: { value: 3 }
    }).apply(wb0);
    const sheet = getSheet(wb, wb.activeSheetId)!;
    // A3 is empty, so the region stops at A2.
    expect(currentRegion(sheet, 0, 1)).toEqual({ sCol: 0, sRow: 1, eCol: 0, eRow: 2 });
  });

  it('adds and removes a chart, and undo restores it', () => {
    const wb0 = createWorkbook();
    const add = new AddChartCommand('Insert chart', wb0.activeSheetId, {
      type: 'column',
      dataRange: 'A1:A4',
      x: 2,
      y: 1,
      width: 320,
      height: 220,
      title: 'Chart'
    } as never);
    const added = add.apply(wb0);
    const sheet = getSheet(added, added.activeSheetId)!;
    expect(sheet.charts).toHaveLength(1);
    const id = sheet.charts[0].id;

    const remove = new RemoveChartCommand('Delete chart', added.activeSheetId, id);
    const removed = remove.apply(added);
    expect(getSheet(removed, removed.activeSheetId)!.charts).toHaveLength(0);

    // Undo path: the command's inverse puts the chart back.
    const restored = remove.inverse().apply(removed);
    expect(getSheet(restored, restored.activeSheetId)!.charts).toHaveLength(1);
  });

  it('adds a note, and empty text removes it entirely', () => {
    const wb0 = createWorkbook();
    const sid = wb0.activeSheetId;
    const added = new SetNoteCommand(sid, 'B2', 'Check this total').apply(wb0);
    const sheet = getSheet(added, sid)!;
    expect(sheet.notes).toHaveLength(1);
    expect(sheet.notes[0]).toMatchObject({ ref: 'B2', text: 'Check this total' });

    // Whitespace-only text must not leave an empty marker behind.
    const cleared = new SetNoteCommand(sid, 'B2', '   ').apply(added);
    expect(getSheet(cleared, sid)!.notes).toHaveLength(0);
  });

  it('protects and unprotects a sheet, and the inverse flips it back', () => {
    const wb0 = createWorkbook();
    const sid = wb0.activeSheetId;
    expect(getSheet(wb0, sid)!.protected).toBeFalsy();

    const protect = new SetSheetProtectedCommand(sid, true);
    expect(protect.isProtecting()).toBe(true);
    const locked = protect.apply(wb0);
    expect(getSheet(locked, sid)!.protected).toBe(true);

    const unlocked = protect.inverse().apply(locked);
    expect(getSheet(unlocked, sid)!.protected).toBe(false);
  });

  it('view commands still apply to a protected sheet', () => {
    // The store allowlists these; this pins the commands that must exist and
    // not throw when a sheet is protected.
    const wb0 = createWorkbook();
    const sid = wb0.activeSheetId;
    const locked = new SetSheetProtectedCommand(sid, true).apply(wb0);
    const zoomed = new SetZoomCommand(sid, 150).apply(locked);
    expect(getSheet(zoomed, sid)!.view?.zoom).toBe(150);
    // The sheet is still protected after a view change.
    expect(getSheet(zoomed, sid)!.protected).toBe(true);
  });
});

describe('page setup', () => {
  it('merges a partial patch and preserves other fields', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const oriented = new SetPageSetupCommand(sid, { orientation: 'landscape' }).apply(wb);
    const scaled = new SetPageSetupCommand(sid, { scale: 150 }).apply(oriented);
    expect(getSheet(scaled, sid)!.pageSetup).toEqual({ orientation: 'landscape', scale: 150 });
  });

  it('undoes a patch back to the previous setup', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const cmd = new SetPageSetupCommand(sid, { paperSize: 'a4', orientation: 'landscape' });
    const applied = cmd.apply(wb);
    const undone = cmd.inverse().apply(applied);
    expect(getSheet(undone, sid)!.pageSetup ?? {}).toEqual({});
  });

  it('restores page setup through the history manager', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const h = new HistoryManager();
    const cmd = new SetPageSetupCommand(sid, { paperSize: 'legal' });
    let doc = cmd.apply(wb);
    h.execute(cmd);
    expect(getSheet(doc, sid)!.pageSetup!.paperSize).toBe('legal');
    // The caller pops and applies the inverse.
    doc = h.popUndo()!.inverse().apply(doc);
    expect(getSheet(doc, sid)!.pageSetup ?? {}).toEqual({});
    h.pushRedo(cmd);
    doc = cmd.apply(doc);
    expect(getSheet(doc, sid)!.pageSetup!.paperSize).toBe('legal');
  });

  it('clears page setup entirely', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const set = new SetPageSetupCommand(sid, { paperSize: 'a4', scale: 90 });
    const cleared = new ClearPageSetupCommand(sid).apply(set.apply(wb));
    expect(getSheet(cleared, sid)!.pageSetup ?? {}).toEqual({});
  });

  it('clears only the print area, keeping other settings', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const withArea = new SetPageSetupCommand(sid, { orientation: 'landscape', printArea: 'A1:C5' }).apply(wb);
    const cleared = new SetPrintAreaCommand(sid, undefined).apply(withArea);
    const setup = getSheet(cleared, sid)!.pageSetup!;
    expect(setup.printArea).toBeUndefined();
    expect(setup.orientation).toBe('landscape');
  });

  it('stores manual page breaks as sorted unique lists', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const cmd = new SetPageSetupCommand(sid, { pageBreaks: { rows: [30, 10, 10], cols: [4] } });
    expect(getSheet(cmd.apply(wb), sid)!.pageSetup!.pageBreaks).toEqual({ rows: [30, 10, 10], cols: [4] });
  });
});

describe('shapes', () => {
  const shape = { id: 's1', kind: 'rect' as const, x: 10, y: 20, width: 120, height: 60 };

  it('adds a shape on top of existing shapes', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const one = new AddShapeCommand(sid, shape).apply(wb);
    const two = new AddShapeCommand(sid, { ...shape, id: 's2' }).apply(one);
    const list = getSheet(two, sid)!.shapes!;
    expect(list).toHaveLength(2);
    expect(list[1].z!).toBeGreaterThan(list[0].z!);
  });

  it('removes a shape and restores it on undo', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const added = new AddShapeCommand(sid, shape).apply(wb);
    const remove = new RemoveShapeCommand(sid, 's1');
    const removed = remove.apply(added);
    expect(getSheet(removed, sid)!.shapes).toEqual([]);
    const restored = remove.inverse().apply(removed);
    expect(getSheet(restored, sid)!.shapes).toHaveLength(1);
  });

  it('patches geometry and reverts it on undo', () => {
    const wb = createWorkbook();
    const sid = wb.activeSheetId;
    const added = new AddShapeCommand(sid, shape).apply(wb);
    const cmd = new UpdateShapeCommand(sid, 's1', { x: 55, fill: '#ff0000' });
    const moved = cmd.apply(added);
    const s = getSheet(moved, sid)!.shapes![0];
    expect(s.x).toBe(55);
    expect(s.fill).toBe('#ff0000');
    expect(s.width).toBe(120);
    const reverted = cmd.inverse().apply(moved);
    const s2 = getSheet(reverted, sid)!.shapes![0];
    expect(s2.x).toBe(10);
    expect(s2.fill).toBeUndefined();
  });

  it('leaves other sheets untouched', () => {
    const wb = createWorkbook();
    const firstId = wb.activeSheetId;
    const added = new AddSheetCommand('Second').apply(wb);
    const other = added.sheets[1].id;
    const withShape = new AddShapeCommand(firstId, shape).apply(added);
    expect(getSheet(withShape, other)!.shapes ?? []).toEqual([]);
  });
});
describe('theme', () => {
  it('falls back to the default theme when a workbook has none', () => {
    const wb = createWorkbook();
    expect(getTheme(wb.theme).name).toBe('Office');
  });

  it('applies a theme patch and keeps unspecified fields', () => {
    const wb = createWorkbook();
    const next = THEMES.find((t) => t.id !== 'office')!;
    const themed = new SetThemeCommand(next).apply(wb);
    expect(themed.theme?.id).toBe(next.id);
    expect(themed.theme?.accents).toEqual(next.accents);
    expect(themed.theme?.name).toBe(next.name);
  });

  it('resolves accent references against the active theme', () => {
    const wb = createWorkbook();
    const set = THEME_COLOR_SETS[1];
    const themed = new SetThemeCommand({ accents: set.accents }).apply(wb);
    expect(accent(themed.theme, 0)).toBe(set.accents[0]);
    expect(accent(themed.theme, 5)).toBe(set.accents[5]);
    // Out-of-range indices wrap instead of returning undefined.
    expect(accent(themed.theme, 99)).toBe(set.accents[99 % set.accents.length]);
  });

  it('round-trips a theme change through undo and redo', () => {
    const history = new HistoryManager();
    let current = createWorkbook();
    const next = THEMES[2];
    const first = new SetThemeCommand(next);
    current = first.apply(current);
    history.execute(first);
    const second = new SetThemeCommand({ accents: THEME_COLOR_SETS[3].accents });
    current = second.apply(current);
    history.execute(second);
    expect(current.theme?.accents).toEqual(THEME_COLOR_SETS[3].accents);

    const undoSecond = history.popUndo()!;
    current = undoSecond.inverse().apply(current);
    history.pushRedo(undoSecond);
    expect(current.theme?.accents).toEqual(next.accents);

    const undoFirst = history.popUndo()!;
    current = undoFirst.inverse().apply(current);
    history.pushRedo(undoFirst);
    expect(current.theme).toBeUndefined();

const redoFirst = history.popRedo()!;
    current = redoFirst.apply(current);
    history.execute(redoFirst);
    expect(current.theme?.id).toBe(next.id);
  });

  it('restores a theme when clearing it', () => {
    const wb = createWorkbook();
    const themed = new SetThemeCommand(THEMES[1]).apply(wb);
    const clear = new ClearThemeCommand();
    const cleared = clear.apply(themed);
    expect(cleared.theme).toBeUndefined();
    const restored = clear.inverse().apply(cleared);
    expect(restored.theme?.id).toBe(THEMES[1].id);
  });

it('survives serialize/deserialize with a theme', () => {
    const themed = new SetThemeCommand(THEMES[3]).apply(createWorkbook());
    const round = JSON.parse(JSON.stringify(themed)) as typeof themed;
    expect(getTheme(round.theme).id).toBe(THEMES[3].id);
  });

  it('round-trips a theme through the native file format', () => {
    const themed = new SetThemeCommand(THEMES[4]).apply(createWorkbook());
    const reopened = deserializeWorkbook(serializeWorkbook(themed));
    expect(getTheme(reopened.theme).id).toBe(THEMES[4].id);
    expect(reopened.theme?.accents).toEqual(THEMES[4].accents);
  });

  it('keeps theme-referencing cell styles across a save/load cycle', () => {
    const wb = new SetThemeCommand(THEMES[2]).apply(createWorkbook());
    const sheetId = wb.activeSheetId;
    const filled = new FormatCellsCommand('fill', sheetId, 1, 1, 1, 1, {
      backgroundTheme: 3,
      backgroundColor: undefined
    }).apply(wb);
    const reopened = deserializeWorkbook(serializeWorkbook(filled));
    const style = getSheet(reopened, reopened.activeSheetId)?.cells.B1?.style;
    expect(style?.backgroundTheme).toBe(3);
    expect(style?.backgroundColor).toBeUndefined();
    expect(accent(reopened.theme, style!.backgroundTheme!)).toBe(THEMES[2].accents[3]);
  });
});

