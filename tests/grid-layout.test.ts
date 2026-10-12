import { describe, it, expect } from 'vitest';
import { SheetGeometry, clampZoom, MIN_ZOOM, MAX_ZOOM, HEADER_SIZE } from '@/grid/layout';
import { createWorkbook } from '@/engine/model/workbook';
import { EditCellsCommand, SetZoomCommand, AddSheetCommand } from '@/engine/commands/commands';

function sheetOf() {
  const wb = createWorkbook();
  return wb.sheets[0];
}

describe('grid geometry zoom', () => {
  it('clamps zoom to the Excel range', () => {
    expect(clampZoom(100)).toBe(100);
    expect(clampZoom(0)).toBe(MIN_ZOOM);
    expect(clampZoom(5000)).toBe(MAX_ZOOM);
    expect(clampZoom(Number.NaN)).toBe(100);
  });

  it('scales pixel offsets, widths and header band', () => {
    const sheet = sheetOf();
    const g100 = new SheetGeometry(sheet, 100);
    const g200 = new SheetGeometry(sheet, 200);
    expect(g200.scale).toBe(2);
    expect(g200.colOffset(3)).toBe(g100.colOffset(3) * 2);
    expect(g200.rowOffset(5)).toBe(g100.rowOffset(5) * 2);
    expect(g200.colWidth(0)).toBe(g100.colWidth(0) * 2);
    expect(g200.rowHeight(1)).toBe(g100.rowHeight(1) * 2);
    expect(g200.headerSize).toBe(HEADER_SIZE * 2);
  });

  it('keeps visible indices zoom-independent', () => {
    const sheet = sheetOf();
    const g100 = new SheetGeometry(sheet, 100);
    const g200 = new SheetGeometry(sheet, 200);
    // Zooming in enlarges cells, so fewer columns fit in the same viewport.
    const c100 = g100.visibleCols(0, 400);
    const c200 = g200.visibleCols(0, 400);
    expect(c200.last).toBeLessThan(c100.last);
    // Scrolling to the same *logical* column resolves identically at any zoom.
    expect(g100.visibleCols(g100.colOffset(6), 300).first).toBe(6);
    expect(g200.visibleCols(g200.colOffset(6), 300).first).toBe(6);
  });

  it('converts rendered pixels back to model units for resizing', () => {
    const sheet = sheetOf();
    const g200 = new SheetGeometry(sheet, 200);
    expect(g200.toModelWidth(g200.colWidth(0))).toBeCloseTo(g200.colWidth(0) / 2);
    expect(g200.toModelHeight(g200.rowHeight(1))).toBeCloseTo(g200.rowHeight(1) / 2);
  });

  it('stores zoom on the sheet and restores it per sheet', () => {
    const wb0 = createWorkbook();
    const withData = new EditCellsCommand('set', wb0.activeSheetId, { A1: { value: 1 } }).apply(wb0);
    const zoomed = new SetZoomCommand(withData.activeSheetId, 150).apply(withData);
    expect(zoomed.sheets[0].view?.zoom).toBe(150);
    const second = new AddSheetCommand('Add sheet').apply(zoomed);
    const newSheet = second.sheets[second.sheets.length - 1];
    expect(newSheet.view?.zoom ?? 100).toBe(100);
    // The first sheet keeps its own zoom.
    expect(second.sheets[0].view?.zoom).toBe(150);
  });
});
