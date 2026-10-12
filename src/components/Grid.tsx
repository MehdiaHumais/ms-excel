/**
 * Spreadsheet grid.
 *
 * Virtualized rendering: only cells inside the viewport (+ a small overscan)
 * are mounted. Header row/column are synchronized via one scroll container.
 * Selection, editing, resize and freeze panes are handled here.
 */

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useWorkbookStore, Selection, SINGLE } from '@/store/useWorkbookStore';
import { useUiStore } from '@/store/useUiStore';
import { useSettingsStore } from '@/store/useSettingsStore';
import { SheetGeometry, HEADER_SIZE, MIN_COL_WIDTH, MAX_COL_WIDTH, isColHidden, isRowHidden } from '@/grid/layout';
import { cellToA1, colToLabel, parseRange, ParsedRange } from '@/engine/model/address';
import { getSheet, getCell } from '@/engine/model/workbook';
import { Worksheet, Cell, ChartDef, TableDef, ShapeDef } from '@/engine/model/types';
import { getTheme, accent } from '@/engine/model/theme';
import { displayCell } from '@/store/display';
import { showContextMenu, closeContextMenu, ContextMenuItem } from '@/components/contextMenu';
import {
  InsertRowsCommand,
  InsertColumnsCommand,
  DeleteRowsCommand,
  DeleteColumnsCommand,
  ClearRangeCommand,
  UpdateChartCommand,
  RemoveChartCommand
} from '@/engine/commands/commands';

const OVERSCAN = 8;
const RESIZE_PROXIMITY = 10;
const MAX_RENDER_COLS = 600;
const MAX_RENDER_ROWS = 3000;

/** Parse a print-area A1 or A1:B2 string into inclusive 1-based/0-based bounds. */
function rangeToA1Box(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const first = ref.split(',')[0].trim();
  const m = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(first);
  if (!m) return null;
  const toCol = (s: string) => {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  };
  const r1 = Number(m[2]);
  const c1 = toCol(m[1]);
  const r2 = m[4] ? Number(m[4]) : r1;
  const c2 = m[3] ? toCol(m[3]) : c1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

interface GridProps {
  height?: number;
}

export function Grid({ height }: GridProps) {
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const selectedChartId = useWorkbookStore((s) => s.selectedChartId);
const selectedShapeId = useWorkbookStore((s) => s.selectedShapeId);
  const showNotes = useWorkbookStore((s) => s.showNotes);
  const editing = useWorkbookStore((s) => s.editing);
  const revision = useWorkbookStore((s) => s.revision);
  const engine = useWorkbookStore((s) => s.engine);
  const setSelection = useWorkbookStore((s) => s.setSelection);
  const commitEdit = useWorkbookStore((s) => s.commitEdit);
  const cancelEdit = useWorkbookStore((s) => s.cancelEdit);
  const startEdit = useWorkbookStore((s) => s.startEdit);
  const setCellRaw = useWorkbookStore((s) => s.setCellRaw);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const openDialog = useUiStore((s) => s.openDialog);
  const showFormulas = useSettingsStore((s) => s.showFormulas);
  const showGridlines = useSettingsStore((s) => s.showGridlines);
const showRuler = useSettingsStore((s) => s.showRuler);
const pageBreakPreview = useUiStore((s) => s.pageBreakPreview);
const split = useUiStore((s) => s.split);
const splitRequest = useUiStore((s) => s.splitRequest);

  const sheet = useMemo<Worksheet | null>(() => {
    if (!workbook || !activeSheetId) return null;
    return getSheet(workbook, activeSheetId) ?? null;
  }, [workbook, activeSheetId]);

  const geometry = useMemo(
  () => (sheet ? new SheetGeometry(sheet, sheet.view?.zoom ?? 100) : new SheetGeometry(emptySheet())),
  [sheet]
);

  /**
   * Merged ranges indexed by every covered cell. Cells that are not the
   * top-left anchor of their merge are skipped so the anchor can span the
   * whole merged area.
   */
  const mergeMap = useMemo(() => {
    const map = new Map<string, ParsedRange>();
    if (!sheet) return map;
    for (const range of sheet.merges) {
      const parsed = parseRange(range);
      if (!parsed) continue;
      const area = (parsed.eRow - parsed.sRow + 1) * (parsed.eCol - parsed.sCol + 1);
      if (area > 20_000) continue;
      for (let r = parsed.sRow; r <= parsed.eRow; r++) {
        for (let c = parsed.sCol; c <= parsed.eCol; c++) {
          map.set(`${c}_${r}`, parsed);
        }
      }
    }
    return map;
  }, [sheet]);

  const rootRef = useRef<HTMLDivElement>(null);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportW, setViewportW] = useState(800);
  const [viewportH, setViewportH] = useState(600);

  // Panes own the scroll containers, so the grid root is what we measure.
  useEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    const measure = () => {
      setViewportW(Math.max(120, el.clientWidth - geometry.headerSize));
      setViewportH(Math.max(80, el.clientHeight - geometry.headerSize - (showRuler ? 18 : 0)));
    };
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    measure();
    return () => ro.disconnect();
  }, [geometry.headerSize, showRuler]);

  // Type-to-edit: when a printable character is typed while not editing,
  // start editing the active cell with that character.
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (editing || e.ctrlKey || e.metaKey || e.altKey) return;
      // Never hijack typing while a modal dialog owns the interaction.
      if (useUiStore.getState().dialog) return;
      const t = e.target as HTMLElement | null;
      if (
        t &&
        (t.tagName === 'INPUT' ||
          t.tagName === 'TEXTAREA' ||
          t.tagName === 'SELECT' ||
          t.tagName === 'BUTTON' ||
          t.isContentEditable)
      ) {
        return;
      }
      if (e.key.length === 1) {
        e.preventDefault();
        useWorkbookStore.getState().startEdit(e.key);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [editing]);

  const columnHeaderRef = useRef<HTMLDivElement>(null);
  const rowHeaderRef = useRef<HTMLDivElement>(null);
  const cornerRef = useRef<HTMLDivElement>(null);

  // View > Split: every pane owns its scroll container and header strips, so the
  // panes scroll independently the way Excel's Split command behaves.
  const [paneScrolls, setPaneScrolls] = useState<Record<string, { left: number; top: number }>>({});
  const paneRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const colHeaderRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const rowHeaderRefs = useRef<Record<string, HTMLDivElement | null>>({});

  const paneRef = (id: string) =>
    (node: HTMLDivElement | null) => {
      paneRefs.current[id] = node;
    };
  const colHeaderRefFor = (id: string) => (node: HTMLDivElement | null) => {
    colHeaderRefs.current[id] = node;
  };
  const rowHeaderRefFor = (id: string) => (node: HTMLDivElement | null) => {
    rowHeaderRefs.current[id] = node;
  };

  const syncHeaders = useCallback((left: number, top: number) => {
    if (columnHeaderRef.current) columnHeaderRef.current.style.transform = `translate3d(${-left}px,0,0)`;
    if (rowHeaderRef.current) rowHeaderRef.current.style.transform = `translate3d(0,${-top}px,0)`;
  }, []);

  const onScroll = useCallback(
    (el: HTMLDivElement) => {
      const left = el.scrollLeft;
      const top = el.scrollTop;
      syncHeaders(left, top);
      setScrollLeft(left);
      setScrollTop(top);
    },
    [syncHeaders]
  );

  const onPaneScroll = useCallback((id: string, el: HTMLDivElement) => {
    const left = el.scrollLeft;
    const top = el.scrollTop;
    const col = colHeaderRefs.current[id];
    const row = rowHeaderRefs.current[id];
    if (col) col.style.transform = `translate3d(${-left}px,0,0)`;
    if (row) row.style.transform = `translate3d(0,${-top}px,0)`;
    if (id === 'main' || id === 'tl') {
      setScrollLeft(left);
      setScrollTop(top);
      syncHeaders(left, top);
    } else {
      setPaneScrolls((cur) => (cur[id]?.left === left && cur[id]?.top === top ? cur : { ...cur, [id]: { left, top } }));
    }
  }, [syncHeaders]);

  if (!workbook || !activeSheetId || !sheet) {
    return <div className="flex-1" />;
  }

  const totalWidth = geometry.colOffset(MAX_RENDER_COLS + 1 + OVERSCAN);
  const totalHeight = geometry.rowOffset(Math.min(MAX_RENDER_ROWS + OVERSCAN + 1, 1048576));

  // Frozen row/col are 1-based indices of the first *unfrozen* row/col.
  const freezeRow = sheet.freeze?.row ?? 0;
  const freezeCol = sheet.freeze?.col ?? 0;

  const visCols = geometry.visibleCols(scrollLeft, viewportW);
  const visRows = geometry.visibleRows(scrollTop, viewportH);
  const firstCol = Math.max(0, visCols.first);
  const lastCol = Math.min(MAX_RENDER_COLS, visCols.last);
  const firstRow = Math.max(1, visRows.first);
  const lastRow = Math.min(MAX_RENDER_ROWS, visRows.last);

  // Prune hidden cols/rows from the rendered span (simplified: skip rendering
  // but keep geometry for offsets).
  const renderCols: number[] = [];
  for (let c = Math.max(0, firstCol - OVERSCAN); c <= Math.min(lastCol + OVERSCAN, MAX_RENDER_COLS); c++) {
    if (!isColHidden(sheet, c)) renderCols.push(c);
  }
  const renderRows: number[] = [];
  for (let r = Math.max(1, firstRow - OVERSCAN); r <= Math.min(lastRow + OVERSCAN, MAX_RENDER_ROWS); r++) {
    if (!isRowHidden(sheet, r)) renderRows.push(r);
  }

  const selInRange = (row: number, col: number) =>
    row >= Math.min(selection.row, selection.endRow) &&
    row <= Math.max(selection.row, selection.endRow) &&
    col >= Math.min(selection.col, selection.endCol) &&
    col <= Math.max(selection.col, selection.endCol);

  const isActive = (row: number, col: number) =>
    row === selection.row && col === selection.col;

  // A grid click must return keyboard focus to the sheet. Cell and header
  // pointer handlers call preventDefault(), which also cancels the browser's
  // default focus move, so a ribbon element left focused (e.g. the font/size
  // <select>) would swallow the next keystroke and type-to-edit would appear
  // broken (nothing would happen until F2/Enter is pressed).
  const releaseRibbonFocus = () => {
    const ae = document.activeElement as HTMLElement | null;
    if (!ae || ae === document.body) return;
    const tag = ae.tagName;
    if (tag === 'SELECT' || tag === 'BUTTON' || tag === 'INPUT' || tag === 'TEXTAREA') ae.blur();
  };

  const handleCellPointerDown = (row: number, col: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (editing) {
      // Commit the pending edit on the cell that was being edited, then move
      // the selection to the freshly clicked cell.
      useWorkbookStore.getState().commitEdit(editing.text);
      setSelection(SINGLE(row, col));
      return;
    }
    releaseRibbonFocus();
    if (e.shiftKey) {
      setSelection({ ...selection, endRow: row, endCol: col });
    } else {
      setSelection(SINGLE(row, col));
    }
  };

  const handleDoubleClick = (row: number, col: number) => {
    setSelection(SINGLE(row, col));
    setTimeout(() => startEdit(), 0);
  };

  /* Whole-row / whole-column selection.
     The engine stores style per cell and `rectRefs` materialises every ref in a
     range, so a literal 1,048,576-row column would build a million strings and
     freeze the tab. Extent is therefore the used range plus a generous margin
     (never less than one full viewport), which stays instant while still
     covering everything the user can see or has data in. */
  const usedExtent = () => {
    let maxRow = 1;
    let maxCol = 0;
    for (const ref of Object.keys(sheet.cells)) {
      const p = parseRange(ref);
      if (!p) continue;
      if (p.eRow > maxRow) maxRow = p.eRow;
      if (p.eCol > maxCol) maxCol = p.eCol;
    }
    return { maxRow, maxCol };
  };

  const selectColumn = (col: number, extend: boolean) => {
    if (editing) useWorkbookStore.getState().commitEdit(editing.text);
    releaseRibbonFocus();
    const endRow = Math.max(usedExtent().maxRow + 500, 1000);
    setSelection(extend ? { row: selection.row, col: selection.col, endRow, endCol: col } : { row: 1, col, endRow, endCol: col });
  };

  const selectRow = (row: number, extend: boolean) => {
    if (editing) useWorkbookStore.getState().commitEdit(editing.text);
    releaseRibbonFocus();
    const endCol = Math.max(usedExtent().maxCol + 50, 200);
    setSelection(extend ? { row: selection.row, col: selection.col, endRow: row, endCol } : { row, col: 0, endRow: row, endCol });
  };

  /* ------------------------- Right-click context menus ------------------------- */

  const cellContextMenu = (row: number, col: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const st = useWorkbookStore.getState();
    if (editing) st.commitEdit(editing.text);
    const sel = st.selection;
    const inSel =
      row >= Math.min(sel.row, sel.endRow) &&
      row <= Math.max(sel.row, sel.endRow) &&
      col >= Math.min(sel.col, sel.endCol) &&
      col <= Math.max(sel.col, sel.endCol);
    if (!inSel) st.setSelection(SINGLE(row, col));
    const s = useWorkbookStore.getState().selection;
    const sRow = Math.min(s.row, s.endRow);
    const eRow = Math.max(s.row, s.endRow);
    const sCol = Math.min(s.col, s.endCol);
    const eCol = Math.max(s.col, s.endCol);
    const sheetId = useWorkbookStore.getState().activeSheetId;
    if (!sheetId) return;
    const rowCount = eRow - sRow + 1;
    const colCount = eCol - sCol + 1;
    const items: ContextMenuItem[] = [
      { label: 'Cut', hint: 'Ctrl+X', onClick: () => useWorkbookStore.getState().cutSelection() },
      { label: 'Copy', hint: 'Ctrl+C', onClick: () => useWorkbookStore.getState().copySelection() },
      { label: 'Paste', hint: 'Ctrl+V', onClick: () => useWorkbookStore.getState().paste() },
      { type: 'separator' },
      { label: 'Insert Rows', onClick: () => applyCommand(new InsertRowsCommand('Insert rows', sheetId, sRow - 1, rowCount)) },
      { label: 'Insert Columns', onClick: () => applyCommand(new InsertColumnsCommand('Insert columns', sheetId, sCol, colCount)) },
      { type: 'separator' },
      { label: 'Delete Rows', onClick: () => applyCommand(new DeleteRowsCommand('Delete rows', sheetId, sRow - 1, rowCount)) },
      { label: 'Delete Columns', onClick: () => applyCommand(new DeleteColumnsCommand('Delete columns', sheetId, sCol, colCount)) },
      { label: 'Clear Contents', onClick: () => applyCommand(new ClearRangeCommand('Clear contents', sheetId, sRow, sCol, eRow, eCol)) },
      { type: 'separator' },
      { label: 'Format Cells…', hint: 'Ctrl+1', onClick: () => openDialog('formatCells') },
      { label: 'Insert Comment', onClick: () => openDialog('insertComment') },
      { label: 'Sort Range…', onClick: () => openDialog('sort') }
    ];
    showContextMenu(items, e.clientX, e.clientY);
  };

  const columnContextMenu = (col: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const st = useWorkbookStore.getState();
    if (editing) st.commitEdit(editing.text);
    const sheetId = st.activeSheetId;
    if (!sheetId) return;
    const endRow = Math.max(usedExtent().maxRow + 500, 1000);
    st.setSelection({ row: 1, col, endRow, endCol: col });
    const items: ContextMenuItem[] = [
      { label: 'Insert Column Left', onClick: () => applyCommand(new InsertColumnsCommand('Insert column left', sheetId, col, 1)) },
      { label: 'Insert Column Right', onClick: () => applyCommand(new InsertColumnsCommand('Insert column right', sheetId, col + 1, 1)) },
      { label: 'Delete Column', onClick: () => applyCommand(new DeleteColumnsCommand('Delete column', sheetId, col, 1)) },
      { type: 'separator' },
      { label: 'Clear Column Contents', onClick: () => applyCommand(new ClearRangeCommand('Clear column', sheetId, 1, col, endRow, col)) }
    ];
    showContextMenu(items, e.clientX, e.clientY);
  };

  const rowContextMenu = (row: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const st = useWorkbookStore.getState();
    if (editing) st.commitEdit(editing.text);
    const sheetId = st.activeSheetId;
    if (!sheetId) return;
    const endCol = Math.max(usedExtent().maxCol + 50, 200);
    st.setSelection({ row, col: 0, endRow: row, endCol });
    const items: ContextMenuItem[] = [
      { label: 'Insert Row Above', onClick: () => applyCommand(new InsertRowsCommand('Insert row above', sheetId, row - 1, 1)) },
      { label: 'Insert Row Below', onClick: () => applyCommand(new InsertRowsCommand('Insert row below', sheetId, row, 1)) },
      { label: 'Delete Row', onClick: () => applyCommand(new DeleteRowsCommand('Delete row', sheetId, row - 1, 1)) },
      { type: 'separator' },
      { label: 'Clear Row Contents', onClick: () => applyCommand(new ClearRangeCommand('Clear row', sheetId, row, 0, row, endCol)) }
    ];
    showContextMenu(items, e.clientX, e.clientY);
  };

  // Cell rendering
  const cellWidth = (c: number) => geometry.colWidth(c);
  const cellHeight = (r: number) => geometry.rowHeight(r);

  // Ruler strip height; 0 hides it so the grid reclaims the space.
  const rulerH = showRuler ? 18 : 0;
  const gridTop = rulerH;

  // View > Split arrives as a cell reference; this grid owns the geometry, so it
  // converts that into pane sizes and clears the request.
  useEffect(() => {
    if (!splitRequest) return;
    const { col, row } = splitRequest;
    // A split in row 1 only produces a vertical divider; col A only a horizontal one.
    const x = col > 0 ? geometry.colOffset(col) + geometry.headerSize : 0;
    const y = row > 1 ? geometry.rowOffset(row) : 0;
    useUiStore.getState().setSplit({
      x: x > 0 ? Math.max(90, Math.min(x, Math.max(90, viewportW - 90))) : 0,
      y: y > 0 ? Math.max(60, Math.min(y, Math.max(60, viewportH - 60))) : 0
    });
    useUiStore.getState().clearSplitRequest();
  }, [splitRequest, geometry, viewportW, viewportH]);

  const paneVisible = (scLeft: number, scTop: number, vw: number, vh: number) => {
    const vc = geometry.visibleCols(scLeft, Math.max(60, vw));
    const vr = geometry.visibleRows(scTop, Math.max(40, vh));
    const lastVisibleCol = Math.min(MAX_RENDER_COLS, vc.last);
    const cols: number[] = [];
    for (let c = Math.max(0, vc.first - OVERSCAN); c <= Math.min(lastVisibleCol + OVERSCAN, MAX_RENDER_COLS); c++) {
      if (!isColHidden(sheet, c)) cols.push(c);
    }
    const rows: number[] = [];
    for (let r = Math.max(1, vr.first - OVERSCAN); r <= Math.min(Math.min(MAX_RENDER_ROWS, vr.last) + OVERSCAN, MAX_RENDER_ROWS); r++) {
      if (!isRowHidden(sheet, r)) rows.push(r);
    }
    return { cols, rows, lastVisibleCol };
  };

  const renderPane = (p: {
    paneId: string;
    box: React.CSSProperties;
    scLeft: number;
    scTop: number;
    vw: number;
    vh: number;
    frozen: boolean;
    objects: boolean;
    breaks: boolean;
  }) => {
    const { cols: paneCols, rows: paneRows, lastVisibleCol } = paneVisible(p.scLeft, p.scTop, p.vw, p.vh);
    return (
      <div className="absolute" style={p.box} data-grid-pane={p.paneId}>
        {/* Corner */}
        <div
          className="absolute z-20 grid-corner"
          style={{ width: geometry.headerSize, height: geometry.headerSize, left: 0, top: 0 }}
        >
          <div className="h-full w-full bg-cell border-b border-r" style={{ borderColor: 'rgb(var(--line-strong))' }} />
        </div>
        {/* Column headers */}
        <div
          className="absolute z-10 overflow-hidden"
          style={{ left: geometry.headerSize, right: 0, top: 0, height: geometry.headerSize }}
          onPointerDown={(e) => e.stopPropagation()}
        >
          <div ref={colHeaderRefFor(p.paneId)} className="absolute flex" style={{ top: 0, height: geometry.headerSize }}>
            {paneCols.map((c) => (
              <ColHeader
                key={c}
                label={colToLabel(c)}
                width={cellWidth(c)}
                offset={geometry.colOffset(c)}
                height={geometry.headerSize}
                active={selection.col === c && selection.row === selection.endRow}
                isSelected={selection.col === c}
                onResizeStart={(e) => resizeCol(e, c, geometry, sheet)}
                onSelect={(e) => selectColumn(c, e.shiftKey)}
                onContextMenu={(e) => columnContextMenu(c, e)}
              />
            ))}
            <div
              style={{
                width: Math.max(0, totalWidth - (geometry.colOffset(lastVisibleCol) + cellWidth(lastVisibleCol))),
                height: 1
              }}
            />
          </div>
        </div>
        {/* Row headers */}
        <div
          className="absolute z-10 overflow-hidden"
          style={{ left: 0, top: geometry.headerSize, width: geometry.headerSize, bottom: 0 }}
          onPointerDown={(e) => e.stopPropagation()}
        >
          <div ref={rowHeaderRefFor(p.paneId)} className="absolute flex flex-col" style={{ left: 0, width: geometry.headerSize }}>
            {paneRows.map((r) => (
              <RowHeader
                key={r}
                index={r}
                width={geometry.headerSize}
                height={cellHeight(r)}
                offset={geometry.rowOffset(r)}
                isSelected={selection.row === r}
                active={selection.row === r && selection.col === selection.endCol}
                onResizeStart={(e) => resizeRow(e, r, geometry, sheet)}
                onSelect={(e) => selectRow(r, e.shiftKey)}
                onContextMenu={(e) => rowContextMenu(r, e)}
              />
            ))}
          </div>
        </div>
        {/* Cells */}
        <div
          ref={paneRef(p.paneId)}
          className="absolute overflow-auto"
          style={{ left: geometry.headerSize, top: geometry.headerSize, right: 0, bottom: 0, cursor: 'default' }}
          onScroll={(e) => onPaneScroll(p.paneId, e.currentTarget)}
          onPointerDown={(e) => {
            if ((e.target as HTMLElement).dataset.cellRef) return;
            if (editing) useWorkbookStore.getState().commitEdit(editing.text);
            releaseRibbonFocus();
            setSelection(SINGLE(1, 0));
          }}
        >
          <div style={{ width: totalWidth, height: totalHeight, position: 'relative' }}>
            {paneRows.map((r) =>
              paneCols.map((c) => {
                const colOffset = geometry.colOffset(c);
                const rowOffset = geometry.rowOffset(r);
                const merge = mergeMap.get(`${c}_${r}`);
                const isMergeAnchor = !!merge && merge.sCol === c && merge.sRow === r;
                // Cells swallowed by a merged range are not rendered separately.
                if (merge && !isMergeAnchor) return null;
                const cell = getCell(sheet, c, r);
                const w = merge ? geometry.colOffset(merge.eCol + 1) - colOffset : cellWidth(c);
                const h = merge ? geometry.rowOffset(merge.eRow + 1) - rowOffset : cellHeight(r);
                return (
                  <div
                    key={`${r}_${c}`}
                    data-cell-ref={cellToA1(c, r)}
                    className={`absolute grid-cell ${merge ? 'grid-cell-merged' : ''} ${selInRange(r, c) ? 'grid-cell-selected' : ''} ${isActive(r, c) ? 'grid-cell-active' : ''}`}
                    style={{
                      left: colOffset,
                      top: rowOffset,
                      width: w,
                      height: h,
                      borderRight: showGridlines ? `1px solid rgb(var(--line))` : 'none',
                      borderBottom: showGridlines ? `1px solid rgb(var(--line))` : 'none',
                      background: selInRange(r, c) ? 'rgb(var(--accent) / 0.10)' : undefined
                    }}
                    onPointerDown={(e) => handleCellPointerDown(r, c, e)}
                    onDoubleClick={() => handleDoubleClick(r, c)}
                    onContextMenu={(e) => cellContextMenu(r, c, e)}
                  >
                    <CellContent
                      cell={cell}
                      row={r}
                      col={c}
                      sheet={sheet}
                      showFormulas={showFormulas}
                      showNotes={showNotes}
                      engine={engine as any}
                      height={h}
                    />
                  </div>
                );
              })
            )}
            {/* Table styling bands (visual only, never intercepts pointer events) */}
            {p.objects &&
              sheet.tables.map((t) => <TableBands key={t.id} table={t} geometry={geometry} />)}
            {/* Charts float above the cells */}
            {p.objects &&
              sheet.charts.map((ch) => (
                <ChartFrame
                  key={ch.id}
                  chart={ch}
                  sheet={sheet}
                  engine={engine as any}
                  geometry={geometry}
                  selected={selectedChartId === ch.id}
                  onSelect={() => {
                    useWorkbookStore.getState().setSelectedChartId(ch.id);
                  }}
                />
              ))}
            {p.objects &&
              (sheet.shapes ?? [])
                .slice()
                .sort((a, b) => (a.z ?? 0) - (b.z ?? 0))
                .map((sp) => (
                  <ShapeFrame
                    key={sp.id}
                    shape={sp}
                    scale={geometry.scale}
                    selected={selectedShapeId === sp.id}
                    onSelect={() => {
                      useWorkbookStore.getState().setSelectedShapeId(sp.id);
                    }}
                    onCommit={(patch) => {
                      useWorkbookStore.getState().updateShape(sp.id, patch);
                    }}
                  />
                ))}
            {/* Fill handle at the bottom-right of the selection */}
            {p.objects &&
              (() => {
                const lastColSel = Math.max(selection.col, selection.endCol);
                const lastRowSel = Math.max(selection.row, selection.endRow);
                return (
                  <div
                    className="absolute"
                    style={{
                      left: geometry.colOffset(lastColSel + 1) - 3,
                      top: geometry.rowOffset(lastRowSel + 1) - 3,
                      zIndex: 30
                    }}
                  >
                    <div className="grid-fill-handle" />
                  </div>
                );
              })()}
            {/* Editing overlay */}
            {p.objects &&
              editing &&
              (() => {
                const merge = mergeMap.get(`${editing.col}_${editing.row}`);
                const isAnchor = !!merge && merge.sCol === editing.col && merge.sRow === editing.row;
                return (
                  <EditorOverlay
                    row={editing.row}
                    col={editing.col}
                    cell={getCell(sheet, editing.col, editing.row)}
                    initialText={editing.text}
                    selectAllOnMount={!editing.seeded}
                    left={geometry.colOffset(editing.col)}
                    top={geometry.rowOffset(editing.row)}
                    width={
                      merge && isAnchor
                        ? geometry.colOffset(merge.eCol + 1) - geometry.colOffset(editing.col)
                        : cellWidth(editing.col)
                    }
                    height={
                      merge && isAnchor
                        ? geometry.rowOffset(merge.eRow + 1) - geometry.rowOffset(editing.row)
                        : cellHeight(editing.row)
                    }
                    onCommit={commitEdit}
                    onCancel={cancelEdit}
                  />
                );
              })()}
          </div>
        </div>
        {/* Page Break Preview: manual breaks + print-area boundary. */}
        {p.breaks &&
          pageBreakPreview &&
          (() => {
            const brk = sheet.pageSetup?.pageBreaks ?? { rows: [], cols: [] };
            const area = sheet.pageSetup?.printArea;
            const areaBox = area ? rangeToA1Box(area) : null;
            return (
              <div
                className="pointer-events-none absolute"
                style={{ left: geometry.headerSize, top: geometry.headerSize, right: 0, bottom: 0, zIndex: 25 }}
              >
                {brk.rows.map((r) => (
                  <div
                    key={`pbr_${r}`}
                    className="absolute"
                    style={{
                      left: 0,
                      right: 0,
                      top: geometry.rowOffset(r) - p.scTop,
                      height: 0,
                      borderTop: '2px dashed rgb(var(--danger))'
                    }}
                  />
                ))}
                {brk.cols.map((c) => (
                  <div
                    key={`pbc_${c}`}
                    className="absolute"
                    style={{
                      top: 0,
                      bottom: 0,
                      left: geometry.colOffset(c) - p.scLeft,
                      width: 0,
                      borderLeft: '2px dashed rgb(var(--danger))'
                    }}
                  />
                ))}
                {areaBox && (
                  <div
                    className="absolute"
                    style={{
                      left: geometry.colOffset(areaBox.c1) - p.scLeft,
                      top: geometry.rowOffset(areaBox.r1) - p.scTop,
                      width: geometry.colOffset(areaBox.c2 + 1) - geometry.colOffset(areaBox.c1),
                      height: geometry.rowOffset(areaBox.r2 + 1) - geometry.rowOffset(areaBox.r1),
                      outline: '2px solid rgb(var(--accent))',
                      background: 'rgb(var(--accent) / 0.06)'
                    }}
                  />
                )}
              </div>
            );
          })()}
        {/* Frozen panes: pinned cells re-rendered above the scrolling body so
            they stay put while the rest of the sheet scrolls underneath. */}
        {p.frozen && freezeCol > 0 && (
          <div
            data-frozen-pane="col"
            className="absolute overflow-hidden"
            style={{
              left: geometry.headerSize,
              top: geometry.headerSize,
              width: geometry.colOffset(freezeCol),
              bottom: 0,
              zIndex: 6,
              boxShadow: '2px 0 4px rgba(0,0,0,0.12)'
            }}
          >
            <div
              style={{
                position: 'absolute',
                left: 0,
                top: -p.scTop,
                width: geometry.colOffset(freezeCol),
                height: totalHeight
              }}
            >
              {paneRows.map((r) =>
                Array.from({ length: freezeCol }, (_unused, c) => {
                  const colOffset = geometry.colOffset(c);
                  const rowOffset = geometry.rowOffset(r);
                  const merge = mergeMap.get(`${c}_${r}`);
                  const isMergeAnchor = !!merge && merge.sCol === c && merge.sRow === r;
                  if (merge && !isMergeAnchor) return null;
                  const cell = getCell(sheet, c, r);
                  const w = merge ? geometry.colOffset(merge.eCol + 1) - colOffset : cellWidth(c);
                  const h = merge ? geometry.rowOffset(merge.eRow + 1) - rowOffset : cellHeight(r);
                  return (
                    <div
                      key={`fz_${r}_${c}`}
                      data-cell-ref={cellToA1(c, r)}
                      className={`absolute grid-cell ${merge ? 'grid-cell-merged' : ''} ${selInRange(r, c) ? 'grid-cell-selected' : ''} ${isActive(r, c) ? 'grid-cell-active' : ''}`}
                      style={{
                        left: colOffset,
                        top: rowOffset,
                        width: w,
                        height: h,
                        borderRight: showGridlines ? `1px solid rgb(var(--line))` : 'none',
                        borderBottom: showGridlines ? `1px solid rgb(var(--line))` : 'none',
                        background: selInRange(r, c) ? 'rgb(var(--accent) / 0.10)' : undefined
                      }}
                      onPointerDown={(e) => handleCellPointerDown(r, c, e)}
                      onDoubleClick={() => handleDoubleClick(r, c)}
                      onContextMenu={(e) => cellContextMenu(r, c, e)}
                    >
                      <CellContent
                        cell={cell}
                        row={r}
                        col={c}
                        sheet={sheet}
                        showFormulas={showFormulas}
                        showNotes={showNotes}
                        engine={engine as any}
                        height={h}
                      />
                    </div>
                  );
                })
              )}
            </div>
          </div>
        )}
        {p.frozen && freezeRow > 1 && (
          <div
            data-frozen-pane="row"
            className="absolute overflow-hidden"
            style={{
              left: geometry.headerSize,
              top: geometry.headerSize,
              right: 0,
              height: geometry.rowOffset(freezeRow) - geometry.rowOffset(1),
              zIndex: 7,
              boxShadow: '0 2px 4px rgba(0,0,0,0.12)'
            }}
          >
            <div
              style={{
                position: 'absolute',
                left: -p.scLeft,
                top: 0,
                width: totalWidth,
                height: geometry.rowOffset(freezeRow)
              }}
            >
              {Array.from({ length: freezeRow - 1 }, (_u, i) => i + 1).map((r) =>
                paneCols.map((c) => {
                  const colOffset = geometry.colOffset(c);
                  const rowOffset = geometry.rowOffset(r);
                  const merge = mergeMap.get(`${c}_${r}`);
                  const isMergeAnchor = !!merge && merge.sCol === c && merge.sRow === r;
                  if (merge && !isMergeAnchor) return null;
                  const cell = getCell(sheet, c, r);
                  const w = merge ? geometry.colOffset(merge.eCol + 1) - colOffset : cellWidth(c);
                  const h = merge ? geometry.rowOffset(merge.eRow + 1) - rowOffset : cellHeight(r);
                  return (
                    <div
                      key={`fr_${r}_${c}`}
                      data-cell-ref={cellToA1(c, r)}
                      className={`absolute grid-cell ${merge ? 'grid-cell-merged' : ''} ${selInRange(r, c) ? 'grid-cell-selected' : ''} ${isActive(r, c) ? 'grid-cell-active' : ''}`}
                      style={{
                        left: colOffset,
                        top: rowOffset,
                        width: w,
                        height: h,
                        borderRight: showGridlines ? `1px solid rgb(var(--line))` : 'none',
                        borderBottom: showGridlines ? `1px solid rgb(var(--line))` : 'none',
                        background: selInRange(r, c) ? 'rgb(var(--accent) / 0.10)' : undefined
                      }}
                      onPointerDown={(e) => handleCellPointerDown(r, c, e)}
                      onDoubleClick={() => handleDoubleClick(r, c)}
                      onContextMenu={(e) => cellContextMenu(r, c, e)}
                    >
                      <CellContent
                        cell={cell}
                        row={r}
                        col={c}
                        sheet={sheet}
                        showFormulas={showFormulas}
                        showNotes={showNotes}
                        engine={engine as any}
                        height={h}
                      />
                    </div>
                  );
                })
              )}
            </div>
          </div>
        )}
      </div>
    );
  };

  /** Drag the split dividers to resize the panes, like Excel's split bars. */
  const startSplitDrag = (axis: 'x' | 'y') => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const root = (e.currentTarget as HTMLElement).closest('.grid-root') as HTMLElement | null;
    const rootRect = (root ?? (e.currentTarget as HTMLElement)).getBoundingClientRect();
    const startX = e.clientX;
    const startY = e.clientY;
    const start = split ?? { x: 0, y: 0 };
    const move = (ev: PointerEvent) => {
      const setSplit = useUiStore.getState().setSplit;
      const minPane = 90;
      if (axis === 'x') {
        const maxX = Math.max(minPane, rootRect.width - geometry.headerSize - minPane);
        setSplit({
          x: Math.round(Math.max(minPane, Math.min(maxX, start.x + ev.clientX - startX))),
          y: start.y
        });
      } else {
        const maxY = Math.max(minPane, rootRect.height - geometry.headerSize - minPane);
        setSplit({
          x: start.x,
          y: Math.round(Math.max(minPane, Math.min(maxY, start.y + ev.clientY - startY)))
        });
      }
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const paneFrozen = freezeCol > 0 || freezeRow > 1;

  // Panes for the current split: Excel shows two panes for a single divider and
  // four when the window is split both horizontally and vertically.
  const splitBoxes = (() => {
    if (!split) return [];
    const hs = geometry.headerSize;
    const { x, y } = split;
    const boxes: { id: string; box: React.CSSProperties; vw: number; vh: number }[] = [];
    if (x > 0 && y > 0) {
      boxes.push({ id: 'tl', box: { left: 0, top: gridTop, width: x, height: y }, vw: x - hs, vh: y - hs });
      boxes.push({ id: 'tr', box: { left: x, top: gridTop, right: 0, height: y }, vw: viewportW - x - hs, vh: y - hs });
      boxes.push({ id: 'bl', box: { left: 0, top: gridTop + y, width: x, bottom: 0 }, vw: x - hs, vh: viewportH - y - hs });
      boxes.push({ id: 'br', box: { left: x, top: gridTop + y, right: 0, bottom: 0 }, vw: viewportW - x - hs, vh: viewportH - y - hs });
    } else if (x > 0) {
      boxes.push({ id: 'left', box: { left: 0, top: gridTop, bottom: 0, width: x }, vw: x - hs, vh: viewportH - hs });
      boxes.push({ id: 'right', box: { left: x, top: gridTop, right: 0, bottom: 0 }, vw: viewportW - x - hs, vh: viewportH - hs });
    } else if (y > 0) {
      boxes.push({ id: 'top', box: { left: 0, top: gridTop, right: 0, height: y }, vw: viewportW - hs, vh: y - hs });
      boxes.push({ id: 'bottom', box: { left: 0, top: gridTop + y, right: 0, bottom: 0 }, vw: viewportW - hs, vh: viewportH - y - hs });
    }
    return boxes;
  })();

  return (
    <div
      ref={rootRef}
      className="grid-root relative h-full overflow-hidden select-none"
      style={{ height }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {showRuler && (
        <>
          {/* Horizontal ruler: ticks sit on column boundaries, like Excel. */}
          <div
            className="absolute z-20 overflow-hidden border-b border-r bg-[rgb(var(--grid-header))]"
            data-ruler="horizontal"
            style={{
              left: geometry.headerSize,
              right: 0,
              top: 0,
              height: rulerH,
              borderColor: 'rgb(var(--line-strong))'
            }}
          >
            <div className="absolute" style={{ left: -scrollLeft, top: 0, height: rulerH, width: totalWidth }}>
              {/* Bottom rule for the ruler strip. */}
              <span
                className="absolute"
                style={{ left: 0, top: rulerH - 6, width: totalWidth, height: 1, background: 'rgb(var(--line))' }}
              />
              {renderCols.map((c) => (
                <span
                  key={`rx_${c}`}
                  className="absolute bottom-0 border-l"
                  style={{ left: geometry.colOffset(c), height: 5, borderColor: 'rgb(var(--line-strong))' }}
                />
              ))}
              {renderCols.map((c) => (
                <span
                  key={`rl_${c}`}
                  className="absolute text-[8px] leading-none text-muted"
                  style={{ left: geometry.colOffset(c) + 2, top: 1 }}
                >
                  {c + 1}
                </span>
              ))}
            </div>
          </div>
          {/* Vertical ruler */}
          <div
            className="absolute z-20 overflow-hidden border-b border-r bg-[rgb(var(--grid-header))]"
            data-ruler="vertical"
            style={{
              left: 0,
              top: gridTop,
              width: geometry.headerSize,
              height: geometry.headerSize,
              borderColor: 'rgb(var(--line-strong))'
            }}
          >
            <div
              className="absolute"
              style={{ left: 0, top: -scrollTop, width: geometry.headerSize, height: totalHeight }}
            >
              {renderRows.map((r) => (
                <span
                  key={`ry_${r}`}
                  className="absolute right-0 border-t"
                  style={{ top: geometry.rowOffset(r), width: 5, borderColor: 'rgb(var(--line-strong))' }}
                />
              ))}
              {renderRows.map((r) => (
                <span
                  key={`rn_${r}`}
                  className="absolute text-[8px] leading-none text-muted"
                  style={{ top: geometry.rowOffset(r) + 1, left: 2 }}
                >
                  {r}
                </span>
              ))}
            </div>
          </div>
        </>
      )}
      {split && (split.x > 0 || split.y > 0) ? (
        <>
          {/* Excel splits horizontally, vertically, or both; both gives four panes. */}
          {splitBoxes.map((p) =>
            renderPane({
              paneId: p.id,
              box: p.box,
              scLeft: p.id === 'tl' || p.id === 'top' || p.id === 'left' ? scrollLeft : paneScrolls[p.id]?.left ?? 0,
              scTop: p.id === 'tl' || p.id === 'top' || p.id === 'left' ? scrollTop : paneScrolls[p.id]?.top ?? 0,
              vw: p.vw,
              vh: p.vh,
              frozen: p.id === 'tl',
              objects: p.id === 'tl',
              breaks: p.id === 'tl'
            })
          )}
          {/* Split bars */}
          {split.x > 0 && (
            <div
              className="absolute z-30 grid-split-bar"
              data-split-bar="vertical"
              onPointerDown={startSplitDrag('x')}
              style={{ left: split.x - 3, top: gridTop, bottom: 0, width: 6, cursor: 'col-resize' }}
            />
          )}
          {split.y > 0 && (
            <div
              className="absolute z-30 grid-split-bar"
              data-split-bar="horizontal"
              onPointerDown={startSplitDrag('y')}
              style={{ left: 0, top: gridTop + split.y - 3, right: 0, height: 6, cursor: 'row-resize' }}
            />
          )}
        </>
      ) : (
        renderPane({
          paneId: 'main',
          box: { left: 0, top: gridTop, right: 0, bottom: 0 },
          scLeft: scrollLeft,
          scTop: scrollTop,
          vw: viewportW,
          vh: viewportH,
          frozen: paneFrozen,
          objects: true,
          breaks: true
        })
      )}
    </div>
  );

  function resizeCol(e: React.MouseEvent, col: number, geo: SheetGeometry, sh: Worksheet) {
    e.stopPropagation();
    e.preventDefault();
    const startX = e.clientX;
    const startW = geo.colWidth(col);
    const move = (ev: MouseEvent) => {
      const d = ev.clientX - startX;
      // Drag distance is in rendered pixels; store the width in unscaled model units.
      const w = Math.max(MIN_COL_WIDTH, Math.min(MAX_COL_WIDTH, Math.round(geo.toModelWidth(startW + d))));
      // apply immediately via a lightweight local override store action
      useWorkbookStore.getState().applyCommand(resizeCommand('col', sh.id, col, w) as any);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  function resizeRow(e: React.MouseEvent, row: number, geo: SheetGeometry, sh: Worksheet) {
    e.stopPropagation();
    e.preventDefault();
    const startY = e.clientY;
    const startH = geo.rowHeight(row);
    const move = (ev: MouseEvent) => {
      const d = ev.clientY - startY;
      const h = Math.max(4, Math.round(geo.toModelHeight(startH + d)));
      useWorkbookStore.getState().applyCommand(resizeCommand('row', sh.id, row, h) as any);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }
}

import { ResizeRowsCommand, ResizeColumnsCommand } from '@/engine/commands/commands';

function resizeCommand(kind: 'col' | 'row', sheetId: string, index: number, size: number) {
  if (kind === 'col') {
    return new ResizeColumnsCommand('Resize column', sheetId, [index], Math.round(size));
  }
  return new ResizeRowsCommand('Resize row', sheetId, [index], Math.round(size));
}

function emptySheet(): Worksheet {
  return {
    id: '',
    name: 'Sheet1',
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
    view: { zoom: 100 }
  };
}

function ShapeFrame({
  shape,
  scale,
  selected,
  onSelect,
  onCommit
}: {
  shape: ShapeDef;
  scale: number;
  selected?: boolean;
  onSelect(): void;
  onCommit(patch: Partial<ShapeDef>): void;
}) {
  // Geometry is stored unscaled; preview while dragging, commit once on release
  // so a drag is a single undo step.
  const [drag, setDrag] = React.useState<{
    mode: 'move' | 'resize';
    startX: number;
    startY: number;
    orig: { x: number; y: number; width: number; height: number };
  } | null>(null);
  const [live, setLive] = React.useState<{ x: number; y: number; width: number; height: number } | null>(null);
  const [editingText, setEditingText] = React.useState(false);
  const textRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    if (editingText && textRef.current) {
      textRef.current.focus();
      textRef.current.select();
    }
  }, [editingText]);

  const box = live ?? { x: shape.x, y: shape.y, width: shape.width, height: shape.height };
  const w = Math.max(8, box.width * scale);
  const h = Math.max(6, box.height * scale);
  const stroke = shape.stroke ?? '#1f2937';
  const strokeW = (shape.strokeWidth ?? 1) * scale;
  const isLine = shape.kind === 'line' || shape.kind === 'arrow';
  // Shape text follows the workbook theme's body font unless overridden.
  const theme = useWorkbookStore((s) => s.workbook?.theme);
  const textFont = getTheme(theme).bodyFont;

  const beginDrag = (e: React.PointerEvent, mode: 'move' | 'resize') => {
    e.stopPropagation();
    e.preventDefault();
    onSelect();
    const orig = { x: shape.x, y: shape.y, width: shape.width, height: shape.height };
    setDrag({ mode, startX: e.clientX, startY: e.clientY, orig });
    setLive(orig);
  };

  const onMove = (ev: PointerEvent) => {
    if (!drag) return;
    const dx = (ev.clientX - drag.startX) / scale;
    const dy = (ev.clientY - drag.startY) / scale;
    if (drag.mode === 'move') {
      setLive({ ...drag.orig, x: drag.orig.x + dx, y: drag.orig.y + dy });
    } else {
      setLive({
        ...drag.orig,
        width: Math.max(20, drag.orig.width + dx),
        height: Math.max(16, drag.orig.height + dy)
      });
    }
  };

  const endDrag = () => {
    if (drag && live) {
      const rounded = {
        x: Math.round(live.x),
        y: Math.round(live.y),
        width: Math.round(live.width),
        height: Math.round(live.height)
      };
      const changed =
        rounded.x !== drag.orig.x ||
        rounded.y !== drag.orig.y ||
        rounded.width !== drag.orig.width ||
        rounded.height !== drag.orig.height;
      if (changed) onCommit(rounded);
    }
    setDrag(null);
    setLive(null);
  };

  React.useEffect(() => {
    if (!drag) return;
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', endDrag);
    return () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', endDrag);
    };
  });

  const body = (
    <div
      className="pointer-events-auto"
      data-shape-id={shape.id}
      data-shape-kind={shape.kind}
      style={{ position: 'absolute', left: box.x * scale, top: box.y * scale, width: w, height: h, zIndex: 30 + (shape.z ?? 0) }}
      onPointerDown={(e) => {
        if (editingText && shape.kind === 'textbox') return;
        beginDrag(e, 'move');
      }}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
        onSelect();
        showContextMenu(
          [
            { label: 'Bring to Front', onClick: () => useWorkbookStore.getState().reorderShape(shape.id, 'front') },
            { label: 'Send to Back', onClick: () => useWorkbookStore.getState().reorderShape(shape.id, 'back') },
            { type: 'separator' },
            { label: 'Delete Shape', danger: true, onClick: () => useWorkbookStore.getState().removeShape(shape.id) }
          ],
          e.clientX,
          e.clientY
        );
      }}
      onDoubleClick={(e) => {
        // Excel: double-click a text box to edit its caption in place.
        if (shape.kind !== 'textbox') return;
        e.stopPropagation();
        onSelect();
        setEditingText(true);
      }}
    >
      {editingText && shape.kind === 'textbox' ? (
        <input
          ref={textRef}
          data-shape-text-editor={shape.id}
          className="h-full w-full bg-transparent px-1 text-center outline-none"
          style={{
            fontSize: (shape.fontSize ?? 13) * scale,
            fontFamily: textFont,
            color: stroke,
            border: '1px dashed rgb(var(--accent))'
          }}
          defaultValue={shape.text ?? ''}
          onPointerDown={(e) => e.stopPropagation()}
          onBlur={(e) => {
            const next = e.currentTarget.value;
            if (next !== (shape.text ?? '')) onCommit({ text: next });
            setEditingText(false);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') e.currentTarget.blur();
            if (e.key === 'Escape') {
              e.stopPropagation();
              setEditingText(false);
            }
            e.stopPropagation();
          }}
        />
      ) : isLine ? (
        <svg width={w} height={h} style={{ overflow: 'visible' }}>
          <line
            x1={0}
            y1={0}
            x2={w}
            y2={h}
            stroke={stroke}
            strokeWidth={Math.max(1, strokeW)}
            markerEnd={shape.kind === 'arrow' ? 'url(#brit-arrow)' : undefined}
          />
          {shape.kind === 'arrow' && (
            <defs>
              <marker id="brit-arrow" markerWidth="8" markerHeight="8" refX="6" refY="3" orient="auto">
                <path d="M0,0 L6,3 L0,6 z" fill={stroke} />
              </marker>
            </defs>
          )}
        </svg>
      ) : (
        <div
          className="flex h-full w-full items-center justify-center overflow-hidden"
          style={{
            background: shape.fill ?? 'transparent',
            border: shape.stroke ? `${Math.max(1, strokeW)}px solid ${stroke}` : 'none',
            borderRadius: shape.kind === 'ellipse' ? '50%' : shape.kind === 'rect' ? 2 : 0,
            color: stroke,
            fontSize: (shape.fontSize ?? 13) * scale,
            fontFamily: textFont,
            textAlign: 'center',
            padding: 2 * scale
          }}
        >
          <span>{shape.text ?? ''}</span>
        </div>
      )}
    </div>
  );

  return (
    <>
      {body}
      {selected && (
        <>
          <div
            className="pointer-events-none absolute"
            style={{ left: box.x * scale - 2, top: box.y * scale - 2, width: w + 4, height: h + 4, zIndex: 40 + (shape.z ?? 0) }}
          >
            <div style={{ position: 'absolute', inset: 0, border: '2px solid rgb(var(--accent))', borderRadius: 3 }} />
            {[
              { left: -4, top: -4 },
              { right: -4, top: -4 },
              { left: -4, bottom: -4 }
            ].map((pos, i) => (
              <span
                key={i}
                className="absolute h-[6px] w-[6px] rounded-[1px]"
                style={{ ...pos, background: 'rgb(var(--accent))', border: '1px solid rgb(var(--cell))' }}
              />
            ))}
            <span
              className="absolute h-[7px] w-[7px] rounded-[1px]"
              style={{ right: -5, bottom: -5, background: 'rgb(var(--accent))', border: '1px solid rgb(var(--cell))', cursor: 'nwse-resize', pointerEvents: 'auto' }}
              onPointerDown={(e) => beginDrag(e, 'resize')}
            />
          </div>
        </>
      )}
    </>
  );
}

function CellContent({
  cell,
  row,
  col,
  sheet,
  showFormulas,
  showNotes,
  engine,
  height
}: {
  cell?: Cell;
  row: number;
  col: number;
  sheet: Worksheet;
  showFormulas: boolean;
  showNotes: boolean;
  engine: { compute(sheetId: string, row: number, col: number): unknown };
  height: number;
}) {
  const computed = cell?.formula ? engine.compute(sheet.id, row, col) : cell?.value;
  const disp = displayCell(cell, showFormulas, computed as never);
  const ref = cellToA1(col, row);
  const note = sheet.notes.find((n) => n.ref === ref);
  const link = cell?.hyperlink;
  // Theme-referencing colors follow the workbook theme, like Excel's theme colors.
  const theme = useWorkbookStore((s) => s.workbook?.theme);
  const themedBg =
    cell?.style?.backgroundTheme != null ? accent(theme, cell.style.backgroundTheme) : undefined;
  const themedFg = cell?.style?.colorTheme != null ? accent(theme, cell.style.colorTheme) : undefined;
  return (
    <div
      className="cell-display h-full w-full overflow-hidden px-1 whitespace-nowrap"
      style={{
        textAlign: disp.align,
        // .cell-display is a flex container, so horizontal placement needs
        // justify-content (text-align alone never moves the flex-item span).
        justifyContent:
          disp.align === 'right' ? 'flex-end' : disp.align === 'center' ? 'center' : 'flex-start',
        fontFamily: cell?.style?.fontFamily ?? `"${getTheme(theme).bodyFont}", Calibri, sans-serif`,
        fontSize: cell?.style?.fontSize,
        fontWeight: cell?.style?.bold ? 600 : 400,
        fontStyle: cell?.style?.italic ? 'italic' : 'normal',
        textDecoration:
          link && !cell?.style?.underline && !cell?.style?.strikethrough
            ? 'underline'
            : cell?.style?.underline && cell?.style?.strikethrough
              ? 'underline line-through'
              : cell?.style?.underline
                ? 'underline'
                : cell?.style?.strikethrough
                  ? 'line-through'
                  : undefined,
        color: themedFg ?? cell?.style?.color ?? (link ? '#0563c1' : undefined),
        background: themedBg ?? cell?.style?.backgroundColor,
        lineHeight: `${height}px`,
        verticalAlign:
          cell?.style?.verticalAlign === 'middle'
            ? 'middle'
            : cell?.style?.verticalAlign === 'bottom'
              ? 'bottom'
              : 'top',
        display: 'flex',
        alignItems:
          cell?.style?.verticalAlign === 'middle'
            ? 'center'
            : cell?.style?.verticalAlign === 'bottom'
              ? 'flex-end'
              : 'flex-start',
        whiteSpace: cell?.style?.wrapText ? 'normal' : 'nowrap',
        overflow: 'hidden',
        position: 'relative'
      }}
      title={note ? `${note.author ? note.author + ':\n' : ''}${note.text}\n\n${disp.text}` : disp.text}
    >
      <span className="min-w-0">{disp.text}</span>
      {note && showNotes && !showFormulas && (
        <span
          aria-label="Note"
          className="absolute right-0 top-0 h-0 w-0"
          style={{ borderTop: '5px solid #d13438', borderLeft: '5px solid transparent' }}
        />
      )}
    </div>
  );
}

function EditorOverlay({
  row,
  col,
  cell,
  left,
  top,
  width,
  height,
  initialText,
  selectAllOnMount,
  onCommit,
  onCancel
}: {
  row: number;
  col: number;
  cell?: Cell;
  left: number;
  top: number;
  width: number;
  height: number;
  initialText: string;
  selectAllOnMount: boolean;
  onCommit(text: string): void;
  onCancel(): void;
}) {
  const [text, setText] = useState(initialText);
  const ref = useRef<HTMLInputElement>(null);
  // Live formatting: the overlay must mirror the cell's current style so that
  // clicking Bold/Font/Align mid-edit updates the text immediately (Excel
  // bolds what you're typing, not only the committed cell).
  const theme = useWorkbookStore((s) => s.workbook?.theme);
  const themedBg =
    cell?.style?.backgroundTheme != null ? accent(theme, cell.style.backgroundTheme) : undefined;
  const themedFg = cell?.style?.colorTheme != null ? accent(theme, cell.style.colorTheme) : undefined;
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    el.focus();
    if (selectAllOnMount) {
      el.select();
    } else {
      // Started by typing: the first character is already seeded, so put the
      // caret after it. Selecting here would make the *next* keystroke replace
      // that character (typing "HELLO" would end up as "ELLO").
      const end = el.value.length;
      el.setSelectionRange(end, end);
    }
  }, []);
  const commit = () => {
    const v = text.replace(/\s+$/g, '');
    onCommit(v);
  };
  return (
    <div
      className="absolute editable-overlay"
      style={{ left, top, width, height, zIndex: 40, padding: 1 }}
    >
      <input
        ref={ref}
        value={text}
        onChange={(e) => {
          const v = e.target.value;
          setText(v);
          useWorkbookStore.setState({ editing: { row, col, text: v } });
        }}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
          } else if (e.key === 'Escape') {
            onCancel();
          } else if (e.key === 'Tab') {
            e.preventDefault();
            const v = text.replace(/\s+$/g, '');
            onCommit(v);
            useWorkbookStore.getState().moveActive(0, e.shiftKey ? -1 : 1);
          } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
            // Excel commits the in-progress edit and moves the active cell.
            e.preventDefault();
            const v = text.replace(/\s+$/g, '');
            const [dRow, dCol] =
              e.key === 'ArrowDown' ? [1, 0] : e.key === 'ArrowUp' ? [-1, 0] : e.key === 'ArrowLeft' ? [0, -1] : [0, 1];
            onCommit(v);
            useWorkbookStore.getState().moveActive(dRow, dCol);
          }
        }}
        className="h-full w-full border bg-cell px-1 outline-none"
        style={{
          borderColor: 'rgb(var(--accent))',
          fontSize: cell?.style?.fontSize ?? 12,
          fontWeight: cell?.style?.bold ? 600 : 400,
          fontStyle: cell?.style?.italic ? 'italic' : 'normal',
          textDecoration:
            cell?.style?.underline && cell?.style?.strikethrough
              ? 'underline line-through'
              : cell?.style?.underline
                ? 'underline'
                : cell?.style?.strikethrough
                  ? 'line-through'
                  : undefined,
          fontFamily: cell?.style?.fontFamily ?? `"${getTheme(theme).bodyFont}", Calibri, sans-serif`,
          color: themedFg ?? cell?.style?.color,
          background: themedBg ?? cell?.style?.backgroundColor,
          textAlign: (cell?.style?.horizontalAlign as 'left' | 'right' | 'center' | undefined)
        }}
      />
    </div>
  );
}

function ColHeader({
  label,
  width,
  offset,
  height,
  active,
  isSelected,
  onResizeStart,
  onSelect,
  onContextMenu
}: {
  label: string;
  width: number;
  offset: number;
  height: number;
  active: boolean;
  isSelected: boolean;
  onResizeStart(e: React.MouseEvent): void;
  onSelect(e: React.PointerEvent): void;
  onContextMenu(e: React.MouseEvent): void;
}) {
  const [hovered, setHovered] = React.useState(false);
  return (
    <div
      className="relative flex items-center justify-center text-[11px] font-medium"
      style={{
        width,
        height,
        backgroundColor: isSelected ? 'rgb(var(--accent-soft))' : 'rgb(var(--grid-header))',
        color: 'rgb(var(--text-muted))',
        borderRight: '1px solid rgb(var(--line))',
        borderBottom: '1px solid rgb(var(--line-strong))',
        position: 'absolute',
        left: offset,
        top: 0
      }}
      onPointerDown={onSelect}
      onContextMenu={onContextMenu}
      title={`Select column ${label}`}
    >
      {label}
      <div
        className="absolute right-0 top-0 h-full"
        style={{ width: RESIZE_PROXIMITY, cursor: 'col-resize', zIndex: 5, touchAction: 'none' }}
        onPointerDown={(e) => {
          e.stopPropagation();
          onResizeStart(e);
        }}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        title="Resize column"
      >
        {hovered && (
          <div
            className="absolute right-0 top-0 h-full"
            style={{ width: 2, background: 'rgb(var(--accent))', opacity: 0.55 }}
          />
        )}
      </div>
    </div>
  );
}

function RowHeader({
  index,
  width,
  height,
  offset,
  active,
  isSelected,
  onResizeStart,
  onSelect,
  onContextMenu
}: {
  index: number;
  width: number;
  height: number;
  offset: number;
  active: boolean;
  isSelected: boolean;
  onResizeStart(e: React.MouseEvent): void;
  onSelect(e: React.PointerEvent): void;
  onContextMenu(e: React.MouseEvent): void;
}) {
  const [hovered, setHovered] = React.useState(false);
  return (
    <div
      className="relative flex items-center justify-center text-[11px] font-medium"
      style={{
        width,
        height,
        backgroundColor: isSelected ? 'rgb(var(--accent-soft))' : 'rgb(var(--grid-header))',
        color: 'rgb(var(--text-muted))',
        borderRight: '1px solid rgb(var(--line-strong))',
        borderBottom: '1px solid rgb(var(--line))',
        position: 'absolute',
        left: 0,
        top: offset
      }}
      onPointerDown={onSelect}
      onContextMenu={onContextMenu}
      title={`Select row ${index}`}
    >
      {index}
      <div
        className="absolute bottom-0 left-0 w-full"
        style={{ height: RESIZE_PROXIMITY, cursor: 'row-resize', zIndex: 5, touchAction: 'none' }}
        onPointerDown={(e) => {
          e.stopPropagation();
          onResizeStart(e);
        }}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        title="Resize row"
      >
        {hovered && (
          <div
            className="absolute bottom-0 left-0 w-full"
            style={{ height: 2, background: 'rgb(var(--accent))', opacity: 0.55 }}
          />
        )}
      </div>
    </div>
  );
}

/* ----------------------------- Tables & charts ----------------------------- */

interface ChartSeries {
  name: string;
  values: number[];
}

interface ChartData {
  categories: string[];
  series: ChartSeries[];
}

function rawValue(
  sheet: Worksheet,
  engine: { compute(sheetId: string, row: number, col: number): unknown },
  row: number,
  col: number
): unknown {
  const cell = getCell(sheet, col, row);
  if (!cell) return undefined;
  return cell.formula ? engine.compute(sheet.id, row, col) : cell.value;
}

function toNumber(v: unknown): number {
  const n = typeof v === 'number' ? v : typeof v === 'string' ? Number(v) : NaN;
  return Number.isFinite(n) ? n : 0;
}

function readChartData(
  sheet: Worksheet,
  engine: { compute(sheetId: string, row: number, col: number): unknown },
  chart: ChartDef
): ChartData | null {
  if (!chart.dataRange) return null;
  const p = parseRange(chart.dataRange);
  if (!p) return null;
  const colCount = p.eCol - p.sCol + 1;
  const rowCount = p.eRow - p.sRow + 1;
  if (colCount < 1 || rowCount < 2) return null;

  const firstDataRow = p.sRow + 1;
  const categories: string[] = [];
  const series: ChartSeries[] = [];

  if (colCount === 1) {
    for (let r = firstDataRow; r <= p.eRow; r++) categories.push(String(r));
    const header = rawValue(sheet, engine, p.sRow, p.sCol);
    series.push({
      name: header == null ? '' : String(header),
      values: Array.from({ length: rowCount - 1 }, (_, i) => toNumber(rawValue(sheet, engine, firstDataRow + i, p.sCol)))
    });
    return { categories, series };
  }

  for (let r = firstDataRow; r <= p.eRow; r++) {
    const v = rawValue(sheet, engine, r, p.sCol);
    categories.push(v == null ? String(r) : String(v));
  }
  for (let c = p.sCol + 1; c <= p.eCol; c++) {
    const header = rawValue(sheet, engine, p.sRow, c);
    series.push({
      name: header == null ? '' : String(header),
      values: Array.from({ length: rowCount - 1 }, (_, i) => toNumber(rawValue(sheet, engine, firstDataRow + i, c)))
    });
  }
  if (!series.length) return null;
  return { categories, series };
}

const CHART_COLORS = ['#3455dd', '#e36c09', '#00b050', '#7030a0', '#d35fb6', '#2e75b6'];

/** Chart series palette follows the workbook theme (Page Layout > Themes). */
function useSeriesColors(): string[] {
  const theme = useWorkbookStore((s) => s.workbook?.theme);
  return theme ? getTheme(theme).accents : CHART_COLORS;
}

function ChartFrame({
  chart,
  sheet,
  engine,
  geometry,
  selected,
  onSelect
}: {
  chart: ChartDef;
  sheet: Worksheet;
  engine: { compute(sheetId: string, row: number, col: number): unknown };
  geometry: SheetGeometry;
  selected?: boolean;
  onSelect?: () => void;
}) {
  const data = readChartData(sheet, engine, chart);
  const colors = useSeriesColors();

  // Charts are anchored to a grid cell but sized in pixels; dragging previews a
  // pixel offset live and commits by snapping back to the nearest grid position.
  const [drag, setDrag] = React.useState<{
    mode: 'move' | 'resize';
    startX: number;
    startY: number;
    origX: number;
    origY: number;
    origW: number;
    origH: number;
    dx: number;
    dy: number;
  } | null>(null);
  const dragRef = React.useRef(drag);
  dragRef.current = drag;

  const baseLeft = geometry.colOffset(Math.max(0, Math.min(chart.x, MAX_RENDER_COLS - 1)));
  const baseTop = geometry.rowOffset(Math.max(1, Math.min(chart.y, MAX_RENDER_ROWS)));
  const left = baseLeft + (drag?.dx ?? 0);
  const top = baseTop + (drag?.dy ?? 0);
  const w = Math.max(120, drag && drag.mode === 'resize' ? drag.origW + drag.dx : chart.width);
  const h = Math.max(90, drag && drag.mode === 'resize' ? drag.origH + drag.dy : chart.height);
  const pad = { top: 24, right: 10, bottom: 26, left: 34 };
  const plotW = w - pad.left - pad.right;
  const plotH = h - pad.top - pad.bottom;

  const beginDrag = (e: React.PointerEvent, mode: 'move' | 'resize') => {
    e.stopPropagation();
    e.preventDefault();
    onSelect?.();
    setDrag({ mode, startX: e.clientX, startY: e.clientY, origX: chart.x, origY: chart.y, origW: chart.width, origH: chart.height, dx: 0, dy: 0 });
  };

  const onMove = (ev: PointerEvent) => {
    setDrag((d) => (d ? { ...d, dx: ev.clientX - d.startX, dy: ev.clientY - d.startY } : d));
  };

  const endDrag = () => {
    const d = dragRef.current;
    if (d) {
      const st = useWorkbookStore.getState();
      if (d.mode === 'move') {
        // Snap the dragged pixel position back onto the column/row grid.
        let targetX = chart.x;
        let best = Infinity;
        for (let c = 0; c <= MAX_RENDER_COLS; c++) {
          const dist = Math.abs(geometry.colOffset(c) - (baseLeft + d.dx));
          if (dist < best) {
            best = dist;
            targetX = c;
          }
        }
        let targetY = chart.y;
        best = Infinity;
        for (let r = 1; r <= MAX_RENDER_ROWS; r++) {
          const dist = Math.abs(geometry.rowOffset(r) - (baseTop + d.dy));
          if (dist < best) {
            best = dist;
            targetY = r;
          }
        }
        if (targetX !== d.origX || targetY !== d.origY) {
          st.applyCommand(new UpdateChartCommand('Move chart', sheet.id, chart.id, { x: targetX, y: targetY }));
        }
      } else {
        const nw = Math.max(120, Math.round(d.origW + d.dx));
        const nh = Math.max(90, Math.round(d.origH + d.dy));
        if (nw !== d.origW || nh !== d.origH) {
          st.applyCommand(new UpdateChartCommand('Resize chart', sheet.id, chart.id, { width: nw, height: nh }));
        }
      }
    }
    setDrag(null);
  };

  React.useEffect(() => {
    if (!drag) return;
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', endDrag);
    return () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', endDrag);
    };
  }, [drag]);

  let max = 0;
  let min = 0;
  if (data) {
    for (const s of data.series) for (const v of s.values) { if (v > max) max = v; if (v < min) min = v; }
  }
  if (max === min) max = min + 1;

  const zeroY = pad.top + (max / (max - min)) * plotH;

  return (
    <div
      className="absolute flex flex-col rounded border bg-[rgb(var(--cell))] shadow"
      data-chart-id={chart.id}
      style={{ left, top, width: w, height: h, zIndex: 25, borderColor: 'rgb(var(--line-strong))' }}
      onPointerDown={(e) => {
        // Charts sit above the cells, so selecting one must not also select a cell.
        e.stopPropagation();
        onSelect?.();
        if (e.button === 0) beginDrag(e, 'move');
      }}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
        onSelect?.();
        showContextMenu(
          [
            { label: 'Delete Chart', danger: true, onClick: () => {
              const st = useWorkbookStore.getState();
              st.setSelectedChartId(null);
              st.applyCommand(new RemoveChartCommand('Delete chart', sheet.id, chart.id));
            } }
          ],
          e.clientX,
          e.clientY
        );
      }}
    >
      {selected && (
        <div
          className="pointer-events-none absolute"
          style={{
            inset: -2,
            border: '2px solid rgb(var(--accent))',
            borderRadius: 3,
            zIndex: 1
          }}
        >
          {[
            { left: -4, top: -4 },
            { right: -4, top: -4 },
            { left: -4, bottom: -4 },
            { right: -4, bottom: -4 }
          ].map((pos, i) => (
            <span
              key={i}
              className="absolute h-[6px] w-[6px] rounded-[1px]"
              style={{ ...pos, background: 'rgb(var(--accent))', border: '1px solid rgb(var(--cell))' }}
            />
          ))}
          <span
            className="absolute h-[7px] w-[7px] rounded-[1px]"
            style={{ right: -5, bottom: -5, background: 'rgb(var(--accent))', border: '1px solid rgb(var(--cell))', cursor: 'nwse-resize', pointerEvents: 'auto' }}
            onPointerDown={(e) => beginDrag(e, 'resize')}
          />
        </div>
      )}
      <div className="truncate px-1 text-center text-[11px] font-semibold">{chart.title ?? 'Chart'}</div>
      <svg width={w} height={h - 18} className="block">
        {data && data.series[0] ? (
          chart.type === 'pie' ? (
            <PiePlot data={data} w={w} h={h - 18} colors={colors} />
          ) : chart.type === 'line' ? (
            <LinePlot data={data} plotW={plotW} plotH={plotH} pad={pad} max={max} min={min} colors={colors} />
          ) : (
            <BarPlot
              data={data}
              plotW={plotW}
              plotH={plotH}
              pad={pad}
              max={max}
              min={min}
              zeroY={zeroY}
              horizontal={chart.type === 'bar'}
              colors={colors}
            />
          )
        ) : (
          <text x={w / 2} y={(h - 18) / 2} textAnchor="middle" fontSize={11} fill="rgb(var(--text-muted))">
            No data in range {chart.dataRange ?? ''}
          </text>
        )}
      </svg>
    </div>
  );
}

function BarPlot({
  data,
  plotW,
  plotH,
  pad,
  max,
  min,
  zeroY,
  horizontal,
  colors
}: {
  data: ChartData;
  plotW: number;
  plotH: number;
  pad: { top: number; right: number; bottom: number; left: number };
  max: number;
  min: number;
  zeroY: number;
  horizontal: boolean;
  colors: string[];
}) {
  const n = data.categories.length || 1;
  const group = (horizontal ? plotH : plotW) / n;
  const bars = data.series.length || 1;
  const barW = Math.max(1, (group * 0.7) / bars);
  const scale = (v: number) => (horizontal ? (v / max) * plotW : ((max - v) / (max - min)) * plotH);
  return (
    <g>
      <line x1={pad.left} y1={zeroY} x2={pad.left + plotW} y2={zeroY} stroke="rgb(var(--line-strong))" />
      {data.series.map((s, si) =>
        s.values.map((v, i) => {
          const color = colors[si % colors.length];
          if (horizontal) {
            const len = Math.abs(scale(v));
            return (
              <rect
                key={`${si}_${i}`}
                x={pad.left}
                y={pad.top + i * group + (group - barW * bars) / 2 + si * barW}
                width={len}
                height={barW}
                fill={color}
              />
            );
          }
          const hgt = Math.abs(scale(v) - scale(0));
          return (
            <rect
              key={`${si}_${i}`}
              x={pad.left + i * group + (group - barW * bars) / 2 + si * barW}
              y={scale(v)}
              width={barW}
              height={Math.max(1, hgt)}
              fill={color}
            />
          );
        })
      )}
      {data.categories.map((c, i) => (
        <text
          key={`lbl_${i}`}
          x={horizontal ? pad.left - 4 : pad.left + i * group + group / 2}
          y={horizontal ? pad.top + i * group + group / 2 + 3 : pad.top + plotH + 12}
          textAnchor={horizontal ? 'end' : 'middle'}
          fontSize={9}
          fill="rgb(var(--text-muted))"
        >
          {c.length > 8 ? c.slice(0, 8) + '…' : c}
        </text>
      ))}
    </g>
  );
}

function LinePlot({
  data,
  plotW,
  plotH,
  pad,
  max,
  min,
  colors
}: {
  data: ChartData;
  plotW: number;
  plotH: number;
  pad: { top: number; right: number; bottom: number; left: number };
  max: number;
  min: number;
  colors: string[];
}) {
  const n = Math.max(1, data.categories.length - 1);
  return (
    <g>
      {data.series.map((s, si) => (
        <polyline
          key={si}
          points={s.values
            .map((v, i) => `${pad.left + (plotW / n) * i + plotW / n / 2},${pad.top + ((max - v) / (max - min)) * plotH}`)
            .join(' ')}
          fill="none"
          stroke={colors[si % colors.length]}
          strokeWidth={1.5}
        />
      ))}
    </g>
  );
}

function PiePlot({ data, w, h, colors }: { data: ChartData; w: number; h: number; colors: string[] }) {
  const values = data.series[0].values;
  const total = values.reduce((a, b) => a + Math.abs(b), 0) || 1;
  const cx = w / 2;
  const cy = h / 2;
  const r = Math.max(8, Math.min(w, h) / 2 - 14);
  let angle = -Math.PI / 2;
  return (
    <g>
      {values.map((v, i) => {
        const slice = (Math.abs(v) / total) * Math.PI * 2;
        const x1 = cx + r * Math.cos(angle);
        const y1 = cy + r * Math.sin(angle);
        angle += slice;
        const x2 = cx + r * Math.cos(angle);
        const y2 = cy + r * Math.sin(angle);
        const large = slice > Math.PI ? 1 : 0;
        return (
          <path
            key={i}
            d={`M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`}
            fill={colors[i % colors.length]}
            stroke="rgb(var(--cell))"
            strokeWidth={1}
          />
        );
      })}
    </g>
  );
}

function TableBands({ table, geometry }: { table: TableDef; geometry: SheetGeometry }) {
  const p = parseRange(table.range);
  if (!p) return null;
  if ((p.eRow - p.sRow + 1) * (p.eCol - p.sCol + 1) > 20_000) return null;
  const left = geometry.colOffset(p.sCol);
  const top = geometry.rowOffset(p.sRow);
  const width = geometry.colOffset(p.eCol + 1) - left;
  const headerH = table.headerRow ? geometry.rowOffset(p.sRow + 1) - top : 0;
  const totalH = geometry.rowOffset(p.eRow + 1) - top;
  const header = table.style.startsWith('TableStyleDark') || table.style === 'TableStyleMedium2' || table.style === 'TableStyleMedium9';
  return (
    <div
      className="absolute"
      style={{ left, top, width, height: totalH, zIndex: -1, pointerEvents: 'none' }}
    >
      {table.headerRow && (
        <div
          className="absolute inset-x-0 top-0"
          style={{ height: headerH, background: header ? 'rgb(var(--accent))' : 'rgb(var(--accent) / 0.16)' }}
        />
      )}
      {Array.from({ length: p.eRow - p.sRow + 1 }, (_, i) => {
        if (table.headerRow && i === 0) return null;
        const r = p.sRow + i;
        const y = geometry.rowOffset(r) - top;
        const h = geometry.rowOffset(r + 1) - geometry.rowOffset(r);
        if (i % 2 === 0) {
          return <div key={i} className="absolute inset-x-0" style={{ top: y, height: h, background: 'rgb(var(--accent) / 0.05)' }} />;
        }
        return null;
      })}
      <div className="absolute inset-0" style={{ border: `1px solid ${header ? 'rgb(var(--accent))' : 'rgb(var(--line-strong))'}` }} />
      {table.showFilter && table.headerRow && (
        <div
          className="absolute flex items-center justify-end pr-1"
          style={{ top: 0, height: headerH, right: 2 }}
        >
          {Array.from({ length: p.eCol - p.sCol + 1 }, (_, i) => (
            <svg key={i} width="7" height="7" viewBox="0 0 10 10" fill="rgb(var(--accent))">
              <path d="M0 0 H10 L5 6 Z" />
            </svg>
          ))}
        </div>
      )}
    </div>
  );
}

export default Grid;

