import React, { useEffect, useState } from 'react';
import { useWorkbookStore } from '@/store/useWorkbookStore';
import { useSettingsStore } from '@/store/useSettingsStore';
import { getSheet, getCell } from '@/engine/model/workbook';

export function StatusBar() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const isProtected = useWorkbookStore(
    (s) => !!(s.activeSheetId && s.workbook && getSheet(s.workbook, s.activeSheetId)?.protected)
  );
  const selection = useWorkbookStore((s) => s.selection);
  const dirty = useWorkbookStore((s) => s.dirty);
  const revision = useWorkbookStore((s) => s.revision);
  const showStatusBar = useSettingsStore((s) => s.showStatusBar);

  const [mode, setMode] = useState('ready');

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') setMode('undo');
      else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') setMode('redo');
      else if (e.key === 'Escape') setMode('ready');
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  // Hooks must be called before any early return, otherwise toggling the
  // status bar changes the hook count and React tears down the whole tree.
  const sum = useSelectionSum(activeSheetId ?? '', selection, revision);

  if (!showStatusBar || !workbook || !activeSheetId) return null;
  const sheet = getSheet(workbook, activeSheetId);
  if (!sheet) return null;

  const { row, col, endRow, endCol } = selection;
  const isRange = row !== endRow || col !== endCol;
  const label = `R${row} C${col + 1}` + (isRange ? ` / R${endRow} C${endCol + 1}` : '');
  const count = isRange ? (Math.abs(endRow - row) + 1) * (Math.abs(endCol - col) + 1) : 1;

  const fmt = (n: number) =>
    Number.isInteger(n) ? String(n) : n.toFixed(2);

  return (
    <div
      className="flex h-[26px] items-center gap-4 px-3 text-[11px] text-muted"
      style={{ background: 'rgb(var(--grid-header))', borderTop: '1px solid rgb(var(--line))' }}
    >
      <span className="w-[74px]">{mode === 'undo' ? 'Undo' : mode === 'redo' ? 'Redo' : 'Ready'}</span>
      {isProtected && (
        <span className="rounded bg-[rgb(var(--accent-soft))] px-1.5 py-0.5 font-medium" title="This sheet is protected">
          Protected
        </span>
      )}
      <span>{label}</span>
      {isRange && <span>Count: {count}</span>}
      {isRange && sum != null && <span>Sum: {fmt(sum)}</span>}
      <div className="flex-1" />
      <span>{dirty ? 'Unsaved' : 'Saved'}</span>
      <span className="rounded px-1.5 py-0.5 ring-1 ring-inset ring-black/10 dark:ring-white/15">
        Normal
      </span>
    </div>
  );
}

function useSelectionSum(sheetId: string, sel: { row: number; col: number; endRow: number; endCol: number }, revision: number): number | null {
  const workbook = useWorkbookStore((s) => s.workbook);
  const engine = useWorkbookStore((s) => s.engine);
  if (!workbook) return null;
  const sheet = getSheet(workbook, sheetId);
  if (!sheet) return null;
  let sum = 0;
  let hasNum = false;
  const sRow = Math.min(sel.row, sel.endRow);
  const eRow = Math.max(sel.row, sel.endRow);
  const sCol = Math.min(sel.col, sel.endCol);
  const eCol = Math.max(sel.col, sel.endCol);
  for (let r = sRow; r <= eRow; r++) {
    for (let c = sCol; c <= eCol; c++) {
      const cell = getCell(sheet, c, r);
      if (!cell) continue;
      let v = cell.value;
      if (typeof v === 'number') hasNum = true;
      else if (cell.formula) v = engine.compute(sheetId, r, c);
      if (typeof v === 'number') {
        hasNum = true;
        sum += v;
      }
    }
  }
  void revision;
  return hasNum ? sum : null;
}

export default StatusBar;