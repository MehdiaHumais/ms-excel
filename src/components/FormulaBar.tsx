import React, { useState } from 'react';
import { useWorkbookStore } from '@/store/useWorkbookStore';
import { useUiStore } from '@/store/useUiStore';
import { cellToA1, colToLabel } from '@/engine/model/address';

export function FormulaBar() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const editing = useWorkbookStore((s) => s.editing);
  const commitEdit = useWorkbookStore((s) => s.commitEdit);
  const startEdit = useWorkbookStore((s) => s.startEdit);
  const cancelEdit = useWorkbookStore((s) => s.cancelEdit);
  const openDialog = useUiStore((s) => s.openDialog);

  const [draft, setDraft] = useState<string>('');

  const sheet = workbook && activeSheetId ? workbook.sheets.find((s) => s.id === activeSheetId) : undefined;
  const activeRef = selection
    ? cellToA1(selection.col, selection.row)
    : '';
  const activeCell = activeRef && sheet ? sheet.cells[activeRef] : undefined;

  const isRange = selection.row !== selection.endRow || selection.col !== selection.endCol;
  const nameBox = isRange
    ? `${colToLabel(Math.min(selection.col, selection.endCol))}${Math.min(selection.row, selection.endRow)}:${colToLabel(Math.max(selection.col, selection.endCol))}${Math.max(selection.row, selection.endRow)}`
    : activeRef;

  const valueText = editing
    ? editing.text
    : activeCell?.formula != null
      ? '=' + activeCell.formula
      : activeCell?.value != null
        ? String(activeCell.value)
        : '';

  const handleFnBoxKey: React.KeyboardEventHandler<HTMLInputElement> = (e) => {
    if (e.key === 'Enter') {
      commitEdit((e.target as HTMLInputElement).value);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitEdit((e.target as HTMLInputElement).value);
      useWorkbookStore.getState().moveActive(0, e.shiftKey ? -1 : 1);
    } else if (
      e.key === 'ArrowDown' ||
      e.key === 'ArrowUp' ||
      e.key === 'ArrowLeft' ||
      e.key === 'ArrowRight'
    ) {
      // Commit the edit and walk the active cell, matching the grid editor.
      e.preventDefault();
      const [dRow, dCol] =
        e.key === 'ArrowDown'
          ? [1, 0]
          : e.key === 'ArrowUp'
            ? [-1, 0]
            : e.key === 'ArrowLeft'
              ? [0, -1]
              : [0, 1];
      commitEdit((e.target as HTMLInputElement).value);
      useWorkbookStore.getState().moveActive(dRow, dCol);
    } else if (e.key === 'Escape') {
      cancelEdit();
    } else {
      setDraft((e.target as HTMLInputElement).value);
    }
  };

  return (
    <div
      className="flex items-center gap-2 border-b px-2 py-[6px]"
      style={{ borderColor: 'rgb(var(--line))', background: 'rgb(var(--grid-header))' }}
    >
      {/* Name box */}
      <input
        className="h-7 w-[120px] rounded border px-2 text-[12px] focus:outline-none"
        style={{ borderColor: 'rgb(var(--line-strong))', background: 'rgb(var(--cell))' }}
        value={nameBox}
        readOnly
        title="Name box"
      />

      <div className="flex items-center justify-center">
        <button
          className="rounded px-2 py-0.5 font-semibold text-[13px] opacity-70 hover:opacity-100"
          title="Insert function (fx)"
          onClick={() => openDialog('insertFunction')}
        >
          fx
        </button>
      </div>

      {/* Formula input */}
      <input
        className="h-7 flex-1 rounded border px-2 text-[13px] focus:outline-none"
        style={{ borderColor: 'rgb(var(--line-strong))', background: 'rgb(var(--cell))' }}
        value={editing ? editing.text : valueText}
        placeholder=""
        onFocus={(e) => {
          if (!editing) {
            e.currentTarget.value = valueText;
            startEdit();
          }
        }}
        onChange={(e) => commitEditPreview(e.target.value)}
        onKeyDown={handleFnBoxKey}
      />
    </div>
  );

  function commitEditPreview(text: string) {
    // While typing in the formula bar, mirror to the editing overlay.
    const st = useWorkbookStore.getState();
    if (st.editing) {
      st.commitEdit(text);
      useWorkbookStore.setState({ editing: { row: st.selection.row, col: st.selection.col, text } });
    } else {
      startEdit();
      useWorkbookStore.setState({ editing: { row: st.selection.row, col: st.selection.col, text } });
    }
  }
}

export default FormulaBar;