import React from 'react';
import { useUiStore } from '@/store/useUiStore';
import { useWorkbookStore, downloadBlob } from '@/store/useWorkbookStore';
import { useSettingsStore } from '@/store/useSettingsStore';
import { X } from 'lucide-react';
import { serializeWorkbook } from '@/engine/io/serialize';
import { workbookToCsvData } from '@/engine/io/csv';
import { cellToA1 } from '@/engine/model/address';
import { DEFAULT_MARGINS, PAPER_SIZES } from '@/engine/model/pageSetup';
import { getSheet } from '@/engine/model/workbook';
import { SortRangeCommand, SetNamedRangeCommand, DeleteNamedRangeCommand, SetValidationCommand, AddTableCommand, SetNoteCommand, EditCellsCommand, RemoveDuplicatesCommand, SetSheetProtectedCommand, SetPageSetupCommand, SetPrintAreaCommand } from '@/engine/commands/commands';
import { DataValidationOperator, DataValidationRule, DataValidationType } from '@/engine/model/types';

export function DialogHost() {
  const dialog = useUiStore((s) => s.dialog);
  const closeDialog = useUiStore((s) => s.closeDialog);

  // Excel closes any modal on Escape. Capture phase so the modal wins over
  // the grid's own Escape handling (chart deselect).
  React.useEffect(() => {
    if (!dialog) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      e.stopPropagation();
      e.preventDefault();
      closeDialog();
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [dialog, closeDialog]);

  if (!dialog) return null;
  return (
    <div
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/30 p-4"
      onPointerDown={() => closeDialog()}
    >
      <div
        className="w-full max-w-lg rounded-lg border bg-cell p-5 shadow-xl"
        style={{ borderColor: 'rgb(var(--line-strong))' }}
        onPointerDown={(e) => e.stopPropagation()}
      >
        {dialog === 'formatCells' && <FormatCellsDialog onClose={closeDialog} />}
        {dialog === 'open' && <OpenDialog onClose={closeDialog} />}
        {dialog === 'saveAs' && <SaveAsDialog onClose={closeDialog} />}
        {dialog === 'about' && <AboutDialog onClose={closeDialog} />}
        {dialog === 'keyboardShortcuts' && <KeyboardShortcutsDialog onClose={closeDialog} />}
        {dialog === 'ai' && <AiDialog onClose={closeDialog} />}
        {dialog === 'findReplace' && <FindReplaceDialog onClose={closeDialog} />}
        {dialog === 'sort' && <SortDialog onClose={closeDialog} />}
        {dialog === 'removeDuplicates' && <RemoveDuplicatesDialog onClose={closeDialog} />}
        {dialog === 'namedRanges' && <NamedRangesDialog onClose={closeDialog} />}
        {dialog === 'defineName' && <DefineNameDialog onClose={closeDialog} />}
        {dialog === 'insertFunction' && <InsertFunctionDialog onClose={closeDialog} />}
        {dialog === 'settings' && <SettingsDialog onClose={closeDialog} />}
        {dialog === 'dataValidation' && <DataValidationDialog onClose={closeDialog} />}
      {dialog === 'insertTable' && <InsertTableDialog onClose={closeDialog} />}
      {dialog === 'insertLink' && <InsertLinkDialog onClose={closeDialog} />}
        {dialog === 'insertComment' && <InsertCommentDialog onClose={closeDialog} />}
        {dialog === 'protectSheet' && <ProtectSheetDialog onClose={closeDialog} />}
        {dialog === 'unprotectSheet' && <UnprotectSheetDialog onClose={closeDialog} />}
        {dialog === 'pageSetup' && <PageSetupDialog onClose={closeDialog} />}
      </div>
    </div>
  );
}

function DialogHeader({ title, onClose }: { title: string; onClose(): void }) {
  return (
    <div className="mb-4 flex items-center justify-between">
      <h2 className="text-base font-semibold">{title}</h2>
      <button onClick={onClose} className="rounded p-1 hover:bg-black/5 dark:hover:bg-white/10">
        <X size={16} />
      </button>
    </div>
  );
}

function FormatCellsDialog({ onClose }: { onClose(): void }) {
  const formatSelection = useWorkbookStore((s) => s.formatSelection);
  const [numberFormat, setNumberFormat] = React.useState('General');
  const [bold, setBold] = React.useState(false);
  const [backgroundColor, setBackgroundColor] = React.useState('');
  const [align, setAlign] = React.useState<'left' | 'center' | 'right' | undefined>(undefined);

  const apply = () => {
    formatSelection({
      numberFormat,
      bold,
      backgroundColor: backgroundColor || undefined,
      horizontalAlign: align
    });
    onClose();
  };

  return (
    <>
      <DialogHeader title="Format Cells" onClose={onClose} />
      <div className="space-y-3">
        <label className="block">
          <span className="mb-1 block text-[12px] text-muted">Number format</span>
          <select
            className="w-full rounded border px-2 py-1.5 text-[13px]"
            style={{ borderColor: 'rgb(var(--line-strong))' }}
            value={numberFormat}
            onChange={(e) => setNumberFormat(e.target.value)}
          >
            <option value="General">General</option>
            <option value="0">Number</option>
            <option value="#,##0">Number with thousands</option>
            <option value="0.00">Number (2 decimals)</option>
            <option value="$#,##0.00">Currency</option>
            <option value="0%">Percent</option>
            <option value="0.00%">Percent (2 decimals)</option>
            <option value="m/d/yyyy">Date</option>
            <option value="yyyy-mm-dd">Date (ISO)</option>
            <option value="h:mm:ss AM/PM">Time</option>
            <option value="0.00E+00">Scientific</option>
            <option value="yyyy-mm-dd hh:mm">Date &amp; time</option>
          </select>
        </label>
        <div className="flex items-center gap-2">
          <label className="flex items-center gap-2 text-[13px]">
            <input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} /> Bold
          </label>
        </div>
        <div className="flex items-center gap-3">
          <label className="flex items-center gap-2 text-[13px]">
            Fill color
            <input type="color" value={backgroundColor || '#ffffff'} onChange={(e) => setBackgroundColor(e.target.value)} />
          </label>
        </div>
        <div className="flex items-center gap-3 text-[13px]">
          <span>Align:</span>
          {(['left', 'center', 'right'] as const).map((a) => (
            <label key={a} className="flex items-center gap-1">
              <input type="radio" name="align" checked={align === a} onChange={() => setAlign(a)} /> {a}
            </label>
          ))}
        </div>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={apply}
        >
          Apply
        </button>
      </div>
    </>
  );
}

function OpenDialog({ onClose }: { onClose(): void }) {
  const setView = useUiStore((s) => s.setView);
  const toast = useUiStore((s) => s.toast);
  const exportCsv = () => {
    const wb = useWorkbookStore.getState().workbook;
    if (!wb) return;
    const rows = workbookToCsvData(wb);
    const csv = rows.map((r) => r.map(escapeCsv).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = wb.metadata.title + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  };
  const openViaTauri = async () => {
    try {
      const { pickOpenFile } = await import('@/platform/bridge');
      const path = await pickOpenFile([
        { name: 'BritSheet Workbook', extensions: ['britsheet'] },
        { name: 'All files', extensions: ['*'] }
      ]);
      if (!path) return;
      const { loadWorkbookFromFile } = await import('@/platform/loaders');
      await loadWorkbookFromFile(path);
      setView('workspace');
      toast('success', `Opened ${path.split(/[\\/]/).pop()}`);
    } catch (err) {
      toast('error', `Could not open file: ${(err as Error).message}`);
    }
    onClose();
  };
  return (
    <>
      <DialogHeader title="Open / Import" onClose={onClose} />
      <p className="mb-3 text-[13px] text-muted">
        Opens native <code>.britsheet</code> files, or imports CSV / JSON. In the desktop build this uses Tauri&#39;s
        native file dialog.
      </p>
      <div className="flex gap-2">
        <button
          className="flex-1 rounded border px-3 py-2 text-[13px]"
          onClick={exportCsv}
        >
          Export current sheet to CSV
        </button>
        <button className="flex-1 rounded border px-3 py-2 text-[13px]" onClick={openViaTauri}>
          Choose file…
        </button>
      </div>
      <p className="mt-3 text-[12px] text-muted">
        XLSX import (except in the Tauri desktop build which loads it with native binaries) uses the JS exceljs reader.
      </p>
    </>
  );
}

function SaveAsDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const workbook = useWorkbookStore((s) => s.workbook);
  const [desktop, setDesktop] = React.useState<boolean | null>(null);

  // The native dialog only exists in the Tauri build; never offer a dead button.
  React.useEffect(() => {
    let alive = true;
    void import('@/platform/bridge').then((m) => alive && setDesktop(m.isTauri()));
    return () => {
      alive = false;
    };
  }, []);

  const finish = async (path: string) => {
    useWorkbookStore.getState().markSaved(path);
    const name = path.split(/[\\/]/).pop() ?? path;
    useUiStore.getState().addRecent({ path, name, modified: Date.now() });
    const bridge = await import('@/platform/bridge');
    void bridge.addRecent(path).catch(() => {});
    toast('success', `Saved ${name}`);
  };

  const saveViaTauri = async () => {
    if (!workbook) return;
    try {
      const { pickSaveFile, saveWorkbookFile } = await import('@/platform/bridge');
      const json = serializeWorkbook(workbook);
      const path = await pickSaveFile(
        jsonSafeName(workbook.metadata.title) + '.britsheet',
        [{ name: 'BritSheet Workbook', extensions: ['britsheet'] }]
      );
      if (!path) return;
      await saveWorkbookFile(path, json);
      await finish(path);
    } catch (err) {
      toast('error', `Could not save file: ${(err as Error).message}`);
    }
    onClose();
  };
  const saveBrowser = async () => {
    if (!workbook) return;
    const json = serializeWorkbook(workbook);
    const name = jsonSafeName(workbook.metadata.title) + '.britsheet';
    downloadBlob(json, name, 'application/json');
    useWorkbookStore.getState().markSaved();
    toast('success', `Downloaded ${name}`);
    onClose();
  };
  return (
    <>
      <DialogHeader title="Save As" onClose={onClose} />
      <p className="mb-4 text-[13px] text-muted">
        Saves as the native <code>.britsheet</code> format (gzip JSON with magic header)
        {desktop === true ? " via BritSheet's own file dialog and an atomic write." : ' as a download.'}
      </p>
      <div className="flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        {desktop === false && (
          <button className="rounded border px-3 py-1.5 text-[13px]" onClick={saveBrowser}>
            Download file
          </button>
        )}
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white disabled:opacity-50"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={saveViaTauri}
          disabled={desktop === null}
        >
          Save
        </button>
      </div>
    </>
  );
}

function DataValidationDialog({ onClose }: { onClose(): void }) {
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const toast = useUiStore((s) => s.toast);

  const sRow = Math.min(selection.row, selection.endRow);
  const sCol = Math.min(selection.col, selection.endCol);
  const eRow = Math.max(selection.row, selection.endRow);
  const eCol = Math.max(selection.col, selection.endCol);
  const rangeLabel =
    sRow === eRow && sCol === eCol
      ? cellToA1(sCol, sRow)
      : `${cellToA1(sCol, sRow)}:${cellToA1(eCol, eRow)}`;

  const [type, setType] = React.useState<DataValidationType>('list');
  const [operator, setOperator] = React.useState<DataValidationOperator>('between');
  const [source, setSource] = React.useState('');
  const [second, setSecond] = React.useState('');
  const [allowBlank, setAllowBlank] = React.useState(true);
  const [errorTitle, setErrorTitle] = React.useState('');
  const [errorMessage, setErrorMessage] = React.useState('');

  const needsSource = type === 'list' || type === 'custom';
  const needsSecond =
    operator === 'between' || operator === 'notBetween';

  const apply = () => {
    if (!activeSheetId) return;
    const rule: DataValidationRule = {
      type,
      operator: needsSource ? undefined : operator,
      formula1: source || undefined,
      formula2: needsSecond ? second || undefined : undefined,
      allowBlank,
      showErrorMessage: true,
      errorTitle: errorTitle || undefined,
      errorMessage: errorMessage || undefined
    };
    applyCommand(
      new SetValidationCommand('Data validation', activeSheetId, sRow, sCol, eRow, eCol, rule)
    );
    toast('success', `Validation applied to ${rangeLabel}`);
    onClose();
  };

  const clear = () => {
    if (!activeSheetId) return;
    applyCommand(
      new SetValidationCommand('Clear validation', activeSheetId, sRow, sCol, eRow, eCol, undefined)
    );
    toast('success', `Validation cleared for ${rangeLabel}`);
    onClose();
  };

  return (
    <>
      <DialogHeader title="Data Validation" onClose={onClose} />
      <div className="space-y-3">
        <label className="block">
          <span className="mb-1 block text-[12px] text-muted">Allow</span>
          <select
            className="w-full rounded border px-2 py-1.5 text-[13px]"
            style={{ borderColor: 'rgb(var(--line-strong))' }}
            value={type}
            onChange={(e) => setType(e.target.value as DataValidationType)}
          >
            <option value="list">List</option>
            <option value="whole">Whole number</option>
            <option value="decimal">Decimal</option>
            <option value="textLength">Text length</option>
            <option value="date">Date</option>
            <option value="time">Time</option>
            <option value="custom">Custom formula</option>
            <option value="none">Any value</option>
          </select>
        </label>

        {type !== 'none' && !needsSource && (
          <label className="block">
            <span className="mb-1 block text-[12px] text-muted">Data</span>
            <select
              className="w-full rounded border px-2 py-1.5 text-[13px]"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              value={operator}
              onChange={(e) => setOperator(e.target.value as DataValidationOperator)}
            >
              <option value="between">between</option>
              <option value="notBetween">not between</option>
              <option value="equal">equal to</option>
              <option value="notEqual">not equal to</option>
              <option value="greaterThan">greater than</option>
              <option value="lessThan">less than</option>
              <option value="greaterThanOrEqual">greater than or equal to</option>
              <option value="lessThanOrEqual">less than or equal to</option>
            </select>
          </label>
        )}

        {type !== 'none' && (
          <label className="block">
            <span className="mb-1 block text-[12px] text-muted">
              {type === 'list' ? 'Source (comma separated values)' : needsSource ? 'Formula' : 'Value'}
            </span>
            <input
              className="w-full rounded border px-2 py-1.5 text-[13px]"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              value={source}
              onChange={(e) => setSource(e.target.value)}
              placeholder={type === 'list' ? 'Yes, No, Maybe' : '0'}
            />
          </label>
        )}

        {type !== 'none' && needsSecond && (
          <label className="block">
            <span className="mb-1 block text-[12px] text-muted">Second value</span>
            <input
              className="w-full rounded border px-2 py-1.5 text-[13px]"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              value={second}
              onChange={(e) => setSecond(e.target.value)}
              placeholder="10"
            />
          </label>
        )}

        {type !== 'none' && (
          <>
            <label className="flex items-center gap-2 text-[13px]">
              <input type="checkbox" checked={allowBlank} onChange={(e) => setAllowBlank(e.target.checked)} />
              Ignore blank
            </label>
            <label className="block">
              <span className="mb-1 block text-[12px] text-muted">Error alert title</span>
              <input
                className="w-full rounded border px-2 py-1.5 text-[13px]"
                style={{ borderColor: 'rgb(var(--line-strong))' }}
                value={errorTitle}
                onChange={(e) => setErrorTitle(e.target.value)}
                placeholder="Invalid entry"
              />
            </label>
            <label className="block">
              <span className="mb-1 block text-[12px] text-muted">Error message</span>
              <input
                className="w-full rounded border px-2 py-1.5 text-[13px]"
                style={{ borderColor: 'rgb(var(--line-strong))' }}
                value={errorMessage}
                onChange={(e) => setErrorMessage(e.target.value)}
                placeholder="This value is not allowed."
              />
            </label>
          </>
        )}

        <p className="text-[12px] text-muted">Applies to {rangeLabel}</p>

        <div className="flex justify-end gap-2 pt-1">
          <button className="rounded border px-3 py-1.5 text-[13px]" onClick={clear}>
            Clear
          </button>
          <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
            Cancel
          </button>
          <button
            className="rounded px-3 py-1.5 text-[13px] text-white"
            style={{ background: 'rgb(var(--accent))' }}
            onClick={apply}
          >
            OK
          </button>
        </div>
      </div>
    </>
  );
}

const SHORTCUT_GROUPS: { caption: string; items: { keys: string; label: string }[] }[] = [
  {
    caption: 'Editing',
    items: [
      { keys: 'F2', label: 'Edit the active cell' },
      { keys: 'Enter', label: 'Commit the edit and move down' },
      { keys: 'Tab', label: 'Commit the edit and move right' },
      { keys: 'Shift+Tab', label: 'Commit the edit and move left' },
      { keys: 'Esc', label: 'Cancel the edit, or deselect a chart' },
      { keys: 'Delete', label: 'Clear cell contents, or delete a selected chart' },
      { keys: 'Backspace', label: 'Clear cell contents, or delete a selected chart' }
    ]
  },
  {
    caption: 'Navigation',
    items: [
      { keys: 'Arrow keys', label: 'Move the active cell' },
      { keys: 'Shift+Arrow keys', label: 'Extend the selection' }
    ]
  },
  {
    caption: 'Clipboard',
    items: [
      { keys: 'Ctrl+C', label: 'Copy' },
      { keys: 'Ctrl+X', label: 'Cut' },
      { keys: 'Ctrl+V', label: 'Paste' }
    ]
  },
  {
    caption: 'Workbook',
    items: [
      { keys: 'Ctrl+Z', label: 'Undo' },
      { keys: 'Ctrl+Y', label: 'Redo' },
      { keys: 'Ctrl+Shift+Z', label: 'Redo' },
      { keys: 'Ctrl+S', label: 'Save' },
      { keys: 'Ctrl+F', label: 'Find' },
      { keys: 'F9', label: 'Recalculate the workbook' },
      { keys: 'Ctrl+`', label: 'Show or hide formulas' }
    ]
  }
];

function KeyboardShortcutsDialog({ onClose }: { onClose(): void }) {
  const mod = navigator.platform.toLowerCase().includes('mac') ? 'Cmd' : 'Ctrl';
  return (
    <>
      <DialogHeader title="Keyboard Shortcuts" onClose={onClose} />
      <div className="max-h-[420px] overflow-auto pr-1">
        {SHORTCUT_GROUPS.map((g) => (
          <div key={g.caption} className="mb-3">
            <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-muted">{g.caption}</div>
            <div className="flex flex-col">
              {g.items.map((it) => (
                <div key={it.keys} className="flex items-center justify-between gap-4 border-b border-[rgb(var(--line))] py-1 last:border-b-0">
                  <span className="text-[12px]">{it.label}</span>
                  <kbd
                    className="shrink-0 rounded border border-[rgb(var(--line-strong))] bg-[rgb(var(--grid-header))] px-1.5 py-0.5 text-[11px] font-medium"
                    title="Shortcut"
                  >
                    {it.keys.replace(/Ctrl/g, mod)}
                  </kbd>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
      <div className="mt-4 flex justify-end">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Close
        </button>
      </div>
    </>
  );
}

function AboutDialog({ onClose }: { onClose(): void }) {
  return (
    <>
      <DialogHeader title="About BritSheet" onClose={onClose} />
      <div className="text-[13px] leading-relaxed">
        <p className="mb-2">
          <span className="font-semibold">BritSheet</span> v0.1.0 — a fast, native spreadsheet for Windows.
        </p>
        <p className="text-muted">
          Built with Tauri 2, React, TypeScript and Rust. Saves in the native <code>.britsheet</code> format,
          imports/exports XLSX and CSV, supports formulas, charts, tables, conditional formatting and more.
        </p>
      </div>
      <div className="mt-5 flex justify-end">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Close
        </button>
      </div>
    </>
  );
}

function AiDialog({ onClose }: { onClose(): void }) {
  const openSettings = () => {
    useUiStore.getState().closeDialog();
    useUiStore.getState().openDialog('settings');
  };
  return (
    <>
      <DialogHeader title="AI Assistant" onClose={onClose} />
      <p className="mb-4 text-[13px] text-muted">
        BritSheet&#39;s built-in copilot can help build formulas, analyze ranges and prepare data. Configure your provider in Settings.
      </p>
      <div className="flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Close
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={openSettings}
        >
          Configure provider
        </button>
      </div>
    </>
  );
}

function FindReplaceDialog({ onClose }: { onClose(): void }) {
  const [findText, setFindText] = React.useState('');
  const [replace, setReplace] = React.useState('');
  const [matchCase, setMatchCase] = React.useState(false);
  const [wholeCell, setWholeCell] = React.useState(false);
  const [results, setResults] = React.useState<Array<{ row: number; col: number; hit: string }> | null>(null);
  const toast = useUiStore((s) => s.toast);
  const findCells = useWorkbookStore((s) => s.findCells);
  const replaceCells = useWorkbookStore((s) => s.replaceCells);
  const runFind = () => {
    if (!findText) return;
    const res = findCells(findText, { matchCase, wholeCell });
    setResults(res);
    toast('info', `Found ${res.length} occurrence${res.length === 1 ? '' : 's'} of "${findText}".`);
  };
  const replaceAll = () => {
    if (!findText) return;
    const n = replaceCells(findText, replace, { matchCase, wholeCell });
    toast('info', `Replaced ${n} occurrence${n === 1 ? '' : 's'}.`);
    setResults(findCells(findText, { matchCase, wholeCell }));
  };
  return (
    <>
      <DialogHeader title="Find & Replace" onClose={onClose} />
      <div className="space-y-3 text-[13px]">
        <input
          className="w-full rounded border px-2 py-1.5"
          style={{ borderColor: 'rgb(var(--line-strong))' }}
          placeholder="Find…"
          value={findText}
          onChange={(e) => setFindText(e.target.value)}
        />
        <input
          className="w-full rounded border px-2 py-1.5"
          style={{ borderColor: 'rgb(var(--line-strong))' }}
          placeholder="Replace with…"
          value={replace}
          onChange={(e) => setReplace(e.target.value)}
        />
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          Match case
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={wholeCell} onChange={(e) => setWholeCell(e.target.checked)} />
          Whole cell
        </label>
        {results && (
          <div className="max-h-40 overflow-y-auto rounded border text-[12px]" style={{ borderColor: 'rgb(var(--line))' }}>
            {results.length === 0 ? (
              <p className="px-2 py-1 text-muted">No matches.</p>
            ) : (
              results.slice(0, 100).map((r) => (
                <div key={`${r.row}-${r.col}`} className="border-b px-2 py-1" style={{ borderColor: 'rgb(var(--line))' }}>
                  <span className="font-medium">Sheet1!{cellToA1(r.col, r.row)}</span> <span className="opacity-70">{r.hit}</span>
                </div>
              ))
            )}
          </div>
        )}
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Close
        </button>
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={replaceAll}>
          Replace All
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={runFind}
        >
          Find
        </button>
      </div>
    </>
  );
}

function SortDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [desc, setDesc] = React.useState(false);
  const [header, setHeader] = React.useState(true);
  const sRow = Math.min(selection.row, selection.endRow);
  const eRow = Math.max(selection.row, selection.endRow);
  const sCol = Math.min(selection.col, selection.endCol);
  const eCol = Math.max(selection.col, selection.endCol);
  const run = () => {
    if (!workbook || !activeSheetId) return;
    if (sRow === eRow || eRow - sRow < 1) {
      toast('info', 'Select a range with at least two data rows to sort.');
      onClose();
      return;
    }
    applyCommand(
      new SortRangeCommand('Sort range', activeSheetId, sRow, sCol, eRow, eCol, [{ col: sCol, desc }], { header })
    );
    toast('info', 'Sorted the selected range.');
    onClose();
  };
  return (
    <>
      <DialogHeader title="Sort" onClose={onClose} />
      <p className="text-[13px] text-muted">
        Sort selected range (col {cellToA1(sCol, 1)} → col {cellToA1(eCol, 1)}) by the first selected column.
      </p>
      <div className="mt-3 space-y-3 text-[13px]">
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={desc} onChange={(e) => setDesc(e.target.checked)} />
          Sort descending
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          My range has a header row
        </label>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={run}
        >
          Sort
        </button>
      </div>
    </>
  );
}

function RemoveDuplicatesDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [header, setHeader] = React.useState(true);
  const sRow = Math.min(selection.row, selection.endRow);
  const eRow = Math.max(selection.row, selection.endRow);
  const sCol = Math.min(selection.col, selection.endCol);
  const eCol = Math.max(selection.col, selection.endCol);
  const rangeLabel = `${cellToA1(sCol, sRow)}:${cellToA1(eCol, eRow)}`;
  const dataRows = eRow - sRow + 1 - (header ? 1 : 0);

  const run = () => {
    if (!workbook || !activeSheetId) return;
    if (dataRows < 2) {
      toast('info', 'Select a range with at least two data rows.');
      onClose();
      return;
    }
    const cmd = new RemoveDuplicatesCommand('Remove duplicates', activeSheetId, sRow, sCol, eRow, eCol, header);
    applyCommand(cmd);
    const removed = cmd.removedCount;
    toast('success', removed ? `Removed ${removed} duplicate ${removed === 1 ? 'row' : 'rows'}.` : 'No duplicate rows found.');
    onClose();
  };

  return (
    <>
      <DialogHeader title="Remove Duplicates" onClose={onClose} />
      <div className="mt-3 space-y-3 text-[13px]">
        <p className="text-muted">
          Remove duplicate rows from <span className="font-medium text-fg">{rangeLabel}</span>, comparing every
          selected column. The first occurrence of each row is kept.
        </p>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          My data has headers
        </label>
        <p className="text-muted">
          {dataRows} data {dataRows === 1 ? 'row' : 'rows'} in range.
        </p>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white disabled:opacity-50"
          style={{ background: 'rgb(var(--accent))' }}
          disabled={dataRows < 2}
          onClick={run}
        >
          OK
        </button>
      </div>
    </>
  );
}

function DefineNameDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [name, setName] = React.useState('');
  const sheetName = workbook?.sheets.find((s) => s.id === activeSheetId)?.name ?? 'Sheet1';
  const initialRef = React.useMemo(() => {
    const sRow = Math.min(selection.row, selection.endRow);
    const eRow = Math.max(selection.row, selection.endRow);
    const sCol = Math.min(selection.col, selection.endCol);
    const eCol = Math.max(selection.col, selection.endCol);
    return `${cellToA1(sCol, sRow)}:${cellToA1(eCol, eRow)}`;
  }, [selection]);
  const [refersTo, setRefersTo] = React.useState(initialRef);

  const submit = () => {
    if (!workbook || !activeSheetId) return;
    const clean = name.trim().replace(/^=/, '');
    if (!clean) {
      toast('info', 'Enter a name (letters, digits and underscores, not starting with a digit).');
      return;
    }
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(clean)) {
      toast('error', 'A name cannot start with a digit or contain spaces.');
      return;
    }
    const ref = refersTo.trim().replace(/^=/, '');
    if (!ref) {
      toast('info', 'Enter the range this name refers to, e.g. A1:B10.');
      return;
    }
    applyCommand(new SetNamedRangeCommand('Define name', clean, ref, sheetName));
    toast('success', `Defined name "${clean}" for ${ref}.`);
    onClose();
  };

  return (
    <>
      <DialogHeader title="Define Name" onClose={onClose} />
      <div className="mt-3 space-y-3 text-[13px]">
        <div>
          <label className="mb-1 block text-muted">Name:</label>
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="e.g. SalesTotal"
            className="w-full rounded border bg-surface px-2 py-1.5"
            style={{ borderColor: 'rgb(var(--line-strong))' }}
          />
        </div>
        <div>
          <label className="mb-1 block text-muted">Refers to:</label>
          <input
            value={refersTo}
            onChange={(e) => setRefersTo(e.target.value)}
            placeholder="=Sheet1!$A$1:$B$10"
            className="w-full rounded border bg-surface px-2 py-1.5 font-mono"
            style={{ borderColor: 'rgb(var(--line-strong))' }}
          />
        </div>
        <p className="text-muted">A name can be used in formulas, e.g. =SUM(SalesTotal).</p>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={submit}
        >
          OK
        </button>
      </div>
    </>
  );
}

function NamedRangesDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [name, setName] = React.useState('');
  const [range, setRange] = React.useState('');
  const namedRanges = workbook?.namedRanges ?? {};
  const sheetName = workbook?.sheets.find((s) => s.id === activeSheetId)?.name ?? 'Sheet1';
  const add = () => {
    const clean = name.trim().replace(/^=/, '');
    if (!clean || !range.trim()) {
      toast('info', 'Enter a name and a range (e.g. A1:B10).');
      return;
    }
    applyCommand(new SetNamedRangeCommand('Define named range', clean, range.trim(), sheetName));
    toast('info', `Named range "${clean}" defined.`);
    setName('');
    setRange('');
  };
  const useSelection = () => {
    const sel = useWorkbookStore.getState().selection ?? { row: 1, col: 0, endRow: 1, endCol: 1 };
    const sRow = Math.min(sel.row, sel.endRow);
    const eRow = Math.max(sel.row, sel.endRow);
    const sCol = Math.min(sel.col, sel.endCol);
    const eCol = Math.max(sel.col, sel.endCol);
    setRange(`${cellToA1(sCol, sRow)}:${cellToA1(eCol, eRow)}`);
    if (sRow === eRow && sCol === eCol) setRange(cellToA1(sCol, sRow));
  };
  return (
    <>
      <DialogHeader title="Named Ranges" onClose={onClose} />
      <div className="space-y-3 text-[13px]">
        <label className="block">
          <span className="mb-1 block text-[12px] text-muted">Range reference</span>
          <div className="flex gap-2">
            <input
              className="flex-1 rounded border px-2 py-1.5"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              placeholder="Name (e.g. Sales)"
              value={name}
              onChange={(e) => setName(e.target.value)}
            />
            <button className="rounded border px-2 py-1.5" onClick={useSelection}>
              Use selection
            </button>
          </div>
          <input
            className="mt-2 w-full rounded border px-2 py-1.5"
            style={{ borderColor: 'rgb(var(--line-strong))' }}
            placeholder="Range (e.g. A1:B10)"
            value={range}
            onChange={(e) => setRange(e.target.value)}
          />
        </label>
        <button
          className="rounded px-3 py-1.5 text-[12px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={add}
        >
          Add
        </button>
        {Object.keys(namedRanges).length === 0 ? (
          <p className="text-[13px] text-muted">No named ranges yet.</p>
        ) : (
          <ul className="space-y-1 text-[13px]">
            {Object.values(namedRanges).map((nr) => (
              <li
                key={nr.name}
                className="flex items-center justify-between gap-2 rounded border px-2 py-1"
                style={{ borderColor: 'rgb(var(--line))' }}
              >
                <button
                  className="flex-1 text-left hover:underline"
                  onClick={() => {
                    const targetId = workbook?.sheets.find((s) => s.name === nr.sheet)?.id;
                    const m = /^([A-Z]+)(\d+)$/.exec(nr.range.split(':')[0]);
                    if (m) {
                      const row = parseInt(m[2], 10);
                      const col = colLabelToIndex(m[1]);
                      if (targetId) useWorkbookStore.getState().setActiveSheet(targetId);
                      useWorkbookStore.getState().setSelection({ row, col, endRow: row, endCol: col });
                    }
                    onClose();
                  }}
                >
                  <span className="font-medium">{nr.name}</span> → {nr.sheet}!{nr.range}
                </button>
                <button
                  className="rounded px-1.5 text-[12px] text-red-500 hover:bg-red-500/10"
                  onClick={() => applyCommand(new DeleteNamedRangeCommand(nr.name))}
                  title="Delete named range"
                >
                  ✕
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Close
        </button>
      </div>
    </>
  );
}

function colLabelToIndex(label: string): number {
  let c = -1;
  for (const ch of label.toUpperCase()) c = (c + 1) * 26 + (ch.charCodeAt(0) - 65);
  return c;
}

export const FUNCTION_LIBRARY: Array<{ name: string; category: string; template: string; about: string }> = [
  { name: 'SUM', category: 'Math', template: 'SUM(number1,[number2],...)', about: 'Adds numbers' },
  { name: 'AVERAGE', category: 'Math', template: 'AVERAGE(number1,[number2],...)', about: 'Returns the average' },
  { name: 'COUNT', category: 'Math', template: 'COUNT(value1,[value2],...)', about: 'Counts numeric values' },
  { name: 'COUNTA', category: 'Math', template: 'COUNTA(value1,[value2],...)', about: 'Counts non-empty cells' },
  { name: 'COUNTIF', category: 'Math', template: 'COUNTIF(range,criteria)', about: 'Counts cells matching criteria' },
  { name: 'SUMIF', category: 'Math', template: 'SUMIF(range,criteria,[sum_range])', about: 'Sums cells matching criteria' },
  { name: 'MIN', category: 'Math', template: 'MIN(number1,[number2],...)', about: 'Returns the smallest number' },
  { name: 'MAX', category: 'Math', template: 'MAX(number1,[number2],...)', about: 'Returns the largest number' },
  { name: 'ROUND', category: 'Math', template: 'ROUND(number,num_digits)', about: 'Rounds a number' },
  { name: 'ROUNDUP', category: 'Math', template: 'ROUNDUP(number,num_digits)', about: 'Rounds a number up' },
  { name: 'ROUNDDOWN', category: 'Math', template: 'ROUNDDOWN(number,num_digits)', about: 'Rounds a number down' },
  { name: 'ABS', category: 'Math', template: 'ABS(number)', about: 'Absolute value' },
  { name: 'SIGN', category: 'Math', template: 'SIGN(number)', about: 'Sign of a number' },
  { name: 'INT', category: 'Math', template: 'INT(number)', about: 'Rounds down to integer' },
  { name: 'SQRT', category: 'Math', template: 'SQRT(number)', about: 'Square root' },
  { name: 'POWER', category: 'Math', template: 'POWER(number,power)', about: 'Raises a number to a power' },
  { name: 'MOD', category: 'Math', template: 'MOD(number,divisor)', about: 'Returns the remainder' },
  { name: 'IF', category: 'Logical', template: 'IF(logical_test,[value_if_true],[value_if_false])', about: 'Conditional evaluation' },
  { name: 'IFERROR', category: 'Logical', template: 'IFERROR(value,value_if_error)', about: 'Returns a fallback on error' },
  { name: 'IFS', category: 'Logical', template: 'IFS(condition1,value1,[condition2,value2],...)', about: 'Multiple conditions' },
  { name: 'AND', category: 'Logical', template: 'AND(logical1,[logical2],...)', about: 'True if all are true' },
  { name: 'OR', category: 'Logical', template: 'OR(logical1,[logical2],...)', about: 'True if any is true' },
  { name: 'NOT', category: 'Logical', template: 'NOT(logical)', about: 'Negates a logical' },
  { name: 'VLOOKUP', category: 'Lookup', template: 'VLOOKUP(lookup_value,table_array,col_index,[range_lookup])', about: 'Vertical lookup' },
  { name: 'INDEX', category: 'Lookup', template: 'INDEX(array,row_num,[column_num])', about: 'Value at a position' },
  { name: 'MATCH', category: 'Lookup', template: 'MATCH(lookup_value,lookup_array,[match_type])', about: 'Position of a value' },
  { name: 'CONCAT', category: 'Text', template: 'CONCAT(text1,[text2],...)', about: 'Concatenates text' },
  { name: 'LEFT', category: 'Text', template: 'LEFT(text,[num_chars])', about: 'Leftmost characters' },
  { name: 'RIGHT', category: 'Text', template: 'RIGHT(text,[num_chars])', about: 'Rightmost characters' },
  { name: 'MID', category: 'Text', template: 'MID(text,start_num,num_chars)', about: 'Middle characters' },
  { name: 'LEN', category: 'Text', template: 'LEN(text)', about: 'Length of text' },
  { name: 'LOWER', category: 'Text', template: 'LOWER(text)', about: 'Lowercase text' },
  { name: 'UPPER', category: 'Text', template: 'UPPER(text)', about: 'Uppercase text' },
  { name: 'TRIM', category: 'Text', template: 'TRIM(text)', about: 'Removes extra spaces' },
  { name: 'SUBSTITUTE', category: 'Text', template: 'SUBSTITUTE(text,old_text,new_text,[instance_num])', about: 'Replaces text' },
  { name: 'TODAY', category: 'Date', template: 'TODAY()', about: 'Current date' },
  { name: 'NOW', category: 'Date', template: 'NOW()', about: 'Current date and time' },
  { name: 'DATE', category: 'Date', template: 'DATE(year,month,day)', about: 'Builds a date' },
  { name: 'YEAR', category: 'Date', template: 'YEAR(date)', about: 'Year of a date' },
  { name: 'MONTH', category: 'Date', template: 'MONTH(date)', about: 'Month of a date' },
  { name: 'DAY', category: 'Date', template: 'DAY(date)', about: 'Day of a date' }
];

function InsertFunctionDialog({ onClose }: { onClose(): void }) {
  const [q, setQ] = React.useState('');
  const [selected, setSelected] = React.useState(FUNCTION_LIBRARY[0].name);
  const toast = useUiStore((s) => s.toast);
  const filtered = FUNCTION_LIBRARY.filter(
    (f) =>
      f.name.toLowerCase().includes(q.toLowerCase()) ||
      f.about.toLowerCase().includes(q.toLowerCase())
  );
  const insert = () => {
    const f = filtered.find((x) => x.name === selected) ?? filtered[0];
    if (!f) return;
    const open = f.template.indexOf('(');
    const base = f.template.slice(0, open);
    useWorkbookStore.getState().setCellRaw(useWorkbookStore.getState().selection.row, useWorkbookStore.getState().selection.col, '=' + base + '()');
    toast('success', `Inserted =${base}() — edit the arguments in the formula bar.`);
    onClose();
  };
  return (
    <>
      <DialogHeader title="Insert Function" onClose={onClose} />
      <div className="space-y-3 text-[13px]">
        <input
          className="w-full rounded border px-2 py-1.5"
          style={{ borderColor: 'rgb(var(--line-strong))' }}
          placeholder="Search functions…"
          value={q}
          onChange={(e) => setQ(e.target.value)}
          autoFocus
        />
        <div className="flex max-h-64 gap-2">
          <div className="max-h-64 flex-1 overflow-y-auto rounded border" style={{ borderColor: 'rgb(var(--line))' }}>
            {filtered.map((f) => (
              <button
                key={f.name}
                className={`flex w-full flex-col px-2 py-1 text-left hover:bg-black/5 dark:hover:bg-white/10 ${selected === f.name ? 'bg-black/5 dark:bg-white/10' : ''}`}
                onClick={() => setSelected(f.name)}
              >
                <span className="font-medium">{f.name}</span>
                <span className="text-[10px] opacity-60">{f.category}</span>
              </button>
            ))}
          </div>
          <div className="flex-1">
            <div className="rounded border p-2" style={{ borderColor: 'rgb(var(--line))' }}>
              {filtered.find((x) => x.name === selected) ? (
                <>
                  <p className="font-semibold">{selected}</p>
                  <p className="mt-1 font-mono text-[11px] opacity-80">
                    ={filtered.find((x) => x.name === selected)?.template}
                  </p>
                  <p className="mt-2 text-[12px] text-muted">{filtered.find((x) => x.name === selected)?.about}</p>
                </>
              ) : (
                <p className="text-muted">No function selected.</p>
              )}
            </div>
          </div>
        </div>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={insert}
        >
          Insert
        </button>
      </div>
    </>
  );
}

function SettingsDialog({ onClose }: { onClose(): void }) {
  const settings = useSettingsStore();
  return (
    <>
      <DialogHeader title="Settings" onClose={onClose} />
      <div className="space-y-4 text-[13px]">
        <label className="block">
          <span className="mb-1 block text-[12px] text-muted">Theme</span>
          <select
            className="w-full rounded border px-2 py-1.5"
            style={{ borderColor: 'rgb(var(--line-strong))' }}
            value={settings.theme}
            onChange={(e) => settings.setTheme(e.target.value as never)}
          >
            <option value="system">System</option>
            <option value="light">Light</option>
            <option value="dark">Dark</option>
          </select>
        </label>
        <label className="flex items-center justify-between">
          <span>Show gridlines</span>
          <input type="checkbox" checked={settings.showGridlines} onChange={(e) => settings.setShowGridlines(e.target.checked)} />
        </label>
        <label className="flex items-center justify-between">
          <span>Show formulas</span>
          <input type="checkbox" checked={settings.showFormulas} onChange={(e) => settings.setShowFormulas(e.target.checked)} />
        </label>
        <label className="flex items-center justify-between">
          <span>Autosave</span>
          <input type="checkbox" checked={settings.autosave} onChange={(e) => settings.setAutosave(e.target.checked)} />
        </label>
        <details className="rounded border p-2" style={{ borderColor: 'rgb(var(--line))' }}>
          <summary className="cursor-pointer">AI provider</summary>
          <div className="mt-2 space-y-2">
            <select
              className="w-full rounded border px-2 py-1.5"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              value={settings.ai.provider}
              onChange={(e) => settings.setAi({ provider: e.target.value as never })}
            >
              <option value="none">None</option>
              <option value="openai">OpenAI</option>
              <option value="anthropic">Anthropic</option>
              <option value="ollama">Ollama (local)</option>
              <option value="custom">Custom endpoint</option>
            </select>
            <input
              className="w-full rounded border px-2 py-1.5"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              placeholder="API key (stored locally)"
              type="password"
              value={settings.ai.apiKey}
              onChange={(e) => settings.setAi({ apiKey: e.target.value })}
            />
            <input
              className="w-full rounded border px-2 py-1.5"
              style={{ borderColor: 'rgb(var(--line-strong))' }}
              placeholder="Model"
              value={settings.ai.model}
              onChange={(e) => settings.setAi({ model: e.target.value })}
            />
          </div>
        </details>
      </div>
      <div className="mt-5 flex justify-end">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Done
        </button>
      </div>
    </>
  );
}

function escapeCsv(v: string): string {
  return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}

export function jsonSafeName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '_');
}

/* ----------------------------- Insert tab dialogs ----------------------------- */

const TABLE_STYLES = [
  'TableStyleMedium2',
  'TableStyleMedium4',
  'TableStyleMedium9',
  'TableStyleMedium16',
  'TableStyleLight1',
  'TableStyleLight9',
  'TableStyleDark1'
] as const;

function InsertTableDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const workbook = useWorkbookStore((s) => s.workbook);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [rows, setRows] = React.useState(5);
  const [cols, setCols] = React.useState(3);
  const [header, setHeader] = React.useState(true);
  const [style, setStyle] = React.useState<string>('TableStyleMedium9');
  const [filters, setFilters] = React.useState(true);

  const run = () => {
    if (!activeSheetId || !workbook) return;
    const nRows = Math.max(1, Math.min(200, Math.floor(rows) || 1));
    const nCols = Math.max(1, Math.min(60, Math.floor(cols) || 1));
    const r1 = Math.min(selection.row, selection.endRow);
    const c1 = Math.min(selection.col, selection.endCol);
    const range = `${cellToA1(c1, r1)}:${cellToA1(c1 + nCols - 1, r1 + nRows - 1)}`;
    const sheet = workbook.sheets.find((s) => s.id === activeSheetId);
    const existing = sheet?.tables.length ?? 0;
    applyCommand(
      new AddTableCommand('Insert table', activeSheetId, {
        name: `Table${existing + 1}`,
        range,
        headerRow: header,
        style,
        showFilter: filters
      } as never)
    );
    toast('info', `Inserted ${range}.`);
    onClose();
  };

  return (
    <>
      <DialogHeader title="Insert Table" onClose={onClose} />
      <p className="text-[13px] text-muted">Choose the size of the table and where its header row sits.</p>
      <div className="mt-4 grid grid-cols-2 gap-3 text-[13px]">
        <label className="flex flex-col gap-1">
          <span className="text-muted">Table columns</span>
          <input
            type="number"
            min={1}
            max={60}
            value={cols}
            onChange={(e) => setCols(Number(e.target.value))}
            className="rounded border px-2 py-1"
          />
        </label>
        <label className="flex flex-col gap-1">
          <span className="text-muted">Table rows</span>
          <input
            type="number"
            min={1}
            max={200}
            value={rows}
            onChange={(e) => setRows(Number(e.target.value))}
            className="rounded border px-2 py-1"
          />
        </label>
      </div>
      <div className="mt-3 space-y-2 text-[13px]">
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          My table has headers
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={filters} onChange={(e) => setFilters(e.target.checked)} />
          Show filter buttons
        </label>
        <label className="flex items-center gap-2">
          <span className="text-muted">Style</span>
          <select value={style} onChange={(e) => setStyle(e.target.value)} className="rounded border px-2 py-1">
            {TABLE_STYLES.map((s) => (
              <option key={s} value={s}>
                {s.replace('TableStyle', '')}
              </option>
            ))}
          </select>
        </label>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button className="rounded px-3 py-1.5 text-[13px] text-white" style={{ background: 'rgb(var(--accent))' }} onClick={run}>
          OK
        </button>
      </div>
    </>
  );
}

function InsertLinkDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [address, setAddress] = React.useState('https://');
  const [text, setText] = React.useState('');

  const run = () => {
    if (!activeSheetId) return;
    const url = address.trim();
    if (!url || url === 'https://') {
      toast('info', 'Enter an address for the link.');
      return;
    }
    const row = Math.min(selection.row, selection.endRow);
    const col = Math.min(selection.col, selection.endCol);
    const ref = cellToA1(col, row);
    applyCommand(
      new EditCellsCommand('Insert link', activeSheetId, {
        [ref]: {
          value: text.trim() || url,
          hyperlink: url,
          style: { color: '#0563c1', underline: true }
        }
      })
    );
    toast('info', `Linked ${ref} to ${url}.`);
    onClose();
  };

  return (
    <>
      <DialogHeader title="Insert Hyperlink" onClose={onClose} />
      <p className="text-[13px] text-muted">Links are stored on the cell and shown underlined in blue.</p>
      <div className="mt-4 space-y-3 text-[13px]">
        <label className="flex flex-col gap-1">
          <span className="text-muted">Text to display</span>
          <input
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder="Optional"
            className="rounded border px-2 py-1"
          />
        </label>
        <label className="flex flex-col gap-1">
          <span className="text-muted">Address</span>
          <input value={address} onChange={(e) => setAddress(e.target.value)} className="rounded border px-2 py-1" />
        </label>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button className="rounded px-3 py-1.5 text-[13px] text-white" style={{ background: 'rgb(var(--accent))' }} onClick={run}>
          OK
        </button>
      </div>
    </>
  );
}

function InsertCommentDialog({ onClose }: { onClose(): void }) {
  const toast = useUiStore((s) => s.toast);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const [note, setNote] = React.useState('');

  const run = () => {
    if (!activeSheetId) return;
    const text = note.trim();
    if (!text) {
      toast('info', 'Write the comment text first.');
      return;
    }
    const ref = cellToA1(
      Math.min(selection.col, selection.endCol),
      Math.min(selection.row, selection.endRow)
    );
    applyCommand(new SetNoteCommand(activeSheetId, ref, text));
    toast('info', `Comment added to ${ref}.`);
    onClose();
  };

  return (
    <>
      <DialogHeader title="New Comment" onClose={onClose} />
      <p className="text-[13px] text-muted">
        Comments show as a red corner marker and appear in the cell tooltip.
      </p>
      <textarea
        value={note}
        onChange={(e) => setNote(e.target.value)}
        rows={5}
        autoFocus
        placeholder="Type your comment"
        className="mt-3 w-full rounded border p-2 text-[13px]"
      />
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button className="rounded px-3 py-1.5 text-[13px] text-white" style={{ background: 'rgb(var(--accent))' }} onClick={run}>
          Insert
        </button>
      </div>
    </>
  );
}

function ProtectSheetDialog({ onClose }: { onClose(): void }) {
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const toast = useUiStore((s) => s.toast);

  const run = () => {
    if (!activeSheetId) return;
    applyCommand(new SetSheetProtectedCommand(activeSheetId, true));
    toast('success', 'Sheet protected');
    onClose();
  };

  return (
    <>
      <DialogHeader title="Protect Sheet" onClose={onClose} />
      <p className="text-[13px] text-muted">
        Protecting this sheet prevents changing its contents and formatting. You can unprotect it again from the
        Review tab.
      </p>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button className="rounded px-3 py-1.5 text-[13px] text-white" style={{ background: 'rgb(var(--accent))' }} onClick={run}>
          OK
        </button>
      </div>
    </>
  );
}

function UnprotectSheetDialog({ onClose }: { onClose(): void }) {
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const toast = useUiStore((s) => s.toast);

  const run = () => {
    if (!activeSheetId) return;
    applyCommand(new SetSheetProtectedCommand(activeSheetId, false));
    toast('success', 'Sheet unprotected');
    onClose();
  };

  return (
    <>
      <DialogHeader title="Unprotect Sheet" onClose={onClose} />
      <p className="text-[13px] text-muted">Allow changes to this sheet again.</p>
      <div className="mt-5 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button className="rounded px-3 py-1.5 text-[13px] text-white" style={{ background: 'rgb(var(--accent))' }} onClick={run}>
          OK
        </button>
      </div>
    </>
  );
}

function PageSetupDialog({ onClose }: { onClose(): void }) {
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const workbook = useWorkbookStore((s) => s.workbook);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const toast = useUiStore((s) => s.toast);
  const current = (activeSheetId && workbook ? getSheet(workbook, activeSheetId)?.pageSetup : undefined) ?? {};

  const [orientation, setOrientation] = React.useState<'portrait' | 'landscape'>(current.orientation ?? 'portrait');
  const [paperSize, setPaperSize] = React.useState(current.paperSize ?? 'letter');
  const [m, setM] = React.useState({ ...DEFAULT_MARGINS, ...(current.margins ?? {}) });
  const [mode, setMode] = React.useState<'scale' | 'fit'>(current.fitToPage ? 'fit' : 'scale');
  const [scale, setScale] = React.useState(current.scale ?? 100);
  const [fitW, setFitW] = React.useState(current.fitToWidth ?? 1);
  const [fitH, setFitH] = React.useState(current.fitToHeight ?? 1);

  const run = () => {
    if (!activeSheetId) return;
    const num = (v: number, fallback: number) => (Number.isFinite(v) && v >= 0 ? v : fallback);
    applyCommand(
      new SetPageSetupCommand(activeSheetId, {
        orientation,
        paperSize,
        margins: {
          top: num(m.top, DEFAULT_MARGINS.top),
          bottom: num(m.bottom, DEFAULT_MARGINS.bottom),
          left: num(m.left, DEFAULT_MARGINS.left),
          right: num(m.right, DEFAULT_MARGINS.right),
          header: num(m.header, DEFAULT_MARGINS.header),
          footer: num(m.footer, DEFAULT_MARGINS.footer)
        },
        fitToPage: mode === 'fit',
        scale: mode === 'scale' ? Math.min(400, Math.max(10, scale)) : current.scale,
        fitToWidth: mode === 'fit' ? Math.max(0, Math.min(10, fitW)) : current.fitToWidth,
        fitToHeight: mode === 'fit' ? Math.max(0, Math.min(10, fitH)) : current.fitToHeight
      })
    );
    toast('success', 'Page setup applied');
    onClose();
  };

  const field = (label: string, value: number, onChange: (v: number) => void) => (
    <label className="flex items-center gap-2 text-[12px]">
      <span className="w-[86px] text-right text-muted">{label}</span>
      <input
        type="number"
        step="0.1"
        min="0"
        value={Number.isFinite(value) ? value : 0}
        onChange={(e) => onChange(parseFloat(e.target.value))}
        className="w-[70px] rounded border px-1.5 py-0.5 text-[12px]"
      />
      <span className="text-muted">inches</span>
    </label>
  );

  return (
    <>
      <DialogHeader title="Page Setup" onClose={onClose} />
      <div className="flex gap-4">
        <div className="flex flex-col gap-2">
          <div className="text-[12px] font-semibold">Margins</div>
          {field('Top:', m.top, (v) => setM({ ...m, top: v }))}
          {field('Bottom:', m.bottom, (v) => setM({ ...m, bottom: v }))}
          {field('Left:', m.left, (v) => setM({ ...m, left: v }))}
          {field('Right:', m.right, (v) => setM({ ...m, right: v }))}
          {field('Header:', m.header, (v) => setM({ ...m, header: v }))}
          {field('Footer:', m.footer, (v) => setM({ ...m, footer: v }))}
        </div>
        <div className="flex flex-col gap-3">
          <div>
            <div className="mb-1 text-[12px] font-semibold">Orientation</div>
            <div className="flex gap-1">
              {(['portrait', 'landscape'] as const).map((o) => (
                <button
                  key={o}
                  onClick={() => setOrientation(o)}
                  className={`rounded border px-2.5 py-1 text-[12px] capitalize ${
                    orientation === o ? 'border-[rgb(var(--accent))] bg-[rgb(var(--accent-soft))]' : ''
                  }`}
                >
                  {o}
                </button>
              ))}
            </div>
          </div>
          <div>
            <div className="mb-1 text-[12px] font-semibold">Paper</div>
            <select
              value={paperSize}
              onChange={(e) => setPaperSize(e.target.value)}
              className="rounded border px-2 py-1 text-[12px]"
            >
              {PAPER_SIZES.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.label}
                </option>
              ))}
            </select>
          </div>
          <div>
            <div className="mb-1 text-[12px] font-semibold">Scaling</div>
            <div className="flex items-center gap-2 text-[12px]">
              <input
                type="radio"
                checked={mode === 'scale'}
                onChange={() => setMode('scale')}
              />
              <span>Adjust to</span>
              <input
                type="number"
                min={10}
                max={400}
                value={scale}
                disabled={mode !== 'scale'}
                onChange={(e) => setScale(parseInt(e.target.value, 10))}
                className="w-[58px] rounded border px-1.5 py-0.5 text-[12px] disabled:opacity-50"
              />
              <span className="text-muted">%</span>
            </div>
            <div className="mt-1.5 flex items-center gap-2 text-[12px]">
              <input
                type="radio"
                checked={mode === 'fit'}
                onChange={() => setMode('fit')}
              />
              <span>Fit to</span>
              <input
                type="number"
                min={0}
                max={10}
                value={fitW}
                disabled={mode !== 'fit'}
                onChange={(e) => setFitW(parseInt(e.target.value, 10))}
                className="w-[46px] rounded border px-1.5 py-0.5 text-[12px] disabled:opacity-50"
              />
              <span className="text-muted">page(s) wide by</span>
              <input
                type="number"
                min={0}
                max={10}
                value={fitH}
                disabled={mode !== 'fit'}
                onChange={(e) => setFitH(parseInt(e.target.value, 10))}
                className="w-[46px] rounded border px-1.5 py-0.5 text-[12px] disabled:opacity-50"
              />
              <span className="text-muted">tall</span>
            </div>
          </div>
        </div>
      </div>
      <p className="mt-3 text-[11px] text-muted">
        Saved with the sheet. Printing and PDF export are not implemented yet, so these settings are stored but not
        applied to output.
      </p>
      <div className="mt-4 flex justify-end gap-2">
        <button className="rounded border px-3 py-1.5 text-[13px]" onClick={onClose}>
          Cancel
        </button>
        <button
          className="rounded px-3 py-1.5 text-[13px] text-white"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={run}
        >
          OK
        </button>
      </div>
    </>
  );
}

export default DialogHost;
