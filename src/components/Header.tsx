import React, { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  Undo2,
  Redo2,
  Copy,
  Scissors,
  ClipboardPaste,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalSpaceBetween,
  ChevronDown,
  Search,
  Rows3,
  Columns3,
  Trash2,
  Sigma,
  Table2,
  BarChart3,
  Link2,
  MessageSquarePlus,
  PieChart,
  TrendingUp,
  FileSpreadsheet,
  FilePlus,
  FolderOpen,
  Save,
  FileDown,
  Clock,
  Percent,
  LayoutGrid,
  Minus,
  Plus,
  Type,
  Grid3x3,
  Palette,
  Sparkles,
  Info,
  Settings,
  Keyboard,
  ListChecks,
  ArrowUpDown,
  CopyMinus,
  Filter,
  CalendarDays,
  Braces,
  Calculator,
  ScrollText,
  Tag,
  Eye,
  GraduationCap,
  Moon,
  Sun,
  Lock,
  LockOpen,
  ChevronUp,
  MessageSquare,
  Printer,
  SquareDashed,
  ChevronsLeft,
  ChevronsRight,
  Paperclip,
  X,
  Square,
  Circle,
  MoveRight,
  Ruler,
  Snowflake
} from 'lucide-react';
import { useWorkbookStore, SINGLE, downloadBlob } from '@/store/useWorkbookStore';
import { useUiStore } from '@/store/useUiStore';
import { useSettingsStore } from '@/store/useSettingsStore';
import {
  InsertRowsCommand,
  InsertColumnsCommand,
  AddTableCommand,
  AddChartCommand,
  MergeCellsCommand,
  UnmergeCellsCommand,
  AddSheetCommand,
  SetZoomCommand,
  RemoveChartCommand,
  SetPageSetupCommand,
  SetPrintAreaCommand,
  SetFreezePanesCommand,
  SetThemeCommand,
  ClearThemeCommand
} from '@/engine/commands/commands';
import { cellToA1, rangeToA1 } from '@/engine/model/address';
import { ShapeKind, PageSetup } from '@/engine/model/types';
import { getSheet, currentRegion } from '@/engine/model/workbook';
import { clampZoom } from '@/grid/layout';
import { DEFAULT_MARGINS, PAPER_SIZES } from '@/engine/model/pageSetup';
import { THEMES, THEME_COLOR_SETS, THEME_FONT_SETS, getTheme } from '@/engine/model/theme';
import { workbookToCsvData } from '@/engine/io/csv';
import { CellStyle } from '@/engine/model/types';
import { FUNCTION_LIBRARY, jsonSafeName } from './DialogHost';

type TabName = 'file' | 'home' | 'insert' | 'draw' | 'page' | 'data' | 'formulas' | 'view' | 'review' | 'help';

const FONTS = ['Calibri', 'Arial', 'Arial Black', 'Comic Sans MS', 'Courier New', 'Gabriola', 'Georgia', 'Impact', 'Microsoft Sans Serif', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72];

const NUMBER_FORMATS = [
  { label: 'General', code: undefined },
  { label: 'Number', code: '0' },
  { label: 'Number (2 dp)', code: '0.00' },
  { label: 'Currency', code: '$#,##0.00' },
  { label: 'Percent', code: '0%' },
  { label: 'Percent (2 dp)', code: '0.00%' },
  { label: 'Scientific', code: '0.00E+00' },
  { label: 'Date (US)', code: 'm/d/yyyy' },
  { label: 'Date (ISO)', code: 'yyyy-mm-dd' },
  { label: 'Time', code: 'h:mm:ss AM/PM' }
];

export function Header() {
  const [tab, setTab] = useState<TabName>('home');
  const [fileMenu, setFileMenu] = useState(false);
  const headerRef = useRef<HTMLDivElement>(null);

  // File opens a dropdown over the ribbon instead of replacing it.
  const selectTab = (name: TabName) => {
    if (name === 'file') {
      setFileMenu((v) => !v);
      return;
    }
    setFileMenu(false);
    setTab(name);
  };

  useEffect(() => {
    if (!fileMenu) return;
    const onDown = (e: PointerEvent) => {
      if (!headerRef.current?.contains(e.target as Node)) setFileMenu(false);
    };
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && setFileMenu(false);
    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('keydown', onKey, true);
    return () => {
      window.removeEventListener('pointerdown', onDown, true);
      window.removeEventListener('keydown', onKey, true);
    };
  }, [fileMenu]);

  return (
    <div
      ref={headerRef}
      className="relative select-none"
      onPointerDown={(e) => e.stopPropagation()}
onMouseDownCapture={(e) => {
           // A ribbon click must not blur an in-progress edit. Once the formula
           // input loses focus the draft is committed and forgotten, so the next
           // keystroke starts a fresh edit and silently replaces what was typed.
           // Swallow the default focus move for buttons only; <select> and
           // <input> in the ribbon still need it to open.
           if (!useWorkbookStore.getState().editing) return;
           if ((e.target as HTMLElement).closest?.('button')) e.preventDefault();
         }}
         onClickCapture={(e) => {
           // Cells are not focusable, so a ribbon button would keep keyboard focus
           // and the grid would then ignore every keystroke. Release it on real
           // pointer clicks only (detail !== 0) so keyboard activation still works.
           const btn = (e.target as HTMLElement).closest?.('button');
           if (btn && e.detail !== 0) btn.blur();
         }}
    >
      <TitleBar onFileMenu={() => selectTab('file')} />
      <TabStrip active={tab} fileMenuOpen={fileMenu} onSelect={selectTab} />
      <RibbonPanel tab={tab} />
      {fileMenu && <FileMenu onClose={() => setFileMenu(false)} />}
    </div>
  );
}

/* ------------------------------- Title bar ------------------------------ */

function TitleBar({ onFileMenu }: { onFileMenu(): void }) {
  const workbook = useWorkbookStore((s) => s.workbook);
  const undo = useWorkbookStore((s) => s.undo);
  const redo = useWorkbookStore((s) => s.redo);
  const canUndo = useWorkbookStore((s) => s.history.canUndo());
  const canRedo = useWorkbookStore((s) => s.history.canRedo());
  const openDialog = useUiStore((s) => s.openDialog);
  const saveWorkbook = useWorkbookStore((s) => s.save);
  const filePath = useWorkbookStore((s) => s.filePath);
  const theme = useSettingsStore((s) => s.theme);
  const setTheme = useSettingsStore((s) => s.setTheme);
  const dirty = useWorkbookStore((s) => s.dirty);

  const name = workbook?.metadata.title ?? 'Untitled';

  return (
    <div
      className="flex items-center border-b px-2"
      style={{ background: 'rgb(var(--panel))', borderColor: 'rgb(var(--line))', height: 36 }}
    >
      <button
        className="flex items-center gap-1.5 rounded px-2 py-1 text-[12px] font-semibold"
        style={{ color: 'rgb(var(--accent))' }}
        onClick={onFileMenu}
        title="File menu"
      >
        <FileSpreadsheet size={16} />
        {name}
        <ChevronDown size={13} />
      </button>

      <div className="mx-1 h-5 w-px" style={{ background: 'rgb(var(--line-strong))' }} />

      <button
        className="rounded p-1 hover:bg-black/5 dark:hover:bg-white/10 disabled:opacity-40"
        onClick={() => void saveWorkbook()}
        title={filePath ? `Save ${filePath.split(/[\\/]/).pop()}` : 'Save (Ctrl+S)'}
      >
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
          <path d="M17 21v-8H7v8" />
          <path d="M7 3v5h8" />
        </svg>
      </button>
      <button
        className="rounded p-1 hover:bg-black/5 dark:hover:bg-white/10 disabled:opacity-40"
        onClick={undo}
        disabled={!canUndo}
        title="Undo (Ctrl+Z)"
      >
        <Undo2 size={16} />
      </button>
      <button
        className="rounded p-1 hover:bg-black/5 dark:hover:bg-white/10 disabled:opacity-40"
        onClick={redo}
        disabled={!canRedo}
        title="Redo (Ctrl+Y)"
      >
        <Redo2 size={16} />
      </button>

      <div className="flex-1" />

      <span
        className="hidden truncate text-[12px] md:block"
        style={{ color: 'rgb(var(--text-muted))' }}
        title={name}
      >
        {name} - BritSheet{dirty ? ' *' : ''}
      </span>

      <div className="flex-1" />

      <button
        className="rounded p-1 hover:bg-black/5 dark:hover:bg-white/10"
        onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
        title={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
      >
        {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
      </button>
    </div>
  );
}

/* -------------------------------- Tab strip ------------------------------ */

const TABS: Array<{ name: TabName; label: string }> = [
  { name: 'file', label: 'File' },
  { name: 'home', label: 'Home' },
  { name: 'insert', label: 'Insert' },
  { name: 'draw', label: 'Draw' },
  { name: 'page', label: 'Page Layout' },
  { name: 'data', label: 'Data' },
  { name: 'formulas', label: 'Formulas' },
  { name: 'review', label: 'Review' },
  { name: 'view', label: 'View' },
  { name: 'help', label: 'Help' }
];

function TabStrip({
  active,
  fileMenuOpen,
  onSelect
}: {
  active: TabName;
  fileMenuOpen: boolean;
  onSelect(t: TabName): void;
}) {
  return (
    <div className="flex items-center px-1" style={{ background: 'rgb(var(--panel))', borderColor: 'rgb(var(--line))' }}>
      {TABS.map((t) => {
        const selected = t.name === 'file' ? fileMenuOpen : active === t.name;
        return (
          <button
            key={t.name}
            onClick={() => onSelect(t.name)}
            className={`relative rounded-t px-2 py-[5px] text-[12px] ${
              selected ? 'font-semibold' : 'opacity-90 hover:opacity-100'
            }`}
            style={selected ? { color: 'rgb(var(--accent))' } : { color: 'rgb(var(--text))' }}
          >
            {t.label}
            {selected && (
              <span
                className="absolute inset-x-1 -bottom-px h-[2px] rounded-full"
                style={{ background: 'rgb(var(--accent))' }}
              />
            )}
          </button>
        );
      })}
    </div>
  );
}

/* -------------------------------- Ribbon -------------------------------- */

function RibbonPanel({ tab }: { tab: TabName }) {
  switch (tab) {
    case 'home':
      return <HomeTab />;
    case 'insert':
      return <InsertTab />;
    case 'draw':
      return <DrawTab />;
    case 'page':
      return <PageLayoutTab />;
    case 'data':
      return <DataTab />;
    case 'formulas':
      return <FormulasTab />;
    case 'view':
      return <ViewTab />;
    case 'review':
      return <ReviewTab />;
    case 'help':
      return <HelpTab />;
  }
}

function FileMenu({ onClose }: { onClose(): void }) {
  const openDialog = useUiStore((s) => s.openDialog);
  const setView = useUiStore((s) => s.setView);
  const toast = useUiStore((s) => s.toast);
  const recentFiles = useUiStore((s) => s.recentFiles);
  const removeRecent = useUiStore((s) => s.removeRecent);
  const filePath = useWorkbookStore((s) => s.filePath);
  const save = useWorkbookStore((s) => s.save);

  const run = (fn: () => void) => () => {
    onClose();
    fn();
  };

  const newWorkbook = () => {
    useWorkbookStore.getState().newWorkbook();
    setView('workspace');
  };

  /** Native open dialog in the desktop build, browser file input otherwise. */
  const openFile = async () => {
    try {
      const { pickOpenFile, isTauri } = await import('@/platform/bridge');
      if (isTauri()) {
        const path = await pickOpenFile([
          { name: 'BritSheet Workbook', extensions: ['britsheet'] },
          { name: 'Excel workbook', extensions: ['xlsx', 'xls'] },
          { name: 'CSV / JSON', extensions: ['csv', 'json'] },
          { name: 'All files', extensions: ['*'] }
        ]);
        if (!path) return;
        const { loadWorkbookFromFile } = await import('@/platform/loaders');
        await loadWorkbookFromFile(path);
        setView('workspace');
        toast('success', `Opened ${path.split(/[\\/]/).pop()}`);
        return;
      }
      // Browser preview: let the user pick a file from disk.
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.britsheet,.bs,.xlsx,.xls,.csv,.txt,.json';
      input.onchange = async () => {
        const file = input.files?.[0];
        if (!file) return;
        const text = await file.text();
        const lower = file.name.toLowerCase();
        try {
          if (lower.endsWith('.csv') || lower.endsWith('.txt')) {
            const { parseCsv, sniffCsvDelimiter, csvRowsToWorkbook } = await import('@/engine/io/csv');
            const wb = csvRowsToWorkbook(parseCsv(text, sniffCsvDelimiter(text)));
            useWorkbookStore.getState().loadWorkbook(wb);
          } else {
            const { deserializeWorkbook } = await import('@/engine/io/serialize');
            useWorkbookStore.getState().loadWorkbook(deserializeWorkbook(text));
          }
          setView('workspace');
          toast('success', `Opened ${file.name}`);
        } catch (err) {
          toast('error', `Could not open ${file.name}: ${(err as Error).message}`);
        }
      };
      input.click();
    } catch (err) {
      toast('error', `Could not open file: ${(err as Error).message}`);
    }
  };

  const openRecent = async (path: string) => {
    try {
      const { loadWorkbookFromFile } = await import('@/platform/loaders');
      await loadWorkbookFromFile(path);
      setView('workspace');
      toast('success', `Opened ${path.split(/[\\/]/).pop()}`);
    } catch (err) {
      toast('error', `Could not open ${path}: ${(err as Error).message}`);
      removeRecent(path);
    }
  };

  const exportCsv = () => {
    const wb = useWorkbookStore.getState().workbook;
    if (!wb) return;
    const rows = workbookToCsvData(wb);
    const csv = rows.map((r) => r.map(escapeCsv).join(',')).join('\n');
    downloadBlob(csv, jsonSafeName(wb.metadata.title) + '.csv', 'text/csv');
    toast('success', 'Exported to CSV');
  };

  const items: Array<{ label: string; hint?: string; icon: React.ReactNode; onClick(): void; muted?: boolean }> = [
    { label: 'New', hint: 'Ctrl+N', icon: <FilePlus size={15} />, onClick: newWorkbook },
    { label: 'Open…', hint: 'Ctrl+O', icon: <FolderOpen size={15} />, onClick: openFile },
    {
      label: 'Save',
      hint: 'Ctrl+S',
      icon: <Save size={15} />,
      muted: true,
      onClick: () => void save()
    },
    { label: 'Save As…', hint: 'F12', icon: <FileDown size={15} />, onClick: () => openDialog('saveAs') },
    { label: 'Export to CSV', icon: <FileDown size={15} />, onClick: exportCsv }
  ];

  return (
    <div
      data-file-menu="true"
      className="absolute left-1 z-50 w-[300px] rounded-b border p-1 shadow-lg"
      style={{ background: 'rgb(var(--panel))', borderColor: 'rgb(var(--line))', boxShadow: 'var(--shadow)' }}
    >
      {items.map((it) => (
        <button
          key={it.label}
          className="flex w-full items-center gap-2 rounded px-2 py-[7px] text-left text-[13px] hover:bg-black/5 dark:hover:bg-white/10"
          onClick={run(it.onClick)}
        >
          <span className="text-[rgb(var(--text-muted))]">{it.icon}</span>
          <span className="flex-1">{it.label}</span>
          {it.muted && <span className="max-w-[150px] truncate text-[10px] opacity-50">{fileName(filePath)}</span>}
          {it.hint && <span className="text-[10px] opacity-50">{it.hint}</span>}
        </button>
      ))}
      {recentFiles.length > 0 && (
        <>
          <div className="my-1 border-t" style={{ borderColor: 'rgb(var(--line))' }} />
          <div className="px-2 pb-1 pt-0.5 text-[10px] uppercase tracking-wide opacity-50">Recent</div>
          {recentFiles.slice(0, 6).map((f) => (
            <div key={f.path} className="group flex items-center gap-1">
              <button
                className="flex min-w-0 flex-1 items-center gap-2 rounded px-2 py-[5px] text-left text-[12px] hover:bg-black/5 dark:hover:bg-white/10"
                onClick={run(() => void openRecent(f.path))}
              >
                <Clock size={13} className="shrink-0 opacity-50" />
                <span className="truncate">{f.name}</span>
              </button>
              <button
                className="rounded p-1 opacity-0 hover:bg-black/5 group-hover:opacity-100 dark:hover:bg-white/10"
                title="Remove from recent files"
                onClick={run(() => removeRecent(f.path))}
              >
                <X size={12} />
              </button>
            </div>
          ))}
        </>
      )}
    </div>
  );
}

/** File name for the current workbook path, or a hint when it has never been saved. */
function fileName(path: string | null): string {
  if (!path) return 'not saved yet';
  return path.split(/[\\/]/).pop() ?? path;
}

/* -------------------------------- Home tab ------------------------------- */

function HomeTab() {
  return (
    <div className="flex items-stretch gap-0 overflow-x-auto border-b px-1" style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}>
      <ClipboardGroup />
      <GroupDivider />
      <FontGroup />
      <GroupDivider />
      <AlignmentGroup />
      <GroupDivider />
      <NumberGroup />
      <GroupDivider />
      <CellsGroup />
      <GroupDivider />
      <EditingGroup />
    </div>
  );
}

function GroupShell({ caption, children }: { caption: string; children: React.ReactNode }) {
  return (
    <div
      data-ribbon-group={caption}
      className="relative flex min-w-[72px] flex-col px-2 pt-1 pb-0.5"
    >
      <div className="flex flex-1 items-start justify-center gap-0.5">{children}</div>
      <div className="mt-1 text-center text-[10px]" style={{ color: 'rgb(var(--text-muted))' }}>
        {caption}
      </div>
    </div>
  );
}

function BigButton({
  icon,
  label,
  onClick,
  active,
  disabled,
  title
}: {
  icon: React.ReactNode;
  label: string;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  title?: string;
}) {
  return (
    <button
      className={`flex flex-col items-center justify-between rounded px-2 py-1 ${
        disabled ? 'opacity-40' : active ? 'bg-[rgb(var(--accent-soft))]' : 'hover:bg-black/5 dark:hover:bg-white/10'
      }`}
      onClick={onClick}
      disabled={disabled}
      title={title ?? label}
    >
      <span className="flex h-6 w-7 items-center justify-center">{icon}</span>
      <span className="text-[10px] leading-tight">{label}</span>
    </button>
  );
}

function SmallButton({
  icon,
  onClick,
  active,
  disabled,
  title,
  compact
}: {
  icon: React.ReactNode;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  title?: string;
  compact?: boolean;
}) {
  return (
    <button
      className={`flex ${compact ? 'h-6 w-7' : 'h-7 w-8'} items-center justify-center rounded ${
        disabled ? 'opacity-40' : active ? 'bg-[rgb(var(--accent-soft))]' : 'hover:bg-black/5 dark:hover:bg-white/10'
      }`}
      onClick={onClick}
      disabled={disabled}
      title={title}
    >
      {icon}
    </button>
  );
}

function ClipboardGroup() {
  const copy = useWorkbookStore((s) => s.copySelection);
  const cut = useWorkbookStore((s) => s.cutSelection);
  const paste = useWorkbookStore((s) => s.paste);
  return (
    <GroupShell caption="Clipboard">
      <BigButton icon={<ClipboardPaste size={18} />} label="Paste" title="Paste (Ctrl+V)" onClick={paste} />
      <div className="flex flex-col gap-0.5 pt-1">
        <SmallButton icon={<Copy size={15} />} title="Copy (Ctrl+C)" onClick={copy} />
        <SmallButton icon={<Scissors size={15} />} title="Cut (Ctrl+X)" onClick={cut} />
      </div>
    </GroupShell>
  );
}

function useActiveCell() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const formatSelection = useWorkbookStore((s) => s.formatSelection);
  const sheet = workbook && activeSheetId ? getSheet(workbook, activeSheetId) : undefined;
  const cell = sheet ? sheet.cells[cellToA1(selection.col, selection.row)] : undefined;
  return { cell, formatSelection };
}

function FontGroup() {
  const { cell, formatSelection } = useActiveCell();
  const [pick, setPick] = useState<'color' | 'fill' | null>(null);
  const [anchor, setAnchor] = useState<{ left: number; top: number } | null>(null);
  const colorRef = useRef<HTMLInputElement>(null);
  const fillRef = useRef<HTMLInputElement>(null);
  const colorBtnRef = useRef<HTMLSpanElement>(null);
  const fillBtnRef = useRef<HTMLSpanElement>(null);

  const style: CellStyle | undefined = cell?.style;
  const toggle = (key: 'bold' | 'italic' | 'underline' | 'strikethrough') =>
    formatSelection({ [key]: !style?.[key] } as Partial<CellStyle>);

  // The ribbon row clips its overflow, so the palette is anchored to the
  // trigger button in viewport coordinates instead of the group box.
  const openPalette = (kind: 'color' | 'fill', btn: React.RefObject<HTMLSpanElement>) => {
    if (pick === kind) {
      setPick(null);
      return;
    }
    const rect = btn.current?.getBoundingClientRect();
    setAnchor(rect ? { left: rect.left, top: rect.bottom + 4 } : null);
    setPick(kind);
  };

  return (
    <GroupShell caption="Font">
      <div className="flex flex-col gap-0.5 pt-0.5">
        <div className="flex items-center gap-0.5">
          <select
            className="h-6 w-[108px] rounded border px-1 text-[11px]"
            style={{ borderColor: 'rgb(var(--line-strong))', background: 'rgb(var(--cell))' }}
            value={style?.fontFamily ?? 'Calibri'}
            onChange={(e) => formatSelection({ fontFamily: e.target.value })}
            title="Font"
          >
            {FONTS.map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
          <select
            className="h-6 w-[52px] rounded border px-1 text-[11px]"
            style={{ borderColor: 'rgb(var(--line-strong))', background: 'rgb(var(--cell))' }}
            value={style?.fontSize ?? 11}
            onChange={(e) => formatSelection({ fontSize: parseInt(e.target.value, 10) })}
            title="Font size"
          >
            {SIZES.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        </div>
        <div className="flex items-center gap-0.5">
          <SmallButton icon={<Bold size={15} />} active={style?.bold} onClick={() => toggle('bold')} title="Bold (Ctrl+B)" />
          <SmallButton icon={<Italic size={15} />} active={style?.italic} onClick={() => toggle('italic')} title="Italic (Ctrl+I)" />
          <SmallButton icon={<Underline size={15} />} active={style?.underline} onClick={() => toggle('underline')} title="Underline (Ctrl+U)" />
          <SmallButton icon={<Strikethrough size={15} />} active={style?.strikethrough} onClick={() => toggle('strikethrough')} title="Strikethrough" />
          <span ref={colorBtnRef}>
            <SmallButton icon={<Palette size={15} />} title="Font color" active={pick === 'color'} onClick={() => openPalette('color', colorBtnRef)} />
          </span>
          <span ref={fillBtnRef}>
            <SmallButton icon={<Grid3x3 size={15} />} title="Fill color" active={pick === 'fill'} onClick={() => openPalette('fill', fillBtnRef)} />
          </span>
        </div>
      </div>
      <input
        ref={colorRef}
        type="color"
        className="hidden"
        onChange={(e) => {
          formatSelection({ color: e.target.value });
          setPick(null);
        }}
      />
      <input
        ref={fillRef}
        type="color"
        className="hidden"
        onChange={(e) => {
          formatSelection({ backgroundColor: e.target.value });
          setPick(null);
        }}
      />
      {pick && (
        <MiniPalette
          anchor={anchor}
          onPick={(hex) => {
            // A literal color clears any theme reference so it wins, like Excel.
            formatSelection(
              pick === 'color' ? { color: hex, colorTheme: undefined } : { backgroundColor: hex, backgroundTheme: undefined }
            );
            setPick(null);
          }}
          onPickTheme={(i) => {
            // Store the accent reference, not the hex, so themes repaint it.
            formatSelection(
              pick === 'color'
                ? { colorTheme: i, color: undefined }
                : { backgroundTheme: i, backgroundColor: undefined }
            );
            setPick(null);
          }}
          onCustom={() => (pick === 'color' ? colorRef : fillRef).current?.click()}
          onClose={() => setPick(null)}
        />
      )}
    </GroupShell>
  );
}

function MiniPalette({
  onPick,
  onPickTheme,
  onCustom,
  onClose,
  anchor
}: {
  onPick(c: string): void;
  onPickTheme(index: number): void;
  onCustom(): void;
  onClose(): void;
  anchor: { left: number; top: number } | null;
}) {
  const colors = ['#000000', '#444444', '#ffffff', '#e7e6e6', '#d35fb6', '#ff0000', '#e36c09', '#ffc000', '#00b050', '#2e75b6', '#7030a0'];
  // Top row is the live theme palette so Page Layout > Themes repaints fills.
  const theme = getTheme(useWorkbookStore((s) => s.workbook?.theme));
  return createPortal(
    <>
      <div className="fixed inset-0 z-[70]" onClick={onClose} />
      <div
        className="fixed z-[80] rounded border bg-[rgb(var(--cell))] p-2 shadow-lg"
        style={{ borderColor: 'rgb(var(--line-strong))', left: anchor?.left ?? 80, top: anchor?.top ?? 120 }}
      >
        <div className="grid w-[108px] grid-cols-6 gap-1">
          {theme.accents.map((c, i) => (
            <button
              key={`t${i}`}
              className="h-3.5 w-3.5 rounded-sm border"
              style={{ background: c, borderColor: 'rgb(var(--line-strong))' }}
              onClick={() => onPickTheme(i)}
              title={`${theme.name} accent ${i + 1}`}
            />
          ))}
          {colors.map((c) => (
            <button
              key={c}
              className="h-3.5 w-3.5 rounded-sm border"
              style={{ background: c, borderColor: 'rgb(var(--line-strong))' }}
              onClick={() => onPick(c)}
              title={c}
            />
          ))}
        </div>
        <button className="mt-1.5 w-full rounded px-1 py-0.5 text-[10px] hover:bg-black/5 dark:hover:bg-white/10" onClick={onCustom}>
          More colors…
        </button>
      </div>
    </>,
    document.body
  );
}

function AlignmentGroup() {
  const { formatSelection } = useActiveCell();
  const selection = useWorkbookStore((s) => s.selection);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);

  const cellsRange = () => {
    const { row, col, endRow, endCol } = selection;
    const a = cellToA1(Math.min(col, endCol), Math.min(row, endRow));
    const b = cellToA1(Math.max(col, endCol), Math.max(row, endRow));
    return a === b ? a : `${a}:${b}`;
  };

  return (
    <GroupShell caption="Alignment">
      <div className="flex flex-col gap-0.5 pt-0.5">
        <div className="flex items-center gap-0.5">
          <SmallButton compact icon={<AlignLeft size={15} />} title="Align left" onClick={() => formatSelection({ horizontalAlign: 'left' })} />
          <SmallButton compact icon={<AlignCenter size={15} />} title="Align center" onClick={() => formatSelection({ horizontalAlign: 'center' })} />
          <SmallButton compact icon={<AlignRight size={15} />} title="Align right" onClick={() => formatSelection({ horizontalAlign: 'right' })} />
        </div>
        <div className="flex items-center gap-0.5">
          <SmallButton compact icon={<AlignVerticalJustifyCenter size={15} />} title="Align top" onClick={() => formatSelection({ verticalAlign: 'top' })} />
          <SmallButton compact icon={<AlignVerticalSpaceBetween size={15} />} title="Align middle" onClick={() => formatSelection({ verticalAlign: 'middle' })} />
          <SmallButton
            compact
            icon={<Type size={15} />}
            title="Wrap text"
            onClick={() => {
              const st = useWorkbookStore.getState();
              const sheet = st.workbook && st.activeSheetId ? getSheet(st.workbook, st.activeSheetId) : undefined;
              const c = sheet?.cells[cellToA1(st.selection.col, st.selection.row)];
              st.formatSelection({ wrapText: !c?.style?.wrapText });
            }}
          />
        </div>
        <div className="flex items-center gap-0.5">
          <SmallButton compact icon={<Grid3x3 size={15} />} title="Merge & center" onClick={() => activeSheetId && applyCommand(new MergeCellsCommand('Merge cells', activeSheetId, cellsRange()))} />
          <SmallButton compact icon={<Grid3x3 size={15} />} title="Unmerge cells" onClick={() => activeSheetId && applyCommand(new UnmergeCellsCommand('Unmerge cells', activeSheetId, cellsRange()))} />
        </div>
      </div>
    </GroupShell>
  );
}

function NumberGroup() {
  const { cell, formatSelection } = useActiveCell();
  const openDialog = useUiStore((s) => s.openDialog);
  const current = cell?.style?.numberFormat;

  return (
    <GroupShell caption="Number">
      <div className="flex flex-col gap-0.5 pt-0.5">
        <select
          className="h-6 w-[118px] rounded border px-1 text-[11px]"
          style={{ borderColor: 'rgb(var(--line-strong))', background: 'rgb(var(--cell))' }}
          value={current ?? 'General'}
          onChange={(e) => formatSelection({ numberFormat: e.target.value === 'General' ? undefined : e.target.value })}
          title="Number format"
        >
          {NUMBER_FORMATS.map((f) => (
            <option key={f.label} value={f.code ?? 'General'}>
              {f.label}
            </option>
          ))}
        </select>
        <div className="flex items-center gap-0.5">
          <SmallButton icon={<Percent size={15} />} title="Percent format" onClick={() => formatSelection({ numberFormat: '0%' })} />
          <SmallButton icon={<FileSpreadsheet size={15} />} title="Currency format" onClick={() => formatSelection({ numberFormat: '$#,##0.00' })} />
          <SmallButton icon={<Search size={15} />} title="Format cells…" onClick={() => openDialog('formatCells')} />
        </div>
      </div>
    </GroupShell>
  );
}

function CellsGroup() {
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const clearSelection = useWorkbookStore((s) => s.clearSelection);
  const undo = useWorkbookStore((s) => s.undo);
  const canUndo = useWorkbookStore((s) => s.history.canUndo());
  const selectedChartId = useWorkbookStore((s) => s.selectedChartId);
  const setSelectedChartId = useWorkbookStore((s) => s.setSelectedChartId);

  /** Excel's Delete: removes a selected chart, otherwise clears cell contents. */
  const deleteOrClear = () => {
    if (selectedChartId) {
      if (activeSheetId) applyCommand(new RemoveChartCommand('Delete chart', activeSheetId, selectedChartId));
      setSelectedChartId(null);
      return;
    }
    clearSelection();
  };

  const insertRow = () =>
    activeSheetId && applyCommand(new InsertRowsCommand('Insert rows', activeSheetId, selection.row - 1, 1));
  const insertCol = () =>
    activeSheetId && applyCommand(new InsertColumnsCommand('Insert columns', activeSheetId, selection.col, 1));

  return (
    <GroupShell caption="Cells">
      <div className="flex flex-col gap-0.5 pt-1">
        <div className="flex items-center gap-0.5">
          <SmallButton icon={<Rows3 size={15} />} title="Insert row above" onClick={insertRow} />
          <SmallButton icon={<Columns3 size={15} />} title="Insert column left" onClick={insertCol} />
        </div>
        <div className="flex items-center gap-0.5">
          <SmallButton
            icon={<Trash2 size={15} />}
            title={selectedChartId ? 'Delete the selected chart' : 'Delete (clear contents)'}
            onClick={deleteOrClear}
          />
          <SmallButton icon={<Undo2 size={15} />} title="Undo (Ctrl+Z)" onClick={undo} disabled={!canUndo} />
        </div>
      </div>
    </GroupShell>
  );
}

function EditingGroup() {
  const openDialog = useUiStore((s) => s.openDialog);
  return (
    <GroupShell caption="Editing">
      <BigButton icon={<ArrowUpDown size={18} />} label="Sort" title="Sort range…" onClick={() => openDialog('sort')} />
      <div className="flex flex-col gap-0.5 pt-1">
        <SmallButton icon={<Search size={15} />} title="Find & Replace (Ctrl+F)" onClick={() => openDialog('findReplace')} />
        <SmallButton icon={<ListChecks size={15} />} title="Data validation…" onClick={() => openDialog('dataValidation')} />
      </div>
    </GroupShell>
  );
}

/* ------------------------------- Other tabs ------------------------------ */

/* --------------------- Excel-style split button + popups --------------------- */

/** A big icon button with a separate caret, mirroring Excel 365 split buttons. */
/** Compact ribbon button that opens a gallery panel (Themes/Colors/Fonts). */
function GalleryButton({
  icon,
  label,
  title,
  children
}: {
  icon: React.ReactNode;
  label: string;
  title: string;
  children: (rect: DOMRect, close: () => void) => React.ReactNode;
}) {
  const [anchor, setAnchor] = useState<DOMRect | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const open = anchor !== null;
  useEffect(() => {
    if (!open) return;
    const close = (e: PointerEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setAnchor(null);
    };
    const esc = (e: KeyboardEvent) => e.key === 'Escape' && setAnchor(null);
    window.addEventListener('pointerdown', close);
    window.addEventListener('keydown', esc);
    return () => {
      window.removeEventListener('pointerdown', close);
      window.removeEventListener('keydown', esc);
    };
  }, [open]);
  return (
    <div ref={ref} className="flex flex-col items-center">
      <button
        className="flex flex-col items-center justify-center rounded px-1.5 py-1 hover:bg-black/5 dark:hover:bg-white/10"
        style={{ width: 56, height: 46 }}
        title={title}
        onClick={(e) => {
          const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
          setAnchor((cur) => (cur ? null : r));
        }}
      >
        {icon}
        <span className="mt-0.5 max-w-[54px] truncate text-[10px] leading-tight">{label}</span>
      </button>
      {open && anchor && createPortal(children(anchor, () => setAnchor(null)), document.body)}
    </div>
  );
}

/** Floating gallery surface anchored under a ribbon button. */
function GalleryPanel({
  menu,
  rect,
  width,
  children
}: {
  menu: string;
  rect: DOMRect;
  width: number;
  children: React.ReactNode;
}) {
  return (
    <div
      data-ribbon-menu={menu}
      className="fixed z-[80] max-h-[70vh] overflow-auto rounded border bg-[rgb(var(--cell))] p-2 shadow-lg"
      style={{
        left: Math.max(4, Math.min(rect.left, window.innerWidth - width - 8)),
        top: rect.bottom + 4,
        width
      }}
    >
      {children}
    </div>
  );
}

function RibbonPanelFooter({ children }: { children: React.ReactNode }) {
  return (
    <div className="mt-2 border-t pt-1" style={{ borderColor: 'rgb(var(--line))' }}>
      {children}
    </div>
  );
}

function SplitButton({
  icon,
  label,
  title,
  onOpen,
  compact,
  children
}: {
  icon: React.ReactNode;
  label: string;
  title: string;
  onOpen(): void;
  compact?: boolean;
  children: (rect: DOMRect, close: () => void) => React.ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const [anchor, setAnchor] = useState<DOMRect | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: PointerEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    const esc = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false);
    window.addEventListener('pointerdown', close);
    window.addEventListener('keydown', esc);
    return () => {
      window.removeEventListener('pointerdown', close);
      window.removeEventListener('keydown', esc);
    };
  }, [open]);
  const openFrom = (e: React.MouseEvent) => {
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setAnchor(r);
    setOpen((v) => !v);
  };
  return (
    <div ref={ref} className="flex flex-col items-center">
      {compact ? (
        <div className="flex items-stretch overflow-hidden rounded hover:bg-black/5 dark:hover:bg-white/10">
          <button
            className="flex items-center gap-1 py-0.5 pl-1 pr-0.5"
            onClick={onOpen}
            title={title}
          >
            {icon}
            <span className="text-[10px] leading-tight">{label}</span>
          </button>
          <button className="flex w-3 items-center justify-center" onClick={openFrom} title={`${title} options`}>
            <ChevronDown size={11} />
          </button>
        </div>
      ) : (
        <div className="flex">
          <button
            className="flex flex-col items-center justify-center rounded-l px-1.5 py-1 hover:bg-black/5 dark:hover:bg-white/10"
            style={{ width: 52, height: 46 }}
            onClick={onOpen}
            title={title}
          >
            {icon}
            <span className="mt-0.5 max-w-[50px] truncate text-[10px] leading-tight">{label}</span>
          </button>
          <button
            className="flex w-4 items-center justify-center rounded-r hover:bg-black/5 dark:hover:bg-white/10"
            style={{ height: 46 }}
            onClick={openFrom}
            title={`${title} options`}
          >
            <ChevronDown size={12} />
          </button>
        </div>
      )}
      {open && anchor && createPortal(children(anchor, () => setOpen(false)), document.body)}
    </div>
  );
}

/** Excel's insert-table size picker: hover to size, click to confirm. */
function TableGridPopup({ rect, onClose, onPick, onDialog }: { rect: DOMRect; onClose(): void; onPick(r: number, c: number): void; onDialog(): void }) {
  const [hover, setHover] = useState({ r: 1, c: 1 });
  return (
    <MenuList x={rect.left} y={rect.bottom + 2} width={198} onClose={onClose}>
      <div className="px-3 pb-1 pt-1 text-[11px]" style={{ color: 'rgb(var(--text-muted))' }}>
        Insert Table
        <span className="float-right text-[12px] font-semibold" style={{ color: 'rgb(var(--text))' }}>
          {hover.r} x {hover.c}
        </span>
      </div>
      <div className="px-3 pb-2">
        <div className="grid grid-cols-10 gap-[2px]">
          {Array.from({ length: 80 }, (_, i) => i + 1).map((n) => {
            const r = Math.ceil(n / 10);
            const c = ((n - 1) % 10) + 1;
            const on = r <= hover.r && c <= hover.c;
            return (
              <button
                key={n}
                className="h-[13px] w-[13px] border"
                style={{ background: on ? 'rgb(var(--accent))' : 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line-strong))' }}
                onMouseEnter={() => setHover({ r, c })}
                onClick={() => {
                  onPick(r, c);
                  onClose();
                }}
                title={`${r} x ${c}`}
              />
            );
          })}
        </div>
        <div className="mt-1 border-t pt-1" style={{ borderColor: 'rgb(var(--line))' }}>
          <button className="w-full rounded px-1 py-0.5 text-left text-[11px] hover:bg-black/5" onClick={onDialog}>
            Insert Table…
          </button>
        </div>
      </div>
    </MenuList>
  );
}

function ChartMenuPopup({ rect, onClose, types, onPick }: { rect: DOMRect; onClose(): void; types: { type: 'column' | 'bar' | 'line' | 'pie'; label: string }[]; onPick(t: 'column' | 'bar' | 'line' | 'pie'): void }) {
  return (
    <MenuList x={rect.left} y={rect.bottom + 2} width={190} onClose={onClose}>
      {types.map((t) => (
        <MenuItem
          key={t.type}
          onClick={() => {
            onPick(t.type);
            onClose();
          }}
        >
          {t.label}
        </MenuItem>
      ))}
    </MenuList>
  );
}

function MenuList({ x, y, width, onClose, children }: { x: number; y: number; width: number; onClose(): void; children: React.ReactNode }) {
  return createPortal(
    <div
      data-ribbon-menu
      className="fixed z-[70] rounded border py-1 shadow-xl"
      style={{
        left: Math.min(x, window.innerWidth - width - 8),
        top: y,
        width,
        background: 'rgb(var(--cell))',
        borderColor: 'rgb(var(--line-strong))'
      }}
      onPointerDown={(e) => e.stopPropagation()}
    >
      {children}
      <button className="hidden" onClick={onClose} />
    </div>,
    document.body
  );
}

function MenuItem({ children, onClick }: { children: React.ReactNode; onClick(): void }) {
  return (
    <button
      className="block w-full px-3 py-1.5 text-left text-[12px] hover:bg-black/5 dark:hover:bg-white/10"
      onClick={onClick}
    >
      {children}
    </button>
  );
}

function ItemMenuPopup({
  rect,
  onClose,
  items,
  onPick,
  extra
}: {
  rect: DOMRect;
  onClose(): void;
  items: string[];
  onPick(item: string): void;
  extra?: { label: string; onClick(): void };
}) {
  return (
    <MenuList x={rect.left} y={rect.bottom + 2} width={196} onClose={onClose}>
      {items.map((n) => (
        <MenuItem
          key={n}
          onClick={() => {
            onPick(n);
            onClose();
          }}
        >
          {n}
        </MenuItem>
      ))}
      {extra && (
        <button
          className="block w-full border-t px-3 py-1.5 text-left text-[12px] hover:bg-black/5 dark:hover:bg-white/10"
          style={{ borderColor: 'rgb(var(--line))' }}
          onClick={() => {
            extra.onClick();
            onClose();
          }}
        >
          {extra.label}
        </button>
      )}
    </MenuList>
  );
}

function InsertTab() {
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const workbook = useWorkbookStore((s) => s.workbook);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const openDialog = useUiStore((s) => s.openDialog);

  const anchorCell = () => ({
    row: Math.min(selection.row, selection.endRow),
    col: Math.min(selection.col, selection.endCol)
  });

  const selectedRange = () => {
    const a = cellToA1(Math.min(selection.col, selection.endCol), Math.min(selection.row, selection.endRow));
    const b = cellToA1(Math.max(selection.col, selection.endCol), Math.max(selection.row, selection.endRow));
    return a === b ? a : `${a}:${b}`;
  };

  const insertTable = (rows: number, cols: number) => {
    if (!activeSheetId || !workbook) return;
    const { row, col } = anchorCell();
    const range = `${cellToA1(col, row)}:${cellToA1(col + cols - 1, row + rows - 1)}`;
    const sheet = workbook.sheets.find((s) => s.id === activeSheetId);
    applyCommand(
      new AddTableCommand('Insert table', activeSheetId, {
        name: `Table${(sheet?.tables.length ?? 0) + 1}`,
        range,
        headerRow: true,
        style: 'TableStyleMedium9',
        showFilter: true
      } as never)
    );
  };

  const insertChart = (type: 'column' | 'bar' | 'line' | 'pie') => {
    if (!activeSheetId || !workbook) return;
    const sheet = workbook.sheets.find((s) => s.id === activeSheetId);
    if (!sheet) return;
    const { row, col } = anchorCell();
    // Excel expands a single-cell selection to the surrounding data region, so
    // inserting from one cell inside a table still plots the table.
    const isSingle = selection.row === selection.endRow && selection.col === selection.endCol;
    let dataRange: string;
    if (isSingle) {
      const region = currentRegion(sheet, col, row);
      dataRange =
        region.sCol === region.eCol && region.sRow === region.eRow
          ? cellToA1(region.sCol, region.sRow)
          : `${cellToA1(region.sCol, region.sRow)}:${cellToA1(region.eCol, region.eRow)}`;
    } else {
      dataRange = selectedRange();
    }
    applyCommand(
      new AddChartCommand('Insert chart', activeSheetId, {
        type,
        dataRange,
        x: col + 2,
        y: row + 1,
        width: 320,
        height: 220,
        title: 'Chart'
      } as never)
    );
  };

  /* Excel's insert-table grid: hover to size it, click to confirm. */
  const tableGrid = (rect: DOMRect, close: () => void) => (
    <TableGridPopup
      rect={rect}
      onClose={close}
      onPick={(r, c) => insertTable(r, c)}
      onDialog={() => {
        openDialog('insertTable');
        close();
      }}
    />
  );

  const chartMenu = (types: { type: 'column' | 'bar' | 'line' | 'pie'; label: string }[], rect: DOMRect, close: () => void) => (
    <ChartMenuPopup rect={rect} onClose={close} types={types} onPick={(t) => insertChart(t)} />
  );

  return (
    <div className="flex items-stretch gap-0 overflow-x-auto border-b px-1" style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 122 }}>
      <GroupShell caption="Tables">
        <SplitButton icon={<Table2 size={18} />} label="Table" title="Insert table" onOpen={() => insertTable(3, 3)}>
          {tableGrid}
        </SplitButton>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Charts">
        <div className="flex flex-col gap-0.5 pt-0.5">
          <SplitButton
            compact
            icon={<BarChart3 size={15} />}
            label="Column or Bar"
            title="Insert column or bar chart"
            onOpen={() => insertChart('column')}
          >
            {(rect, close) =>
              chartMenu(
                [
                  { type: 'column', label: 'Clustered Column' },
                  { type: 'bar', label: 'Clustered Bar' }
                ],
                rect,
                close
              )
            }
          </SplitButton>
          <SplitButton compact icon={<TrendingUp size={15} />} label="Line" title="Insert line chart" onOpen={() => insertChart('line')}>
            {(rect, close) => chartMenu([{ type: 'line', label: 'Line' }], rect, close)}
          </SplitButton>
          <SplitButton compact icon={<PieChart size={15} />} label="Pie" title="Insert pie chart" onOpen={() => insertChart('pie')}>
            {(rect, close) => chartMenu([{ type: 'pie', label: 'Pie' }], rect, close)}
          </SplitButton>
        </div>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Links">
        <BigButton icon={<Link2 size={18} />} label="Link" title="Insert link" onClick={() => openDialog('insertLink')} />
        <BigButton icon={<MessageSquarePlus size={18} />} label="Comment" title="Insert comment" onClick={() => openDialog('insertComment')} />
      </GroupShell>
    </div>
  );
}

/* Excel's Function Library groups, in ribbon order. Every entry maps to a
   function the engine actually registers. */
const FUNCTION_GROUPS: Array<{ key: string; label: string; icon: React.ReactNode }> = [
  { key: 'Math', label: 'Math & Trig', icon: <Sigma size={14} /> },
  { key: 'Text', label: 'Text', icon: <Type size={14} /> },
  { key: 'Date', label: 'Date & Time', icon: <CalendarDays size={14} /> },
  { key: 'Lookup', label: 'Lookup & Reference', icon: <Search size={14} /> },
  { key: 'Logical', label: 'Logical', icon: <Braces size={14} /> }
];

const AUTOSUM_ITEMS: Array<{ label: string; fn: string }> = [
  { label: 'Sum', fn: 'SUM' },
  { label: 'Average', fn: 'AVERAGE' },
  { label: 'Count Numbers', fn: 'COUNT' },
  { label: 'Max', fn: 'MAX' },
  { label: 'Min', fn: 'MIN' }
];

function DataTab() {
  const openDialog = useUiStore((s) => s.openDialog);
  return (
    <div
      className="flex items-stretch gap-0 overflow-x-auto border-b px-1"
      style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}
    >
      <GroupShell caption="Sort & Filter">
        <BigButton icon={<ArrowUpDown size={18} />} label="Sort" title="Sort range…" onClick={() => openDialog('sort')} />
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Data Tools">
        <BigButton
          icon={<ListChecks size={18} />}
          label="Data Validation"
          title="Data validation…"
          onClick={() => openDialog('dataValidation')}
        />
        <BigButton
          icon={<CopyMinus size={18} />}
          label="Remove Duplicates"
          title="Remove duplicates…"
          onClick={() => openDialog('removeDuplicates')}
        />
      </GroupShell>
    </div>
  );
}

function FormulasTab() {
  const openDialog = useUiStore((s) => s.openDialog);
  const recalculate = useWorkbookStore((s) => s.recalculate);
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selection = useWorkbookStore((s) => s.selection);
  const setCellRaw = useWorkbookStore((s) => s.setCellRaw);
  const startEdit = useWorkbookStore((s) => s.startEdit);
  const showFormulas = useSettingsStore((s) => s.showFormulas);
  const setShowFormulas = useSettingsStore((s) => s.setShowFormulas);

  const anchor = () => ({
    row: Math.min(selection.row, selection.endRow),
    col: Math.min(selection.col, selection.endCol)
  });

  const selectedRange = () => {
    const a = cellToA1(Math.min(selection.col, selection.endCol), Math.min(selection.row, selection.endRow));
    const b = cellToA1(Math.max(selection.col, selection.endCol), Math.max(selection.row, selection.endRow));
    return a === b ? null : `${a}:${b}`;
  };

  /**
   * AutoSum's range guess, same idea as Excel: use the selection, otherwise the
   * run of numbers directly above the cell, otherwise the run to its left.
   */
  const guessSumRange = (): string | null => {
    const sel = selectedRange();
    if (sel) return sel;
    if (!workbook || !activeSheetId) return null;
    const sheet = getSheet(workbook, activeSheetId);
    if (!sheet) return null;
    const { row, col } = anchor();
    const isNumber = (r: number, c: number) => {
      const cell = sheet.cells[cellToA1(c, r)];
      if (!cell) return false;
      const v = cell.formula != null ? cell.computed ?? cell.value : cell.value;
      return typeof v === 'number';
    };
    if (row - 1 >= 1 && isNumber(row - 1, col)) {
      let start = row - 1;
      while (start - 1 >= 1 && isNumber(start - 1, col)) start -= 1;
      return `${cellToA1(col, start)}:${cellToA1(col, row - 1)}`;
    }
    if (col - 1 >= 0 && isNumber(row, col - 1)) {
      let start = col - 1;
      while (start - 1 >= 0 && isNumber(row, start - 1)) start -= 1;
      return `${cellToA1(start, row)}:${cellToA1(col - 1, row)}`;
    }
    return null;
  };

  const autoSum = (fn: string) => {
    const { row, col } = anchor();
    const range = guessSumRange();
    setCellRaw(row, col, range ? `=${fn}(${range})` : `=${fn}()`);
  };

  /** Zero-argument functions commit straight away; the rest drop into edit mode. */
  const pickFunction = (name: string) => {
    const { row, col } = anchor();
    const entry = FUNCTION_LIBRARY.find((f) => f.name === name);
    const zeroArg = entry ? /\(\s*\)/.test(entry.template) : true;
    if (zeroArg) setCellRaw(row, col, `=${name}()`);
    else startEdit(`=${name}()`);
  };

  const functionsIn = (key: string) => FUNCTION_LIBRARY.filter((f) => f.category === key).map((f) => f.name);

  return (
    <div
      className="flex items-stretch gap-0 overflow-x-auto border-b px-1"
      style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}
    >
      <GroupShell caption="Function Library">
        <div className="flex flex-col items-center gap-1">
          <div className="flex items-start gap-1">
            <BigButton
              icon={<Calculator size={18} />}
              label="Insert Function"
              title="Insert function…"
              onClick={() => openDialog('insertFunction')}
            />
            <SplitButton icon={<Sigma size={18} />} label="AutoSum" title="AutoSum" onOpen={() => autoSum('SUM')}>
              {(rect, close) => (
                <ItemMenuPopup
                  rect={rect}
                  onClose={close}
                  items={AUTOSUM_ITEMS.map((a) => a.label)}
                  onPick={(label) => autoSum(AUTOSUM_ITEMS.find((a) => a.label === label)?.fn ?? 'SUM')}
                  extra={{
                    label: 'More Functions…',
                    onClick: () => openDialog('insertFunction')
                  }}
                />
              )}
            </SplitButton>
          </div>
          <div className="flex items-start gap-1">
            {FUNCTION_GROUPS.map((g) => (
              <SplitButton
                key={g.key}
                compact
                icon={g.icon}
                label={g.label}
                title={`Insert a ${g.label} function`}
                onOpen={() => pickFunction(functionsIn(g.key)[0])}
              >
                {(rect, close) => (
                  <ItemMenuPopup rect={rect} onClose={close} items={functionsIn(g.key)} onPick={pickFunction} />
                )}
              </SplitButton>
            ))}
          </div>
        </div>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Formula Auditing">
        <BigButton
          icon={<Eye size={18} />}
          label="Show Formulas"
          title="Show formulas in cells (Ctrl+`)"
          active={showFormulas}
          onClick={() => setShowFormulas(!showFormulas)}
        />
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Calculation">
        <BigButton
          icon={<GraduationCap size={18} />}
          label="Calculate Now"
          title="Recalculate the workbook (F9)"
          onClick={recalculate}
        />
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Defined Names">
        <BigButton
          icon={<ScrollText size={18} />}
          label="Name Manager"
          title="Manage named ranges…"
          onClick={() => openDialog('namedRanges')}
        />
        <BigButton icon={<Tag size={18} />} label="Define Name" title="Define a name…" onClick={() => openDialog('defineName')} />
      </GroupShell>
    </div>
  );
}

function ViewTab() {
  const showGridlines = useSettingsStore((s) => s.showGridlines);
  const showFormulas = useSettingsStore((s) => s.showFormulas);
  const showStatusBar = useSettingsStore((s) => s.showStatusBar);
  const showRuler = useSettingsStore((s) => s.showRuler);
  const setShowGridlines = useSettingsStore((s) => s.setShowGridlines);
  const setShowFormulas = useSettingsStore((s) => s.setShowFormulas);
  const setShowStatusBar = useSettingsStore((s) => s.setShowStatusBar);
  const setShowRuler = useSettingsStore((s) => s.setShowRuler);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const setSelection = useWorkbookStore((s) => s.setSelection);
  const selection = useWorkbookStore((s) => s.selection);
  const workbook = useWorkbookStore((s) => s.workbook);
  const toast = useUiStore((s) => s.toast);
  const zoom = useWorkbookStore((s) => (s.workbook && s.activeSheetId ? getSheet(s.workbook, s.activeSheetId)?.view?.zoom ?? 100 : 100));

  const setZoom = (z: number) => {
    if (activeSheetId) applyCommand(new SetZoomCommand(activeSheetId, clampZoom(z)));
  };
  const isHundred = zoom === 100;

  const sheet = activeSheetId && workbook ? getSheet(workbook, activeSheetId) : null;
  const freeze = sheet?.freeze ?? {};
  const isFrozen = (freeze.row ?? 0) > 0 || (freeze.col ?? 0) > 0;
  const breaks = sheet?.pageSetup?.pageBreaks ?? { rows: [], cols: [] };
  const printArea = sheet?.pageSetup?.printArea;
  const breakPreview = useUiStore((s) => s.pageBreakPreview);
  const setBreakPreview = useUiStore((s) => s.setPageBreakPreview);
  const split = useUiStore((s) => s.split);
  const setSplit = useUiStore((s) => s.setSplit);
  const requestSplit = useUiStore((s) => s.requestSplit);

  const setFreeze = (row: number, col: number) => {
    if (!activeSheetId) return;
    applyCommand(new SetFreezePanesCommand(activeSheetId, { row, col }));
    toast('success', row > 1 || col > 1 ? `Panes frozen at ${cellToA1(Math.max(0, col - 1), row)}` : 'Panes frozen at the active cell');
  };

  return (
    <div className="flex items-stretch gap-0 border-b px-1" style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}>
      <GroupShell caption="Workbook Views">
        <BigButton
          icon={<LayoutGrid size={18} />}
          label="Normal"
          title="Show cells in a normal grid. Resets the zoom to 100% and returns to cell A1."
          active={isHundred && !breakPreview}
          onClick={() => {
            setBreakPreview(false);
            setZoom(100);
            setSelection(SINGLE(1, 0));
          }}
        />
        <BigButton
          icon={<SquareDashed size={18} />}
          label="Page Break Preview"
          title="Show manual page breaks and the print area, and switch the ruler to page units."
          active={breakPreview}
          onClick={() => {
            setBreakPreview(!breakPreview);
            if (!breakPreview) {
              const total = breaks.rows.length + breaks.cols.length;
              toast('info', printArea ? `Print area ${printArea}` : total ? `${total} manual page break(s)` : 'No page breaks set yet');
            }
          }}
        />
      </GroupShell>
      <GroupShell caption="Show">
        <BigButton icon={<Grid3x3 size={18} />} label="Gridlines" title="Show or hide cell gridlines." active={showGridlines} onClick={() => setShowGridlines(!showGridlines)} />
        <BigButton icon={<Type size={18} />} label="Formulas" title="Show formulas in cells instead of their values (Ctrl+`)." active={showFormulas} onClick={() => setShowFormulas(!showFormulas)} />
        <BigButton icon={<AlignVerticalSpaceBetween size={18} />} label="Status Bar" title="Show or hide the status bar." active={showStatusBar} onClick={() => setShowStatusBar(!showStatusBar)} />
        <BigButton icon={<Ruler size={18} />} label="Ruler" title="Show or hide the horizontal and vertical ruler." active={showRuler} onClick={() => setShowRuler(!showRuler)} />
      </GroupShell>
      <GroupShell caption="Window">
        <SplitButton
          icon={<Snowflake size={18} />}
          label="Freeze Panes"
          title="Freeze rows and columns at the active cell."
          onOpen={() => setFreeze(Math.max(1, selection.row), Math.max(1, selection.col))}
        >
          {(rect, close) => (
            <ItemMenuPopup
              rect={rect}
              onClose={close}
              items={['Freeze Panes', 'Freeze Top Row', 'Freeze First Column']}
              onPick={(item) => {
                if (item === 'Freeze Top Row') setFreeze(2, 0);
                else if (item === 'Freeze First Column') setFreeze(0, 2);
                else setFreeze(Math.max(1, selection.row), Math.max(1, selection.col));
              }}
              extra={
                isFrozen
                  ? {
                      label: 'Unfreeze Panes',
                      onClick: () => {
                        if (activeSheetId) applyCommand(new SetFreezePanesCommand(activeSheetId, {}));
                        toast('info', 'Panes unfrozen');
                      }
                    }
                  : undefined
              }
            />
          )}
        </SplitButton>
        <BigButton
          icon={<Columns3 size={18} />}
          label="Split"
          title="Split the window into panes that scroll independently at the active cell."
          active={!!split}
          onClick={() => {
            if (split) {
              setSplit(null);
              toast('info', 'Split removed');
              return;
            }
            const col = Math.max(0, selection.col);
            const row = Math.max(1, selection.row);
            if (col === 0 && row === 1) {
              toast('info', 'Select a cell other than A1 before splitting the window.');
              return;
            }
            // Split and Freeze Panes are mutually exclusive in Excel.
            if (activeSheetId) applyCommand(new SetFreezePanesCommand(activeSheetId, {}));
            // The grid owns the geometry, so it converts the cell into pixels.
            requestSplit(col, row);
            toast('success', `Window split at ${cellToA1(col, row)}`);
          }}
        />
      </GroupShell>
      <GroupShell caption="Zoom">
        <BigButton icon={<Percent size={18} />} label={`${zoom}%`} title="Reset the zoom to 100%." onClick={() => setZoom(100)} />
        <div className="flex flex-col gap-1 self-center" style={{ marginLeft: 6 }}>
          <SmallButton compact icon={<Minus size={15} />} title="Zoom out (decrease by 10%)" onClick={() => setZoom(zoom - 10)} />
          <SmallButton compact icon={<Plus size={15} />} title="Zoom in (increase by 10%)" onClick={() => setZoom(zoom + 10)} />
        </div>
      </GroupShell>
    </div>
  );
}

function HelpTab() {
  const openDialog = useUiStore((s) => s.openDialog);
  return (
    <div className="flex items-stretch gap-0 border-b px-1" style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}>
      <GroupShell caption="Help">
        <BigButton
          icon={<Keyboard size={18} />}
          label="Shortcuts"
          title="View the keyboard shortcuts this workbook supports."
          onClick={() => openDialog('keyboardShortcuts')}
        />
        <BigButton
          icon={<Search size={18} />}
          label="Find"
          title="Find text in the worksheet (Ctrl+F)."
          onClick={() => openDialog('findReplace')}
        />
      </GroupShell>
      <GroupShell caption="Information">
        <BigButton icon={<Info size={18} />} label="About" title="About BritSheet." onClick={() => openDialog('about')} />
        <BigButton icon={<Settings size={18} />} label="Settings" title="Change application settings." onClick={() => openDialog('settings')} />
        <BigButton icon={<Sparkles size={18} />} label="AI Copilot" title="Ask the AI assistant about this workbook." onClick={() => openDialog('ai')} />
      </GroupShell>
    </div>
  );
}

function PageLayoutTab() {
  const openDialog = useUiStore((s) => s.openDialog);
  const toast = useUiStore((s) => s.toast);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const workbook = useWorkbookStore((s) => s.workbook);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const selection = useWorkbookStore((s) => s.selection);
  // Re-render on every workbook mutation so the ribbon reflects stored settings.
  useWorkbookStore((s) => s.revision);

  const setup =
    (activeSheetId && workbook ? getSheet(workbook, activeSheetId)?.pageSetup : undefined) ?? {};
  const margins = { ...DEFAULT_MARGINS, ...(setup.margins ?? {}) };
  const orientation = setup.orientation ?? 'portrait';
  const paperSize = setup.paperSize ?? 'letter';
  const fitToPage = setup.fitToPage === true;
  const printGridlines = setup.printGridlines === true;
  const printHeadings = setup.printHeadings === true;
  const breaks = setup.pageBreaks ?? { rows: [], cols: [] };

  const patch = (p: Partial<PageSetup>, msg: string) => {
    if (!activeSheetId) return;
    applyCommand(new SetPageSetupCommand(activeSheetId, p));
    toast('success', msg);
  };

  const MARGIN_PRESETS: { label: string; values: typeof margins }[] = [
    { label: 'Normal', values: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
    { label: 'Narrow', values: { top: 0.5, bottom: 0.5, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
    { label: 'Moderate', values: { top: 1, bottom: 1, left: 0.75, right: 0.75, header: 0.5, footer: 0.5 } },
    { label: 'Wide', values: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } }
  ];
  const currentPreset =
    MARGIN_PRESETS.find(
      (p) =>
        Math.abs(p.values.top - margins.top) < 0.01 &&
        Math.abs(p.values.left - margins.left) < 0.01 &&
        Math.abs(p.values.bottom - margins.bottom) < 0.01 &&
        Math.abs(p.values.right - margins.right) < 0.01
    )?.label ?? 'Custom';

  const printAreaRef =
    selection.row === selection.endRow && selection.col === selection.endCol
      ? cellToA1(selection.col, selection.row)
      : rangeToA1(selection.col, selection.row, selection.endCol, selection.endRow);

  const insertBreak = (axis: 'row' | 'col') => {
    if (!activeSheetId) return;
    const idx = axis === 'row' ? selection.row : selection.col;
    if (idx <= 1) {
      toast('info', 'Select a cell that is not in the first row or column to insert a page break there.');
      return;
    }
    const next = { rows: [...breaks.rows], cols: [...breaks.cols] };
    const list = axis === 'row' ? next.rows : next.cols;
    if (!list.includes(idx)) list.push(idx);
    list.sort((a, b) => a - b);
    patch({ pageBreaks: next }, `${axis === 'row' ? 'Horizontal' : 'Vertical'} page break added`);
  };

  const removeAllBreaks = () => patch({ pageBreaks: { rows: [], cols: [] } }, 'All page breaks removed');

  const theme = getTheme(workbook?.theme);
  const themeName = theme.name;

  return (
    <div
      className="flex items-stretch gap-0 border-b px-1"
      style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}
    >
      <GroupShell caption="Themes">
        <GalleryButton icon={<Palette size={18} />} label={themeName} title="Themes">
          {(rect, close) => (
            <GalleryPanel menu="themes" rect={rect} width={250}>
              <div className="grid grid-cols-2 gap-1">
                {THEMES.map((t) => (
                  <button
                    key={t.id}
                    type="button"
                    title={`${t.name} theme`}
                    onClick={() => {
                      applyCommand(new SetThemeCommand(t));
                      toast('success', `Theme: ${t.name}`);
                      close();
                    }}
                    className={`flex flex-col items-start gap-1 rounded border p-1 text-left hover:bg-black/5 dark:hover:bg-white/10 ${
                      theme.id === t.id ? 'border-[rgb(var(--accent))]' : ''
                    }`}
                    style={{ borderColor: theme.id === t.id ? 'rgb(var(--accent))' : 'rgb(var(--line-strong))' }}
                  >
                    <span className="flex h-3 w-full overflow-hidden">
                      {t.accents.map((a) => (
                        <span key={a} style={{ background: a, flex: 1 }} />
                      ))}
                    </span>
                    <span className="text-[10px]">{t.name}</span>
                  </button>
                ))}
              </div>
              <RibbonPanelFooter>
                <button
                  type="button"
                  className="w-full rounded px-2 py-1 text-left text-[11px] hover:bg-black/5 dark:hover:bg-white/10"
                  onClick={() => {
                    applyCommand(new ClearThemeCommand());
                    toast('success', 'Theme reset to Office');
                    close();
                  }}
                >
                  Reset to Office theme
                </button>
              </RibbonPanelFooter>
            </GalleryPanel>
          )}
        </GalleryButton>
      </GroupShell>
      <GroupShell caption="Colors">
        <GalleryButton icon={<Palette size={18} />} label="Colors" title="Colors">
          {(rect, close) => (
            <GalleryPanel menu="colors" rect={rect} width={300}>
              <div className="grid grid-cols-3 gap-1">
                {THEME_COLOR_SETS.map((cs) => (
                  <button
                    key={cs.name}
                    type="button"
                    title={`${cs.name} colors`}
                    onClick={() => {
                      applyCommand(new SetThemeCommand({ accents: cs.accents }));
                      toast('success', `Colors: ${cs.name}`);
                      close();
                    }}
                    className="flex flex-col items-start gap-1 rounded border p-1 text-left hover:bg-black/5 dark:hover:bg-white/10"
                    style={{ borderColor: 'rgb(var(--line-strong))' }}
                  >
                    <span className="flex h-3 w-full overflow-hidden">
                      {cs.accents.map((a) => (
                        <span key={a} style={{ background: a, flex: 1 }} />
                      ))}
                    </span>
                    <span className="text-[10px]">{cs.name}</span>
                  </button>
                ))}
              </div>
            </GalleryPanel>
          )}
        </GalleryButton>
      </GroupShell>
      <GroupShell caption="Fonts">
        <GalleryButton icon={<Type size={18} />} label="Fonts" title="Fonts">
          {(rect, close) => (
            <GalleryPanel menu="fonts" rect={rect} width={240}>
              <div className="flex flex-col gap-[2px]">
                {THEME_FONT_SETS.map((f) => (
                  <button
                    key={f.name}
                    type="button"
                    title={`${f.name}: heading ${f.heading}, body ${f.body}`}
                    onClick={() => {
                      applyCommand(new SetThemeCommand({ headingFont: f.heading, bodyFont: f.body }));
                      toast('success', `Fonts: ${f.name}`);
                      close();
                    }}
                    className="flex h-[20px] items-center justify-between rounded border px-2 text-[11px] hover:bg-black/5 dark:hover:bg-white/10"
                    style={{ borderColor: 'rgb(var(--line-strong))' }}
                  >
                    <span style={{ fontFamily: f.heading }}>AaBbCc</span>
                    <span style={{ fontFamily: f.body }} className="text-muted">
                      {f.name}
                    </span>
                  </button>
                ))}
              </div>
            </GalleryPanel>
          )}
        </GalleryButton>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Margins">
        <SplitButton
          icon={<LayoutGrid size={18} />}
          label="Margins"
          title="Set the page margins."
          onOpen={() => openDialog('pageSetup')}
        >
          {(rect, close) => (
            <ItemMenuPopup
              rect={rect}
              onClose={close}
              items={MARGIN_PRESETS.map((p) => p.label)}
              onPick={(label) => {
                const found = MARGIN_PRESETS.find((p) => p.label === label);
                if (found) patch({ margins: found.values }, `Margins: ${label}`);
              }}
              extra={{ label: 'Custom Margins...', onClick: () => openDialog('pageSetup') }}
            />
          )}
        </SplitButton>
        <span className="self-center text-[10px] text-muted">{currentPreset}</span>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Orientation">
        <SplitButton
          icon={<PageIcon orientation={orientation} />}
          label={orientation === 'portrait' ? 'Portrait' : 'Landscape'}
          title="Set the page orientation."
          onOpen={() => patch({ orientation: orientation === 'portrait' ? 'landscape' : 'portrait' }, 'Orientation changed')}
        >
          {(rect, close) => (
            <ItemMenuPopup
              rect={rect}
              onClose={close}
              items={['Portrait', 'Landscape']}
              onPick={(o) => patch({ orientation: o === 'Landscape' ? 'landscape' : 'portrait' }, `Orientation: ${o}`)}
            />
          )}
        </SplitButton>
      </GroupShell>
      <GroupShell caption="Size">
        <SplitButton
          icon={<Paperclip size={18} />}
          label={PAPER_LABEL[paperSize] ?? paperSize}
          title="Set the paper size."
          onOpen={() => openDialog('pageSetup')}
        >
          {(rect, close) => (
            <ItemMenuPopup
              rect={rect}
              onClose={close}
              items={PAPER_SIZES.map((p) => p.label)}
              onPick={(label) => {
                const found = PAPER_SIZES.find((p) => p.label === label);
                if (found) patch({ paperSize: found.id }, `Paper size: ${label}`);
              }}
              extra={{ label: 'More Paper Sizes...', onClick: () => openDialog('pageSetup') }}
            />
          )}
        </SplitButton>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Print Area">
        <SmallButton
          icon={<Printer size={15} />}
          title={`Set the print area to ${printAreaRef}.`}
          onClick={() => patch({ printArea: printAreaRef }, `Print area: ${printAreaRef}`)}
        />
        <SmallButton
          icon={<SquareDashed size={15} />}
          title="Add the current selection to the print area."
          onClick={() => {
            const existing = setup.printArea;
            const next = existing ? `${existing},${printAreaRef}` : printAreaRef;
            patch({ printArea: next }, 'Added to print area');
          }}
          disabled={!setup.printArea}
        />
        <SmallButton
          icon={<X size={15} />}
          title="Clear the print area."
          onClick={() => {
            if (!activeSheetId) return;
            applyCommand(new SetPrintAreaCommand(activeSheetId, undefined));
            toast('info', 'Print area cleared');
          }}
          disabled={!setup.printArea}
        />
      </GroupShell>
      <GroupShell caption="Breaks">
        <SmallButton
          icon={<Scissors size={15} />}
          title="Insert a page break above the active row."
          onClick={() => insertBreak('row')}
        />
        <SmallButton
          icon={<Scissors size={15} className="rotate-90" />}
          title="Insert a page break to the left of the active column."
          onClick={() => insertBreak('col')}
        />
        <SmallButton
          icon={<Minus size={15} />}
          title="Remove all manual page breaks."
          onClick={removeAllBreaks}
          disabled={breaks.rows.length + breaks.cols.length === 0}
        />
        <span className="self-center text-[10px] text-muted">
          {breaks.rows.length + breaks.cols.length === 0 ? 'None' : `${breaks.rows.length + breaks.cols.length} set`}
        </span>
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Scale to Fit">
        <SmallButton
          icon={<ChevronsLeft size={15} />}
          title="Fit the sheet to one page wide (automatic height)."
          onClick={() => patch({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 }, 'Fit to 1 page wide')}
          active={fitToPage && setup.fitToWidth === 1 && (setup.fitToHeight ?? 0) === 0}
        />
        <SmallButton
          icon={<ChevronsRight size={15} />}
          title="Fit the sheet to one page tall (automatic width)."
          onClick={() => patch({ fitToPage: true, fitToHeight: 1, fitToWidth: 0 }, 'Fit to 1 page tall')}
          active={fitToPage && setup.fitToHeight === 1 && (setup.fitToWidth ?? 0) === 0}
        />
        <SmallButton
          icon={<Percent size={15} />}
          title="Set a fixed print scale percentage."
          onClick={() => openDialog('pageSetup')}
        />
      </GroupShell>
      <GroupShell caption="Sheet Options">
        <SmallButton
          icon={<Grid3x3 size={15} />}
          title="Print gridlines."
          onClick={() => patch({ printGridlines: !printGridlines }, printGridlines ? 'Print gridlines off' : 'Print gridlines on')}
          active={printGridlines}
        />
        <SmallButton
          icon={<Type size={15} />}
          title="Print row and column headings."
          onClick={() => patch({ printHeadings: !printHeadings }, printHeadings ? 'Print headings off' : 'Print headings on')}
          active={printHeadings}
        />
      </GroupShell>
    </div>
  );
}

const PAPER_LABEL: Record<string, string> = Object.fromEntries(PAPER_SIZES.map((p) => [p.id, p.label]));

function PageIcon({ orientation }: { orientation: 'portrait' | 'landscape' }) {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6">
      {orientation === 'portrait' ? (
        <rect x="7" y="3" width="10" height="18" rx="1" />
      ) : (
        <rect x="3" y="7" width="18" height="10" rx="1" />
      )}
    </svg>
  );
}

function DrawTab() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const selectedShapeId = useWorkbookStore((s) => s.selectedShapeId);
  const addShape = useWorkbookStore((s) => s.addShape);
  const updateShape = useWorkbookStore((s) => s.updateShape);
  const removeShape = useWorkbookStore((s) => s.removeShape);
  const reorderShape = useWorkbookStore((s) => s.reorderShape);
  const isSheetProtected = useWorkbookStore((s) => s.isSheetProtected);
  const toast = useUiStore((s) => s.toast);

  const sheet = activeSheetId && workbook ? getSheet(workbook, activeSheetId) : null;
  const selected = selectedShapeId ? sheet?.shapes?.find((s) => s.id === selectedShapeId) ?? null : null;
  const protectedNow = isSheetProtected();

  // Shapes are anchored in unscaled sheet pixels; drop new ones in the middle
  // of the sheet area so they are immediately visible and grabbable.
  const insert = (kind: ShapeKind) => {
    if (protectedNow) {
      toast('error', 'This sheet is protected. Unprotect it to insert shapes.');
      return;
    }
    const x = 60 + Math.round(Math.random() * 160);
    const y = 40 + Math.round(Math.random() * 120);
    addShape(kind, x, y, kind === 'textbox' ? 140 : 160, kind === 'textbox' ? 28 : 96);
  };

  const restyle = (patch: Record<string, unknown>, describe: string) => {
    if (!selected) {
      toast('info', 'Select a shape first.');
      return;
    }
    if (protectedNow) {
      toast('error', 'This sheet is protected. Unprotect it to edit shapes.');
      return;
    }
    updateShape(selected.id, patch);
    toast('info', describe);
  };

  const SWATCHES = ['#1f2937', '#dc2626', '#2563eb', '#16a34a', '#eab308', '#9333ea'];

  return (
    <div
      className="flex items-stretch gap-0 border-b px-1"
      style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}
    >
      <GroupShell caption="Shapes">
        <BigButton icon={<Square size={18} />} label="Rectangle" onClick={() => insert('rect')} />
        <BigButton icon={<Circle size={18} />} label="Oval" onClick={() => insert('ellipse')} />
        <BigButton icon={<Minus size={18} />} label="Line" onClick={() => insert('line')} />
        <BigButton icon={<MoveRight size={18} />} label="Arrow" onClick={() => insert('arrow')} />
        <BigButton icon={<Type size={18} />} label="Text Box" onClick={() => insert('textbox')} />
      </GroupShell>
      <GroupShell caption="Shape Fill">
        <div className="flex w-[104px] flex-wrap gap-[3px] px-1">
          {SWATCHES.map((c) => (
            <button
              key={c}
              type="button"
              title={`Shape fill ${c}`}
              disabled={!selected}
              onClick={() => restyle({ fill: c }, `Shape fill set to ${c}.`)}
              className="h-[14px] w-[14px] rounded-[2px] border disabled:opacity-40"
              style={{ background: c, borderColor: 'rgb(var(--line-strong))' }}
            />
          ))}
          <button
            type="button"
            title="No fill"
            disabled={!selected}
            onClick={() => restyle({ fill: undefined }, 'Shape fill removed.')}
            className="h-[14px] w-[14px] rounded-[2px] border disabled:opacity-40"
            style={{ background: 'transparent', borderColor: 'rgb(var(--line-strong))' }}
          >
            <X size={10} className="mx-auto text-[rgb(var(--danger))]" />
          </button>
        </div>
        <div className="flex gap-1 px-1">
          {[1, 2, 4].map((w) => (
            <button
              key={w}
              type="button"
              title={`${w} pt outline`}
              disabled={!selected}
              onClick={() => restyle({ strokeWidth: w }, `Shape outline set to ${w} pt.`)}
              className="flex h-[18px] w-[26px] items-center justify-center rounded-[2px] border border-[rgb(var(--line))] bg-[rgb(var(--ribbon-btn))] disabled:opacity-40"
            >
              <span style={{ height: Math.max(1, w), width: 16, background: 'rgb(var(--text))' }} />
            </button>
          ))}
        </div>
      </GroupShell>

      <GroupShell caption="Shape Outline">
        <div className="flex w-[104px] flex-wrap gap-[3px] px-1">
          {SWATCHES.map((c) => (
            <button
              key={c}
              type="button"
              title={`Shape outline ${c}`}
              disabled={!selected}
              onClick={() => restyle({ stroke: c }, `Shape outline set to ${c}.`)}
              className="h-[14px] w-[14px] rounded-[2px] border disabled:opacity-40"
              style={{ background: c, borderColor: 'rgb(var(--line-strong))' }}
            />
          ))}
          <button
            type="button"
            title="No outline"
            disabled={!selected}
            onClick={() => restyle({ stroke: undefined }, 'Shape outline removed.')}
            className="h-[14px] w-[14px] rounded-[2px] border disabled:opacity-40"
            style={{ background: 'transparent', borderColor: 'rgb(var(--line-strong))' }}
          >
            <X size={10} className="mx-auto text-[rgb(var(--danger))]" />
          </button>
        </div>
      </GroupShell>

      <GroupShell caption="Arrange">
        <div className="flex flex-col gap-1 px-1 pt-1">
          <button
            type="button"
            disabled={!selected}
            onClick={() => {
              reorderShape(selected!.id, 'forward');
              toast('info', 'Shape brought forward.');
            }}
            className="flex h-[22px] items-center gap-1.5 rounded-[2px] border border-[rgb(var(--line))] bg-[rgb(var(--ribbon-btn))] px-2 text-[11px] disabled:opacity-40"
          >
            <ChevronUp size={13} /> Bring Forward
          </button>
          <button
            type="button"
            disabled={!selected}
            onClick={() => {
              reorderShape(selected!.id, 'backward');
              toast('info', 'Shape sent backward.');
            }}
            className="flex h-[22px] items-center gap-1.5 rounded-[2px] border border-[rgb(var(--line))] bg-[rgb(var(--ribbon-btn))] px-2 text-[11px] disabled:opacity-40"
          >
            <ChevronDown size={13} /> Send Backward
          </button>
        </div>
      </GroupShell>

      <GroupShell caption="Shape">
        <BigButton
          icon={<Trash2 size={18} />}
          label="Delete Shape"
          disabled={!selected}
          onClick={() => {
            if (!selected) return;
            removeShape(selected.id);
            toast('info', 'Shape deleted.');
          }}
        />
      </GroupShell>
    </div>
  );
}

function ReviewTab() {
  const openDialog = useUiStore((s) => s.openDialog);
  const toast = useUiStore((s) => s.toast);
  const selection = useWorkbookStore((s) => s.selection);
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const revision = useWorkbookStore((s) => s.revision);
  const deleteNoteAtSelection = useWorkbookStore((s) => s.deleteNoteAtSelection);
  const gotoNote = useWorkbookStore((s) => s.gotoNote);
  const isSheetProtected = useWorkbookStore((s) => s.isSheetProtected);
  const showNotes = useWorkbookStore((s) => s.showNotes);

  const sheet = activeSheetId && workbook ? getSheet(workbook, activeSheetId) : null;
  const activeRef = cellToA1(
    Math.min(selection.col, selection.endCol),
    Math.min(selection.row, selection.endRow)
  );
  const noteHere = sheet?.notes.find((n) => n.ref === activeRef) ?? null;
  const noteCount = sheet?.notes.length ?? 0;
  const isProtected = isSheetProtected();

  const protectToggle = () => {
    if (isProtected) {
      openDialog('unprotectSheet');
      return;
    }
    openDialog('protectSheet');
  };

  const moveNote = (dir: -1 | 1) => {
    if (!gotoNote(dir)) toast('info', dir === 1 ? 'No later comment on this sheet.' : 'No earlier comment on this sheet.');
  };

  return (
    <div
      className="flex items-stretch gap-0 border-b px-1"
      style={{ background: 'rgb(var(--grid-header))', borderColor: 'rgb(var(--line))', height: 104 }}
    >
      <GroupShell caption="Comments">
        <BigButton
          icon={<MessageSquarePlus size={18} />}
          label="New Comment"
          title="Insert a comment on the active cell."
          onClick={() => openDialog('insertComment')}
        />
        <BigButton
          icon={<Trash2 size={18} />}
          label="Delete"
          title="Delete the comment on the active cell."
          onClick={deleteNoteAtSelection}
          disabled={!noteHere}
        />
        <BigButton
          icon={<ChevronUp size={18} />}
          label="Previous"
          title="Go to the previous comment on this sheet."
          onClick={() => moveNote(-1)}
          disabled={noteCount === 0}
        />
        <BigButton
          icon={<ChevronDown size={18} />}
          label="Next"
          title="Go to the next comment on this sheet."
          onClick={() => moveNote(1)}
          disabled={noteCount === 0}
        />
        <BigButton
          icon={<MessageSquare size={18} />}
          label={showNotes ? 'Hide Comments' : 'Show Comments'}
          title="Show or hide the comment markers on the sheet."
          onClick={() => useWorkbookStore.getState().setShowNotes(!showNotes)}
          active={showNotes}
        />
      </GroupShell>
      <GroupDivider />
      <GroupShell caption="Protect">
        <BigButton
          icon={isProtected ? <LockOpen size={18} /> : <Lock size={18} />}
          label={isProtected ? 'Unprotect Sheet' : 'Protect Sheet'}
          title={
            isProtected
              ? 'Allow changes to the contents and formatting of this sheet.'
              : 'Prevent changes to the contents and formatting of this sheet.'
          }
          onClick={protectToggle}
        />
      </GroupShell>
    </div>
  );
}

function GroupDivider() {
  return <div className="mx-1 my-1 w-px bg-[rgb(var(--line-strong))] opacity-60" />;
}

function escapeCsv(v: string): string {
  if (/[",\n\r]/.test(v)) return `"${v.replace(/"/g, '""')}"`;
  return v;
}

export default Header;

