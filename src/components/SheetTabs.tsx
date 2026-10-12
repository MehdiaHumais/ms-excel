import React from 'react';
import { useWorkbookStore } from '@/store/useWorkbookStore';
import {
  SetSheetColorCommand,
  RenameSheetCommand,
  AddSheetCommand,
  DuplicateSheetCommand,
  RemoveSheetCommand,
  MoveSheetCommand,
  SetSheetStateCommand
} from '@/engine/commands/commands';
import { Worksheet } from '@/engine/model/types';

export function SheetTabs() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const applyCommand = useWorkbookStore((s) => s.applyCommand);
  const setActiveSheet = useWorkbookStore((s) => s.setActiveSheet);

  if (!workbook) return null;

  const sheets = workbook.sheets;

  const addSheet = () => {
    applyCommand(new AddSheetCommand());
  };

  return (
    <div
      className="flex items-center px-1 py-0.5"
      style={{ background: 'rgb(var(--grid-header))', borderTop: '1px solid rgb(var(--line))' }}
      onContextMenu={(e) => e.preventDefault()}
      onClickCapture={(e) => {
        // Same as the ribbon: a clicked tab must not hold keyboard focus, or the
        // grid would refuse to accept typing afterwards.
        const btn = (e.target as HTMLElement).closest?.('button');
        if (btn && e.detail !== 0) btn.blur();
      }}
    >
      <div className="flex flex-1 items-center gap-0.5 overflow-x-auto">
        {sheets.map((s, i) => (
          <SheetTab
            key={s.id}
            sheet={s}
            active={s.id === activeSheetId}
            onClick={() => setActiveSheet(s.id)}
            canMoveLeft={i > 0}
            canMoveRight={i < sheets.length - 1}
            onRename={(name) => applyCommand(new RenameSheetCommand(s.id, name))}
            onColor={(color) => applyCommand(new SetSheetColorCommand(s.id, color))}
            onDuplicate={() => applyCommand(new DuplicateSheetCommand(s.id))}
            onDelete={() => applyCommand(new RemoveSheetCommand(s.id))}
            onMoveLeft={() => applyCommand(new MoveSheetCommand(s.id, i - 1))}
            onMoveRight={() => applyCommand(new MoveSheetCommand(s.id, i + 2))}
            onHide={() => applyCommand(new SetSheetStateCommand(s.id, 'hidden'))}
          />
        ))}
        <button
          className="ml-1 flex h-7 w-8 items-center justify-center rounded text-lg hover:bg-black/5 dark:hover:bg-white/10"
          onClick={addSheet}
          title="Add sheet"
        >
          +
        </button>
      </div>
      <div className="flex items-center gap-1 px-2 text-xs text-muted">
        <span>{sheets.length} sheets</span>
      </div>
    </div>
  );
}

function SheetTab({
  sheet,
  active,
  onClick,
  canMoveLeft,
  canMoveRight,
  onRename,
  onColor,
  onDuplicate,
  onDelete,
  onMoveLeft,
  onMoveRight,
  onHide
}: {
  sheet: Worksheet;
  active: boolean;
  onClick(): void;
  canMoveLeft: boolean;
  canMoveRight: boolean;
  onRename(name: string): void;
  onColor(color: string | undefined): void;
  onDuplicate(): void;
  onDelete(): void;
  onMoveLeft(): void;
  onMoveRight(): void;
  onHide(): void;
}) {
  const [editing, setEditing] = React.useState(false);
  const [draft, setDraft] = React.useState(sheet.name);

  return (
    <div
      onClick={onClick}
      onDoubleClick={() => {
        setDraft(sheet.name);
        setEditing(true);
      }}
      className={`relative flex h-7 max-w-[160px] cursor-pointer items-center px-3 text-[12px] ${
        active
          ? 'text-text'
          : 'text-muted hover:bg-black/5 dark:hover:bg-white/5'
      }`}
      style={{
        borderTop: active ? '2px solid rgb(var(--accent))' : '2px solid transparent',
        background: active ? 'rgb(var(--cell))' : 'transparent',
        borderLeft: '1px solid rgb(var(--line))',
        borderRight: '1px solid rgb(var(--line))'
      }}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
        showSheetMenu({
          x: e.clientX,
          y: e.clientY,
          canMoveLeft,
          canMoveRight,
          onRename: () => {
            setDraft(sheet.name);
            setEditing(true);
          },
          onColor,
          onDuplicate,
          onDelete,
          onMoveLeft,
          onMoveRight,
          onHide
        });
      }}
    >
      {editing ? (
        <input
          autoFocus
          className="w-full bg-cell outline-none"
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={() => {
            if (draft.trim()) onRename(draft.trim());
            setEditing(false);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
            if (e.key === 'Escape') {
              setEditing(false);
            }
          }}
        />
      ) : (
        <span className="truncate">
          <span
            className="mr-1 inline-block h-2.5 w-2.5 rounded-full align-middle"
            style={{ background: sheet.color ?? 'transparent', border: sheet.color ? 'none' : '1px solid rgb(var(--line-strong))' }}
          />
          {sheet.name}
        </span>
      )}
    </div>
  );
}

const COLORS = [undefined, '#d93025', '#e37400', '#f9ab00', '#188038', '#1a73e8', '#7b1fa2', '#5f6368'];

function showSheetMenu(opts: {
  x: number;
  y: number;
  canMoveLeft: boolean;
  canMoveRight: boolean;
  onRename(): void;
  onColor(c: string | undefined): void;
  onDuplicate(): void;
  onDelete(): void;
  onMoveLeft(): void;
  onMoveRight(): void;
  onHide(): void;
}) {
  // Lightweight context menu built inline (no dependency).
  const menu = document.createElement('div');
  menu.className = 'sheet-context-menu fixed z-50 min-w-[150px] rounded-md border bg-cell py-1 shadow-lg text-[13px]';
  menu.style.borderColor = 'rgb(var(--line-strong))';
  menu.style.background = 'rgb(var(--cell))';
  menu.style.left = `${opts.x}px`;
  menu.style.top = `${opts.y}px`;

  const items: Array<{ label: string; onClick(): void; disabled?: boolean }> = [
    { label: 'Rename', onClick: opts.onRename },
    { label: 'Duplicate', onClick: opts.onDuplicate },
    { label: 'Move left', onClick: opts.onMoveLeft, disabled: !opts.canMoveLeft },
    { label: 'Move right', onClick: opts.onMoveRight, disabled: !opts.canMoveRight },
    { label: 'Hide', onClick: opts.onHide },
    { label: 'Delete', onClick: opts.onDelete }
  ];

  for (const item of items) {
    const btn = document.createElement('div');
    btn.textContent = item.label;
    btn.className =
      'px-3 py-1 ' + (item.disabled ? 'opacity-40' : 'cursor-pointer hover:bg-black/5 dark:hover:bg-white/10');
    btn.addEventListener('click', () => {
      menu.remove();
      item.onClick();
    });
    menu.appendChild(btn);
  }

  const colorRow = document.createElement('div');
  colorRow.className = 'flex items-center gap-1 border-t px-2 pt-1.5 mt-1';
  colorRow.style.borderColor = 'rgb(var(--line))';
  for (const c of COLORS) {
    const swatch = document.createElement('button');
    swatch.className = 'h-4 w-4 rounded-full';
    swatch.style.background = c ?? 'transparent';
    swatch.style.border = c ? 'none' : '1px solid rgb(var(--line-strong))';
    swatch.addEventListener('click', () => {
      menu.remove();
      opts.onColor(c);
    });
    colorRow.appendChild(swatch);
  }
  menu.appendChild(colorRow);

  document.body.appendChild(menu);
  const close = () => menu.remove();
  window.setTimeout(() => {
    window.addEventListener('pointerdown', close, { once: true });
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') close();
    });
  }, 0);
}

export default SheetTabs;