/**
 * Lightweight DOM context menu (no dependencies), mirroring the styling of the
 * sheet-tab right-click menu but reusable anywhere (cells, headers, charts...).
 *
 * Usage:
 *   showContextMenu(items, e.clientX, e.clientY);
 *
 * The menu is appended to document.body as a fixed overlay and closes on:
 * pointerdown elsewhere, Escape, window resize, or a second showContextMenu.
 */

export type ContextMenuItem =
  | { type: 'separator' }
  | {
      label: string;
      hint?: string;
      onClick?: () => void;
      disabled?: boolean;
      danger?: boolean;
    };

let current: HTMLDivElement | null = null;

export function closeContextMenu(): void {
  if (current) {
    current.remove();
    current = null;
  }
}

export function showContextMenu(items: ContextMenuItem[], x: number, y: number): void {
  closeContextMenu();
  const menu = document.createElement('div');
  menu.className =
    'sheet-context-menu fixed z-[100] min-w-[170px] rounded-md border bg-cell py-1 text-[13px] shadow-lg';
  menu.style.borderColor = 'rgb(var(--line-strong))';
  menu.style.background = 'rgb(var(--cell))';
  menu.style.left = '0px';
  menu.style.top = '0px';

  const onOutside = (ev: Event) => {
    if (menu.contains(ev.target as Node)) return;
    closeContextMenu();
  };
  const onEscape = (ev: KeyboardEvent) => {
    if (ev.key === 'Escape') closeContextMenu();
  };
  const closeAll = () => {
    closeContextMenu();
  };

  for (const raw of items) {
    if ((raw as { type?: string }).type === 'separator') {
      const hr = document.createElement('div');
      hr.className = 'my-1 border-t';
      hr.style.borderColor = 'rgb(var(--line))';
      menu.appendChild(hr);
      continue;
    }
    const item = raw as { label: string; hint?: string; onClick?: () => void; disabled?: boolean; danger?: boolean };
    const row = document.createElement('div');
    row.className =
      'flex items-center justify-between gap-3 px-3 py-1 leading-tight ' +
      (item.disabled
        ? 'opacity-40'
        : 'cursor-pointer hover:bg-black/5 dark:hover:bg-white/10');
    row.style.color = item.danger ? 'rgb(var(--danger))' : '';
    const label = document.createElement('span');
    label.textContent = item.label;
    row.appendChild(label);
    if (item.hint) {
      const hint = document.createElement('span');
      hint.textContent = item.hint;
      hint.style.color = 'rgb(var(--text-muted))';
      hint.style.fontSize = '11px';
      row.appendChild(hint);
    }
    row.addEventListener(
      'click',
      () => {
        closeContextMenu();
        if (!item.disabled && item.onClick) item.onClick();
      },
      { once: true }
    );
    menu.appendChild(row);
  }

  // Keep the menu on screen (menu opens below/right of the cursor by default).
  document.body.appendChild(menu);
  const rect = menu.getBoundingClientRect();
  const px = Math.min(x, Math.max(8, window.innerWidth - rect.width - 8));
  const py = Math.min(y, Math.max(8, window.innerHeight - rect.height - 8));
  menu.style.left = `${px}px`;
  menu.style.top = `${py}px`;

  window.addEventListener('pointerdown', onOutside, true);
  window.addEventListener('click', onOutside, true);
  window.addEventListener('contextmenu', closeAll, true);
  window.addEventListener('keydown', onEscape, true);
  window.addEventListener('resize', closeAll, true);
  window.addEventListener('blur', closeAll, true);
  window.addEventListener('scroll', closeAll, true);

  current = menu;
  const removeListeners = () => {
    window.removeEventListener('pointerdown', onOutside, true);
    window.removeEventListener('click', onOutside, true);
    window.removeEventListener('contextmenu', closeAll, true);
    window.removeEventListener('keydown', onEscape, true);
    window.removeEventListener('resize', closeAll, true);
    window.removeEventListener('blur', closeAll, true);
    window.removeEventListener('scroll', closeAll, true);
  };
  menu.addEventListener('DOMNodeRemoved', () => {
    removeListeners();
    if (current === menu) current = null;
  });
}