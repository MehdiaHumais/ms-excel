/**
 * UI state that isn't the workbook: active view (home vs workspace), dialogs,
 * recent files, toasts.
 */

import { create } from 'zustand';

export type ViewMode = 'home' | 'workspace';

export type DialogName =
  | null
  | 'open'
  | 'saveAs'
  | 'formatCells'
  | 'findReplace'
  | 'sort'
  | 'namedRanges'
  | 'defineName'
  | 'dataValidation'
  | 'removeDuplicates'
  | 'insertFunction'
  | 'insertTable'
  | 'insertLink'
  | 'insertComment'
  | 'protectSheet'
  | 'unprotectSheet'
  | 'pageSetup'
  | 'settings'
  | 'about'
  | 'keyboardShortcuts'
  | 'ai';

export interface RecentFile {
  path: string;
  name: string;
  modified: number;
}

interface Toast {
  id: number;
  kind: 'info' | 'success' | 'error';
  message: string;
}

interface UiState {
  view: ViewMode;
  dialog: DialogName;
  dialogPayload: Record<string, unknown>;
  recentFiles: RecentFile[];
  toasts: Toast[];
  undoCount: number;
  redoCount: number;
  /** View > Page Break Preview: show break lines and the print area boundary. */
  pageBreakPreview: boolean;
  /** View > Split: pixel size of the top-left pane, or null when not split. */
  split: { x: number; y: number } | null;
  /** Pending View > Split request; the grid turns it into pixel sizes. */
  splitRequest: { col: number; row: number } | null;

  setView(v: ViewMode): void;
  setPageBreakPreview(b: boolean): void;
  setSplit(next: { x: number; y: number } | null): void;
  requestSplit(col: number, row: number): void;
  clearSplitRequest(): void;
  openDialog(name: Exclude<DialogName, null>, payload?: Record<string, unknown>): void;
  closeDialog(): void;
  addRecent(file: RecentFile): void;
  removeRecent(path: string): void;
  toast(kind: Toast['kind'], message: string): void;
  dismissToast(id: number): void;
  bumpHistory(): void;
}

const RECENT_KEY = 'britsheet.recent.v1';

function loadRecent(): RecentFile[] {
  try {
    const raw = localStorage.getItem(RECENT_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as RecentFile[];
    if (!Array.isArray(parsed)) return [];
    return parsed.slice(0, 12);
  } catch {
    return [];
  }
}

function persistRecent(files: RecentFile[]): void {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(files.slice(0, 12)));
  } catch {
    /* ignore */
  }
}

let toastId = 0;

export const useUiStore = create<UiState>((set, get) => ({
  view: 'workspace',
  dialog: null,
  dialogPayload: {},
  recentFiles: loadRecent(),
  toasts: [],
  undoCount: 0,
  redoCount: 0,
  pageBreakPreview: false,
  split: null,
  splitRequest: null,

  setView(view) {
    set({ view });
  },

  setPageBreakPreview(b) {
    set({ pageBreakPreview: b });
  },

  setSplit(next) {
    set({ split: next });
  },

  requestSplit(col, row) {
    set({ splitRequest: { col, row } });
  },

  clearSplitRequest() {
    set({ splitRequest: null });
  },

  openDialog(dialog, payload = {}) {
    set({ dialog, dialogPayload: payload });
  },

  closeDialog() {
    set({ dialog: null, dialogPayload: {} });
  },

  addRecent(file) {
    const next = [file, ...get().recentFiles.filter((f) => f.path !== file.path)];
    persistRecent(next);
    set({ recentFiles: next });
  },

  removeRecent(path) {
    const next = get().recentFiles.filter((f) => f.path !== path);
    persistRecent(next);
    set({ recentFiles: next });
  },

  toast(kind, message) {
    const id = ++toastId;
    const toasts = [...get().toasts, { id, kind, message }].slice(-4);
    set({ toasts });
    window.setTimeout(() => get().dismissToast(id), 4200);
  },

  dismissToast(id) {
    set({ toasts: get().toasts.filter((t) => t.id !== id) });
  },

  bumpHistory() {
    const { undoCount, redoCount } = get();
    set({ undoCount: undoCount + 1, redoCount: redoCount + 1 });
  }
}));