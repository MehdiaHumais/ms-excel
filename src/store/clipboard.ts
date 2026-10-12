/**
 * Application clipboard (in-memory, not OS clipboard yet).
 * Holds a rectangular snapshot of cell value/formula/style so copy/cut/paste
 * work without touching the native clipboard.
 */

import { CellPrimitive, CellStyle } from '@/engine/model/types';

export interface ClipboardCell {
  value?: CellPrimitive;
  formula?: string;
  style?: CellStyle;
}

export interface ClipboardData {
  rows: number;
  cols: number;
  cells: Array<ClipboardCell | null>;
  /** True when the source range should be cleared after paste (cut). */
  cut: boolean;
}

let current: ClipboardData | null = null;

export function setClipboard(data: ClipboardData | null): void {
  current = data;
}

export function getClipboard(): ClipboardData | null {
  return current;
}

export function clipboardHasData(): boolean {
  return current !== null;
}