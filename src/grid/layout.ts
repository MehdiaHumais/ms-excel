/**
 * Grid geometry: compute pixel positions for columns/rows and which columns/rows
 * are visible in a viewport. Sizes are default-based with sparse per-index
 * overrides so we never iterate the full sheet (1M rows).
 */

import { Worksheet, RowSettings, ColumnSettings } from '@/engine/model/types';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '@/engine/model/types';

export const MIN_COL_WIDTH = 10;
export const MAX_COL_WIDTH = 1024;
export const HEADER_SIZE = 28;
export const MIN_ZOOM = 10;
export const MAX_ZOOM = 400;

/** Clamp a zoom percentage to the range Excel allows. */
export function clampZoom(z: number): number {
  if (!Number.isFinite(z)) return 100;
  return Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(z)));
}

type Sizes = Map<number, number>;

function buildSizes(sheet: Worksheet, kind: 'col' | 'row'): Sizes {
  const map: Sizes = new Map();
  if (kind === 'col') {
    const src = sheet.columns as Record<number, ColumnSettings>;
    for (const key of Object.keys(src)) {
      const idx = Number(key);
      const w = src[idx]?.width;
      if (w !== undefined && w !== null) map.set(idx, Math.max(MIN_COL_WIDTH, Math.min(MAX_COL_WIDTH, w)));
    }
  } else {
    const src = sheet.rows as Record<number, RowSettings>;
    for (const key of Object.keys(src)) {
      const idx = Number(key);
      const h = src[idx]?.height;
      if (h !== undefined && h !== null) map.set(idx, Math.max(4, h));
    }
  }
  return map;
}

export class SheetGeometry {
  readonly colSizes: Sizes;
  readonly rowSizes: Sizes;
  readonly defaultCol = DEFAULT_COL_WIDTH;
  readonly defaultRow = DEFAULT_ROW_HEIGHT;
  /** Zoom scale (1 = 100%). Every pixel output below is multiplied by it. */
  readonly scale: number;

  constructor(sheet: Worksheet, zoom = 100) {
    this.colSizes = buildSizes(sheet, 'col');
    this.rowSizes = buildSizes(sheet, 'row');
    this.scale = clampZoom(zoom) / 100;
  }

  /** Header band thickness in device pixels for the current zoom. */
  get headerSize(): number {
    return HEADER_SIZE * this.scale;
  }

  /** Convert a rendered pixel size back into unscaled model units. */
  toModelWidth(px: number): number {
    return px / this.scale;
  }

  toModelHeight(px: number): number {
    return px / this.scale;
  }

  /** Pixel offset of column `col` from the grid origin (0-based col). */
  colOffset(col: number): number {
    return this.offsetWithDefaults(this.colSizes, col, this.defaultCol) * this.scale;
  }

  rowOffset(row: number): number {
    // rows are 1-based
    return this.offsetWithDefaults(this.rowSizes, row - 1, this.defaultRow) * this.scale;
  }

  rowOffset0(row0: number): number {
    return this.offsetWithDefaults(this.rowSizes, row0, this.defaultRow) * this.scale;
  }

  private offsetWithDefaults(sizes: Sizes, idx: number, def: number): number {
    let total = idx * def;
    for (const [i, size] of sizes) {
      if (i >= idx) break;
      total += size - def;
    }
    return total;
  }

  colWidth(col: number): number {
    return (this.colSizes.get(col) ?? this.defaultCol) * this.scale;
  }

  rowHeight(row: number): number {
    return (this.rowSizes.get(row) ?? this.defaultRow) * this.scale;
  }

  /** Visible col range [first, last] inclusive for a horizontal scroll offset within `width`. */
  visibleCols(scrollLeft: number, width: number): { first: number; last: number } {
    // Resolve indices in unscaled space so they stay zoom-independent.
    return this.visible(this.colSizes, this.defaultCol, scrollLeft / this.scale, width / this.scale, 0);
  }

  visibleRows(scrollTop: number, height: number): { first: number; last: number } {
    const r = this.visible(this.rowSizes, this.defaultRow, scrollTop / this.scale, height / this.scale, 0);
    return { first: r.first + 1, last: r.last + 1 };
  }

  private visible(sizes: Sizes, def: number, offset: number, span: number, minIndex: number): { first: number; last: number } {
    let first = Math.max(minIndex, Math.floor(offset / def));
    // Clamp first to a size-override-aware position (rare cases where an early
    // override shifts offsets before `first`).
    let acc = this.offsetWithDefaults(sizes, first, def);
    while (first > minIndex && acc > offset) {
      first--;
      acc = this.offsetWithDefaults(sizes, first, def);
    }
    // Walk forward accumulating size until we cover the viewport.
    const lastStart = first;
    let cursor = acc;
    let last = first;
    let i = first;
    for (let guard = 0; guard < 512; guard++) {
      if (cursor >= offset + span) break;
      const w = sizes.get(i) ?? def;
      cursor += w;
      last = i;
      i++;
    }
    if (last < Math.max(minIndex, lastStart) || (this.offsetWithDefaults(sizes, first, def) > offset + span && span === 0)) {
      last = first;
    }
    return { first, last };
  }

  /** Convert a scroll offset into the nearest row number (1-based). */
  rowAt(scrollTop: number): number {
    const { first } = this.visible(this.rowSizes, this.defaultRow, scrollTop, 0, 0);
    return first + 1;
  }

  colAt(scrollLeft: number): number {
    const { first } = this.visible(this.colSizes, this.defaultCol, scrollLeft, 0, 0);
    return first;
  }
}

export function isColHidden(sheet: Worksheet, col: number): boolean {
  return sheet.columns[col]?.hidden === true;
}
export function isRowHidden(sheet: Worksheet, row: number): boolean {
  return sheet.rows[row]?.hidden === true;
}