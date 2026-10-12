/**
 * Serialization for the native .britsheet format.
 *
 * The on-disk format is a single JSON payload describing the whole Workbook.
 * Compression and atomic write are handled by the Rust backend; in the browser
 * (dev / web build) we fall back to plain JSON for save/open.
 *
 * A magic header marks real BritSheet files: the 10-byte ASCII string
 * "BRITSHEET1" followed by a version byte, then the (optionally gzipped) JSON.
 * Those bytes are appended by the Tauri command `save_file`; here we only
 * produce/consume the JSON document.
 */

import { WorkbookDoc } from '../model/types';
import { cloneWorkbook } from '../model/workbook';

export const BRITSHEET_MAGIC = 'BRITSHEET1';
export const BRITSHEET_VERSION = 1;
export const BRITSHEET_EXTENSION = '.britsheet';

/** Version migration entry points. */
const MIGRATORS: Array<(doc: unknown) => unknown> = [
  // v0 -> v1: nothing yet; kept as the seam for future format evolution.
  (doc) => doc
];

export function serializeWorkbook(wb: WorkbookDoc): string {
  // Clean the model: drop transient 'computed' fields so files stay small and
  // deterministic.
  const clean = cloneWorkbook(wb);
  for (const sheet of clean.sheets) {
    for (const ref of Object.keys(sheet.cells)) {
      const cell = sheet.cells[ref];
      if (cell && 'computed' in cell) delete (cell as Record<string, unknown>).computed;
    }
  }
  clean.metadata.modified = Date.now();
  return JSON.stringify(clean);
}

export function deserializeWorkbook(json: string): WorkbookDoc {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('BritSheet file is not valid JSON.');
  }
  if (!isWorkbookLike(parsed)) {
    throw new Error('BritSheet file does not look like a workbook document.');
  }
  const migrated = MIGRATORS.reduce((acc, fn) => fn(acc), parsed as unknown);
  const wb = migrated as WorkbookDoc;
  // Recompute the document schema version on load; keep existing metadata
  // unless it's missing required fields.
  wb.version = BRITSHEET_VERSION;
  wb.metadata = {
    ...(wb.metadata ?? {}),
    fileFormatVersion: BRITSHEET_VERSION,
    app: 'BritSheet',
    modified: Date.now()
  };
  wb.metadata.title = wb.metadata.title || 'Untitled';
  wb.metadata.author = wb.metadata.author ?? '';
  wb.metadata.created = wb.metadata.created ?? Date.now();
  wb.sheets = Array.isArray(wb.sheets) && wb.sheets.length > 0 ? wb.sheets : [];
  if (wb.sheets.length === 0) {
    throw new Error('BritSheet file contains no sheets.');
  }
  if (!wb.activeSheetId || !wb.sheets.some((s) => s.id === wb.activeSheetId)) {
    wb.activeSheetId = wb.sheets[0].id;
  }
  wb.namedRanges = wb.namedRanges ?? {};
  wb.settings = wb.settings ?? {};
  return normalize(wb);
}

function isWorkbookLike(v: unknown): v is Record<string, unknown> {
  return (
    typeof v === 'object' &&
    v !== null &&
    'sheets' in v &&
    Array.isArray((v as Record<string, unknown>).sheets)
  );
}

/** Fill in anything older files might lack so the rest of the app can rely on a complete shape. */
function normalize(wb: WorkbookDoc): WorkbookDoc {
  for (const sheet of wb.sheets) {
    sheet.cells = sheet.cells ?? {};
    sheet.rows = sheet.rows ?? {};
    sheet.columns = sheet.columns ?? {};
    sheet.merges = sheet.merges ?? [];
    sheet.charts = sheet.charts ?? [];
    sheet.shapes = sheet.shapes ?? [];
    sheet.tables = sheet.tables ?? [];
    sheet.conditionalFormats = sheet.conditionalFormats ?? [];
    sheet.dataValidations = sheet.dataValidations ?? {};
    sheet.notes = sheet.notes ?? [];
    sheet.state = sheet.state ?? 'visible';
    sheet.gridlines = sheet.gridlines ?? true;
    sheet.showRowColHeaders = sheet.showRowColHeaders ?? true;
    sheet.view = sheet.view ?? { zoom: 100 };
  }
  return wb;
}

/** Build a fresh placeholder document for new/empty workbooks. */
export function emptyWorkbook(title = 'Untitled'): WorkbookDoc {
  return cloneWorkbook(createFreshWorkbook(title));
}

// Local, dependency-free version of createWorkbook to avoid a cycle; kept
// minimal and in sync with engine/model/workbook.
import { createWorkbook as modelCreateWorkbook } from '../model/workbook';
function createFreshWorkbook(title: string): WorkbookDoc {
  return modelCreateWorkbook({ title });
}