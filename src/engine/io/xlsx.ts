/**
 * XLSX import/export via exceljs.
 *
 * Round-trips the parts of Excel we model: values, formulas source text,
 * styles (font/bold/italic/size/color, fill, alignment, number formats,
 * borders), column widths / row heights, merged ranges, hidden rows/cols,
 * freeze panes, data validation and table detection is limited by the library.
 *
 * API is async (workbook buffers), matching exceljs. Callers wire this to the
 * file system through Tauri / browser downloads.
 */

import ExcelJS from 'exceljs';
import { WorkbookDoc, Worksheet, Cell, CellPrimitive, CellStyle, Border } from '../model/types';
import { createWorkbook } from '../model/workbook';
import { cellToA1, parseRange, colToLabel } from '../model/address';

export class XlsxImportError extends Error {
  constructor(message: string, public readonly cause?: unknown) {
    super(message);
  }
}

export class XlsxExportError extends Error {
  constructor(message: string, public readonly cause?: unknown) {
    super(message);
  }
}

/* ------------------------------- Import ------------------------------- */

export async function workbookFromXlsx(data: Uint8Array | ArrayBuffer): Promise<WorkbookDoc> {
  const wbWorkbook = new ExcelJS.Workbook();
  try {
    await wbWorkbook.xlsx.load(data as ArrayBuffer);
  } catch (e) {
    throw new XlsxImportError('Could not read the XLSX file.', e);
  }
  if (wbWorkbook.worksheets.length === 0) {
    throw new XlsxImportError('The XLSX file contains no worksheets.');
  }
  const docs: Array<{ name: string }> = wbWorkbook.worksheets.map((ws) => ({ name: ws.name }));
  const wb = createWorkbook({ title: wbWorkbook.title || 'Imported', sheets: docs });
  wb.metadata.author = wbWorkbook.creator ?? '';
  wb.sheets = wbWorkbook.worksheets.map((ws, i) => sheetFromExcelws(ws, wb.sheets[i]?.id ?? `import_${i}`));
  wb.activeSheetId = wb.sheets[wbWorkbook.worksheets[0]?.id ?? 0]?.id ?? wb.sheets[0].id;
  return wb;
}

function sheetFromExcelws(ws: ExcelJS.Worksheet, fallbackId: string): Worksheet {
  const sheet: Worksheet = {
    id: fallbackId,
    name: ws.name,
    state: ws.state === 'hidden' ? 'hidden' : ws.state === 'veryHidden' ? 'veryHidden' : 'visible',
    cells: {},
    rows: {},
    columns: {},
    merges: [],
    charts: [],
    tables: [],
    conditionalFormats: [],
    dataValidations: {},
    notes: [],
    view: { zoom: 100 }
  };

  const maxCol = ws.columnCount || 0;
  const maxRow = ws.rowCount || 0;
  for (let r = 1; r <= Math.max(1, maxRow); r++) {
    const exRow = ws.getRow(r);
    if (exRow.height) sheet.rows[r] = { height: exRow.height };
    if (exRow.hidden) sheet.rows[r] = { ...sheet.rows[r], hidden: true };
  }
  for (let c = 1; c <= Math.max(1, maxCol); c++) {
    const exCol = ws.getColumn(c);
    if (exCol.width) sheet.columns[c - 1] = { width: exCol.width * 7.5 };
    if (exCol.hidden) sheet.columns[c - 1] = { ...sheet.columns[c - 1], hidden: true };
  }

  const merges = (ws.model.merges ?? []).map((m) => {
    const r = m as unknown as { tl: { col: number; row: number }; br: { col: number; row: number } };
    return `${colToLabel(r.tl.col)}${r.tl.row}:${colToLabel(r.br.col)}${r.br.row}`;
  });
  sheet.merges = merges;

  ws.eachRow((row: ExcelJS.Row, rowNumber: number) => {
    if (rowNumber < 1) return;
    row.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
      const model: Cell = {};
      const value = cell.value;
      if (value && typeof value === 'object' && 'formula' in value && typeof value.formula === 'string') {
        model.formula = value.formula.startsWith('=') ? value.formula.slice(1) : value.formula;
        if (value.result !== undefined) model.value = normalizeResult(value.result);
      } else if (value && typeof value === 'object' && 'sharedFormula' in (value as object)) {
        // exceljs gives us raw shared formulas; best-effort expansion is complex,
        // so skip the formula and keep the cached value.
        const sf = value as unknown as { result?: unknown };
        model.formula = undefined;
        if (sf.result !== undefined) model.value = normalizeResult(sf.result);
      } else {
        model.value = normalizeResult(value);
      }
      const style = styleFromExcel(cell);
      if (style) model.style = style;
      sheet.cells[cellToA1(colNumber - 1, rowNumber)] = model;
    });
  });

  const freeze = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number };
  if (freeze?.state === 'frozen') {
    sheet.freeze = {
      row: (freeze.ySplit ?? 0) + 1,
      col: freeze.xSplit ?? 0
    };
  }
  return sheet;
}

function normalizeResult(v: unknown): CellPrimitive {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'boolean' || typeof v === 'number') return v;
  if (v instanceof Date) return v.getTime();
  if (v instanceof Error) return '#VALUE!' as CellPrimitive;
  return null;
}

function styleFromExcel(cell: ExcelJS.Cell): CellStyle | undefined {
  const s: CellStyle = {};
  const f = cell.font;
  if (f) {
    if (f.name) s.fontFamily = f.name;
    if (f.size) s.fontSize = f.size;
    if (f.bold) s.bold = true;
    if (f.italic) s.italic = true;
    if (f.underline) s.underline = true;
    if (f.strike) s.strikethrough = true;
    if (f.color?.argb) s.color = '#' + (f.color.argb as string).slice(2);
  }
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor?.argb) {
    s.backgroundColor = '#' + (fill.fgColor.argb as string).slice(2);
  }
  const al = cell.alignment;
  if (al) {
    if (al.horizontal) s.horizontalAlign = al.horizontal as CellStyle['horizontalAlign'];
    if (al.vertical) s.verticalAlign = al.vertical as CellStyle['verticalAlign'];
    if (al.wrapText) s.wrapText = true;
    if (al.indent) s.indent = al.indent;
  }
  if (cell.numFmt && cell.numFmt !== 'General') s.numberFormat = cell.numFmt;
  const b = cell.border;
  if (b && (b.top || b.left || b.bottom || b.right)) {
    s.borders = {};
    for (const side of ['top', 'left', 'bottom', 'right'] as const) {
      const bd = rawBorder(b[side]);
      if (bd) s.borders[side] = bd;
    }
  }
  return Object.keys(s).length ? s : undefined;
}

function rawBorder(
  b: Partial<ExcelJS.Border> | undefined
): { style?: Exclude<Border['style'], undefined>; color?: string } | undefined {
  if (!b) return undefined;
  const out: { style?: Exclude<Border['style'], undefined>; color?: string } = {};
  if (b.style) out.style = b.style as Exclude<Border['style'], undefined>;
  if (b.color?.argb) out.color = '#' + (b.color.argb as string).slice(2);
  return Object.keys(out).length ? out : undefined;
}

/* ------------------------------- Export ------------------------------- */

export interface XlsxExportOptions {
  /** Use cached computed values when formulas exist but results are unknown. */
  includeFormulas?: boolean;
}

export async function workbookToXlsxBuffer(wb: WorkbookDoc, opts: XlsxExportOptions = {}): Promise<Uint8Array> {
  const wbWorkbook = new ExcelJS.Workbook();
  wbWorkbook.title = wb.metadata.title;
  wbWorkbook.creator = wb.metadata.author || 'BritSheet';
  try {
    for (const sheet of wb.sheets) {
      const ws = wbWorkbook.addWorksheet(sheet.name, {
        state: sheet.state === 'visible' ? undefined : (sheet.state as 'hidden' | 'veryHidden'),
        views:
          sheet.freeze && (sheet.freeze.row || sheet.freeze.col)
            ? [
                {
                  state: 'frozen',
                  ySplit: Math.max(0, (sheet.freeze.row ?? 0) - 1),
                  xSplit: sheet.freeze.col ?? 0
                }
              ]
            : undefined
      });
      writeSheetToExcelws(ws, sheet, opts);
    }
    const buffer = await wbWorkbook.xlsx.writeBuffer();
    return new Uint8Array(buffer);
  } catch (e) {
    throw new XlsxExportError('Could not write the XLSX file.', e);
  }
}

function writeSheetToExcelws(ws: ExcelJS.Worksheet, sheet: Worksheet, opts: XlsxExportOptions): void {
  // Column widths (exceljs is in character units; our values are px-ish).
  for (const [colKey, col] of Object.entries(sheet.columns)) {
    const c = Number(colKey);
    const wbCol = ws.getColumn(c + 1);
    if (col.width) wbCol.width = col.width / 7.5;
    if (col.hidden) wbCol.hidden = true;
  }
  for (const [rowKey, row] of Object.entries(sheet.rows)) {
    const r = Number(rowKey);
    const wbRow = ws.getRow(r);
    if (row.height) wbRow.height = row.height;
    if (row.hidden) wbRow.hidden = true;
  }

  for (const ref of Object.keys(sheet.cells)) {
    const cell = sheet.cells[ref];
    const p = parseRange(ref);
    if (!p) continue;
    const wsCell = ws.getCell(p.sRow, p.sCol + 1);
    writeCell(wsCell, ref, cell, opts);
  }

  for (const m of sheet.merges) {
    const p = parseRange(m);
    if (!p) continue;
    const topLeft = ws.getCell(p.sRow, p.sCol + 1);
    const bottomRight = ws.getCell(p.eRow, p.eCol + 1);
    ws.mergeCells(topLeft.address + ':' + bottomRight.address);
  }
}

function writeCell(wsCell: ExcelJS.Cell, ref: string, model: Cell, opts: XlsxExportOptions): void {
  if (model.formula !== undefined) {
    const cached = model.computed ?? model.value;
    wsCell.value = {
      formula: `=${model.formula}`,
      result: cached ?? undefined
    } as unknown as ExcelJS.CellValue;
  } else if (model.value !== undefined && model.value !== null) {
    if (typeof model.value === 'string') {
      // Numeric-looking strings stay strings (Excel would otherwise parse).
      wsCell.value = model.value;
    } else if (typeof model.value === 'boolean') {
      wsCell.value = model.value;
    } else if (typeof model.value === 'number') {
      // Serial dates: values with a date numberFormat get emitted as numbers too,
      // preserving the raw cell; Excel renders them locally.
      wsCell.value = model.value;
    }
  }
  if (model.style) applyStyle(wsCell, model.style);
}

function applyStyle(wsCell: ExcelJS.Cell, s: CellStyle): void {
  const font: Partial<ExcelJS.Font> = {};
  if (s.fontFamily) font.name = s.fontFamily;
  if (s.fontSize) font.size = s.fontSize;
  if (s.bold) font.bold = true;
  if (s.italic) font.italic = true;
  if (s.underline) font.underline = true;
  if (s.strikethrough) font.strike = true;
  if (s.color) font.color = { argb: s.color.replace('#', 'FF') };
  if (Object.keys(font).length) wsCell.font = font as ExcelJS.Font;

  if (s.backgroundColor) {
    wsCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: s.backgroundColor.replace('#', 'FF') }
    };
  }
  const alignment: Partial<ExcelJS.Alignment> = {};
  if (s.horizontalAlign) alignment.horizontal = s.horizontalAlign as 'left' | 'center' | 'right';
  if (s.verticalAlign) alignment.vertical = s.verticalAlign as 'top' | 'middle' | 'bottom';
  if (s.wrapText) alignment.wrapText = true;
  if (s.indent) alignment.indent = s.indent;
  if (Object.keys(alignment).length) wsCell.alignment = alignment as ExcelJS.Alignment;

  if (s.numberFormat && s.numberFormat !== 'General') {
    wsCell.numFmt = s.numberFormat;
  }
  if (s.borders) {
    const borders: Partial<Record<'top' | 'left' | 'bottom' | 'right', ExcelJS.Border>> = {};
    for (const side of ['top', 'left', 'bottom', 'right'] as const) {
      const bd = s.borders[side] as Border | undefined;
      if (bd) {
        const border: Partial<ExcelJS.Border> = {
          style: (bd.style ?? 'thin') as ExcelJS.BorderStyle
        };
        if (bd.color) border.color = { argb: bd.color.replace('#', 'FF') };
        borders[side] = border as ExcelJS.Border;
      }
    }
    if (Object.keys(borders).length) wsCell.border = borders;
  }
}