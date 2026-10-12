/**
 * BritSheet core data model types.
 *
 * The model is deliberately framework-independent: UI layers read and mutate a
 * plain-typed Workbook, and all serialization (native .britsheet, XLSX, CSV)
 * maps to and from these types.
 */

import type { ThemeDef } from './theme';

/** Type of a cell. */
export type CellPrimitive = number | string | boolean | null;

/** Spreadsheet error values, stored and displayed like Excel. */
export type ErrorValue =
  | '#DIV/0!'
  | '#NAME?'
  | '#VALUE!'
  | '#REF!'
  | '#NUM!'
  | '#N/A'
  | '#NULL!'
  | '#CIRC!'
  | '#SPILL!';

export const ALL_ERRORS: ErrorValue[] = [
  '#DIV/0!',
  '#NAME?',
  '#VALUE!',
  '#REF!',
  '#NUM!',
  '#N/A',
  '#NULL!',
  '#CIRC!',
  '#SPILL!'
];

export function isErrorValue(v: unknown): v is ErrorValue {
  return typeof v === 'string' && (ALL_ERRORS as string[]).includes(v);
}

/** Border style for a side. */
export interface Border {
  style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
  color?: string;
}

export interface CellBorders {
  top?: Border;
  right?: Border;
  bottom?: Border;
  left?: Border;
}

/**
 * Cell visual + semantic formatting. All fields optional so partial updates
 * are trivial. Formatting belongs to the model, not only to React.
 */
export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  backgroundColor?: string;
  /** 0-based accent index; overrides the literal color so themes repaint it. */
  colorTheme?: number;
  backgroundTheme?: number;
  horizontalAlign?: 'left' | 'center' | 'right' | 'justify' | 'fill';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: string;
  borders?: CellBorders;
  indent?: number;
  textRotation?: number;
  locked?: boolean;
}

export type HorizontalAlign = NonNullable<CellStyle['horizontalAlign']>;
export type VerticalAlign = NonNullable<CellStyle['verticalAlign']>;

/** One cell. Value + optional formula + optional style. */
export interface Cell {
  value?: CellPrimitive;
  /** Formula text WITHOUT the leading '='. */
  formula?: string;
  /** Cached computed display value, recomputed by the formula engine. */
  computed?: CellPrimitive | ErrorValue;
  style?: CellStyle;
  note?: string;
  hyperlink?: string;
}

export interface RowSettings {
  height?: number;
  hidden?: boolean;
}

export interface ColumnSettings {
  width?: number;
  hidden?: boolean;
}

export interface MergeDef {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export type SheetState = 'visible' | 'hidden' | 'veryHidden';

export type ChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'pie'
  | 'area'
  | 'scatter';

export interface ChartDef {
  id: string;
  type: ChartType;
  /** Source range in A1 form within this sheet. */
  dataRange?: string;
  /** Optional title; optional secondary category range (scatter). */
  title?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  options?: {
    showLegend?: boolean;
    stacked?: boolean;
    fillColor?: string;
    gridlines?: boolean;
  };
}

export type TableStyleName =
  | 'TableStyleLight1'
  | 'TableStyleLight9'
  | 'TableStyleMedium2'
  | 'TableStyleMedium9'
  | 'TableStyleDark1'
  | 'None';

export interface TableDef {
  id: string;
  name: string;
  /** A1:Tn range including header. */
  range: string;
  headerRow: boolean;
  totalRow?: boolean;
  style: TableStyleName;
  showFilter?: boolean;
}

export type ConditionalFormatRule =
  | { type: 'cellValue'; operator: '>' | '<' | '>=' | '<=' | '=' | '!='; value: CellPrimitive; style: CellStyle; range: string }
  | { type: 'duplicates'; style: CellStyle; range: string; unique?: boolean }
  | { type: 'blank'; style: CellStyle; range: string; notBlank?: boolean }
  | { type: 'formula'; formula: string; style: CellStyle; range: string };

export type DataValidationType =
  | 'list'
  | 'whole'
  | 'decimal'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom'
  | 'none';

export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

export interface DataValidationRule {
  type: DataValidationType;
  operator?: DataValidationOperator;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  errorTitle?: string;
  errorMessage?: string;
  showErrorMessage?: boolean;
  promptTitle?: string;
  promptMessage?: string;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paperSize?: string;
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  margins?: { top: number; bottom: number; left: number; right: number; header: number; footer: number };
  printArea?: string;
  printTitles?: { rows?: string; cols?: string };
  header?: string;
  footer?: string;
  fitToPage?: boolean;
  /** Manual page breaks: 1-based row/col indices the break sits *above/left of*. */
  pageBreaks?: { rows: number[]; cols: number[] };
  printGridlines?: boolean;
  printHeadings?: boolean;
}

export interface FreezePane {
  /** First non-frozen row (1-based), 0 = none frozen. */
  row?: number;
  /** First non-frozen column (1-based), 0 = none frozen. */
  col?: number;
}

export interface NoteDef {
  ref: string;
  text: string;
  author?: string;
}

/** A drawing object (Draw tab). Positioned in unscaled sheet pixels. */
export type ShapeKind = 'rect' | 'ellipse' | 'line' | 'arrow' | 'textbox';

export interface ShapeDef {
  id: string;
  kind: ShapeKind;
  x: number;
  y: number;
  width: number;
  height: number;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  text?: string;
  fontSize?: number;
  /** Draw order; higher is in front. */
  z?: number;
}

export interface NamedRange {
  name: string;
  sheet: string;
  range: string;
  comment?: string;
}

export interface SheetViewState {
  zoom: number;
  activeCell?: string;
  scrollTop?: number;
  scrollLeft?: number;
  /** Show note markers/tooltips (Excel: "Show Comments"). */
  showNotes?: boolean;
}

export interface Worksheet {
  id: string;
  name: string;
  color?: string;
  state: SheetState;
  /** Sparse map A1 ref -> Cell. */
  cells: Record<string, Cell>;
  /** 1-based row index -> settings. */
  rows: Record<number, RowSettings>;
  /** 0-based column index -> settings. */
  columns: Record<number, ColumnSettings>;
  /** Merged ranges stored as normalized "A1:B2" strings. */
  merges: string[];
  charts: ChartDef[];
  shapes?: ShapeDef[];
  tables: TableDef[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: Record<string, DataValidationRule>;
  notes: NoteDef[];
  autofilter?: string;
  freeze?: FreezePane;
  view?: SheetViewState;
  pageSetup?: PageSetup;
  gridlines?: boolean;
  showRowColHeaders?: boolean;
  /** Sheet is protected: editing and formatting are blocked. */
  protected?: boolean;
}

export interface WorkbookMetadata {
  title: string;
  author: string;
  created: number;
  modified: number;
  app: string;
  /** Stored for future compatibility / migration. */
  fileFormatVersion: number;
}

export interface WorkbookDoc {
  /** Schema version for the in-memory model. */
  version: number;
  metadata: WorkbookMetadata;
  sheets: Worksheet[];
  activeSheetId: string;
  namedRanges: Record<string, NamedRange>;
  /** Workbook-level settings that ride along with the file. */
  settings: {
    calculationMode?: 'automatic' | 'manual';
    iteration?: boolean;
    maxIterations?: number;
    standardFont?: string;
    standardFontSize?: number;
  };
  printSettings?: PageSetup;
  /** Page Layout > Themes / Colors / Fonts. */
  theme?: ThemeDef;
}

export const SHEET_SIZE = {
  rows: 1_048_576,
  cols: 16_384
};

export const DEFAULT_COL_WIDTH = 100;
export const DEFAULT_ROW_HEIGHT = 24;
export const DEFAULT_FONT_FAMILY = 'Segoe UI';
export const DEFAULT_FONT_SIZE = 12;