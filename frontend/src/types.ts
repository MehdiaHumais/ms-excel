export interface SheetCreate {
  name: string;
}

export interface SheetResponse {
  id: number;
  name: string;
}

export interface CellResponse {
  row: number;
  col: number;
  value: string | null;
}

export interface CellUpdate {
  row: number;
  col: number;
  value: string | null;
}

export interface BulkCellUpdate {
  cells: CellUpdate[];
}