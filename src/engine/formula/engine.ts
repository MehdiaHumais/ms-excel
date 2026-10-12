import { WorkbookDoc, Worksheet, CellPrimitive, ErrorValue } from '../model/types';
import { getSheet, getSheetByName } from '../model/workbook';
import { cellToA1 } from '../model/address';
import { Evaluator, FormulaResult, FormulaSource, isVolatile } from './core';

/** Placeholder values rendered while a computation is in progress. */
export interface FormulaEngine {
  /** Compute a cell's displayed value (formulas evaluated, else raw). */
  compute(sheetId: string, row: number, col: number): CellPrimitive | ErrorValue | null;
  /** Invalidate the whole computation cache. */
  invalidateAll(): void;
  /** Evaluate an expression without storing it in a cell. */
  evaluateExpression(expr: string, sheetId: string): FormulaResult;
}

/**
 * Workbook-backed formula engine.
 *
 * Cells with formulas are evaluated lazily when `compute(...)` is called and
 * the result cached. The whole cache is invalidated on any change; this is
 * correct and simple for typical workbooks. Cross-sheet references and named
 * ranges are supported; circular references resolve to #CIRC!.
 */
export class WorkbookFormulaEngine implements FormulaEngine {
  private wb: WorkbookDoc | null = null;
  private activeSheetId = '';
  private cache = new Map<string, CellPrimitive | ErrorValue | null>();
  private currentEvaluator: Evaluator | null = null;

  private source: FormulaSource = {
    getCell: (sheetName, row, col) => {
      const wb = this.wb;
      if (!wb) return null;
      const sheet = getSheetByName(wb, sheetName);
      if (!sheet) return '#REF!' as ErrorValue;
      const cell = sheet.cells[cellToA1(col, row)];
      if (!cell) return null;
      if (cell.formula != null) {
        const ev = this.currentEvaluator;
        if (ev) {
          const res = ev.evaluateFormulaKeepStack(cell.formula, sheet.name);
          return res.kind === 'v' ? res.value : res.value;
        }
        // Fall back to an isolated evaluation (no circular protection).
        const eval2 = new Evaluator(this.sourceDirect());
        const r2 = eval2.evaluateFormula(cell.formula, sheet.name);
        return r2.kind === 'v' ? r2.value : r2.value;
      }
      return cell.value ?? null;
    },
    resolveSheet: (name) => {
      const wb = this.wb;
      if (!wb) return null;
      const s = getSheetByName(wb, name);
      return s ? s.id : null;
    },
    resolveNamedRange: (name) => {
      const wb = this.wb;
      if (!wb) return null;
      // Excel names are case-insensitive, and SetNamedRangeCommand stores them
      // under an UPPERCASE key, so try the raw name first, then the folded key.
      const nr = wb.namedRanges[name] ?? wb.namedRanges[name.toUpperCase()];
      if (!nr) return null;
      if (!getSheetByName(wb, nr.sheet)) return null;
      return { sheet: nr.sheet, range: nr.range };
    },
    sheetExists: (name) => {
      const wb = this.wb;
      if (!wb) return false;
      return getSheetByName(wb, name) !== undefined;
    }
  };

  private sourceDirect(): FormulaSource {
    return this.source;
  }

  attach(workbook: WorkbookDoc, activeSheetId: string): void {
    this.wb = workbook;
    this.activeSheetId = activeSheetId;
    this.invalidateAll();
  }

  invalidateAll(): void {
    this.cache.clear();
  }

  compute(sheetId: string, row: number, col: number): CellPrimitive | ErrorValue | null {
    const wb = this.wb;
    if (!wb) return null;
    const sheet = getSheet(wb, sheetId);
    if (!sheet) return null;
    const cell = sheet.cells[cellToA1(col, row)];
    if (!cell) return null;
    if (cell.formula == null) return cell.value ?? null;
    const key = `${sheetId}!${row},${col}`;
    if (this.cache.has(key)) return this.cache.get(key) ?? null;
    const evaluator = new Evaluator(this.source);
    this.currentEvaluator = evaluator;
    const volatile = isVolatile(cell.formula);
    const result = evaluator.evaluateFormulaKeepStack(cell.formula, sheet.name);
    this.currentEvaluator = null;
    const val = result.kind === 'v' ? result.value : result.value;
    if (!volatile) this.cache.set(key, val);
    return val;
  }

  evaluateExpression(expr: string, sheetId: string): FormulaResult {
    const wb = this.wb;
    if (!wb) return { kind: 'e', value: '#VALUE!' };
    const sheet = getSheet(wb, sheetId);
    const evaluator = new Evaluator(this.source);
    this.currentEvaluator = evaluator;
    const sheetName = sheet ? sheet.name : (wb.sheets[0]?.name ?? 'Sheet1');
    const result = evaluator.evaluateFormula(expr, sheetName);
    this.currentEvaluator = null;
    return result;
  }

  /** Convenience for tests: evaluate against a constructed workbook. */
  static evaluateWithWorkbook(wb: WorkbookDoc, formula: string, sheetId?: string): FormulaResult {
    const engine = new WorkbookFormulaEngine();
    const sheet = sheetId ? wb.sheets.find((s) => s.id === sheetId) : wb.sheets[0];
    engine.attach(wb, sheet?.id ?? wb.activeSheetId);
    return engine.evaluateExpression(formula, sheet?.id ?? wb.activeSheetId);
  }
}

export function createFormulaEngine(workbook: WorkbookDoc, activeSheetId: string): FormulaEngine {
  const engine = new WorkbookFormulaEngine();
  engine.attach(workbook, activeSheetId);
  return engine;
}

export { Evaluator, isVolatile };