/**
 * Cell display logic: turn a cell's computed/raw value plus its style into the
 * text that the grid renders, using the number-format engine. Also decides
 * horizontal alignment defaults (numbers right, boolean center, errors left).
 */

import { Cell, CellPrimitive, ErrorValue, isErrorValue } from '@/engine/model/types';
import { formatValue, isDateFormat } from '@/engine/format/numberFormat';
import { serialToDate } from '@/engine/model/serial';

export interface DisplayResult {
  text: string;
  align: 'left' | 'center' | 'right';
  isFormula: boolean;
  value: CellPrimitive | ErrorValue | null;
}

export function displayCell(
  cell: Cell | undefined,
  showFormulas: boolean,
  computed?: CellPrimitive | ErrorValue | null | undefined
): DisplayResult {
  if (!cell) return { text: '', align: 'left', isFormula: false, value: null };

  const isFormula = cell.formula != null;
  if (isFormula && showFormulas) {
    return {
      text: '=' + cell.formula,
      align: 'left',
      isFormula: true,
      value: cell.value ?? null
    };
  }

  const value: CellPrimitive | ErrorValue | null =
    isFormula ? (computed !== undefined ? computed ?? null : cell.value ?? null) : (cell.value ?? null);

  if (value == null) return { text: '', align: 'left', isFormula, value: null };

  if (isErrorValue(value)) {
    return { text: value, align: 'left', isFormula, value };
  }

  if (typeof value === 'boolean') {
    return { text: value ? 'TRUE' : 'FALSE', align: 'center', isFormula, value };
  }

  if (typeof value === 'number') {
    const fmt = cell.style?.numberFormat;
    let text: string;
    if (fmt) {
      text = formatValue(value, fmt);
    } else if (isDateLikeDefault(value)) {
      text = formatValue(value, 'm/d/yyyy');
    } else {
      text = formatValue(value, 'General');
    }
    return { text, align: 'right', isFormula, value };
  }

  // string
  return { text: value, align: (cell.style?.horizontalAlign as 'left' | 'right' | 'center') ?? 'left', isFormula, value };
}

const DATE_EPOCH = 24107; // 1966-01-01 serial; before this, treat as plain number.

function isDateLikeDefault(n: number): boolean {
  return n > DATE_EPOCH && n < 800000;
}

export function computedValueForCell(cell: Cell | undefined): CellPrimitive | ErrorValue | null {
  return cell?.value ?? null;
}