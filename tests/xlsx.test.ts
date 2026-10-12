import { describe, it, expect } from 'vitest';
import { createWorkbook } from '@/engine/model/workbook';
import { workbookToXlsxBuffer, workbookFromXlsx } from '@/engine/io/xlsx';

describe('XLSX round-trip', () => {
  it('exports and re-imports values and formulas', async () => {
    const wb = createWorkbook({ title: 'XLSX Test' });
    const sheet = wb.sheets[0];
    sheet.cells.A1 = { value: 10 };
    sheet.cells.A2 = { value: 20 };
    sheet.cells.B1 = { formula: 'SUM(A1:A2)' };

    const bytes = await workbookToXlsxBuffer(wb);
    expect(bytes.length).toBeGreaterThan(100);

    const back = await workbookFromXlsx(bytes);
    const s = back.sheets[0];
    expect(s.cells.A1.value).toBe(10);
    expect(s.cells.A2.value).toBe(20);
    expect(s.cells.B1.formula ?? s.cells.B1.value).toBeDefined();
  });

  it('handles an empty workbook', async () => {
    const wb = createWorkbook();
    const bytes = await workbookToXlsxBuffer(wb);
    const back = await workbookFromXlsx(bytes);
    expect(back.sheets).toHaveLength(1);
  });
});