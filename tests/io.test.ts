import { describe, it, expect } from 'vitest';
import { createWorkbook } from '@/engine/model/workbook';
import { serializeWorkbook, deserializeWorkbook, emptyWorkbook, BRITSHEET_MAGIC, BRITSHEET_VERSION } from '@/engine/io/serialize';
import { parseCsv, sniffCsvDelimiter, workbookToCsvData, csvRowsToWorkbook, coerceCellValue } from '@/engine/io/csv';

describe('native .britsheet serialization', () => {
  it('round-trips a workbook through JSON', () => {
    const wb = createWorkbook({ title: 'Round Trip' });
    const json = serializeWorkbook(wb);
    expect(json.includes('"Round Trip"')).toBe(true);
    const back = deserializeWorkbook(json);
    expect(back.metadata.title).toBe('Round Trip');
    expect(back.sheets).toHaveLength(1);
    expect(back.activeSheetId).toBe(wb.activeSheetId);
  });

  it('serializes to clean JSON (magic header lives in the Rust backend)', () => {
    const wb = createWorkbook({ title: 'M' });
    const text = serializeWorkbook(wb);
    const parsed = JSON.parse(text);
    expect(parsed.metadata.title).toBe('M');
    expect(parsed.sheets).toHaveLength(1);
    // The native container prefix is applied by the Tauri `save_file` command.
    expect(BRITSHEET_MAGIC).toBe('BRITSHEET1');
    expect(BRITSHEET_VERSION).toBe(1);
  });

  it('produces an empty workbook from emptyWorkbook()', () => {
    const empty = emptyWorkbook();
    expect(empty.sheets.length).toBeGreaterThanOrEqual(1);
    expect(typeof empty.metadata.title).toBe('string');
  });

  it('deserialize rejects malformed input', () => {
    expect(() => deserializeWorkbook('not a workbook')).toThrow();
  });
});

describe('CSV parsing / export', () => {
  it('parses quoted CSV with embedded commas and newlines', () => {
    const rows = parseCsv('a,"b,c"\n"d\ne",f\n', ',');
    expect(rows).toEqual([
      ['a', 'b,c'],
      ['d\ne', 'f'],
    ]);
  });

  it('sniffs a semicolon delimiter', () => {
    expect(sniffCsvDelimiter('1;2;3\n4;5;6\n')).toBe(';');
    expect(sniffCsvDelimiter('1,2,3\n4,5,6\n')).toBe(',');
  });

  it('coerces cell text into typed values', () => {
    expect(coerceCellValue('123')).toBe(123);
    expect(coerceCellValue(' 42 ')).toBe(42);
    expect(coerceCellValue('true')).toBe(true);
    expect(coerceCellValue('hello')).toBe('hello');
    expect(coerceCellValue('')).toBeNull();
  });

  it('exports a workbook to CSV rows and imports them back', () => {
    const wb = createWorkbook();
    const csv = workbookToCsvData(wb);
    expect(Array.isArray(csv)).toBe(true);
    const wb2 = csvRowsToWorkbook([['alpha', 'beta'], ['1', '2']]);
    expect(wb2.sheets[0].cells.A1.value).toBe('alpha');
    expect(wb2.sheets[0].cells.B2.value).toBe(2);
  });
});