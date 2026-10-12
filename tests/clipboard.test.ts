import { describe, it, expect, beforeEach } from 'vitest';
import { ClipboardData, clipboardHasData, setClipboard, getClipboard } from '@/store/clipboard';

describe('clipboard store', () => {
  beforeEach(() => {
    setClipboard(null);
  });

  it('stores a cell grid and reports content', () => {
    const data: ClipboardData = {
      rows: 2,
      cols: 2,
      cut: false,
      cells: [
        { value: 1 },
        { value: 'x', style: { bold: true } },
        { value: 2 },
        null,
      ],
    };
    setClipboard(data);
    expect(clipboardHasData()).toBe(true);
    const got = getClipboard();
    expect(got?.rows).toBe(2);
    expect(got?.cols).toBe(2);
    expect(got?.cells[1]?.style?.bold).toBe(true);
    expect(got?.cells[3]).toBeNull();
  });

  it('clears the clipboard via null payload', () => {
    setClipboard(null);
    expect(clipboardHasData()).toBe(false);
  });
});