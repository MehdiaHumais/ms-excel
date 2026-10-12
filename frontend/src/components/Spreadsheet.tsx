import React, { useEffect, useState } from 'react';
import { SheetResponse, CellResponse, BulkCellUpdate } from '../types';
import './Spreadsheet.css';

type Props = {
  sheet: SheetResponse;
};

const API_BASE = 'http://localhost:8000';

const Spreadsheet: React.FC<Props> = ({ sheet }) => {
  const [cells, setCells] = useState<CellResponse[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fetchCells = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API_BASE}/sheets/${sheet.id}/cells`);
      if (!res.ok) throw new Error('Failed to fetch cells');
      const data: CellResponse[] = await res.json();
      setCells(data);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchCells();
  }, [sheet.id]);

  const handleCellChange = (row: number, col: number, value: string) => {
    setCells((prev) =>
      prev.map((c) =>
        c.row === row && c.col === col ? { ...c, value } : c
      )
    );
  };

  const saveChanges = async () => {
    const payload: BulkCellUpdate = {
      cells: cells.map((c) => ({
        row: c.row,
        col: c.col,
        value: c.value ?? '',
      })),
    };
    try {
      const res = await fetch(`${API_BASE}/sheets/${sheet.id}/cells`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Failed to save cells');
      }
      await fetchCells(); // refresh
    } catch (e: any) {
      setError(e.message);
    }
  };

  // Simple 10x10 grid for demo
  const rows = Array.from({ length: 10 }, (_, i) => i);
  const cols = Array.from({ length: 10 }, (_, i) => i);

  return (
    <div className="spreadsheet">
      {loading && <p>Loading...</p>}
      {error && <p className="error">{error}</p>}
      <table>
        <thead>
          <tr>
            <th></th>
            {cols.map((c) => (
              <th key={c}>{String.fromCharCode(65 + c)}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r}>
              <th>{r + 1}</th>
              {cols.map((c) => {
                const cell = cells.find((c) => c.row === r && c.col === c);
                return (
                  <td key={c}>
                    <input
                      type="text"
                      value={cell?.value ?? ''}
                      onChange={(e) =>
                        handleCellChange(r, c, e.target.value)
                      }
                      className="cell-input"
                    />
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      <button className="primary" onClick={saveChanges}>
        Save
      </button>
    </div>
  );
};

export default Spreadsheet;