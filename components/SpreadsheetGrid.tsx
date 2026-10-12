'use client';

import { useState, useEffect } from 'react';

type Cell = {
  value: string;
};

type GridData = Cell[][];

export default function SpreadsheetGrid() {
  const [grid, setGrid] = useState<GridData>([]);
  const [selected, setSelected] = useState<{ row: number; col: number } | null>(null);
  const rows = 20;
  const cols = 10;

  // Initialize grid with empty cells
  useEffect(() => {
    const initial: GridData = Array.from({ length: rows }, () =>
      Array.from({ length: cols }, () => ({ value: '' }))
    );
    setGrid(initial);
  }, []);

  const handleCellClick = (row: number, col: number) => {
    setSelected({ row, col });
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>, row: number, col: number) => {
    const newValue = e.target.value;
    setGrid((prev) => {
      const newGrid = prev.map((r) => r.slice());
      newGrid[row][col] = { value: newValue };
      return newGrid;
    });
  };

  if (grid.length === 0) {
    return <p className="p-4">Loading grid...</p>;
  }

  return (
    <div className="overflow-auto border border-gray-300 rounded-2xl shadow-card">
      <table className="w-full border-collapse">
        <thead>
          <tr>
            <th className="w-10"></th>
            {Array.from({ length: cols }, (_, c) => (
              <th key={c} className="px-2 py-1 bg-gray-100 border border-gray-300 text-center">
                {String.fromCharCode(65 + c)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {grid.map((rowData, r) => (
            <tr key={r}>
              <th className="px-2 py-1 bg-gray-100 border border-gray-300 text-center">{r + 1}</th>
              {rowData.map((cell, c) => {
                const isSelected = selected?.row === r && selected?.col === c;
                return (
                  <td
                    key={c}
                    className={`border border-gray-300 p-0 ${isSelected ? 'bg-primary-100' : ''}`}
                    onClick={() => handleCellClick(r, c)}
                  >
                    <input
                      type="text"
                      className="w-full h-full p-1 focus:outline-none focus:ring-2 focus:ring-primary"
                      value={cell.value}
                      onChange={(e) => handleChange(e, r, c)}
                    />
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}