import React, { useEffect, useState } from 'react';
import Spreadsheet from './components/Spreadsheet';
import { SheetResponse, CellResponse, SheetCreate, BulkCellUpdate } from './types';
import './App.css';

const API_BASE = 'http://localhost:8000';

function App() {
  const [sheets, setSheets] = useState<SheetResponse[]>([]);
  const [activeSheet, setActiveSheet] = useState<SheetResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [newSheetName, setNewSheetName] = useState('');

  const fetchSheets = async () => {
    try {
      const res = await fetch(`${API_BASE}/sheets`);
      if (!res.ok) throw new Error('Failed to fetch sheets');
      const data = await res.json();
      setSheets(data);
    } catch (e: any) {
      setError(e.message);
    }
  };

  useEffect(() => {
    fetchSheets();
  }, []);

  const createSheet = async () => {
    if (!newSheetName.trim()) {
      setError('Sheet name cannot be empty');
      return;
    }
    const payload: SheetCreate = { name: newSheetName };
    try {
      const res = await fetch(`${API_BASE}/sheets`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Failed to create sheet');
      }
      const sheet: SheetResponse = await res.json();
      setSheets((prev) => [...prev, sheet]);
      setNewSheetName('');
    } catch (e: any) {
      setError(e.message);
    }
  };

  const deleteSheet = async (id: number) => {
    try {
      const res = await fetch(`${API_BASE}/sheets/${id}`, {
        method: 'DELETE',
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Failed to delete sheet');
      }
      setSheets((prev) => prev.filter((s) => s.id !== id));
      if (activeSheet?.id === id) setActiveSheet(null);
    } catch (e: any) {
      setError(e.message);
    }
  };

  return (
    <div className="app">
      <header className="header">
        <h1>MS Excel Clone</h1>
      </header>
      <main className="main">
        <section className="sheet-list">
          <h2>Sheets</h2>
          <ul>
            {sheets.map((s) => (
              <li key={s.id}>
                <button
                  className={activeSheet?.id === s.id ? 'active' : ''}
                  onClick={() => setActiveSheet(s)}
                >
                  {s.name}
                </button>
                <button className="danger" onClick={() => deleteSheet(s.id)}>
                  Delete
                </button>
              </li>
            ))}
          </ul>
          <div className="new-sheet">
            <input
              type="text"
              placeholder="New sheet name"
              value={newSheetName}
              onChange={(e) => setNewSheetName(e.target.value)}
            />
            <button onClick={createSheet}>Create Sheet</button>
          </div>
          {error && <p className="error">{error}</p>}
        </section>
        <section className="sheet-view">
          {activeSheet ? (
            <Spreadsheet sheet={activeSheet} />
          ) : (
            <p>Select a sheet to view its contents.</p>
          )}
        </section>
      </main>
    </div>
  );
}

export default App;