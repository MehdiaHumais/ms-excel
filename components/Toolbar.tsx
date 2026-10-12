import { useState } from 'react';

export default function Toolbar() {
  const [activeTool, setActiveTool] = useState<string>('select');

  const tools = [
    { id: 'select', label: 'Select' },
    { id: 'format', label: 'Format' },
    { id: 'insert', label: 'Insert' },
    { id: 'data', label: 'Data' },
  ];

  return (
    <nav className="bg-surface p-2 flex space-x-2 shadow-card">
      {tools.map((tool) => (
        <button
          key={tool.id}
          onClick={() => setActiveTool(tool.id)}
          className={`px-4 py-2 rounded-xl transition-colors ${
            activeTool === tool.id
              ? 'bg-primary text-white'
              : 'bg-gray-200 hover:bg-gray-300'
          }`}
        >
          {tool.label}
        </button>
      ))}
    </nav>
  );
}