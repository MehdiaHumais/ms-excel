import React from 'react';
import { FilePlus2, FolderOpen, FileSpreadsheet, Settings, Sparkles } from 'lucide-react';
import { useWorkbookStore } from '@/store/useWorkbookStore';
import { useUiStore } from '@/store/useUiStore';

export function HomeScreen() {
  const newWorkbook = useWorkbookStore((s) => s.newWorkbook);
  const setView = useUiStore((s) => s.setView);
  const openDialog = useUiStore((s) => s.openDialog);
  const recentFiles = useUiStore((s) => s.recentFiles);

  const blank = () => {
    newWorkbook();
    setView('workspace');
  };

  return (
    <div className="flex h-full" style={{ background: 'rgb(var(--panel))' }}>
      <div className="flex h-full w-56 flex-col border-r p-3" style={{ borderColor: 'rgb(var(--line))' }}>
        <div className="mb-4 flex items-center gap-2 px-1 pt-2">
          <FileSpreadsheet size={22} style={{ color: 'rgb(var(--accent))' }} />
          <span className="text-lg font-semibold">BritSheet</span>
        </div>
        <nav className="space-y-1">
          <NavItem icon={<FilePlus2 size={16} />} label="New" active onClick={blank} />
          <NavItem icon={<FolderOpen size={16} />} label="Open" onClick={() => openDialog('open')} />
          <NavItem icon={<Sparkles size={16} />} label="AI Assistant" onClick={() => openDialog('ai')} />
          <NavItem icon={<Settings size={16} />} label="Settings" onClick={() => openDialog('settings')} />
        </nav>
      </div>

      <div className="flex-1 overflow-y-auto p-8">
        <h1 className="mb-6 text-2xl font-semibold">Start</h1>
        <button
          className="mb-8 flex items-center justify-center gap-3 rounded-lg px-8 py-4 text-base font-medium text-white shadow transition hover:opacity-90"
          style={{ background: 'rgb(var(--accent))' }}
          onClick={blank}
        >
          <FilePlus2 size={20} />
          Blank workbook
        </button>

        {recentFiles.length > 0 && (
          <>
            <h2 className="mb-2 text-lg font-medium">Recent</h2>
            <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
              {recentFiles.map((f) => (
                <button
                  key={f.path}
                  className="flex items-center gap-2 rounded-lg border p-3 text-left hover:bg-black/5 dark:hover:bg-white/5"
                  style={{ borderColor: 'rgb(var(--line))' }}
                  onClick={() => {
                    openDialog('open');
                  }}
                >
                  <FileSpreadsheet size={20} style={{ color: 'rgb(var(--accent))' }} />
                  <div className="min-w-0">
                    <div className="truncate text-[13px] font-medium">{f.name}</div>
                    <div className="truncate text-[11px] opacity-50">{new Date(f.modified).toLocaleDateString()}</div>
                  </div>
                </button>
              ))}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function NavItem({
  icon,
  label,
  active,
  onClick
}: {
  icon: React.ReactNode;
  label: string;
  active?: boolean;
  onClick(): void;
}) {
  return (
    <button
      className={`flex w-full items-center gap-2 rounded-md px-3 py-2 text-[13px] ${
        active ? 'text-white' : 'hover:bg-black/5 dark:hover:bg-white/10'
      }`}
      style={active ? { background: 'rgb(var(--accent))' } : undefined}
      onClick={onClick}
    >
      {icon}
      {label}
    </button>
  );
}

export default HomeScreen;