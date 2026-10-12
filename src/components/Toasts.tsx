import React from 'react';
import { useUiStore } from '@/store/useUiStore';
import { CheckCircle2, Info, XCircle } from 'lucide-react';

export function Toasts() {
  const toasts = useUiStore((s) => s.toasts);
  if (toasts.length === 0) return null;
  return (
    <div className="fixed bottom-4 right-4 z-[70] flex flex-col gap-2">
      {toasts.map((t) => (
        <div
          key={t.id}
          className="flex items-center gap-2 rounded-lg border bg-cell px-3 py-2 text-[13px] shadow-lg"
          style={{ borderColor: 'rgb(var(--line-strong))' }}
        >
          {t.kind === 'success' && <CheckCircle2 size={16} className="text-green-600" />}
          {t.kind === 'error' && <XCircle size={16} className="text-red-600" />}
          {t.kind === 'info' && <Info size={16} className="text-blue-600" />}
          <span>{t.message}</span>
        </div>
      ))}
    </div>
  );
}

export default Toasts;