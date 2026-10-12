/**
 * Loader helpers: read a workbook file (native .britsheet, CSV, JSON or XLSX)
 * from a path using the Tauri backend in the desktop build, driving the
 * workbook store afterwards.
 */
import { openWorkbookFile, readTextFile } from '@/platform/bridge';

export interface LoadedFile {
  name: string;
}

/** Parse file content by extension into a workbook and load it into the store. */
export async function loadWorkbookFromFile(path: string, name?: string): Promise<void> {
  const fileName = name ?? path.split(/[\\/]/).pop() ?? path;
  const ext = fileName.toLowerCase().split('.').pop() ?? '';

  if (ext === 'britsheet' || ext === 'bs') {
    const json = await openWorkbookFile(path);
    const { deserializeWorkbook } = await import('@/engine/io/serialize');
    const wb = deserializeWorkbook(json);
    useWorkbookStore.getState().loadWorkbook(wb);
    remember(path, fileName);
    return;
  }

  if (ext === 'csv' || ext === 'txt') {
    const text = await readTextFile(path);
    const { parseCsv, sniffCsvDelimiter, csvRowsToWorkbook } = await import('@/engine/io/csv');
    const rows = parseCsv(text, sniffCsvDelimiter(text));
    const wb = csvRowsToWorkbook(rows);
    useWorkbookStore.getState().loadWorkbook(wb);
    remember(path, fileName);
    return;
  }

  if (ext === 'xlsx' || ext === 'xls') {
    const { workbookFromXlsx } = await import('@/engine/io/xlsx');
    const bytes = await readXlsxBytes(path);
    const wb = await workbookFromXlsx(bytes);
    useWorkbookStore.getState().loadWorkbook(wb);
    remember(path, fileName);
    return;
  }

  if (ext === 'json') {
    const text = await readTextFile(path);
    const { deserializeWorkbook } = await import('@/engine/io/serialize');
    const wb = deserializeWorkbook(text);
    useWorkbookStore.getState().loadWorkbook(wb);
    remember(path, fileName);
    return;
  }

  throw new Error(`Unsupported file type: .${ext}`);
}

/** Keep the path so File > Save / Ctrl+S can write back without a dialog. */
function remember(path: string, name: string): void {
  useWorkbookStore.getState().setFilePath(path);
  useUiStore.getState().addRecent({ path, name, modified: Date.now() });
  // Mirror into the desktop recent-files database so the list survives restarts.
  void import('@/platform/bridge')
    .then((m) => (m.isTauri() ? m.addRecent(path) : undefined))
    .catch(() => {});
}

async function readXlsxBytes(path: string): Promise<Uint8Array> {
  // Tauri backend returns base64 for binary files; decode in the webview.
  const b64 = await import('@tauri-apps/api/core').then((m) => m.invoke<string>('read_binary_file', { path }));
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

// Local import placeholder for the store used in this module.
import { useWorkbookStore } from '@/store/useWorkbookStore';
import { useUiStore } from '@/store/useUiStore';