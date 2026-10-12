/**
 * Platform bridge: detects whether BritSheet runs inside the Tauri desktop
 * shell or as a browser preview, and forwards file/open/save/dialog calls to
 * the appropriate backend. The UI code underneath is identical either way.
 */
import { invoke } from '@tauri-apps/api/core';
import { open, save } from '@tauri-apps/plugin-dialog';

let cached: boolean | null = null;

/** True when running inside the Tauri webview (desktop build). */
export function isTauri(): boolean {
  if (cached !== null) return cached;
  cached = typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;
  return cached;
}

export async function openWorkbookFile(path: string): Promise<string> {
  return invoke<string>('open_workbook', { path });
}

export async function saveWorkbookFile(path: string, json: string): Promise<void> {
  await invoke('save_workbook', { path, json });
}

export async function readTextFile(path: string): Promise<string> {
  return invoke<string>('read_text_file', { path });
}

export async function listRecent(): Promise<Array<{ path: string; name: string; modified: number }>> {
  return invoke('list_recent');
}

export async function addRecent(path: string): Promise<void> {
  await invoke('add_recent', { path });
}

export async function removeRecent(path: string): Promise<void> {
  await invoke('remove_recent', { path });
}

export async function defaultDataPath(): Promise<string> {
  return invoke<string>('default_data_path');
}

export async function pickOpenFile(
  filters?: Array<{ name: string; extensions: string[] }>
): Promise<string | null> {
  if (!isTauri()) return null;
  const result = await open({ multiple: false, directory: false, filters: filters ?? [] });
  return typeof result === 'string' ? result : null;
}

export async function pickSaveFile(
  defaultPath: string,
  filters?: Array<{ name: string; extensions: string[] }>
): Promise<string | null> {
  if (!isTauri()) return null;
  const result = await save({ defaultPath, filters: filters ?? [] });
  return typeof result === 'string' ? result : null;
}

/**
 * Path this process was launched with, if Windows opened a workbook via the
 * registered file association. `null` when started normally.
 */
export async function launchFile(): Promise<string | null> {
  if (!isTauri()) return null;
  try {
    return await invoke<string | null>('launch_file');
  } catch {
    return null;
  }
}

/**
 * Fires when a second launch (e.g. double-clicking a workbook) hands a path to
 * the instance that is already running. Returns an unlisten function.
 */
export async function onOpenFileRequest(cb: (path: string) => void): Promise<() => void> {
  const { listen } = await import('@tauri-apps/api/event');
  const unlisten = await listen<string>('britsheet://open-file', (e) => cb(e.payload));
  return unlisten;
}