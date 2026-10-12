/**
 * App-wide settings store: theme, zoom prefs, AI provider config, etc.
 * Persisted to localStorage when available.
 */

import { create } from 'zustand';

export type Theme = 'light' | 'dark' | 'system';
export type AiProvider = 'none' | 'openai' | 'anthropic' | 'ollama' | 'custom';

export interface AiSettings {
  provider: AiProvider;
  apiKey: string;
  model: string;
  baseUrl: string;
  temperature: number;
}

interface SettingsState {
  theme: Theme;
  editorFontSize: number;
  showGridlines: boolean;
  showFormulas: boolean;
  autosave: boolean;
  showStatusBar: boolean;
  showRuler: boolean;
  ai: AiSettings;

  setTheme(t: Theme): void;
  setEditorFontSize(n: number): void;
  setShowGridlines(b: boolean): void;
  setShowFormulas(b: boolean): void;
  setAutosave(b: boolean): void;
  setShowStatusBar(b: boolean): void;
  setShowRuler(b: boolean): void;
  setAi(patch: Partial<AiSettings>): void;
  applyTheme(): void;
}

const STORAGE_KEY = 'britsheet.settings.v1';

function load(): Partial<SettingsState> {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

const persisted = load();

export const useSettingsStore = create<SettingsState>((set, get) => ({
  theme: persisted.theme ?? 'light',
  editorFontSize: persisted.editorFontSize ?? 12,
  showGridlines: persisted.showGridlines ?? true,
  showFormulas: persisted.showFormulas ?? false,
  autosave: persisted.autosave ?? false,
  showStatusBar: persisted.showStatusBar ?? true,
  showRuler: persisted.showRuler ?? false,
  ai: {
    provider: persisted.ai?.provider ?? 'none',
    apiKey: persisted.ai?.apiKey ?? '',
    model: persisted.ai?.model ?? 'gpt-4o-mini',
    baseUrl: persisted.ai?.baseUrl ?? '',
    temperature: persisted.ai?.temperature ?? 0.7
  },

  setTheme(t) {
    set({ theme: t });
    get().applyTheme();
    persist(get());
  },
  setEditorFontSize(n) {
    set({ editorFontSize: n });
    persist(get());
  },
  setShowGridlines(b) {
    set({ showGridlines: b });
    persist(get());
  },
  setShowFormulas(b) {
    set({ showFormulas: b });
    persist(get());
  },
  setAutosave(b) {
    set({ autosave: b });
    persist(get());
  },
  setShowStatusBar(b) {
    set({ showStatusBar: b });
    persist(get());
  },
  setShowRuler(b) {
    set({ showRuler: b });
    persist(get());
  },
  setAi(patch) {
    set({ ai: { ...get().ai, ...patch } });
    persist(get());
  },
  applyTheme() {
    const { theme } = get();
    const dark =
      theme === 'dark' ||
      (theme === 'system' && window.matchMedia('(prefers-color-scheme: dark)').matches);
    document.documentElement.classList.toggle('dark', dark);
  }
}));

function persist(s: SettingsState): void {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        theme: s.theme,
        editorFontSize: s.editorFontSize,
        showGridlines: s.showGridlines,
        showFormulas: s.showFormulas,
        autosave: s.autosave,
        showStatusBar: s.showStatusBar,
        showRuler: s.showRuler,
        ai: s.ai
      })
    );
  } catch {
    /* ignore quota / privacy errors */
  }
}

// Apply the theme immediately on load.
if (typeof document !== 'undefined') {
  useSettingsStore.getState().applyTheme();
  window
    .matchMedia('(prefers-color-scheme: dark)')
    .addEventListener('change', () => useSettingsStore.getState().applyTheme());
}