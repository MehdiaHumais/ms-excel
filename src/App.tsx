import React, { useEffect } from 'react';
import { useWorkbookStore } from '@/store/useWorkbookStore';
import { useUiStore } from '@/store/useUiStore';
import { useSettingsStore } from '@/store/useSettingsStore';
import { Header } from '@/components/Header';
import { FormulaBar } from '@/components/FormulaBar';
import { Grid } from '@/components/Grid';
import { SheetTabs } from '@/components/SheetTabs';
import { StatusBar } from '@/components/StatusBar';
import { HomeScreen } from '@/components/HomeScreen';
import { DialogHost } from '@/components/DialogHost';
import { Toasts } from '@/components/Toasts';
import { ErrorBoundary } from '@/components/ErrorBoundary';
import { RemoveChartCommand } from '@/engine/commands/commands';
import { applyThemeToChrome, getTheme } from '@/engine/model/theme';
import { WorkbookDoc } from '@/engine/model/types';

export default function App() {
  const view = useUiStore((s) => s.view);
  const workbook = useWorkbookStore((s) => s.workbook);

  useGlobalShortcuts();
  useWorkbookTheme(workbook?.theme);
  useFileAssociationOpen();

  return (
    <div className="flex h-full flex-col">
      {view === 'home' || !workbook ? (
        <ErrorBoundary label="The start screen">
          <HomeScreen />
        </ErrorBoundary>
      ) : (
        <div className="flex h-full flex-col">
          <ErrorBoundary label="The ribbon">
            <Header />
          </ErrorBoundary>
          <ErrorBoundary label="The formula bar">
            <FormulaBar />
          </ErrorBoundary>
          <div className="min-h-0 flex-1">
            <ErrorBoundary label="The worksheet">
              <Grid />
            </ErrorBoundary>
          </div>
          <ErrorBoundary label="The sheet tabs">
            <SheetTabs />
          </ErrorBoundary>
          <ErrorBoundary label="The status bar">
            <StatusBar />
          </ErrorBoundary>
        </div>
      )}
      <DialogHost />
      <Toasts />
    </div>
  );
}

function useWorkbookTheme(theme: WorkbookDoc['theme']) {
  // Page Layout > Themes / Colors / Fonts repaint the chrome, like Excel.
  useEffect(() => {
    applyThemeToChrome(getTheme(theme));
  }, [theme]);
}

/**
 * Opening a workbook from Explorer.
 *
 * Windows hands the path to the app either on the launch command line (first
 * instance) or to the running instance via the single-instance plugin (double
 * clicked while the app is already open). Both end up loading the workbook and
 * jumping out of the start screen.
 */
function useFileAssociationOpen() {
  useEffect(() => {
    let disposed = false;
    let unlisten: (() => void) | undefined;

    const open = async (path: string) => {
      try {
        const { loadWorkbookFromFile } = await import('@/platform/loaders');
        await loadWorkbookFromFile(path);
        if (!disposed) useUiStore.getState().setView('workspace');
      } catch (err) {
        console.error('Could not open workbook from command line:', err);
        useUiStore.getState().toast('error', `Could not open ${path}`);
      }
    };

    void (async () => {
      const { isTauri, launchFile, onOpenFileRequest } = await import('@/platform/bridge');
      if (!isTauri()) return;
      const initial = await launchFile();
      if (initial && !disposed) await open(initial);
      const stop = await onOpenFileRequest((path) => void open(path));
      if (disposed) stop();
      else unlisten = stop;
    })();

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, []);
}

function useGlobalShortcuts() {
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement;
      const mod = e.ctrlKey || e.metaKey;
      // Dropdowns and ribbon buttons own their own keyboard interaction:
      // hijacking Arrow/Enter here breaks font, size and number-format lists.
      if (
        t &&
        (t.tagName === 'INPUT' ||
          t.tagName === 'TEXTAREA' ||
          t.tagName === 'SELECT' ||
          t.isContentEditable)
      ) {
        return;
      }
      // A focused ribbon button must keep Enter/Space/Tab/arrows for itself, but
      // that must not disable Ctrl+C/X/V/Z/F, F2, F9 or Ctrl+` after a ribbon click.
      if (t && t.tagName === 'BUTTON' && !mod) {
        if (e.key === 'Enter' || e.key === ' ' || e.key === 'Tab' || e.key === 'Escape' || e.key.startsWith('Arrow')) {
          return;
        }
      }
      const st = useWorkbookStore.getState();
      if (mod && (e.key === 'z' || e.key === 'Z')) {
        e.preventDefault();
        if (e.shiftKey) st.redo();
        else st.undo();
        return;
      }
      if (mod && (e.key === 'y' || e.key === 'Y')) {
        e.preventDefault();
        st.redo();
        return;
      }
      if (mod && (e.key === 'c' || e.key === 'C')) {
        e.preventDefault();
        st.copySelection();
        return;
      }
      if (mod && (e.key === 'x' || e.key === 'X')) {
        e.preventDefault();
        st.cutSelection();
        return;
      }
      if (mod && (e.key === 'v' || e.key === 'V')) {
        e.preventDefault();
        st.paste();
        return;
      }
      if (mod && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault();
        useUiStore.getState().openDialog('findReplace');
        return;
      }
      if (mod && (e.key === 's' || e.key === 'S')) {
        // The toolbar and File menu already advertise Ctrl+S; make it real.
        e.preventDefault();
        void st.save();
        return;
      }
      if (e.key === 'F12') {
        e.preventDefault();
        useUiStore.getState().openDialog('saveAs');
        return;
      }
      if (e.key === 'F2') {
        e.preventDefault();
        st.startEdit();
        return;
      }
      if (e.key === 'F9') {
        e.preventDefault();
        st.recalculate();
        return;
      }
      if (mod && (e.key === '`' || e.code === 'Backquote')) {
        e.preventDefault();
        const settings = useSettingsStore.getState();
        settings.setShowFormulas(!settings.showFormulas);
        return;
      }
      if (e.key === 'Escape') {
        // Escape deselects a selected shape or chart before anything else.
        if (st.selectedShapeId) {
          e.preventDefault();
          st.setSelectedShapeId(null);
          return;
        }
        if (st.selectedChartId) {
          e.preventDefault();
          st.setSelectedChartId(null);
          return;
        }
      }
      if (e.key === 'Delete') {
        e.preventDefault();
        if (st.selectedShapeId) {
          st.removeShape(st.selectedShapeId);
          return;
        }
        if (st.selectedChartId) {
          if (st.activeSheetId) st.applyCommand(new RemoveChartCommand('Delete chart', st.activeSheetId, st.selectedChartId));
          st.setSelectedChartId(null);
          return;
        }
        st.clearSelection();
        return;
      }
      if (e.key === 'Backspace') {
        e.preventDefault();
        if (st.selectedShapeId) {
          st.removeShape(st.selectedShapeId);
          return;
        }
        if (st.selectedChartId) {
          if (st.activeSheetId) st.applyCommand(new RemoveChartCommand('Delete chart', st.activeSheetId, st.selectedChartId));
          st.setSelectedChartId(null);
          return;
        }
        st.clearSelection();
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        const sel = st.selection;
        if (st.editing) st.commitEdit(st.editing.text);
        else st.moveActive(1, 0);
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        st.moveActive(0, 1);
        return;
      }
      const arrow = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[
        e.key as string
      ] as [number, number] | undefined;
      if (arrow) {
        e.preventDefault();
        st.moveActive(arrow[0], arrow[1], e.shiftKey);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);
}