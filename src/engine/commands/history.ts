/**
 * Undo/redo command history.
 *
 * A command is an object with `label`, an `apply(wb) => wb` transform and an
 * `inverse()` that yields a command canceling the effect. New workbooks are
 * produced immutably; history inverts by replaying captured previous state.
 */

import { WorkbookDoc } from '../model/types';

export interface Command {
  readonly label: string;
  apply(wb: WorkbookDoc): WorkbookDoc;
  inverse(): Command;
}

export interface HistoryState {
  undoStack: Command[];
  redoStack: Command[];
}

export const MAX_HISTORY = 200;

export class HistoryManager {
  private undoStack: Command[] = [];
  private redoStack: Command[] = [];

  execute(command: Command): void {
    this.undoStack.push(command);
    if (this.undoStack.length > MAX_HISTORY) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Pop the command to undo; the caller applies its inverse. */
  popUndo(): Command | null {
    return this.undoStack.pop() ?? null;
  }

  pushRedo(command: Command): void {
    this.redoStack.push(command);
    if (this.redoStack.length > MAX_HISTORY) this.redoStack.shift();
  }

  popRedo(): Command | null {
    return this.redoStack.pop() ?? null;
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }

  snapshot(): HistoryState {
    return { undoStack: [...this.undoStack], redoStack: [...this.redoStack] };
  }

  restore(state: HistoryState): void {
    this.undoStack = state.undoStack;
    this.redoStack = state.redoStack;
  }
}

/** Group several commands into a single undoable step. */
export class CompositeCommand implements Command {
  constructor(
    public readonly label: string,
    private commands: Command[]
  ) {}

  apply(wb: WorkbookDoc): WorkbookDoc {
    return this.commands.reduce((acc, c) => c.apply(acc), wb);
  }

  inverse(): Command {
    return new CompositeCommand(
      this.label,
      this.commands
        .map((c) => c.inverse())
        .reverse()
    );
  }
}

export function composite(label: string, commands: Command[]): Command {
  return new CompositeCommand(label, commands);
}

/** No-op command (e.g. for actions that changed nothing). */
export class NoopCommand implements Command {
  constructor(public readonly label = 'No changes') {}
  apply(wb: WorkbookDoc): WorkbookDoc {
    return wb;
  }
  inverse(): Command {
    return new NoopCommand(this.label);
  }
}