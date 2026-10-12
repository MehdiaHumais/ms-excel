import React from 'react';

interface Props {
  children: React.ReactNode;
  label?: string;
}

interface State {
  error: Error | null;
}

/**
 * Keeps one broken panel from white-screening the whole app. Without this a
 * single render error (for example a hook-order mistake) unmounts everything,
 * which looks like "the app is dead" rather than "this control is broken".
 */
export class ErrorBoundary extends React.Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error('BritSheet crashed in ' + (this.props.label ?? 'a panel'), error, info.componentStack);
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="flex h-full flex-col items-center justify-center gap-3 p-6 text-center">
        <div className="text-sm font-semibold text-danger">
          {(this.props.label ?? 'This panel') + ' stopped working.'}
        </div>
        <div className="max-w-lg text-xs text-muted break-words">{this.state.error.message}</div>
        <button
          className="rounded border border-line px-3 py-1.5 text-xs font-medium hover:bg-[rgb(var(--accent-soft))]"
          onClick={() => this.setState({ error: null })}
        >
          Try again
        </button>
      </div>
    );
  }
}

export default ErrorBoundary;
