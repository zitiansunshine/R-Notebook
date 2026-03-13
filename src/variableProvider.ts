// =============================================================================
// variableProvider.ts — Notebook Variable Provider for R Notebook kernels.
//
// Integrates with VS Code 1.90+'s native Variables panel
// (vscode.notebooks.registerVariableProvider).  The provider is registered
// in extension.ts with a runtime feature check so it gracefully degrades on
// older VS Code versions.
//
// Both RNotebookController and PyNotebookController call notifyChanged()
// after each cell execution so the Variables panel stays in sync.
// =============================================================================

import * as vscode from 'vscode';
import { getSession }   from './rSessionManager';
import { getPySession } from './pySessionManager';
import { NOTEBOOK_TYPE as R_NOTEBOOK_TYPE, PY_NOTEBOOK_TYPE } from './extensionIds';

export class RNotebookVariableProvider {

  private readonly _emitter = new vscode.EventEmitter<vscode.NotebookDocument>();

  /** VS Code calls this event to know when to refresh the Variables panel. */
  readonly onDidChangeVariables = this._emitter.event;

  /** Called by R/Python controllers after every cell execution. */
  notifyChanged(notebook: vscode.NotebookDocument): void {
    this._emitter.fire(notebook);
  }

  /**
   * Called by VS Code whenever the Variables panel is open and needs data.
   * Returns a flat list of top-level variables (parent === undefined).
   * Ignores paging/indexed children (kind !== 'named' would be empty).
   */
  async provideVariables(
    notebook: vscode.NotebookDocument,
    parent: unknown,
    _kind: unknown,
    _start: number,
    token: vscode.CancellationToken,
  ): Promise<unknown[]> {
    if (parent !== undefined) return [];   // no nested drill-down

    const uri    = notebook.uri.toString();
    const nbType = notebook.notebookType;
    let vars: { name: string; type: string; size: string; value: string }[] = [];

    try {
      if (nbType === R_NOTEBOOK_TYPE) {
        const s = getSession(uri);
        if (s && !token.isCancellationRequested) {
          const r = s.isBusy() ? s.cachedVars() : await s.vars();
          vars = r.vars ?? [];
        }
      } else if (nbType === PY_NOTEBOOK_TYPE) {
        const s = getPySession(uri);
        if (s && !token.isCancellationRequested) {
          const r = s.isBusy() ? s.cachedVars() : await s.vars();
          vars = r.vars ?? [];
        }
      }
    } catch {
      if (nbType === R_NOTEBOOK_TYPE) {
        vars = getSession(uri)?.cachedVars().vars ?? [];
      } else if (nbType === PY_NOTEBOOK_TYPE) {
        vars = getPySession(uri)?.cachedVars().vars ?? [];
      }
    }

    if (token.isCancellationRequested) return [];

    return vars.map(v => ({
      name:                 v.name  ?? '',
      // Prefer size (e.g., "32 × 11") over raw repr for the value column
      value:                v.size  ? v.size  : (v.value ?? ''),
      type:                 v.type  ?? '',
      language:             nbType === R_NOTEBOOK_TYPE ? 'r' : 'python',
      expression:           v.name  ?? '',
      hasNamedChildren:     false,
      indexedChildrenCount: 0,
    }));
  }

  dispose(): void {
    this._emitter.dispose();
  }
}
