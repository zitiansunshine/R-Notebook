import * as vscode from 'vscode';

import {
  COMMAND_IDS,
  NOTEBOOK_TYPE,
  PY_NOTEBOOK_TYPE,
} from './extensionIds';

type NotebookKernelSourceAction = {
  label: string;
  documentation?: vscode.Uri;
  command: string | vscode.Command;
};

type NotebookKernelSourceProvider = {
  onDidChangeNotebookKernelSourceActions?: vscode.Event<void>;
  provideNotebookKernelSourceActions: (
    _token: vscode.CancellationToken,
  ) => readonly NotebookKernelSourceAction[] | Thenable<readonly NotebookKernelSourceAction[]>;
};

type ProposedNotebookApi = typeof vscode.notebooks & {
  registerKernelSourceActionProvider?: (
    notebookType: string,
    provider: NotebookKernelSourceProvider,
  ) => vscode.Disposable;
};

function availableKernelsAction(command: string): NotebookKernelSourceAction {
  return {
    label: 'Available Kernels',
    command: {
      command,
      title: 'Available Kernels',
    },
  };
}

export function registerNotebookKernelSourceProviders(
  ctx: vscode.ExtensionContext,
): void {
  const notebooks = vscode.notebooks as ProposedNotebookApi;
  if (typeof notebooks.registerKernelSourceActionProvider !== 'function') return;

  try {
    ctx.subscriptions.push(
      notebooks.registerKernelSourceActionProvider(NOTEBOOK_TYPE, {
        provideNotebookKernelSourceActions: () => [
          availableKernelsAction(COMMAND_IDS.notebookSelectAvailableRKernels),
        ],
      }),
      notebooks.registerKernelSourceActionProvider(PY_NOTEBOOK_TYPE, {
        provideNotebookKernelSourceActions: () => [
          availableKernelsAction(COMMAND_IDS.notebookSelectAvailablePythonKernels),
        ],
      }),
    );
  } catch {
    // Cursor can expose the proposed kernel-source hook while still rejecting
    // extension usage on stable builds. Skipping this keeps notebook activation
    // alive so the standard toolbar commands continue to work.
  }
}
