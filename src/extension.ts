// =============================================================================
// extension.ts — entrypoint for the R Notebook extension
// =============================================================================

import * as vscode from 'vscode';

import { RMarkdownEditorProvider }                                     from './rmarkdownEditor';
import { disposeSession, getAllSessions, purgeSessionState }           from './rSessionManager';
import { disposePySession, getAllPySessions }                          from './pySessionManager';
import { buildNotebookCellsFromRmdText, RmdNotebookSerializer }        from './rmdNotebookSerializer';
import { IpynbSerializer }                                             from './ipynbSerializer';
import { registerNotebookKernelSourceProviders }                       from './notebookKernelSourceProvider';
import { RNotebookController, NOTEBOOK_TYPE, registerFigureOptions }   from './rNotebookController';
import { registerRNotebookCompletionProvider }                         from './rNotebookCompletionProvider';
import { RmdOutputStore }                                              from './rmdOutputStore';
import { PyNotebookController, PY_NOTEBOOK_TYPE, registerPyFigureOptions } from './pyNotebookController';
import { RNotebookVariableProvider }                                   from './variableProvider';
import { forgetRNotebookKernel }                                       from './notebookKernelState';
import {
  cleanupClearedNotebookChange,
  clearNotebookOutputs,
  exportDocumentUri,
} from './notebookActions';
import { hasPersistedRmdState, splitRmdSourceAndState } from './rmdPersistedState';
import {
  COMMAND_IDS,
  EXTENSION_BRAND,
  VARIABLES_PANEL_VIEW_TYPE,
} from './extensionIds';

// ---------------------------------------------------------------------------

let lastActiveNotebookUri: string | undefined;

export function activate(ctx: vscode.ExtensionContext): void {
  const syncingNotebookUrisFromText = new Set<string>();
  const pendingNotebookTextSyncs = new Set<string>();

  // ── Variable provider (VS Code 1.90+ native Variables panel) ────────────
  const varProvider = new RNotebookVariableProvider();
  ctx.subscriptions.push(varProvider);

  const nbns = vscode.notebooks as any;
  if (typeof nbns.registerVariableProvider === 'function') {
    ctx.subscriptions.push(
      nbns.registerVariableProvider({ notebookType: NOTEBOOK_TYPE }, varProvider),
      nbns.registerVariableProvider({ notebookType: PY_NOTEBOOK_TYPE }, varProvider),
    );
  }

  ctx.subscriptions.push(
    vscode.workspace.registerNotebookSerializer(
      NOTEBOOK_TYPE,
      new RmdNotebookSerializer(),
      { transientOutputs: false },
    ),
  );

  const rmdOutputStore = new RmdOutputStore();
  const rController = new RNotebookController(ctx, rmdOutputStore, varProvider);
  ctx.subscriptions.push(rController);
  ctx.subscriptions.push(registerRNotebookCompletionProvider());
  registerFigureOptions(ctx);

  ctx.subscriptions.push(
    vscode.workspace.registerNotebookSerializer(
      PY_NOTEBOOK_TYPE,
      new IpynbSerializer(),
      { transientOutputs: false },
    ),
  );

  const pyController = new PyNotebookController(ctx, varProvider);
  ctx.subscriptions.push(pyController);
  registerPyFigureOptions(ctx);
  registerNotebookKernelSourceProviders(ctx);

  rememberNotebookDocument(vscode.window.activeNotebookEditor?.notebook);
  ctx.subscriptions.push(
    vscode.window.onDidChangeActiveNotebookEditor((editor) => {
      rememberNotebookDocument(editor?.notebook);
      void ensureManagedNotebookTextDocument(editor?.notebook);
    }),
    vscode.workspace.onDidOpenNotebookDocument((notebook) => {
      void ensureManagedNotebookTextDocument(notebook);
    }),
    vscode.workspace.onDidChangeTextDocument((event) => {
      void syncNotebookDocumentFromTextChange(
        event,
        syncingNotebookUrisFromText,
        pendingNotebookTextSyncs,
        rmdOutputStore,
        rController,
        pyController,
      );
    }),
    vscode.workspace.onDidChangeNotebookDocument((event) => {
      void cleanupClearedNotebookChange(event, rmdOutputStore);
      void flushPendingNotebookTextSync(
        event.notebook,
        syncingNotebookUrisFromText,
        pendingNotebookTextSyncs,
        rmdOutputStore,
        rController,
        pyController,
      );
    }),
  );
  registerNotebookStateWatchers(ctx, rmdOutputStore, pyController);
  for (const notebook of vscode.workspace.notebookDocuments) {
    void ensureManagedNotebookTextDocument(notebook);
  }

  if (typeof vscode.notebooks.createRendererMessaging === 'function') {
    const rendererMessaging = vscode.notebooks.createRendererMessaging('rNotebook.execResultRenderer');
    ctx.subscriptions.push(
      rendererMessaging.onDidReceiveMessage(async ({ editor, message }) => {
        if (!message || typeof message !== 'object') return;
        if (
          message.type !== 'open_console_in_tab'
          && message.type !== 'open_output_in_tab'
          && message.type !== 'copy_console'
        ) return;

        const chunkId = typeof message.chunkId === 'string' && message.chunkId
          ? message.chunkId
          : 'console';
        const title = typeof message.title === 'string' && message.title
          ? message.title
          : message.type === 'open_output_in_tab'
            ? `Output: ${chunkId}`
            : `Console: ${chunkId}`;
        const content = typeof message.content === 'string'
          ? message.content
          : String(message.content ?? '');
        if (message.type === 'copy_console') {
          await vscode.env.clipboard.writeText(content);
          void vscode.window.setStatusBarMessage('Copied to ClipBoard', 2000);
          return;
        }

        const panel = vscode.window.createWebviewPanel(
          'rNotebook.consoleOutput',
          title,
          editor.viewColumn ?? vscode.ViewColumn.Beside,
          { enableScripts: true, retainContextWhenHidden: true },
        );
        const copyDisposable = panel.webview.onDidReceiveMessage(async (panelMessage) => {
          if (!panelMessage || typeof panelMessage !== 'object') return;
          if (panelMessage.type !== 'copy_console') return;
          const panelContent = typeof panelMessage.content === 'string'
            ? panelMessage.content
            : String(panelMessage.content ?? '');
          await vscode.env.clipboard.writeText(panelContent);
          void vscode.window.setStatusBarMessage('Copied to ClipBoard', 2000);
          await panel.webview.postMessage({ type: 'copy_console_done' });
        });
        panel.onDidDispose(() => copyDisposable.dispose());
        panel.webview.html = buildConsoleTabHtml(title, content);
      }),
    );
  }

  ctx.subscriptions.push(
    vscode.workspace.onDidCloseNotebookDocument((notebook) => {
      const uri = notebook.uri.toString();
      pendingNotebookTextSyncs.delete(uri);
      syncingNotebookUrisFromText.delete(uri);
      if (lastActiveNotebookUri === uri) lastActiveNotebookUri = undefined;
      if (notebook.notebookType === NOTEBOOK_TYPE) {
        // Reopened notebook cells get fresh document URIs, so stale cached output
        // mappings must not survive close/reopen cycles.
        rmdOutputStore.clear(uri);
        forgetRNotebookKernel(uri);
        void disposeSession(uri);
      } else if (notebook.notebookType === PY_NOTEBOOK_TYPE) {
        void disposePySession(uri);
      }
    }),
  );

  // ── Rich webview editor (.Rmd, available via "Open With") ────────────────
  ctx.subscriptions.push(RMarkdownEditorProvider.register(ctx));

  // ── Global status bar (active session count) ─────────────────────────────
  const statusBar = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Right, 100,
  );
  statusBar.tooltip = 'Active R / Python sessions (click to show)';
  ctx.subscriptions.push(statusBar);

  const updateStatus = () => {
    const n = getAllSessions().size + getAllPySessions().size;
    statusBar.text = n > 0 ? `$(beaker) ${EXTENSION_BRAND}: ${n}` : '';
    n > 0 ? statusBar.show() : statusBar.hide();
  };
  const interval = setInterval(updateStatus, 5000);
  ctx.subscriptions.push({ dispose: () => clearInterval(interval) });

  const rKeepAlive = setInterval(() => {
    for (const session of getAllSessions().values()) {
      if (session.isBusy()) continue;
      void session.keepAlive().catch(() => undefined);
    }
  }, 60_000);
  ctx.subscriptions.push({ dispose: () => clearInterval(rKeepAlive) });

  // ── Webview commands ──────────────────────────────────────────────────────
  ctx.subscriptions.push(
    vscode.commands.registerCommand(COMMAND_IDS.notebookSelectAvailableRKernels, async (target?: unknown) => {
      await rController.showKernelPicker(resolveNotebookDocument(target));
    }),
    vscode.commands.registerCommand(COMMAND_IDS.notebookSelectAvailablePythonKernels, async (target?: unknown) => {
      await pyController.showKernelPicker(resolveNotebookDocument(target), { includeManual: true });
    }),
    vscode.commands.registerCommand(COMMAND_IDS.rRunChunk, () =>
      vscode.commands.executeCommand('workbench.action.webview.postMessage',
        { type: 'run_focused_chunk' })),
    vscode.commands.registerCommand(COMMAND_IDS.rRunAll, () =>
      vscode.commands.executeCommand('workbench.action.webview.postMessage',
        { type: 'run_all_trigger' })),
    vscode.commands.registerCommand(COMMAND_IDS.rRestartSession, () =>
      vscode.commands.executeCommand('workbench.action.webview.postMessage',
        { type: 'reset_session' })),
  );

  // ── Notebook toolbar: Interrupt ──────────────────────────────────────────
  ctx.subscriptions.push(
    vscode.commands.registerCommand(COMMAND_IDS.notebookInterrupt, async (target?: unknown) => {
      const notebook = resolveNotebookDocument(target);
      let interrupted = false;

      if (notebook) {
        rememberNotebookDocument(notebook);
        const uri  = notebook.uri.toString();
        const type = notebook.notebookType;

        if (type === NOTEBOOK_TYPE) {
          interrupted = rController.interruptNotebook(uri);
        } else if (type === PY_NOTEBOOK_TYPE) {
          interrupted = pyController.interruptNotebook(uri);
        }
      }

      if (!interrupted) {
        for (const [uri, session] of getAllSessions()) {
          if (!session.isBusy()) continue;
          interrupted = rController.interruptNotebook(uri) || interrupted;
        }
        for (const [uri, session] of getAllPySessions()) {
          if (!session.isBusy()) continue;
          interrupted = pyController.interruptNotebook(uri) || interrupted;
        }
      }

      if (!interrupted) {
        vscode.window.showWarningMessage('No running kernel to interrupt.');
      }
    }),
  );

  // ── Notebook toolbar: Restart kernel ─────────────────────────────────────
  ctx.subscriptions.push(
    vscode.commands.registerCommand(COMMAND_IDS.notebookRestart, async (target?: unknown) => {
      const notebook = resolveNotebookDocument(target);
      if (!notebook) {
        vscode.window.showWarningMessage('Open an R Notebook or Python Notebook to restart its kernel.');
        return;
      }
      rememberNotebookDocument(notebook);
      const type = notebook.notebookType;

      try {
        let restarted = false;
        if (type === NOTEBOOK_TYPE) {
          restarted = await rController.restartNotebook(notebook);
        } else if (type === PY_NOTEBOOK_TYPE) {
          restarted = await pyController.restartNotebook(notebook);
        }
        if (!restarted) {
          vscode.window.showWarningMessage('Unable to determine which kernel to restart for this notebook.');
          return;
        }
        vscode.window.showInformationMessage('Kernel restarted.');
      } catch (err: any) {
        vscode.window.showErrorMessage(`Kernel restart failed: ${err.message}`);
      }
    }),
  );

  // ── Notebook toolbar: Show Variables (webview panel) ─────────────────────
  let varPanel: vscode.WebviewPanel | undefined;

  ctx.subscriptions.push(
    vscode.commands.registerCommand(COMMAND_IDS.notebookShowVariables, async (target?: unknown) => {
      const notebook = resolveNotebookDocument(target);
      rememberNotebookDocument(notebook);
      // Collect vars from ALL active R and Python sessions
      interface SessionVars {
        language: 'R' | 'Python';
        doc: string;
        busy?: boolean;
        vars: { name: string; type: string; size: string; value: string }[];
      }
      const sections: SessionVars[] = [];

      for (const [uri, session] of getAllSessions()) {
        try {
          const r = session.isBusy() ? session.cachedVars() : await session.vars();
          sections.push({ language: 'R', doc: uri, busy: session.isBusy(), vars: r?.vars ?? [] });
        } catch {
          const cached = session.cachedVars();
          if ((cached.vars?.length ?? 0) > 0 || session.isBusy()) {
            sections.push({ language: 'R', doc: uri, busy: session.isBusy(), vars: cached.vars ?? [] });
          }
        }
      }
      for (const [uri, session] of getAllPySessions()) {
        try {
          const r = session.isBusy() ? session.cachedVars() : await session.vars();
          sections.push({ language: 'Python', doc: uri, busy: session.isBusy(), vars: r?.vars ?? [] });
        } catch {
          const cached = session.cachedVars();
          if ((cached.vars?.length ?? 0) > 0 || session.isBusy()) {
            sections.push({ language: 'Python', doc: uri, busy: session.isBusy(), vars: cached.vars ?? [] });
          }
        }
      }

      if (notebook) {
        const notebookUri = notebook.uri.toString();
        sections.sort((a, b) => {
          const aRank = a.doc === notebookUri ? 0 : 1;
          const bRank = b.doc === notebookUri ? 0 : 1;
          return aRank - bRank || a.doc.localeCompare(b.doc);
        });
      }

      const html = buildVarsPanelHtml(sections);

      if (varPanel) {
        varPanel.reveal(vscode.ViewColumn.Beside);
        varPanel.webview.html = html;
      } else {
        varPanel = vscode.window.createWebviewPanel(
          VARIABLES_PANEL_VIEW_TYPE,
          'Session Variables',
          vscode.ViewColumn.Beside,
          { enableScripts: false },
        );
        varPanel.webview.html = html;
        varPanel.onDidDispose(() => { varPanel = undefined; }, null, ctx.subscriptions);
      }
    }),
  );

  ctx.subscriptions.push(
    vscode.commands.registerCommand(COMMAND_IDS.notebookClearOutputs, async (target?: unknown) => {
      const notebook = resolveNotebookDocument(target);
      if (!notebook) {
        vscode.window.showWarningMessage('Open an R Notebook, Quarto Notebook, or Python Notebook to clear outputs.');
        return;
      }
      rememberNotebookDocument(notebook);

      try {
        const changed = await clearNotebookOutputs(notebook, rmdOutputStore);
        vscode.window.showInformationMessage(
          changed ? 'All outputs cleared and notebook saved.' : 'No outputs to clear.',
        );
      } catch (err: any) {
        vscode.window.showErrorMessage(`Clearing outputs failed: ${err.message}`);
      }
    }),
    vscode.commands.registerCommand(COMMAND_IDS.notebookExport, async (target?: unknown) => {
      const notebook = resolveNotebookDocument(target);
      const textDocument = resolveExportableTextDocument();

      try {
        if (notebook) {
          rememberNotebookDocument(notebook);
          await saveNotebookDocument(notebook);
          await exportDocumentUri(notebook.uri);
          return;
        }

        if (textDocument) {
          await textDocument.save();
          await exportDocumentUri(textDocument.uri);
          return;
        }

        vscode.window.showWarningMessage('Open an .Rmd, .qmd, or .ipynb file to export it.');
      } catch (err: any) {
        vscode.window.showErrorMessage(`Export failed: ${err.message}`);
      }
    }),
  );

}

export function deactivate(): void { /* sessions disposed per-document */ }

function resolveNotebookDocument(target?: unknown): vscode.NotebookDocument | undefined {
  return resolveNotebookTarget(target)
    ?? managedNotebook(vscode.window.activeNotebookEditor?.notebook)
    ?? findNotebookByUri(lastActiveNotebookUri)
    ?? vscode.window.visibleNotebookEditors
      .map((editor) => editor.notebook)
      .find((notebook) => isManagedNotebookDocument(notebook))
    ?? vscode.workspace.notebookDocuments.find((notebook) => isManagedNotebookDocument(notebook));
}

function resolveNotebookTarget(target?: unknown): vscode.NotebookDocument | undefined {
  const directNotebook = resolveLiveManagedNotebookDocument(target);
  if (directNotebook) return directNotebook;

  if (Array.isArray(target)) {
    for (const entry of target) {
      const resolved = resolveNotebookTarget(entry);
      if (resolved) return resolved;
    }
    return undefined;
  }

  const targetUri = reviveUriLike(target);
  if (targetUri) {
    return findNotebookByUri(targetUri);
  }

  if (target && typeof target === 'object') {
    const candidate = target as {
      notebook?: unknown;
      notebookEditor?: { notebook?: unknown };
      cell?: { notebook?: unknown };
      uri?: unknown;
      resource?: unknown;
      notebookType?: string;
      notebookUri?: unknown;
    };
    const notebook = resolveLiveManagedNotebookDocument(candidate.notebook)
      ?? resolveLiveManagedNotebookDocument(candidate.notebookEditor?.notebook)
      ?? resolveLiveManagedNotebookDocument(candidate.cell?.notebook);
    if (notebook) return notebook;

    const resource = reviveUriLike(candidate.notebookUri ?? candidate.resource ?? candidate.uri);
    if (resource) return findNotebookByUri(resource);
  }

  return undefined;
}

function rememberNotebookDocument(notebook?: vscode.NotebookDocument): void {
  if (!isManagedNotebookDocument(notebook)) return;
  lastActiveNotebookUri = notebook.uri.toString();
}

function resolveLiveManagedNotebookDocument(target?: unknown): vscode.NotebookDocument | undefined {
  if (!target || typeof target !== 'object') return undefined;
  const candidate = target as {
    uri?: unknown;
    notebookType?: unknown;
    getCells?: unknown;
    cellAt?: unknown;
  };
  const notebookType = candidate.notebookType;
  if (notebookType !== NOTEBOOK_TYPE && notebookType !== PY_NOTEBOOK_TYPE) return undefined;

  const uri = reviveUriLike(candidate.uri);
  if (uri) {
    return findNotebookByUri(uri)
      ?? (
        typeof candidate.getCells === 'function' || typeof candidate.cellAt === 'function'
          ? target as vscode.NotebookDocument
          : undefined
      );
  }

  return typeof candidate.getCells === 'function' || typeof candidate.cellAt === 'function'
    ? target as vscode.NotebookDocument
    : undefined;
}

function reviveUriLike(value: unknown): vscode.Uri | undefined {
  if (!value) return undefined;
  if (value instanceof vscode.Uri) return value;
  if (typeof value !== 'object') return undefined;

  const candidate = value as {
    scheme?: unknown;
    path?: unknown;
  };
  if (typeof candidate.scheme !== 'string' || typeof candidate.path !== 'string') {
    return undefined;
  }

  try {
    return vscode.Uri.revive(value as vscode.UriComponents);
  } catch {
    return undefined;
  }
}

function findNotebookByUri(uri?: string | vscode.Uri): vscode.NotebookDocument | undefined {
  const uriKey = typeof uri === 'string' ? uri : uri?.toString();
  if (!uriKey) return undefined;
  const notebook = vscode.workspace.notebookDocuments.find((doc) => doc.uri.toString() === uriKey);
  return managedNotebook(notebook);
}

function managedNotebook(
  notebook?: vscode.NotebookDocument,
): vscode.NotebookDocument | undefined {
  return isManagedNotebookDocument(notebook) ? notebook : undefined;
}

function isManagedNotebookDocument(notebook?: vscode.NotebookDocument): boolean {
  return notebook?.notebookType === NOTEBOOK_TYPE
    || notebook?.notebookType === PY_NOTEBOOK_TYPE;
}

function shouldTrackBackingTextDocument(
  notebook?: vscode.NotebookDocument,
): boolean {
  return notebook?.notebookType === NOTEBOOK_TYPE;
}

async function ensureManagedNotebookTextDocument(
  notebook?: vscode.NotebookDocument,
): Promise<void> {
  if (!shouldTrackBackingTextDocument(notebook)) return;
  if (!isManagedNotebookDocument(notebook)) return;
  try {
    await vscode.workspace.openTextDocument(notebook.uri);
  } catch {
    // Best effort only; the notebook can still function without the backing text
    // document being open, but external edit syncing will be unavailable.
  }
}

async function syncNotebookDocumentFromTextChange(
  event: vscode.TextDocumentChangeEvent,
  syncingNotebookUrisFromText: Set<string>,
  pendingNotebookTextSyncs: Set<string>,
  rmdOutputStore: RmdOutputStore,
  rController: RNotebookController,
  pyController: PyNotebookController,
): Promise<void> {
  const notebook = findNotebookByUri(event.document.uri.toString());
  if (!notebook) return;
  if (!isManagedNotebookUri(event.document.uri)) return;
  if (
    notebook.notebookType === PY_NOTEBOOK_TYPE
    && !vscode.window.visibleTextEditors.some((editor) => editor.document.uri.toString() === event.document.uri.toString())
  ) {
    return;
  }

  const notebookUri = notebook.uri.toString();
  const text = event.document.getText();
  if (isRmdPersistedStateOnlyChange(notebook, text)) {
    pendingNotebookTextSyncs.delete(notebookUri);
    return;
  }
  if (syncingNotebookUrisFromText.has(notebookUri)) {
    pendingNotebookTextSyncs.add(notebookUri);
    return;
  }
  if (hasPendingExecutionForNotebook(notebook, rController, pyController)) {
    pendingNotebookTextSyncs.add(notebookUri);
    return;
  }

  pendingNotebookTextSyncs.delete(notebookUri);
  await syncNotebookDocumentFromText(
    notebook,
    text,
    syncingNotebookUrisFromText,
    pendingNotebookTextSyncs,
    rmdOutputStore,
    rController,
    pyController,
  );
}

async function flushPendingNotebookTextSync(
  notebook: vscode.NotebookDocument,
  syncingNotebookUrisFromText: Set<string>,
  pendingNotebookTextSyncs: Set<string>,
  rmdOutputStore: RmdOutputStore,
  rController: RNotebookController,
  pyController: PyNotebookController,
): Promise<void> {
  if (!isManagedNotebookDocument(notebook)) return;
  const notebookUri = notebook.uri.toString();
  if (!pendingNotebookTextSyncs.has(notebookUri)) return;
  if (syncingNotebookUrisFromText.has(notebookUri)) return;
  if (hasPendingExecutionForNotebook(notebook, rController, pyController)) return;

  pendingNotebookTextSyncs.delete(notebookUri);
  try {
    const textDocument = await vscode.workspace.openTextDocument(notebook.uri);
    const text = textDocument.getText();
    if (isRmdPersistedStateOnlyChange(notebook, text)) {
      return;
    }
    await syncNotebookDocumentFromText(
      notebook,
      text,
      syncingNotebookUrisFromText,
      pendingNotebookTextSyncs,
      rmdOutputStore,
      rController,
      pyController,
    );
  } catch {
    pendingNotebookTextSyncs.add(notebookUri);
  }
}

type NotebookTextSyncState = {
  cells: readonly vscode.NotebookCellData[];
  metadata?: vscode.NotebookDocumentMetadata;
};

async function syncNotebookDocumentFromText(
  notebook: vscode.NotebookDocument,
  text: string,
  syncingNotebookUrisFromText: Set<string>,
  pendingNotebookTextSyncs: Set<string>,
  rmdOutputStore: RmdOutputStore,
  rController: RNotebookController,
  pyController: PyNotebookController,
): Promise<void> {
  const notebookUri = notebook.uri.toString();
  let nextState: NotebookTextSyncState;
  try {
    nextState = buildNotebookTextSyncState(notebook, text);
  } catch {
    pendingNotebookTextSyncs.add(notebookUri);
    return;
  }

  if (hasPendingExecutionForNotebook(notebook, rController, pyController)) {
    pendingNotebookTextSyncs.add(notebookUri);
    return;
  }

  const shouldHardResetRmdOutputs = shouldTreatRmdSourceAsFresh(
    notebook,
    text,
    rmdOutputStore,
  );
  if (shouldHardResetRmdOutputs) {
    rmdOutputStore.markHardReset(notebookUri);
  }

  if (notebookMatchesSyncState(notebook, nextState)) {
    if (shouldHardResetRmdOutputs) {
      await clearNotebookCodeOutputs(notebook);
      rmdOutputStore.finishHardReset(notebookUri);
    }
    return;
  }

  syncingNotebookUrisFromText.add(notebookUri);
  try {
    const edits: vscode.NotebookEdit[] = [];
    if (!notebookMetadataEqual(notebook.metadata, nextState.metadata)) {
      edits.push(buildUpdateNotebookMetadataEdit(nextState.metadata ?? {}));
    }
    edits.push(buildReplaceNotebookCellsEdit(notebook, nextState.cells));

    const edit = new vscode.WorkspaceEdit();
    edit.set(notebook.uri, edits);
    await vscode.workspace.applyEdit(edit);
  } finally {
    syncingNotebookUrisFromText.delete(notebookUri);
  }

  await flushPendingNotebookTextSync(
    notebook,
    syncingNotebookUrisFromText,
    pendingNotebookTextSyncs,
    rmdOutputStore,
    rController,
    pyController,
  );
}

function isRmdPersistedStateOnlyChange(
  notebook: vscode.NotebookDocument,
  text: string,
): boolean {
  if (notebook.notebookType !== NOTEBOOK_TYPE) return false;
  if (!hasPersistedRmdState(text)) return false;

  let nextState: NotebookTextSyncState;
  try {
    nextState = buildNotebookTextSyncState(notebook, text);
  } catch {
    return false;
  }
  return notebookMatchesSyncState(notebook, nextState);
}

function buildNotebookTextSyncState(
  notebook: vscode.NotebookDocument,
  text: string,
): NotebookTextSyncState {
  if (notebook.notebookType === NOTEBOOK_TYPE) {
    const { source } = splitRmdSourceAndState(text);
    return {
      cells: buildNotebookCellsFromRmdText(source).map((cell) => stripCodeCellOutputs(cell)),
    };
  }

  JSON.parse(text);
  const serializer = new IpynbSerializer();
  const notebookData = serializer.deserializeNotebook(new TextEncoder().encode(text));
  return {
    cells: notebookData.cells.map((cell, index) => cloneIpynbSyncCellData(notebook, cell, index)),
    metadata: notebookData.metadata as vscode.NotebookDocumentMetadata | undefined,
  };
}

function cloneNotebookCellData(cell: vscode.NotebookCellData): vscode.NotebookCellData {
  const clone = new vscode.NotebookCellData(cell.kind, cell.value, cell.languageId);
  clone.metadata = cell.metadata as vscode.NotebookCellMetadata | undefined;
  clone.outputs = [...(cell.outputs ?? [])];
  return clone;
}

function cloneIpynbSyncCellData(
  notebook: vscode.NotebookDocument,
  cell: vscode.NotebookCellData,
  index: number,
): vscode.NotebookCellData {
  const clone = cloneNotebookCellData(cell);
  if (clone.kind !== vscode.NotebookCellKind.Code) return clone;

  const currentCell = notebook.getCells()[index];
  if (!currentCell || currentCell.kind !== vscode.NotebookCellKind.Code) return clone;

  // Backing-file sync should update source/metadata without blowing away the
  // notebook's current rendered outputs. This matches Jupyter's behavior more
  // closely and prevents long-running Python notebook results from vanishing
  // when the hidden .ipynb text document changes underneath the open notebook.
  clone.outputs = [...currentCell.outputs];
  return clone;
}

function notebookMatchesSyncState(
  notebook: vscode.NotebookDocument,
  nextState: NotebookTextSyncState,
): boolean {
  return notebookCellsMatch(notebook, nextState.cells)
    && notebookMetadataEqual(notebook.metadata, nextState.metadata);
}

function stripCodeCellOutputs(cell: vscode.NotebookCellData): vscode.NotebookCellData {
  const next = new vscode.NotebookCellData(cell.kind, cell.value, cell.languageId);
  next.metadata = cell.metadata as vscode.NotebookCellMetadata | undefined;
  next.outputs = cell.kind === vscode.NotebookCellKind.Code ? [] : [...cell.outputs];
  return next;
}

function shouldTreatRmdSourceAsFresh(
  notebook: vscode.NotebookDocument,
  text: string,
  rmdOutputStore: RmdOutputStore,
): boolean {
  if (notebook.notebookType !== NOTEBOOK_TYPE) return false;
  if (hasPersistedRmdState(text)) return false;

  const docUri = notebook.uri.toString();
  return rmdOutputStore.hasHardReset(docUri) || notebookHasVisibleCodeOutputs(notebook);
}

function notebookHasVisibleCodeOutputs(notebook: vscode.NotebookDocument): boolean {
  return notebook
    .getCells()
    .some((cell) => cell.kind === vscode.NotebookCellKind.Code && cell.outputs.length > 0);
}

function buildReplaceNotebookCellsEdit(
  notebook: vscode.NotebookDocument,
  cells: readonly vscode.NotebookCellData[],
): vscode.NotebookEdit {
  const notebookEdit = vscode.NotebookEdit as typeof vscode.NotebookEdit & {
    replaceCells?: (
      range: vscode.NotebookRange,
      newCells: readonly vscode.NotebookCellData[],
    ) => vscode.NotebookEdit;
  };
  if (typeof notebookEdit.replaceCells !== 'function') {
    throw new Error('Notebook replace API is unavailable in this editor build.');
  }
  return notebookEdit.replaceCells(
    new vscode.NotebookRange(0, notebook.getCells().length),
    cells,
  );
}

async function clearNotebookCodeOutputs(
  notebook: vscode.NotebookDocument,
): Promise<void> {
  const edits: vscode.NotebookEdit[] = [];
  for (const cell of notebook.getCells()) {
    if (cell.kind !== vscode.NotebookCellKind.Code) continue;
    if (cell.outputs.length === 0) continue;
    edits.push(buildUpdateCellOutputsEdit(cell, []));
  }
  if (edits.length === 0) return;

  const edit = new vscode.WorkspaceEdit();
  edit.set(notebook.uri, edits);
  await vscode.workspace.applyEdit(edit);
}

function buildUpdateCellOutputsEdit(
  cell: vscode.NotebookCell,
  outputs: readonly vscode.NotebookCellOutput[],
): vscode.NotebookEdit {
  const notebookEdit = vscode.NotebookEdit as typeof vscode.NotebookEdit & {
    updateCellOutputs?: (
      index: number,
      outputs: readonly vscode.NotebookCellOutput[],
    ) => vscode.NotebookEdit;
    replaceCells?: (
      range: vscode.NotebookRange,
      newCells: readonly vscode.NotebookCellData[],
    ) => vscode.NotebookEdit;
  };

  if (typeof notebookEdit.updateCellOutputs === 'function') {
    return notebookEdit.updateCellOutputs(cell.index, outputs);
  }
  if (typeof notebookEdit.replaceCells !== 'function') {
    throw new Error('Notebook output edit API is unavailable in this editor build.');
  }

  const replacement = new vscode.NotebookCellData(
    cell.kind,
    cell.document.getText(),
    cell.document.languageId,
  );
  replacement.metadata = cell.metadata as vscode.NotebookCellMetadata | undefined;
  replacement.outputs = [...outputs];
  return notebookEdit.replaceCells(
    new vscode.NotebookRange(cell.index, cell.index + 1),
    [replacement],
  );
}

function buildUpdateNotebookMetadataEdit(
  metadata: vscode.NotebookDocumentMetadata,
): vscode.NotebookEdit {
  return vscode.NotebookEdit.updateNotebookMetadata(metadata);
}

function notebookCellsMatch(
  notebook: vscode.NotebookDocument,
  nextCells: readonly vscode.NotebookCellData[],
): boolean {
  const currentCells = notebook.getCells();
  if (currentCells.length !== nextCells.length) return false;

  for (let index = 0; index < currentCells.length; index += 1) {
    const current = currentCells[index];
    const next = nextCells[index];
    if (current.kind !== next.kind) return false;
    if (current.document.languageId !== next.languageId) return false;
    if (current.document.getText() !== next.value) return false;
    if (!notebookCellMetadataEqual(current.metadata, next.metadata)) return false;
    if (
      notebook.notebookType === PY_NOTEBOOK_TYPE
      && !notebookCellOutputsEqual(current.outputs, next.outputs ?? [])
    ) {
      return false;
    }
  }

  return true;
}

function notebookMetadataEqual(
  left: vscode.NotebookDocumentMetadata | undefined,
  right: vscode.NotebookDocumentMetadata | undefined,
): boolean {
  return stableStringify(left ?? {}) === stableStringify(right ?? {});
}

function notebookCellMetadataEqual(
  left: vscode.NotebookCellMetadata | undefined,
  right: vscode.NotebookCellMetadata | undefined,
): boolean {
  return stableStringify(normalizeNotebookCellMetadata(left))
    === stableStringify(normalizeNotebookCellMetadata(right));
}

function notebookCellOutputsEqual(
  left: readonly vscode.NotebookCellOutput[],
  right: readonly vscode.NotebookCellOutput[],
): boolean {
  return stableStringify(left.map(serializeNotebookCellOutput))
    === stableStringify(right.map(serializeNotebookCellOutput));
}

function serializeNotebookCellOutput(output: vscode.NotebookCellOutput): {
  metadata?: { [key: string]: any };
  items: { mime: string; data: string }[];
} {
  return {
    metadata: output.metadata,
    items: output.items.map((item) => ({
      mime: item.mime,
      data: Buffer.from(item.data).toString('base64'),
    })),
  };
}

function normalizeNotebookCellMetadata(
  metadata: vscode.NotebookCellMetadata | undefined,
): Record<string, unknown> {
  if (!metadata || typeof metadata !== 'object') return {};
  const source = metadata as Record<string, unknown>;
  const hasRmdNotebookKeys =
    source.kind !== undefined
    || source.optionStyle !== undefined
    || source.options !== undefined;
  if (!hasRmdNotebookKeys) return source;
  const normalized: Record<string, unknown> = {};
  if (source.kind !== undefined) normalized.kind = source.kind;
  if (source.optionStyle !== undefined) normalized.optionStyle = source.optionStyle;
  if (source.options !== undefined) normalized.options = source.options;
  return normalized;
}

function stableStringify(value: unknown): string {
  return JSON.stringify(normalizeForComparison(value));
}

function normalizeForComparison(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeForComparison(entry));
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([, entry]) => entry !== undefined)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, normalizeForComparison(entry)]),
    );
  }
  return value;
}

function isManagedNotebookUri(uri: vscode.Uri): boolean {
  const fsPath = uri.fsPath.toLowerCase();
  return fsPath.endsWith('.rmd') || fsPath.endsWith('.qmd') || fsPath.endsWith('.ipynb');
}

function hasPendingExecutionForNotebook(
  notebook: vscode.NotebookDocument,
  rController: RNotebookController,
  pyController: PyNotebookController,
): boolean {
  const docUri = notebook.uri.toString();
  if (notebook.notebookType === NOTEBOOK_TYPE) {
    return rController.hasPendingExecution(docUri);
  }
  if (notebook.notebookType === PY_NOTEBOOK_TYPE) {
    return pyController.hasPendingExecution(docUri);
  }
  return false;
}

function resolveExportableTextDocument(): vscode.TextDocument | undefined {
  const document = vscode.window.activeTextEditor?.document;
  if (!document) return undefined;
  const ext = document.uri.fsPath.toLowerCase();
  if (ext.endsWith('.rmd') || ext.endsWith('.qmd') || ext.endsWith('.ipynb')) {
    return document;
  }
  return undefined;
}

function scheduleNotebookConflictSave(
  notebook: vscode.NotebookDocument,
  pendingNotebookConflictSaves: Set<string>,
  notebookConflictSaveTimers: Map<string, ReturnType<typeof setTimeout>>,
  syncingNotebookUrisFromText: Set<string>,
  pendingNotebookTextSyncs: Set<string>,
  rController: RNotebookController,
  pyController: PyNotebookController,
): void {
  if (!isManagedNotebookDocument(notebook)) return;

  const notebookUri = notebook.uri.toString();
  if (!notebook.isDirty) {
    clearScheduledNotebookConflictSave(
      notebookUri,
      pendingNotebookConflictSaves,
      notebookConflictSaveTimers,
    );
    return;
  }

  const hasVisibleBackingEditor = vscode.window.visibleTextEditors.some((editor) =>
    editor.document.uri.toString() === notebookUri,
  );
  if (hasVisibleBackingEditor) return;

  pendingNotebookConflictSaves.add(notebookUri);
  const existingTimer = notebookConflictSaveTimers.get(notebookUri);
  if (existingTimer) clearTimeout(existingTimer);

  notebookConflictSaveTimers.set(
    notebookUri,
    setTimeout(() => {
      void attemptNotebookConflictSave(
        notebookUri,
        pendingNotebookConflictSaves,
        notebookConflictSaveTimers,
        syncingNotebookUrisFromText,
        pendingNotebookTextSyncs,
        rController,
        pyController,
      );
    }, 1200),
  );
}

function clearScheduledNotebookConflictSave(
  notebookUri: string,
  pendingNotebookConflictSaves: Set<string>,
  notebookConflictSaveTimers: Map<string, ReturnType<typeof setTimeout>>,
): void {
  pendingNotebookConflictSaves.delete(notebookUri);
  const timer = notebookConflictSaveTimers.get(notebookUri);
  if (timer) clearTimeout(timer);
  notebookConflictSaveTimers.delete(notebookUri);
}

async function attemptNotebookConflictSave(
  notebookUri: string,
  pendingNotebookConflictSaves: Set<string>,
  notebookConflictSaveTimers: Map<string, ReturnType<typeof setTimeout>>,
  syncingNotebookUrisFromText: Set<string>,
  pendingNotebookTextSyncs: Set<string>,
  rController: RNotebookController,
  pyController: PyNotebookController,
): Promise<void> {
  const notebook = findNotebookByUri(notebookUri);
  if (!notebook || !pendingNotebookConflictSaves.has(notebookUri)) {
    clearScheduledNotebookConflictSave(
      notebookUri,
      pendingNotebookConflictSaves,
      notebookConflictSaveTimers,
    );
    return;
  }

  if (!notebook.isDirty) {
    clearScheduledNotebookConflictSave(
      notebookUri,
      pendingNotebookConflictSaves,
      notebookConflictSaveTimers,
    );
    return;
  }

  const hasVisibleBackingEditor = vscode.window.visibleTextEditors.some((editor) =>
    editor.document.uri.toString() === notebookUri,
  );
  if (hasVisibleBackingEditor) {
    clearScheduledNotebookConflictSave(
      notebookUri,
      pendingNotebookConflictSaves,
      notebookConflictSaveTimers,
    );
    return;
  }

  if (
    syncingNotebookUrisFromText.has(notebookUri)
    || pendingNotebookTextSyncs.has(notebookUri)
    || hasPendingExecutionForNotebook(notebook, rController, pyController)
  ) {
    scheduleNotebookConflictSave(
      notebook,
      pendingNotebookConflictSaves,
      notebookConflictSaveTimers,
      syncingNotebookUrisFromText,
      pendingNotebookTextSyncs,
      rController,
      pyController,
    );
    return;
  }

  try {
    await saveNotebookDocument(notebook);
  } catch {
    // If the conflict prompt is still open, retry after a short delay.
  }

  if (!notebook.isDirty) {
    clearScheduledNotebookConflictSave(
      notebookUri,
      pendingNotebookConflictSaves,
      notebookConflictSaveTimers,
    );
    return;
  }

  scheduleNotebookConflictSave(
    notebook,
    pendingNotebookConflictSaves,
    notebookConflictSaveTimers,
    syncingNotebookUrisFromText,
    pendingNotebookTextSyncs,
    rController,
    pyController,
  );
}

function registerNotebookStateWatchers(
  ctx: vscode.ExtensionContext,
  rmdOutputStore: RmdOutputStore,
  pyController: PyNotebookController,
): void {
  const patterns = ['**/*.Rmd', '**/*.rmd', '**/*.Qmd', '**/*.qmd', '**/*.ipynb'];
  const handleUri = (uri: vscode.Uri) => {
    void purgeNotebookUriState(uri, rmdOutputStore, pyController);
  };

  for (const pattern of patterns) {
    const watcher = vscode.workspace.createFileSystemWatcher(pattern);
    ctx.subscriptions.push(
      watcher,
      watcher.onDidDelete(handleUri),
      watcher.onDidCreate(handleUri),
    );
  }
}

async function purgeNotebookUriState(
  uri: vscode.Uri,
  rmdOutputStore: RmdOutputStore,
  pyController: PyNotebookController,
): Promise<void> {
  const docUri = uri.toString();
  rmdOutputStore.markHardReset(docUri);
  forgetRNotebookKernel(docUri);
  await pyController.forgetNotebookSelection(docUri);
  await purgeSessionState(docUri);
  await disposePySession(docUri);
  if (lastActiveNotebookUri === docUri) lastActiveNotebookUri = undefined;
}

async function saveNotebookDocument(notebook: vscode.NotebookDocument): Promise<void> {
  const savable = notebook as vscode.NotebookDocument & { save?: () => Thenable<boolean> };
  if (typeof savable.save !== 'function') return;
  await savable.save();
}

function buildConsoleTabHtml(chunkId: string, content: string): string {
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<style>
  body {
    font-family: var(--vscode-font-family, -apple-system, sans-serif);
    font-size: var(--vscode-font-size, 13px);
    color: var(--vscode-editor-foreground);
    background: var(--vscode-editor-background);
    padding: 16px 20px;
    margin: 0;
  }
  h1 {
    font-size: 15px;
    font-weight: 600;
    margin: 0 0 12px;
  }
  .toolbar {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 12px;
  }
  button {
    background: var(--vscode-button-background, #0078d4);
    color: var(--vscode-button-foreground, #fff);
    border: 1px solid var(--vscode-button-background, #0078d4);
    border-radius: 2px;
    padding: 3px 9px;
    font-size: 12px;
    cursor: pointer;
  }
  #copy-status {
    color: var(--vscode-descriptionForeground);
    font-size: 12px;
  }
  pre {
    margin: 0;
    white-space: pre-wrap;
    word-break: break-word;
    font-family: var(--vscode-editor-font-family, monospace);
    font-size: 12px;
    line-height: 1.45;
  }
	</style>
	</head><body>
	<h1>${escapeHtml(chunkId)}</h1>
	<div class="toolbar">
	  <button id="copy-console" type="button">Copy</button>
	  <span id="copy-status"></span>
	</div>
	<pre>${escapeHtml(content)}</pre>
	<script>
	(function () {
	  const content = ${escapeScriptJson(content)};
	  const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : null;
	  const button = document.getElementById('copy-console');
	  const status = document.getElementById('copy-status');
	  function showStatus(text) {
	    if (!status) return;
	    status.textContent = text;
	    setTimeout(() => { status.textContent = ''; }, 1800);
	  }
	  button?.addEventListener('click', async () => {
	    try {
	      if (vscode) {
	        vscode.postMessage({ type: 'copy_console', content });
	        return;
	      }
	      await navigator.clipboard.writeText(content);
	      showStatus('Copied to ClipBoard');
	    } catch {
	      showStatus('Copy failed');
	    }
	  });
	  window.addEventListener('message', (event) => {
	    if (event.data && event.data.type === 'copy_console_done') {
	      showStatus('Copied to ClipBoard');
	    }
	  });
	}());
	</script>
	</body></html>`;
}

function escapeHtml(value: string): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeScriptJson(value: string): string {
  return JSON.stringify(String(value ?? ''))
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

// ---------------------------------------------------------------------------
// Variables panel HTML

function buildVarsPanelHtml(
  sections: { language: string; doc: string; busy?: boolean; vars: { name: string; type: string; size: string; value: string }[] }[],
): string {
  function esc(s: string): string {
    return String(s ?? '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function shortDoc(uri: string): string {
    const last = uri.split(/[\\/]/).pop() ?? uri;
    return last.length > 50 ? '…' + last.slice(-47) : last;
  }

  function buildTable(
    vars: { name: string; type: string; size: string; value: string }[],
    busy = false,
  ): string {
    if (vars.length === 0) {
      return busy
        ? '<p class="empty">Session is running. No cached variables available yet.</p>'
        : '<p class="empty">No variables in session.</p>';
    }
    const rows = vars.map(v => `
      <tr>
        <td class="name">${esc(v.name)}</td>
        <td class="type">${esc(v.type)}</td>
        <td class="size">${esc(v.size)}</td>
        <td class="value">${esc(v.value)}</td>
      </tr>`).join('');
    return `<table>
      <thead><tr><th>Name</th><th>Type</th><th>Size</th><th>Value</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  }

  const rSections    = sections.filter(s => s.language === 'R');
  const pySections   = sections.filter(s => s.language === 'Python');
  const totalR       = rSections.reduce((n, s) => n + s.vars.length, 0);
  const totalPy      = pySections.reduce((n, s) => n + s.vars.length, 0);

  function buildLangSection(lang: string, secs: typeof sections): string {
    if (secs.length === 0) return '';
    const icon = lang === 'R' ? '🔵' : '🟡';
    const inner = secs.map(s => {
      const docLabel = secs.length > 1
        ? `<div class="doc-label">${esc(shortDoc(s.doc))}</div>`
        : '';
      const busyLabel = s.busy
        ? '<div class="session-note">Running now. Showing latest cached snapshot.</div>'
        : '';
      return docLabel + busyLabel + buildTable(s.vars, s.busy);
    }).join('');
    const count = secs.reduce((n, s) => n + s.vars.length, 0);
    return `
    <section>
      <h2>${icon} ${esc(lang)} <span class="count">(${count} variable${count !== 1 ? 's' : ''})</span></h2>
      ${inner}
    </section>`;
  }

  const body = sections.length === 0
    ? '<p class="empty">No active sessions. Run a cell to start a kernel.</p>'
    : buildLangSection('R', rSections) + buildLangSection('Python', pySections);

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<style>
  body {
    font-family: var(--vscode-font-family, -apple-system, sans-serif);
    font-size: var(--vscode-font-size, 13px);
    color: var(--vscode-editor-foreground);
    background: var(--vscode-editor-background);
    padding: 16px 20px;
    margin: 0;
  }
  h1 { font-size: 15px; font-weight: 600; margin: 0 0 16px; }
  h2 { font-size: 13px; font-weight: 600; margin: 20px 0 8px;
       border-bottom: 1px solid var(--vscode-editorGroup-border); padding-bottom: 4px; }
  .count { font-weight: 400; color: var(--vscode-descriptionForeground); }
  .doc-label { font-size: 11px; color: var(--vscode-descriptionForeground);
               margin-bottom: 4px; font-style: italic; }
  .session-note { font-size: 11px; color: var(--vscode-descriptionForeground); margin: 2px 0 6px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 12px; }
  thead th {
    text-align: left; padding: 4px 10px;
    background: var(--vscode-editor-background);
    border-bottom: 2px solid var(--vscode-editorGroup-border);
    font-weight: 600; position: sticky; top: 0;
  }
  td { padding: 3px 10px; border-bottom: 1px solid var(--vscode-editorGroup-border); }
  td.name  { font-weight: 600; white-space: nowrap; }
  td.type  { color: var(--vscode-descriptionForeground); white-space: nowrap; }
  td.size  { color: var(--vscode-descriptionForeground); white-space: nowrap; }
  td.value { font-family: var(--vscode-editor-font-family, monospace); font-size: 11px;
             max-width: 280px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  tr:hover td { background: var(--vscode-list-hoverBackground); }
  .empty { color: var(--vscode-descriptionForeground); font-style: italic; }
  section { margin-bottom: 8px; }
</style>
</head><body>
<h1>Session Variables</h1>
${body}
</body></html>`;
}
