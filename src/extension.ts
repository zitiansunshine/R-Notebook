// =============================================================================
// extension.ts — entrypoint for the R Notebook extension
// =============================================================================

import * as vscode from 'vscode';

import { RMarkdownEditorProvider }                                     from './rmarkdownEditor';
import { getAllSessions, purgeSessionState }                           from './rSessionManager';
import { disposePySession, getAllPySessions }                          from './pySessionManager';
import { buildNotebookCellsFromRmdText, RmdNotebookSerializer }        from './rmdNotebookSerializer';
import { IpynbSerializer }                                             from './ipynbSerializer';
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
import {
  COMMAND_IDS,
  EXTENSION_BRAND,
  VARIABLES_PANEL_VIEW_TYPE,
} from './extensionIds';

// ---------------------------------------------------------------------------

let lastActiveNotebookUri: string | undefined;

export function activate(ctx: vscode.ExtensionContext): void {
  const syncingNotebookUrisFromText = new Set<string>();

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

  rememberNotebookDocument(vscode.window.activeNotebookEditor?.notebook);
  ctx.subscriptions.push(
    vscode.window.onDidChangeActiveNotebookEditor((editor) => {
      rememberNotebookDocument(editor?.notebook);
      void ensureRmdNotebookTextDocument(editor?.notebook);
    }),
    vscode.workspace.onDidOpenNotebookDocument((notebook) => {
      void ensureRmdNotebookTextDocument(notebook);
    }),
    vscode.workspace.onDidChangeTextDocument((event) => {
      void syncNotebookDocumentFromTextChange(event, syncingNotebookUrisFromText);
    }),
    vscode.workspace.onDidChangeNotebookDocument((event) => {
      void cleanupClearedNotebookChange(event, rmdOutputStore);
    }),
  );
  registerNotebookStateWatchers(ctx, rmdOutputStore);
  for (const notebook of vscode.workspace.notebookDocuments) {
    void ensureRmdNotebookTextDocument(notebook);
  }

  if (typeof vscode.notebooks.createRendererMessaging === 'function') {
    const rendererMessaging = vscode.notebooks.createRendererMessaging('rNotebook.execResultRenderer');
    ctx.subscriptions.push(
      rendererMessaging.onDidReceiveMessage(async ({ editor, message }) => {
        if (!message || typeof message !== 'object') return;
        if (message.type !== 'open_console_in_tab') return;

        const chunkId = typeof message.chunkId === 'string' && message.chunkId
          ? message.chunkId
          : 'console';
        const content = typeof message.content === 'string'
          ? message.content
          : String(message.content ?? '');
        const panel = vscode.window.createWebviewPanel(
          'rNotebook.consoleOutput',
          `Console: ${chunkId}`,
          editor.viewColumn ?? vscode.ViewColumn.Beside,
          { enableScripts: false, retainContextWhenHidden: true },
        );
        panel.webview.html = buildConsoleTabHtml(chunkId, content);
      }),
    );
  }

  ctx.subscriptions.push(
    vscode.workspace.onDidCloseNotebookDocument((notebook) => {
      const uri = notebook.uri.toString();
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
  if (Array.isArray(target)) {
    for (const entry of target) {
      const resolved = resolveNotebookTarget(entry);
      if (resolved) return resolved;
    }
    return undefined;
  }

  if (target instanceof vscode.Uri) {
    return findNotebookByUri(target.toString());
  }

  if (target && typeof target === 'object') {
    const candidate = target as {
      notebook?: vscode.NotebookDocument;
      notebookEditor?: vscode.NotebookEditor;
      cell?: vscode.NotebookCell;
      uri?: vscode.Uri;
      resource?: vscode.Uri;
      notebookType?: string;
      notebookUri?: vscode.Uri;
    };
    if (isManagedNotebookDocument(candidate.notebook)) return candidate.notebook;
    if (isManagedNotebookDocument(candidate.notebookEditor?.notebook)) return candidate.notebookEditor?.notebook;
    if (isManagedNotebookDocument(candidate.cell?.notebook)) return candidate.cell?.notebook;
    if (candidate.uri && candidate.notebookType) {
      const notebook = candidate as unknown as vscode.NotebookDocument;
      return isManagedNotebookDocument(notebook) ? notebook : undefined;
    }
    const resource = candidate.notebookUri ?? candidate.resource ?? candidate.uri;
    if (resource) return findNotebookByUri(resource.toString());
  }

  return undefined;
}

function rememberNotebookDocument(notebook?: vscode.NotebookDocument): void {
  if (!isManagedNotebookDocument(notebook)) return;
  lastActiveNotebookUri = notebook.uri.toString();
}

function findNotebookByUri(uri?: string): vscode.NotebookDocument | undefined {
  if (!uri) return undefined;
  const notebook = vscode.workspace.notebookDocuments.find((doc) => doc.uri.toString() === uri);
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

async function ensureRmdNotebookTextDocument(
  notebook?: vscode.NotebookDocument,
): Promise<void> {
  if (notebook?.notebookType !== NOTEBOOK_TYPE) return;
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
): Promise<void> {
  const notebook = findNotebookByUri(event.document.uri.toString());
  if (!notebook || notebook.notebookType !== NOTEBOOK_TYPE) return;
  if (!isRmdLikeUri(event.document.uri)) return;

  const notebookUri = notebook.uri.toString();
  if (syncingNotebookUrisFromText.has(notebookUri)) return;

  const nextCells = buildNotebookCellsFromRmdText(event.document.getText())
    .map((cell) => stripCodeCellOutputs(cell));
  if (notebookCellsMatch(notebook, nextCells)) return;

  syncingNotebookUrisFromText.add(notebookUri);
  try {
    const edit = new vscode.WorkspaceEdit();
    edit.set(notebook.uri, [
      buildReplaceNotebookCellsEdit(notebook, nextCells),
    ]);
    await vscode.workspace.applyEdit(edit);
  } finally {
    syncingNotebookUrisFromText.delete(notebookUri);
  }
}

function stripCodeCellOutputs(cell: vscode.NotebookCellData): vscode.NotebookCellData {
  const next = new vscode.NotebookCellData(cell.kind, cell.value, cell.languageId);
  next.metadata = cell.metadata as vscode.NotebookCellMetadata | undefined;
  next.outputs = cell.kind === vscode.NotebookCellKind.Code ? [] : [...cell.outputs];
  return next;
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
  }

  return true;
}

function notebookCellMetadataEqual(
  left: vscode.NotebookCellMetadata | undefined,
  right: vscode.NotebookCellMetadata | undefined,
): boolean {
  return stableStringify(left ?? {}) === stableStringify(right ?? {});
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

function isRmdLikeUri(uri: vscode.Uri): boolean {
  const fsPath = uri.fsPath.toLowerCase();
  return fsPath.endsWith('.rmd') || fsPath.endsWith('.qmd');
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

function registerNotebookStateWatchers(
  ctx: vscode.ExtensionContext,
  rmdOutputStore: RmdOutputStore,
): void {
  const patterns = ['**/*.Rmd', '**/*.rmd', '**/*.Qmd', '**/*.qmd', '**/*.ipynb'];
  const handleUri = (uri: vscode.Uri) => {
    void purgeNotebookUriState(uri, rmdOutputStore);
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
): Promise<void> {
  const docUri = uri.toString();
  rmdOutputStore.markHardReset(docUri);
  forgetRNotebookKernel(docUri);
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
<pre>${escapeHtml(content)}</pre>
</body></html>`;
}

function escapeHtml(value: string): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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
