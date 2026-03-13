// =============================================================================
// extension.ts — entrypoint for the R Notebook extension
// =============================================================================

import * as vscode from 'vscode';

import { RMarkdownEditorProvider }                                     from './rmarkdownEditor';
import { disposeSession, getAllSessions, getSession }                  from './rSessionManager';
import { disposePySession, getAllPySessions, getPySession }            from './pySessionManager';
import { RmdNotebookSerializer }                                       from './rmdNotebookSerializer';
import { IpynbSerializer }                                             from './ipynbSerializer';
import { RNotebookController, NOTEBOOK_TYPE, registerFigureOptions }   from './rNotebookController';
import { registerRNotebookCompletionProvider }                         from './rNotebookCompletionProvider';
import { RmdOutputStore }                                              from './rmdOutputStore';
import { PyNotebookController, PY_NOTEBOOK_TYPE, registerPyFigureOptions } from './pyNotebookController';
import { RNotebookVariableProvider }                                   from './variableProvider';
import { forgetRNotebookKernel }                                       from './notebookKernelState';
import { registerNotebookKernelSourceProviders }                       from './notebookKernelSourceProvider';
import {
  COMMAND_IDS,
  EXTENSION_BRAND,
  VARIABLES_PANEL_VIEW_TYPE,
} from './extensionIds';

// ---------------------------------------------------------------------------

export function activate(ctx: vscode.ExtensionContext): void {

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

  ctx.subscriptions.push(
    vscode.commands.registerCommand(
      COMMAND_IDS.notebookSelectAvailableRKernels,
      async (target?: unknown) => {
        const notebook = resolveNotebookDocument(target);
        if (notebook?.notebookType && notebook.notebookType !== NOTEBOOK_TYPE) return;
        await rController.showKernelPicker(notebook);
      },
    ),
    vscode.commands.registerCommand(
      COMMAND_IDS.notebookSelectAvailablePythonKernels,
      async (target?: unknown) => {
        const notebook = resolveNotebookDocument(target);
        if (notebook?.notebookType && notebook.notebookType !== PY_NOTEBOOK_TYPE) return;
        await pyController.showKernelPicker(notebook, {
          includeManual: false,
          title: 'Select Python Kernel',
        });
      },
    ),
  );
  registerNotebookKernelSourceProviders(ctx);

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
        const uri  = notebook.uri.toString();
        const type = notebook.notebookType;

        if (type === NOTEBOOK_TYPE) {
          rController.interruptNotebook(uri);
          interrupted = true;
        } else if (type === PY_NOTEBOOK_TYPE) {
          pyController.interruptNotebook(uri);
          interrupted = true;
        }
      }

      if (!interrupted) {
        for (const [uri, session] of getAllSessions()) {
          if (!session.isBusy()) continue;
          rController.interruptNotebook(uri);
          interrupted = true;
        }
        for (const [uri, session] of getAllPySessions()) {
          if (!session.isBusy()) continue;
          pyController.interruptNotebook(uri);
          interrupted = true;
        }
      }
    }),
  );

  // ── Notebook toolbar: Restart kernel ─────────────────────────────────────
  ctx.subscriptions.push(
    vscode.commands.registerCommand(COMMAND_IDS.notebookRestart, async (target?: unknown) => {
      const notebook = resolveNotebookDocument(target);
      if (!notebook) return;
      const uri  = notebook.uri.toString();
      const type = notebook.notebookType;

      try {
        if (type === NOTEBOOK_TYPE) {
          await getSession(uri)?.restart();
        } else if (type === PY_NOTEBOOK_TYPE) {
          await getPySession(uri)?.restart();
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
    vscode.commands.registerCommand(COMMAND_IDS.notebookShowVariables, async () => {
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

}

export function deactivate(): void { /* sessions disposed per-document */ }

function resolveNotebookDocument(target?: unknown): vscode.NotebookDocument | undefined {
  if (target && typeof target === 'object') {
    const candidate = target as {
      notebook?: vscode.NotebookDocument;
      notebookEditor?: vscode.NotebookEditor;
      cell?: vscode.NotebookCell;
      uri?: vscode.Uri;
      notebookType?: string;
      notebookUri?: vscode.Uri;
    };
    if (candidate.notebook) return candidate.notebook;
    if (candidate.notebookEditor?.notebook) return candidate.notebookEditor.notebook;
    if (candidate.cell?.notebook) return candidate.cell.notebook;
    if (candidate.uri && candidate.notebookType) {
      return candidate as unknown as vscode.NotebookDocument;
    }
    const resource = candidate.notebookUri ?? candidate.uri;
    if (resource) {
      const match = vscode.workspace.notebookDocuments.find((doc) => doc.uri.toString() === resource.toString());
      if (match) return match;
    }
  }
  return vscode.window.activeNotebookEditor?.notebook;
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
