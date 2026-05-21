// =============================================================================
// rmarkdownEditor.ts — VSCode CustomEditorProvider for .Rmd files
// =============================================================================

import * as vscode from 'vscode';
import { parseRmd, RmdChunk } from './rmdParser';
import { getOrCreateSession, disposeSession } from './rSessionManager';
import { ExecResult, DfDataResult } from './kernelProtocol';
import {
  RMARKDOWN_EDITOR_VIEW_TYPE,
  getRAdditionalExecutablePaths,
  getRConfigValue,
  getRFigureDefaults,
  mergeRFigureOptions,
  rememberRExecutablePath,
  updateRConfigValue,
} from './extensionIds';
import { pickRKernelPath } from './kernelDiscovery';
import { clearPersistedRmdOutputs, exportDocumentUri } from './notebookActions';
import {
  buildPersistedRmdStateFromChunks,
  mergeRmdSourceAndState,
  restoreChunkResults,
  splitRmdSourceAndState,
} from './rmdPersistedState';

// ---------------------------------------------------------------------------

export class RMarkdownEditorProvider implements vscode.CustomTextEditorProvider {

  public static readonly viewType = RMARKDOWN_EDITOR_VIEW_TYPE;
  private readonly runQueues = new Map<string, Promise<void>>();
  private readonly runEpochs = new Map<string, number>();
  private readonly outputCaches = new Map<string, Map<string, ExecResult>>();
  private readonly documentWriteQueues = new Map<string, Promise<void>>();
  private readonly webviewTextVersions = new Map<string, number>();

  public static register(ctx: vscode.ExtensionContext): vscode.Disposable {
    const provider = new RMarkdownEditorProvider(ctx);
    return vscode.window.registerCustomEditorProvider(
      RMarkdownEditorProvider.viewType,
      provider,
      { webviewOptions: { retainContextWhenHidden: true } },
    );
  }

  constructor(private readonly ctx: vscode.ExtensionContext) {
    ctx.subscriptions.push(
      vscode.workspace.onWillSaveTextDocument((event) => this.handleWillSaveTextDocument(event)),
    );
  }

  // ---- CustomTextEditorProvider --------------------------------------------

  async resolveCustomTextEditor(
    document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
    _token: vscode.CancellationToken,
  ): Promise<void> {
    const docUri = document.uri.toString();
    const outputCache = this.outputCache(docUri);
    panel.webview.options = { enableScripts: true };
    panel.webview.html = this.buildHtml(panel.webview);

    // Initial render (always show the document, even if R fails to start)
    const { source, state } = splitRmdSourceAndState(document.getText());
    const chunks = parseRmd(source);
    this.resetOutputCache(outputCache, restoreChunkResults(chunks, state));
    this.postMessage(panel, {
      type: 'init',
      chunks,
      outputs: this.serialiseCache(outputCache),
      fig_defaults: getRFigureDefaults(),
    });

    // Spin up R session (non-fatal: errors shown in the webview).
    // Store named listener references so they can be removed when the panel closes,
    // preventing zombie listeners if the session object is reused (e.g. split editor).
    const session = getOrCreateSession(docUri, this.rBinPath(), this.execTimeoutMs());
    const onSessionError = (err: Error) => {
      this.postMessage(panel, { type: 'kernel_error', message: err.message });
    };
    const onSessionStderr = (d: string) => {
      this.postMessage(panel, { type: 'kernel_stderr', text: d });
    };
    const onSessionExit = () => {
      this.postMessage(panel, { type: 'kernel_exit' });
    };
    const onSessionProgress = (msg: any) => {
      this.postMessage(panel, {
        type: 'chunk_progress',
        chunk_id: msg.chunk_id,
        line: msg.line,
        total: msg.total,
      });
    };
    const onSessionStream = (msg: any) => {
      if (!msg?.chunk_id || !msg?.text) return;
      const current = outputCache.get(msg.chunk_id) ?? emptyExecResult(msg.chunk_id);
      if (msg.stream === 'stderr') {
        current.stderr += msg.text;
      } else {
        current.stdout += msg.text;
      }
      current.console = `${current.console ?? ''}${msg.text}`;
      outputCache.set(msg.chunk_id, current);
      this.postMessage(panel, { type: 'chunk_stream', chunk_id: msg.chunk_id, result: current });
    };
    session.on('error',    onSessionError);
    session.on('stderr',   onSessionStderr);
    session.on('exit',     onSessionExit);
    session.on('progress', onSessionProgress);
    session.on('stream',   onSessionStream);
    try {
      await session.start();
    } catch (err: any) {
      this.postMessage(panel, { type: 'kernel_error', message: err.message });
    }

    // Document changes
    const changeDisp = vscode.workspace.onDidChangeTextDocument(e => {
      if (e.document.uri.toString() === docUri) {
        const { source: updatedSource, state: updatedState } = splitRmdSourceAndState(e.document.getText());
        const updated = parseRmd(updatedSource);
        this.resetOutputCache(outputCache, restoreChunkResults(updated, updatedState));
        this.postMessage(panel, {
          type: 'chunks_updated',
          chunks: updated,
          outputs: this.serialiseCache(outputCache),
        });
      }
    });

    // Messages from WebView
    panel.webview.onDidReceiveMessage(async msg => {
      switch (msg.type) {

        case 'run_chunk': {
          const chunk: RmdChunk = msg.chunk;
          const sourceText = typeof msg.source === 'string'
            ? msg.source
            : splitRmdSourceAndState(document.getText()).source;
          const chunks = Array.isArray(msg.chunks) ? msg.chunks as RmdChunk[] : parseRmd(sourceText);
          void this.enqueueDocumentRun(docUri, (runEpoch) =>
            this.executeChunk(panel, session, document, outputCache, chunk, chunks, sourceText, runEpoch),
          ).catch((err: any) => {
            this.postMessage(panel, {
              type: 'chunk_result',
              chunk_id: chunk.id,
              error: err?.message ?? String(err),
            });
          });
          break;
        }

        case 'run_all': {
          const sourceText = typeof msg.source === 'string'
            ? msg.source
            : splitRmdSourceAndState(document.getText()).source;
          const chunks: RmdChunk[] = Array.isArray(msg.chunks) ? msg.chunks : parseRmd(sourceText);
          void this.enqueueDocumentRun(docUri, (runEpoch) =>
            this.executeChunks(panel, session, document, outputCache, chunks, sourceText, runEpoch),
          ).catch((err: any) => {
            this.postMessage(panel, { type: 'error', message: err?.message ?? String(err) });
          });
          break;
        }

        case 'df_page': {
          try {
            const result: DfDataResult = await session.dfPage(
              msg.chunk_id, msg.name, msg.page, msg.page_size,
            );
            this.postMessage(panel, { ...result, type: 'df_data' });
          } catch (err: any) {
            this.postMessage(panel, { type: 'error', message: err.message });
          }
          break;
        }

        case 'copy_console': {
          const content = typeof msg.content === 'string'
            ? msg.content
            : String(msg.content ?? '');
          await vscode.env.clipboard.writeText(content);
          void vscode.window.setStatusBarMessage('Copied to ClipBoard', 2000);
          break;
        }

        case 'open_output_in_tab': {
          const content = typeof msg.content === 'string'
            ? msg.content
            : String(msg.content ?? '');
          const title = typeof msg.title === 'string' && msg.title
            ? msg.title
            : 'R Notebook Output';
          const outputPanel = vscode.window.createWebviewPanel(
            'rNotebook.consoleOutput',
            title,
            vscode.ViewColumn.Beside,
            { enableScripts: true, retainContextWhenHidden: true },
          );
          const copyDisposable = outputPanel.webview.onDidReceiveMessage(async (panelMessage) => {
            if (!panelMessage || typeof panelMessage !== 'object') return;
            if (panelMessage.type !== 'copy_console') return;
            const panelContent = typeof panelMessage.content === 'string'
              ? panelMessage.content
              : String(panelMessage.content ?? '');
            await vscode.env.clipboard.writeText(panelContent);
            void vscode.window.setStatusBarMessage('Copied to ClipBoard', 2000);
            await outputPanel.webview.postMessage({ type: 'copy_console_done' });
          });
          outputPanel.onDidDispose(() => copyDisposable.dispose());
          outputPanel.webview.html = buildTextTabHtml(title, content);
          break;
        }

        case 'set_r_path': {
          const current = getRConfigValue('rPath', 'Rscript');
          const newPath = await pickRKernelPath({
            currentPath: current,
            additionalPaths: getRAdditionalExecutablePaths(),
            title: 'Select R Kernel (Rscript path)',
            placeHolder: current,
          });
          if (newPath !== undefined && newPath !== current) {
            await updateRConfigValue('rPath', newPath, vscode.ConfigurationTarget.Global);
            await rememberRExecutablePath(newPath);
            session.setExecutablePath(newPath);
            await session.restart();
            this.postMessage(panel, { type: 'session_reset' });
          }
          break;
        }

        case 'interrupt_kernel': {
          this.bumpRunEpoch(docUri);
          session.interrupt();
          break;
        }

        case 'reset_session': {
          // Kill and restart the R process so it's truly fresh.
          try { await session.restart(); } catch (err: any) {
            this.postMessage(panel, { type: 'kernel_error', message: `Restart failed: ${err.message}` });
          }
          outputCache.clear();
          const { source: currentSource } = splitRmdSourceAndState(document.getText());
          await this.persistDocumentState(
            document,
            currentSource,
            parseRmd(currentSource),
            outputCache,
          );
          this.postMessage(panel, { type: 'session_reset' });
          break;
        }

        case 'clear_outputs': {
          const sourceText = typeof msg.fullText === 'string'
            ? msg.fullText
            : splitRmdSourceAndState(document.getText()).source;
          outputCache.clear();
          const nextText = clearPersistedRmdOutputs(sourceText);
          const edit = new vscode.WorkspaceEdit();
          edit.replace(
            document.uri,
            new vscode.Range(0, 0, document.lineCount, 0),
            nextText,
          );
          await vscode.workspace.applyEdit(edit);
          await document.save();
          this.postMessage(panel, { type: 'outputs_cleared' });
          break;
        }

        case 'export_document': {
          const sourceText = typeof msg.fullText === 'string'
            ? msg.fullText
            : splitRmdSourceAndState(document.getText()).source;
          const chunks = Array.isArray(msg.chunks) ? msg.chunks as RmdChunk[] : parseRmd(sourceText);
          const nextText = mergeRmdSourceAndState(
            sourceText,
            buildPersistedRmdStateFromChunks(chunks, outputCache),
          );
          const edit = new vscode.WorkspaceEdit();
          edit.replace(
            document.uri,
            new vscode.Range(0, 0, document.lineCount, 0),
            nextText,
          );
          await vscode.workspace.applyEdit(edit);
          await document.save();
          await exportDocumentUri(document.uri);
          break;
        }

        case 'get_vars': {
          try {
            const result = await session.vars();
            this.postMessage(panel, { type: 'vars_result', vars: result.vars });
          } catch {
            this.postMessage(panel, { type: 'vars_result', vars: [] });
          }
          break;
        }

        case 'get_completions': {
          try {
            const completions = await session.complete(msg.chunk_id, msg.code, msg.cursor_pos);
            this.postMessage(panel, { type: 'completions_result', chunk_id: msg.chunk_id, completions, cursor_pos: msg.cursor_pos });
          } catch {
            this.postMessage(panel, { type: 'completions_result', chunk_id: msg.chunk_id, completions: [], cursor_pos: msg.cursor_pos });
          }
          break;
        }

        case 'code_changed': {
          await this.applyWebviewDocumentText(document, msg.fullText, msg.chunks, outputCache, msg.version, {
            persistOutputs: false,
          });
          break;
        }

        case 'save_document': {
          await this.applyWebviewDocumentText(document, msg.fullText, msg.chunks, outputCache, msg.version, {
            persistOutputs: false,
          });
          await document.save();
          break;
        }
      }
    });

    panel.onDidDispose(async () => {
      changeDisp.dispose();
      session.removeListener('error',    onSessionError);
      session.removeListener('stderr',   onSessionStderr);
      session.removeListener('exit',     onSessionExit);
      session.removeListener('progress', onSessionProgress);
      session.removeListener('stream',   onSessionStream);
      this.outputCaches.delete(docUri);
      await disposeSession(docUri);
    });
  }

  // ---- Helpers -------------------------------------------------------------

  private postMessage(panel: vscode.WebviewPanel, msg: object): void {
    panel.webview.postMessage(msg);
  }

  private serialiseCache(cache: Map<string, ExecResult>): Record<string, ExecResult> {
    return Object.fromEntries(cache.entries());
  }

  private rBinPath(): string {
    return getRConfigValue('rPath', 'Rscript');
  }

  private execTimeoutMs(): number {
    return getRConfigValue('execTimeoutMs', 3_000_000) ?? 3_000_000;
  }

  private runEpoch(docUri: string): number {
    return this.runEpochs.get(docUri) ?? 0;
  }

  private bumpRunEpoch(docUri: string): number {
    const next = this.runEpoch(docUri) + 1;
    this.runEpochs.set(docUri, next);
    return next;
  }

  private isRunInterrupted(docUri: string, runEpoch: number): boolean {
    return this.runEpoch(docUri) !== runEpoch;
  }

  private enqueueDocumentRun(
    docUri: string,
    task: (runEpoch: number) => Promise<void>,
  ): Promise<void> {
    const runEpoch = this.runEpoch(docUri);
    const previous = this.runQueues.get(docUri) ?? Promise.resolve();
    const next = previous.catch(() => undefined).then(async () => {
      if (this.isRunInterrupted(docUri, runEpoch)) return;
      await task(runEpoch);
    });
    const tracked = next.finally(() => {
      if (this.runQueues.get(docUri) === tracked) {
        this.runQueues.delete(docUri);
      }
    });
    this.runQueues.set(docUri, tracked);
    return tracked;
  }

  private async executeChunks(
    panel: vscode.WebviewPanel,
    session: ReturnType<typeof getOrCreateSession>,
    document: vscode.TextDocument,
    outputCache: Map<string, ExecResult>,
    chunks: RmdChunk[],
    sourceText: string,
    runEpoch: number,
  ): Promise<void> {
    session.setExecTimeoutMs(this.execTimeoutMs());
    for (const chunk of chunks) {
      if (this.isRunInterrupted(document.uri.toString(), runEpoch)) break;
      if (chunk.kind !== 'code' || chunk.language !== 'r') continue;
      if (chunk.options.eval === false) continue;
      await this.executeChunk(panel, session, document, outputCache, chunk, chunks, sourceText, runEpoch);
      if (this.isRunInterrupted(document.uri.toString(), runEpoch)) break;
    }
  }

  private async executeChunk(
    panel: vscode.WebviewPanel,
    session: ReturnType<typeof getOrCreateSession>,
    document: vscode.TextDocument,
    outputCache: Map<string, ExecResult>,
    chunk: RmdChunk,
    chunks: RmdChunk[],
    sourceText: string,
    runEpoch: number,
  ): Promise<void> {
    if (this.isRunInterrupted(document.uri.toString(), runEpoch)) return;
    if (chunk.language !== 'r') {
      this.postMessage(panel, {
        type: 'chunk_result',
        chunk_id: chunk.id,
        error: `Language '${chunk.language}' not supported yet`,
      });
      return;
    }

    session.setExecTimeoutMs(this.execTimeoutMs());
    const runningResult = emptyExecResult(chunk.id, chunk.code);
    outputCache.set(chunk.id, runningResult);
    this.postMessage(panel, { type: 'chunk_running', chunk_id: chunk.id, result: runningResult });

    try {
      const liveResult = outputCache.get(chunk.id) ?? emptyExecResult(chunk.id, chunk.code);
      const result = mergeConsoleResult(
        await session.exec(chunk.id, chunk.code, mergeRFigureOptions(chunk.options as Record<string, unknown>)),
        liveResult,
      );
      outputCache.set(chunk.id, result);
      this.postMessage(panel, { type: 'chunk_result', chunk_id: chunk.id, result });
      await this.persistDocumentState(document, sourceText, chunks, outputCache);
    } catch (err: any) {
      outputCache.set(chunk.id, { ...emptyExecResult(chunk.id, chunk.code), error: err.message });
      this.postMessage(panel, {
        type: 'chunk_result',
        chunk_id: chunk.id,
        error: err.message,
      });
      await this.persistDocumentState(document, sourceText, chunks, outputCache);
    }
  }

  private outputCache(docUri: string): Map<string, ExecResult> {
    const cached = this.outputCaches.get(docUri);
    if (cached) return cached;
    const next = new Map<string, ExecResult>();
    this.outputCaches.set(docUri, next);
    return next;
  }

  private resetOutputCache(
    target: Map<string, ExecResult>,
    next: Map<string, ExecResult>,
  ): void {
    target.clear();
    for (const [key, value] of next) target.set(key, value);
  }

  private enqueueDocumentWrite(
    docUri: string,
    task: () => Promise<void>,
  ): Promise<void> {
    const previous = this.documentWriteQueues.get(docUri) ?? Promise.resolve();
    const next = previous.catch(() => undefined).then(task);
    const tracked = next.finally(() => {
      if (this.documentWriteQueues.get(docUri) === tracked) {
        this.documentWriteQueues.delete(docUri);
      }
    });
    this.documentWriteQueues.set(docUri, tracked);
    return tracked;
  }

  private async applyWebviewDocumentText(
    document: vscode.TextDocument,
    fullText: unknown,
    chunksValue: unknown,
    outputCache: Map<string, ExecResult>,
    versionValue?: unknown,
    options?: { persistOutputs?: boolean },
  ): Promise<void> {
    const sourceText = typeof fullText === 'string'
      ? fullText
      : splitRmdSourceAndState(document.getText()).source;
    const chunks = options?.persistOutputs
      ? Array.isArray(chunksValue)
        ? chunksValue as RmdChunk[]
        : parseRmd(sourceText)
      : [];
    const docUri = document.uri.toString();
    const version = typeof versionValue === 'number' && Number.isFinite(versionValue)
      ? versionValue
      : undefined;
    if (version !== undefined && version < (this.webviewTextVersions.get(docUri) ?? -1)) {
      return;
    }
    await this.enqueueDocumentWrite(docUri, async () => {
      if (version !== undefined && version < (this.webviewTextVersions.get(docUri) ?? -1)) {
        return;
      }
      const nextText = options?.persistOutputs
        ? mergeRmdSourceAndState(
          sourceText,
          buildPersistedRmdStateFromChunks(chunks, outputCache),
        )
        : this.buildDocumentTextWithExistingPersistedState(document, sourceText);
      if (document.getText() === nextText) {
        if (version !== undefined) {
          this.webviewTextVersions.set(docUri, version);
        }
        return;
      }
      const edit = new vscode.WorkspaceEdit();
      edit.replace(
        document.uri,
        new vscode.Range(0, 0, document.lineCount, 0),
        nextText,
      );
      const applied = await vscode.workspace.applyEdit(edit);
      if (!applied) throw new Error('Failed to write R Markdown document changes.');
      if (version !== undefined) {
        this.webviewTextVersions.set(docUri, version);
      }
    });
  }

  private async persistDocumentState(
    document: vscode.TextDocument,
    _sourceText: string,
    _chunks: RmdChunk[],
    outputCache: Map<string, ExecResult>,
  ): Promise<void> {
    const docUri = document.uri.toString();
    await this.enqueueDocumentWrite(docUri, async () => {
      const nextText = this.buildPersistedDocumentText(document, outputCache);
      if (!nextText) return;
      const edit = new vscode.WorkspaceEdit();
      edit.replace(
        document.uri,
        new vscode.Range(0, 0, document.lineCount, 0),
        nextText,
      );
      const applied = await vscode.workspace.applyEdit(edit);
      if (!applied) throw new Error('Failed to persist R Markdown output state.');
    });
  }

  private handleWillSaveTextDocument(event: vscode.TextDocumentWillSaveEvent): void {
    const docUri = event.document.uri.toString();
    const outputCache = this.outputCaches.get(docUri);
    if (!outputCache) return;
    if (this.runQueues.has(docUri)) return;

    event.waitUntil((async () => {
      const pending = this.documentWriteQueues.get(docUri);
      if (pending) await pending.catch(() => undefined);
      const nextText = this.buildPersistedDocumentText(event.document, outputCache);
      if (!nextText) return [];
      return [
        vscode.TextEdit.replace(
          new vscode.Range(0, 0, event.document.lineCount, 0),
          nextText,
        ),
      ];
    })());
  }

  private buildPersistedDocumentText(
    document: vscode.TextDocument,
    outputCache: Map<string, ExecResult>,
  ): string | null {
    const { source: currentSource } = splitRmdSourceAndState(document.getText());
    const currentChunks = parseRmd(currentSource);
    const nextText = mergeRmdSourceAndState(
      currentSource,
      buildPersistedRmdStateFromChunks(currentChunks, outputCache),
    );
    return document.getText() === nextText ? null : nextText;
  }

  private buildDocumentTextWithExistingPersistedState(
    document: vscode.TextDocument,
    sourceText: string,
  ): string {
    const { state } = splitRmdSourceAndState(document.getText());
    return mergeRmdSourceAndState(sourceText, state);
  }

  private buildHtml(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.ctx.extensionUri, 'webview', 'rmarkdownPanel.js'),
    );
    const styleUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.ctx.extensionUri, 'webview', 'rmarkdownPanel.css'),
    );
    const nonce = getNonce();
    return /* html */`
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <meta http-equiv="Content-Security-Policy"
    content="default-src 'none';
             img-src ${webview.cspSource} data: blob:;
             style-src ${webview.cspSource} 'unsafe-inline';
             script-src 'nonce-${nonce}';">
  <link rel="stylesheet" href="${styleUri}"/>
  <title>RMarkdown</title>
</head>
<body>
  <div id="toolbar">
    <button id="btn-run-all">▶ Run All</button>
    <button id="btn-clear-outputs" title="Remove all saved outputs from this notebook">⌫ Clear Outputs</button>
    <button id="btn-export" title="Export this notebook to HTML or PDF">⇪ Export</button>
    <button id="btn-interrupt" title="Interrupt running R code">⏹ Interrupt</button>
    <button id="btn-reset">↺ Restart R</button>
    <button id="btn-r-path" title="Set path to Rscript executable">⚙ R Path</button>
    <button id="btn-line-nums" title="Toggle line numbers">⑂ Lines</button>
    <span id="kernel-status" class="status-idle">● Idle</span>
    <span class="add-chunk-split"><button id="btn-add-chunk" title="Add R chunk after selected chunk">+ Chunk</button><button id="btn-add-chunk-arrow" title="Choose chunk type to add">▾</button></span>
    <button id="btn-del-chunk" title="Delete selected chunk">✂ Delete</button>
    <button id="btn-vars" title="Toggle variable inspector">⊡ Vars</button>
    <span style="margin-left:8px;font-size:9px;opacity:0.35;font-family:monospace">v0.5.1</span>
  </div>
  <div id="main-layout">
    <div id="rmd-container"></div>
    <div id="var-panel">
      <div class="var-panel-header">
        <span class="var-panel-title">Variables</span>
        <button id="btn-vars-refresh" title="Refresh">↻</button>
        <button id="btn-vars-close" title="Close">✕</button>
      </div>
      <div class="var-panel-body">
        <table class="var-table">
          <thead><tr><th>Name</th><th>Type</th><th>Size</th><th>Value</th></tr></thead>
          <tbody id="var-table-body"><tr><td colspan="4" class="var-empty">No variables yet — run a chunk to populate.</td></tr></tbody>
        </table>
      </div>
    </div>
  </div>
  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

function getNonce(): string {
  let text = '';
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++)
    text += chars.charAt(Math.floor(Math.random() * chars.length));
  return text;
}

function emptyExecResult(chunkId: string, sourceCode = ''): ExecResult {
  return {
    type: 'result',
    chunk_id: chunkId,
    source_code: sourceCode,
    console_segments: sourceCode ? [{ code: sourceCode, output: '' }] : [],
    console: '',
    stdout: '',
    stderr: '',
    plots: [],
    plots_html: [],
    dataframes: [],
    error: null,
  };
}

function mergeConsoleResult(result: ExecResult, liveResult: ExecResult): ExecResult {
  const liveConsole = liveResult.console ?? '';
  return {
    ...result,
    source_code: result.source_code ?? liveResult.source_code,
    console_segments: result.console_segments ?? liveResult.console_segments,
    console: result.console || liveConsole || '',
  };
}

function buildTextTabHtml(title: string, content: string): string {
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
h1 { font-size: 15px; font-weight: 600; margin: 0 0 12px; }
.toolbar { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
button {
  background: var(--vscode-button-background, #0078d4);
  color: var(--vscode-button-foreground, #fff);
  border: 1px solid var(--vscode-button-background, #0078d4);
  border-radius: 2px;
  padding: 3px 9px;
  font-size: 12px;
  cursor: pointer;
}
#copy-status { color: var(--vscode-descriptionForeground); font-size: 12px; }
pre {
  margin: 0;
  white-space: pre-wrap;
  word-break: break-word;
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 12px;
  line-height: 1.45;
}
</style></head><body>
<h1>${escapeHtml(title)}</h1>
<div class="toolbar">
  <button id="copy-output" type="button">Copy</button>
  <span id="copy-status"></span>
</div>
<pre>${escapeHtml(content)}</pre>
<script>
(function () {
  const content = ${escapeScriptJson(content)};
  const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : null;
  const button = document.getElementById('copy-output');
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
