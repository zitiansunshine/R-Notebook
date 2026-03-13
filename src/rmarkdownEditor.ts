// =============================================================================
// rmarkdownEditor.ts — VSCode CustomEditorProvider for .Rmd files
// =============================================================================

import * as vscode from 'vscode';
import { parseRmd, RmdChunk } from './rmdParser';
import { getOrCreateSession, disposeSession } from './rSessionManager';
import { ExecResult, DfDataResult } from './kernelProtocol';
import {
  RMARKDOWN_EDITOR_VIEW_TYPE,
  getRConfigValue,
  updateRConfigValue,
} from './extensionIds';
import { pickRKernelPath } from './kernelDiscovery';
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

  public static register(ctx: vscode.ExtensionContext): vscode.Disposable {
    const provider = new RMarkdownEditorProvider(ctx);
    return vscode.window.registerCustomEditorProvider(
      RMarkdownEditorProvider.viewType,
      provider,
      { webviewOptions: { retainContextWhenHidden: true } },
    );
  }

  constructor(private readonly ctx: vscode.ExtensionContext) {}

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
    this.postMessage(panel, { type: 'init', chunks, outputs: this.serialiseCache(outputCache) });

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

        case 'set_r_path': {
          const current = getRConfigValue('rPath', 'Rscript');
          const newPath = await pickRKernelPath({
            currentPath: current,
            title: 'Select R Kernel (Rscript path)',
            placeHolder: current,
          });
          if (newPath !== undefined && newPath !== current) {
            await updateRConfigValue('rPath', newPath, vscode.ConfigurationTarget.Global);
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
          const chunks = Array.isArray(msg.chunks)
            ? msg.chunks as RmdChunk[]
            : parseRmd(msg.fullText);
          const nextText = mergeRmdSourceAndState(
            msg.fullText,
            buildPersistedRmdStateFromChunks(chunks, outputCache),
          );
          // Reflect webview edits back to the TextDocument
          const edit = new vscode.WorkspaceEdit();
          edit.replace(
            document.uri,
            new vscode.Range(0, 0, document.lineCount, 0),
            nextText,
          );
          vscode.workspace.applyEdit(edit);
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
    outputCache.set(chunk.id, emptyExecResult(chunk.id));
    this.postMessage(panel, { type: 'chunk_running', chunk_id: chunk.id });

    try {
      const liveResult = outputCache.get(chunk.id) ?? emptyExecResult(chunk.id);
      const result = mergeConsoleResult(
        await session.exec(chunk.id, chunk.code, chunk.options as any),
        liveResult,
      );
      outputCache.set(chunk.id, result);
      this.postMessage(panel, { type: 'chunk_result', chunk_id: chunk.id, result });
      await this.persistDocumentState(document, sourceText, chunks, outputCache);
      await session.vars().catch(() => undefined);
    } catch (err: any) {
      outputCache.set(chunk.id, { ...emptyExecResult(chunk.id), error: err.message });
      this.postMessage(panel, {
        type: 'chunk_result',
        chunk_id: chunk.id,
        error: err.message,
      });
      await this.persistDocumentState(document, sourceText, chunks, outputCache);
      await session.vars().catch(() => undefined);
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

  private async persistDocumentState(
    document: vscode.TextDocument,
    sourceText: string,
    chunks: RmdChunk[],
    outputCache: Map<string, ExecResult>,
  ): Promise<void> {
    const { source: currentSource } = splitRmdSourceAndState(document.getText());
    const currentChunks = parseRmd(currentSource);
    const nextText = mergeRmdSourceAndState(
      currentSource,
      buildPersistedRmdStateFromChunks(currentChunks, outputCache),
    );
    if (document.getText() === nextText) return;

    const edit = new vscode.WorkspaceEdit();
    edit.replace(
      document.uri,
      new vscode.Range(0, 0, document.lineCount, 0),
      nextText,
    );
    await vscode.workspace.applyEdit(edit);
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

function emptyExecResult(chunkId: string): ExecResult {
  return {
    type: 'result',
    chunk_id: chunkId,
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
  if (result.console || !liveConsole) return result;
  return { ...result, console: liveConsole };
}
