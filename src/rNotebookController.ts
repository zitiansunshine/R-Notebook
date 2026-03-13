// =============================================================================
// rNotebookController.ts — VS Code NotebookController that runs R cells via
// the R Notebook R kernel subprocess.
//
// Each cell's outputs are combined into a single text/html block whose layout
// matches the webview (rmarkdownPanel.css) exactly:
//   • text stdout  → .output-text     (left border, monospace)
//   • stderr       → .output-stderr   (warning colours)
//   • errors       → .output-error    (error colours)
//   • plots        → .plot-wrap img   (with hover "Save PNG" button)
//   • dataframes   → .df-viewer       (title bar, sticky-header table)
//   • 2+ outputs   → CSS-only thumbnail strip identical to the webview,
//                    using radio inputs so tabs work without JavaScript.
//
// Per-cell figure options (fig.width / fig.height / dpi) are exposed via a
// status bar item on each R cell. Clicking it opens a quick-pick editor that
// stores the values in cell.metadata.options, picked up on next execution.
// =============================================================================

import * as vscode from 'vscode';
import { getOrCreateSession, getSession } from './rSessionManager';
import {
  COMMAND_IDS,
  NOTEBOOK_TYPE as R_NOTEBOOK_TYPE,
  affectsRConfig,
  getRConfigValue,
  updateRConfigValue,
} from './extensionIds';
import { ExecResult, ProgressMessage, StreamMessage, StreamOutputMessage } from './kernelProtocol';
import { discoverRKernels, RKernelDescriptor } from './kernelDiscovery';
import { getRememberedRNotebookKernel, rememberRNotebookKernel } from './notebookKernelState';
import {
  notebookOutputItemsFromExecResult,
  notebookOutputFromExecResult,
  RAW_EXEC_RESULT_MIME,
  RmdOutputStore,
} from './rmdOutputStore';
import type { VariableNotifier } from './pyNotebookController';

export const NOTEBOOK_TYPE = R_NOTEBOOK_TYPE;

type LiveExecutionState = {
  chunkId: string;
  cellDocUri: string;
  execution: vscode.NotebookCellExecution;
  output: vscode.NotebookCellOutput | null;
  result: ExecResult;
  dirty: boolean;      // has unrendered stream output
  rendering: boolean;  // replaceOutput in-flight
  renderPromise: Promise<void> | null;
  completed: boolean;
};

type RControllerEntry = {
  controller: vscode.NotebookController;
  descriptor: RKernelDescriptor;
  selectionDisposable: vscode.Disposable;
};

export class RNotebookController {
  private readonly controllers = new Map<string, RControllerEntry>();
  private readonly liveExecutions = new Map<string, LiveExecutionState>();
  private readonly liveCellDocuments = new Set<string>();
  private readonly deletedLiveCellDocuments = new Set<string>();
  private readonly notebookQueues = new Map<string, Promise<void>>();
  private readonly queueEpochs = new Map<string, number>();
  private readonly boundSessions = new WeakSet<object>();
  private readonly restoreInFlight = new Set<string>();
  private readonly disposables: vscode.Disposable[] = [];
  private readonly extensionId: string;
  private kernelDescriptors: RKernelDescriptor[] = [];
  private renderInterval: ReturnType<typeof setInterval> | null = null;
  private varNotifier?: VariableNotifier;

  constructor(
    ctx: vscode.ExtensionContext,
    private readonly outputStore: RmdOutputStore,
    varNotifier?: VariableNotifier,
  ) {
    this.extensionId = ctx.extension.id;
    this.varNotifier = varNotifier;
    this.refreshKernelCatalog();
    this.disposables.push(
      vscode.workspace.onDidOpenNotebookDocument((notebook) => {
        if (notebook.notebookType !== NOTEBOOK_TYPE) return;
        this.outputStore.syncNotebook(notebook);
        this.syncNotebookAffinities(notebook);
        void this.restoreNotebookOutputs(notebook);
      }),
      vscode.workspace.onDidChangeNotebookDocument((event) => {
        if (event.notebook.notebookType !== NOTEBOOK_TYPE) return;
        this.handleDeletedRunningCells(event);
        this.outputStore.applyNotebookChange(event);
        if (this.restoreInFlight.has(event.notebook.uri.toString())) return;
        if (!this.shouldRestoreNotebookOutputs(event)) return;
        void this.restoreNotebookOutputs(event.notebook);
      }),
      vscode.workspace.onDidChangeConfiguration((event) => {
        if (affectsRConfig(event, 'rPath')) {
          this.refreshKernelCatalog();
        }
      }),
    );
  }

  public async showKernelPicker(notebook?: vscode.NotebookDocument): Promise<void> {
    const activeNotebook = vscode.window.activeNotebookEditor?.notebook;
    const targetNotebook = notebook?.notebookType === NOTEBOOK_TYPE
      ? notebook
      : activeNotebook?.notebookType === NOTEBOOK_TYPE
        ? activeNotebook
        : undefined;

    const currentDescriptor = targetNotebook
      ? this.resolveDescriptorForNotebook(targetNotebook)
      : this.resolveDefaultDescriptor();
    const descriptors = this.availableDescriptors();
    const selectableDescriptors = descriptors.length > 0
      ? descriptors
      : [currentDescriptor];

    type PickItem = vscode.QuickPickItem & {
      descriptor: RKernelDescriptor;
    };

    const items: PickItem[] = selectableDescriptors.map((descriptor) => ({
      label: descriptor.displayName,
      description: descriptor.id === currentDescriptor.id
        ? `${descriptor.description ?? 'R kernel'} · current`
        : descriptor.description,
      detail: descriptor.rPath,
      descriptor,
    }));

    const pick = await vscode.window.showQuickPick(items, {
      title: 'Select R Kernel',
      placeHolder: currentDescriptor.rPath,
      matchOnDescription: true,
      matchOnDetail: true,
    });
    if (!pick) return;

    if (targetNotebook) {
      await this.selectDescriptorForNotebook(targetNotebook, pick.descriptor);
      vscode.window.showInformationMessage(`R kernel set to ${pick.descriptor.displayName}.`);
      return;
    }

    await updateRConfigValue('rPath', pick.descriptor.rPath, vscode.ConfigurationTarget.Global);
    this.refreshKernelCatalog();
    vscode.window.showInformationMessage(`R kernel path set to: ${pick.descriptor.rPath}`);
  }

  private rBinPath(): string {
    return getRConfigValue('rPath', 'Rscript');
  }

  private execTimeoutMs(): number {
    return getRConfigValue('execTimeoutMs', 0) ?? 0;
  }

  private availableDescriptors(): RKernelDescriptor[] {
    if (this.kernelDescriptors.length === 0) {
      this.refreshKernelCatalog();
    }
    return this.kernelDescriptors;
  }

  private refreshKernelCatalog(): void {
    const configuredR = this.rBinPath();
    const discovered = discoverRKernels(configuredR);
    this.kernelDescriptors = discovered.length > 0
      ? discovered
      : [{
        id: `executable:${configuredR}`,
        label: configuredR,
        description: 'R executable',
        detail: configuredR,
        rPath: configuredR,
        displayName: configuredR,
        source: 'executable',
      }];
    const nextIds = new Set<string>();

    for (const descriptor of this.kernelDescriptors) {
      const controllerId = this.controllerIdForDescriptor(descriptor);
      nextIds.add(controllerId);

      const existing = this.controllers.get(controllerId);
      if (existing) {
        existing.descriptor = descriptor;
        existing.controller.label = descriptor.displayName;
        existing.controller.description = descriptor.description ?? 'R kernel';
        existing.controller.detail = descriptor.rPath;
        continue;
      }

      const controller = vscode.notebooks.createNotebookController(
        controllerId,
        NOTEBOOK_TYPE,
        descriptor.displayName,
      );
      controller.supportedLanguages = ['r', 'R'];
      controller.supportsExecutionOrder = false;
      controller.description = descriptor.description ?? 'R kernel';
      controller.detail = descriptor.rPath;
      controller.executeHandler = this._execute.bind(this);
      controller.interruptHandler = (notebook) => {
        this.interruptNotebook(notebook.uri.toString());
      };

      const selectionDisposable = controller.onDidChangeSelectedNotebooks((event) => {
        if (!event.selected) return;
        const selectedDescriptor = this.controllers.get(controller.id)?.descriptor ?? descriptor;
        void this.handleControllerSelection(event.notebook, selectedDescriptor, controller.id);
      });

      this.controllers.set(controllerId, { controller, descriptor, selectionDisposable });
    }

    for (const [controllerId, entry] of this.controllers) {
      if (nextIds.has(controllerId)) continue;
      entry.selectionDisposable.dispose();
      entry.controller.dispose();
      this.controllers.delete(controllerId);
    }

    this.syncAllNotebookAffinities();
  }

  private controllerIdForDescriptor(descriptor: RKernelDescriptor): string {
    return `r-notebook-r-kernel:${descriptor.id}`;
  }

  private controllerEntryForController(
    controller: vscode.NotebookController,
  ): RControllerEntry | undefined {
    return this.controllers.get(controller.id);
  }

  private preferredControllerForNotebook(
    notebook: vscode.NotebookDocument,
  ): vscode.NotebookController | undefined {
    const descriptor = this.resolveDescriptorForNotebook(notebook);
    return this.controllers.get(this.controllerIdForDescriptor(descriptor))?.controller
      ?? this.controllers.values().next().value?.controller;
  }

  private syncAllNotebookAffinities(): void {
    for (const notebook of vscode.workspace.notebookDocuments) {
      this.syncNotebookAffinities(notebook);
    }
  }

  private syncNotebookAffinities(notebook: vscode.NotebookDocument): void {
    if (notebook.notebookType !== NOTEBOOK_TYPE) return;
    const preferred = this.resolveDescriptorForNotebook(notebook);
    for (const entry of this.controllers.values()) {
      entry.controller.updateNotebookAffinity(
        notebook,
        entry.descriptor.id === preferred.id
          ? vscode.NotebookControllerAffinity.Preferred
          : vscode.NotebookControllerAffinity.Default,
      );
    }
  }

  private resolveDefaultDescriptor(): RKernelDescriptor {
    const configuredR = this.rBinPath();
    return this.availableDescriptors().find((descriptor) => descriptor.rPath === configuredR)
      ?? this.availableDescriptors()[0]
      ?? {
        id: `executable:${configuredR}`,
        label: configuredR,
        description: 'R executable',
        detail: configuredR,
        rPath: configuredR,
        displayName: configuredR,
        source: 'executable',
      };
  }

  private resolveDescriptorForNotebook(notebook: vscode.NotebookDocument): RKernelDescriptor {
    const remembered = getRememberedRNotebookKernel(notebook.uri.toString());
    return this.availableDescriptors().find((descriptor) =>
      this.controllerIdForDescriptor(descriptor) === remembered?.controllerId
      || descriptor.kernelspecName === remembered?.kernelspecName
      || descriptor.displayName === remembered?.displayName
      || descriptor.rPath === remembered?.rPath,
    ) ?? this.resolveDefaultDescriptor();
  }

  private async handleControllerSelection(
    notebook: vscode.NotebookDocument,
    descriptor: RKernelDescriptor,
    controllerId: string,
  ): Promise<void> {
    if (notebook.notebookType !== NOTEBOOK_TYPE) return;
    rememberRNotebookKernel(notebook.uri.toString(), {
      controllerId,
      displayName: descriptor.displayName,
      kernelspecName: descriptor.kernelspecName,
      rPath: descriptor.rPath,
    });

    const session = getSession(notebook.uri.toString());
    if (session) {
      session.setExecutablePath(descriptor.rPath);
      await session.restart().catch(() => undefined);
      this.varNotifier?.notifyChanged(notebook);
    }

    this.syncNotebookAffinities(notebook);
  }

  private async selectDescriptorForNotebook(
    notebook: vscode.NotebookDocument,
    descriptor: RKernelDescriptor,
  ): Promise<void> {
    const controllerId = this.controllerIdForDescriptor(descriptor);
    rememberRNotebookKernel(notebook.uri.toString(), {
      controllerId,
      displayName: descriptor.displayName,
      kernelspecName: descriptor.kernelspecName,
      rPath: descriptor.rPath,
    });

    const didSelect = await this.selectNotebookController(notebook, controllerId);
    if (!didSelect) {
      await this.handleControllerSelection(notebook, descriptor, controllerId);
      return;
    }

    this.syncNotebookAffinities(notebook);
  }

  private async selectNotebookController(
    notebook: vscode.NotebookDocument,
    controllerId: string,
  ): Promise<boolean> {
    if (!this.controllers.has(controllerId)) return false;
    const editor = await this.resolveNotebookEditor(notebook);
    if (!editor) return false;

    await vscode.commands.executeCommand('notebook.selectKernel', {
      editor,
      id: controllerId,
      extension: this.extensionId,
    });
    return true;
  }

  private async resolveNotebookEditor(
    notebook: vscode.NotebookDocument,
  ): Promise<vscode.NotebookEditor | undefined> {
    const key = notebook.uri.toString();
    const active = vscode.window.activeNotebookEditor;
    if (active?.notebook.uri.toString() === key) return active;

    const visible = vscode.window.visibleNotebookEditors.find((editor) =>
      editor.notebook.uri.toString() === key,
    );
    if (visible) return visible;

    try {
      return await vscode.window.showNotebookDocument(notebook, {
        preserveFocus: true,
        preview: false,
        viewColumn: vscode.ViewColumn.Active,
      });
    } catch {
      return undefined;
    }
  }

  private async _execute(
    cells: vscode.NotebookCell[],
    notebook: vscode.NotebookDocument,
    ctrl: vscode.NotebookController,
  ): Promise<void> {
    const docUri = notebook.uri.toString();
    return this.enqueueNotebookExecution(docUri, (runEpoch) =>
      this.runQueuedExecution(docUri, runEpoch, cells, notebook, ctrl),
    );
  }

  private async runQueuedExecution(
    docUri: string,
    runEpoch: number,
    cells: vscode.NotebookCell[],
    notebook: vscode.NotebookDocument,
    ctrl: vscode.NotebookController,
  ): Promise<void> {
    const descriptor = this.controllerEntryForController(ctrl)?.descriptor
      ?? this.resolveDescriptorForNotebook(notebook);
    rememberRNotebookKernel(docUri, {
      controllerId: this.controllerIdForDescriptor(descriptor),
      displayName: descriptor.displayName,
      kernelspecName: descriptor.kernelspecName,
      rPath: descriptor.rPath,
    });
    const session = getOrCreateSession(docUri, descriptor.rPath, this.execTimeoutMs());
    this.bindSession(docUri, session);

    let startError: Error | undefined;
    try {
      await session.start();
      if (session.cachedVars().vars.length === 0) {
        await this.refreshVariableCache(session);
      }
    } catch (err: any) {
      startError = err instanceof Error ? err : new Error(String(err));
    }

    for (const cell of cells) {
      if (this.isQueueInterrupted(docUri, runEpoch)) break;

      const execution = ctrl.createNotebookCellExecution(cell);
      execution.start(Date.now());

      const chunkId = `nb-${cell.index}`;
      const code    = cell.document.getText();
      const opts    = (cell.metadata?.options ?? {}) as Record<string, unknown>;
      const liveKey = this.liveKey(docUri, chunkId);
      const cellDocUri = cell.document.uri.toString();
      const liveState: LiveExecutionState = {
        chunkId,
        cellDocUri,
        execution,
        output: null,
        result: emptyExecResult(chunkId, code),
        dirty: false,
        rendering: false,
        renderPromise: null,
        completed: false,
      };
      this.liveExecutions.set(liveKey, liveState);
      this.liveCellDocuments.add(cellDocUri);
      const cancelDisp = execution.token.onCancellationRequested(() => session.interrupt());

      try {
        await this.replaceExecutionOutput(liveState, liveState.result, chunkId, { running: true });
        if (startError) throw startError;
        const result = mergeConsoleResult(await session.exec(chunkId, code, {
          fig_width:  opts['fig_width']  as number | undefined,
          fig_height: opts['fig_height'] as number | undefined,
          dpi:        opts['dpi']        as number | undefined,
        }), liveState.result);
        liveState.completed = true;
        await this.waitForPendingRender(liveState);
        if (this.deletedLiveCellDocuments.has(cellDocUri)) continue;
        await this.replaceExecutionOutput(liveState, result, chunkId);
        this.outputStore.setNotebookCellResult(
          notebook,
          cell.index,
          notebookOutputFromExecResult(result, chunkId).length > 0 ? result : null,
        );
        execution.end(!result.error, Date.now());
        await this.refreshVariableCache(session);
      } catch (err: any) {
        const errorResult: ExecResult = {
          ...liveState.result,
          error: err.message,
        };
        liveState.completed = true;
        await this.waitForPendingRender(liveState);
        if (this.deletedLiveCellDocuments.has(cellDocUri)) continue;
        await this.replaceExecutionOutput(liveState, errorResult, chunkId);
        this.outputStore.setNotebookCellResult(
          notebook,
          cell.index,
          notebookOutputFromExecResult(errorResult, chunkId).length > 0 ? errorResult : null,
        );
        execution.end(false, Date.now());
        await this.refreshVariableCache(session);
      } finally {
        cancelDisp.dispose();
        this.liveCellDocuments.delete(cellDocUri);
        this.deletedLiveCellDocuments.delete(cellDocUri);
        this.liveExecutions.delete(liveKey);
      }

      if (this.isQueueInterrupted(docUri, runEpoch)) break;
    }

    // Notify variable provider so the Variables panel refreshes
    this.varNotifier?.notifyChanged(notebook);
  }

  private enqueueNotebookExecution(
    docUri: string,
    task: (runEpoch: number) => Promise<void>,
  ): Promise<void> {
    const runEpoch = this.queueEpoch(docUri);
    const previous = this.notebookQueues.get(docUri) ?? Promise.resolve();
    const next = previous.catch(() => undefined).then(async () => {
      if (this.isQueueInterrupted(docUri, runEpoch)) return;
      await task(runEpoch);
    });
    const tracked = next.finally(() => {
      if (this.notebookQueues.get(docUri) === tracked) {
        this.notebookQueues.delete(docUri);
      }
    });
    this.notebookQueues.set(docUri, tracked);
    return tracked;
  }

  public interruptNotebook(docUri: string): boolean {
    const session = getSession(docUri);
    const interrupted = session?.isBusy() ?? false;
    this.bumpQueueEpoch(docUri);
    session?.interrupt();
    return interrupted;
  }

  public async restartNotebook(notebook: vscode.NotebookDocument): Promise<boolean> {
    if (notebook.notebookType !== NOTEBOOK_TYPE) return false;
    const docUri = notebook.uri.toString();
    const descriptor = this.resolveDescriptorForNotebook(notebook);
    const session = getOrCreateSession(docUri, descriptor.rPath, this.execTimeoutMs());
    session.setExecutablePath(descriptor.rPath);
    this.bindSession(docUri, session);
    await session.restart();
    this.varNotifier?.notifyChanged(notebook);
    return true;
  }

  private bindSession(docUri: string, session: ReturnType<typeof getOrCreateSession>): void {
    if (this.boundSessions.has(session as object)) return;
    this.boundSessions.add(session as object);
    session.on('progress', (msg: ProgressMessage) => this.handleProgress(docUri, msg));
    session.on('stream', (msg: StreamMessage) => this.handleStream(docUri, msg));
    session.on('stream_output', (msg: StreamOutputMessage) => this.handleStreamOutput(docUri, msg));
    session.on('stderr', (text: string) => this.handleKernelStderr(docUri, text));
  }

  private handleProgress(docUri: string, msg: ProgressMessage): void {
    const state = this.liveExecutions.get(this.liveKey(docUri, msg.chunk_id));
    if (!state || state.completed) return;
    if (!msg.expr_code) return;
    state.result.console_segments = [
      ...(state.result.console_segments ?? []),
      { code: msg.expr_code, output: '' },
    ];
    state.dirty = true;
    this.startRenderLoop();
  }

  private handleStreamOutput(docUri: string, msg: StreamOutputMessage): void {
    const state = this.liveExecutions.get(this.liveKey(docUri, msg.chunk_id));
    if (!state || state.completed) return;
    if (msg.kind === 'plot' && msg.b64) {
      state.result.plots = [...(state.result.plots ?? [])];
      state.result.plots[msg.index] = msg.b64;
    } else if (msg.kind === 'df' && msg.df) {
      state.result.dataframes = [...(state.result.dataframes ?? [])];
      state.result.dataframes[msg.index] = msg.df;
    }
    state.result.output_order = [
      ...(state.result.output_order ?? []),
      { type: msg.kind, index: msg.index, name: msg.name },
    ];
    state.dirty = true;
    this.startRenderLoop();
  }

  private handleStream(docUri: string, msg: StreamMessage): void {
    const state = this.liveExecutions.get(this.liveKey(docUri, msg.chunk_id));
    if (!state || state.completed || !msg.text) return;
    if (this.deletedLiveCellDocuments.has(state.cellDocUri)) return;
    this.appendLiveConsoleText(state, msg.stream, msg.text);
  }

  private handleKernelStderr(docUri: string, text: string): void {
    const state = this.activeLiveExecutionForDocument(docUri);
    if (!state || state.completed || !text) return;
    if (this.deletedLiveCellDocuments.has(state.cellDocUri)) return;
    this.appendLiveConsoleText(state, 'stderr', text);
  }

  private activeLiveExecutionForDocument(docUri: string): LiveExecutionState | undefined {
    const prefix = `${docUri}::`;
    for (const [liveKey, state] of this.liveExecutions) {
      if (liveKey.startsWith(prefix)) return state;
    }
    return undefined;
  }

  private appendLiveConsoleText(
    state: LiveExecutionState,
    stream: 'stdout' | 'stderr',
    text: string,
  ): void {
    const currentConsole = state.result.console ?? '';
    const needsSeparator =
      currentConsole.length > 0 &&
      !currentConsole.endsWith('\n') &&
      !text.startsWith('\n');
    if (stream === 'stderr') {
      state.result.stderr += text;
    } else {
      state.result.stdout += text;
    }
    state.result.console = `${currentConsole}${needsSeparator ? '\n' : ''}${text}`;
    if (text && state.result.console_segments && state.result.console_segments.length > 0) {
      const segments = state.result.console_segments.map((segment) => ({ ...segment }));
      const lastSegment = segments[segments.length - 1];
      const existingOutput = lastSegment.output ?? '';
      const separator = existingOutput.length > 0 ? '\n' : '';
      lastSegment.output = `${existingOutput}${separator}${text}`;
      state.result.console_segments = segments;
    }
    // Mark dirty and let the render loop flush at its own pace (≤10×/sec).
    // This prevents flooding VS Code's extension host with one replaceOutput
    // call per stream line, which causes lag and dropped execution state.
    state.dirty = true;
    this.startRenderLoop();
  }

  /** Starts a shared 100 ms render interval that flushes dirty live states. */
  private startRenderLoop(): void {
    if (this.renderInterval) return;
    this.renderInterval = setInterval(() => {
      if (this.liveExecutions.size === 0) {
        clearInterval(this.renderInterval!);
        this.renderInterval = null;
        return;
      }
      for (const state of this.liveExecutions.values()) {
        if (!state.dirty || state.completed || state.rendering) continue;
        if (this.deletedLiveCellDocuments.has(state.cellDocUri)) continue;
        state.dirty = false;
        state.rendering = true;
        const renderPromise = this.replaceExecutionOutput(
          state,
          state.result,
          state.chunkId,
          { running: true },
        );
        state.renderPromise = renderPromise;
        void renderPromise
          .catch(() => undefined)
          .finally(() => {
            state.rendering = false;
            if (state.renderPromise === renderPromise) {
              state.renderPromise = null;
            }
          });
      }
    }, 100);
  }

  private async waitForPendingRender(state: LiveExecutionState): Promise<void> {
    if (!state.renderPromise) return;
    try {
      await state.renderPromise;
    } catch {}
  }

  private liveKey(docUri: string, chunkId: string): string {
    return `${docUri}::${chunkId}`;
  }

  private queueEpoch(docUri: string): number {
    return this.queueEpochs.get(docUri) ?? 0;
  }

  private bumpQueueEpoch(docUri: string): number {
    const next = this.queueEpoch(docUri) + 1;
    this.queueEpochs.set(docUri, next);
    return next;
  }

  private isQueueInterrupted(docUri: string, runEpoch: number): boolean {
    return this.queueEpoch(docUri) !== runEpoch;
  }

  private handleDeletedRunningCells(event: vscode.NotebookDocumentChangeEvent): void {
    let interrupted = false;
    for (const change of event.contentChanges) {
      for (const cell of change.removedCells) {
        const cellDocUri = cell.document.uri.toString();
        if (!this.liveCellDocuments.has(cellDocUri)) continue;
        this.deletedLiveCellDocuments.add(cellDocUri);
        for (const state of this.liveExecutions.values()) {
          if (state.cellDocUri !== cellDocUri) continue;
          state.completed = true;
        }
        interrupted = true;
      }
    }
    if (interrupted) {
      this.interruptNotebook(event.notebook.uri.toString());
    }
  }

  private async replaceExecutionOutput(
    state: LiveExecutionState,
    result: ExecResult,
    chunkId: string,
    options?: { running?: boolean },
  ): Promise<void> {
    const items = notebookOutputItemsFromExecResult(result, chunkId, options);
    if (!items) {
      state.output = null;
      await state.execution.clearOutput();
      return;
    }
    if (!state.output) {
      state.output = new vscode.NotebookCellOutput(items, {
        chunkId,
        running: Boolean(options?.running),
      });
      await state.execution.replaceOutput([state.output]);
      return;
    }
    state.output.metadata = {
      chunkId,
      running: Boolean(options?.running),
    };
    await state.execution.replaceOutputItems(items, state.output);
  }

  private async restoreNotebookOutputs(notebook: vscode.NotebookDocument): Promise<void> {
    const docUri = notebook.uri.toString();
    if (this.restoreInFlight.has(docUri)) return;
    const cached = this.outputStore.getForNotebook(notebook);
    if (cached.size === 0) return;

    // Do not create new cell executions while any notebook is actively running.
    // Calling createNotebookCellExecution on this controller while other cells are
    // executing can cause VS Code to cancel those in-progress executions.
    if (this.liveExecutions.size > 0) return;

    this.restoreInFlight.add(docUri);
    try {
      for (const cell of notebook.getCells()) {
        if (cell.kind !== vscode.NotebookCellKind.Code) continue;
        if (cell.outputs.length > 0) continue;
        if (this.liveCellDocuments.has(cell.document.uri.toString())) continue;
        const result = cached.get(cell.index);
        if (!result) continue;
        const controller = this.preferredControllerForNotebook(notebook);
        if (!controller) continue;
        const execution = controller.createNotebookCellExecution(cell);
        execution.start(Date.now());
        await execution.replaceOutput(notebookOutputFromExecResult(result, `nb-${cell.index}`));
        execution.end(!result.error, Date.now());
      }
    } finally {
      this.restoreInFlight.delete(docUri);
    }
  }

  private shouldRestoreNotebookOutputs(event: vscode.NotebookDocumentChangeEvent): boolean {
    if (event.contentChanges.length > 0) return true;
    return event.cellChanges.some((change) => Boolean(change.document || change.metadata));
  }

  private async refreshVariableCache(
    session: ReturnType<typeof getOrCreateSession>,
  ): Promise<void> {
    try {
      await session.vars();
    } catch {}
  }

  /** Clean up when the extension deactivates */
  dispose(): void {
    if (this.renderInterval) {
      clearInterval(this.renderInterval);
      this.renderInterval = null;
    }
    for (const disposable of this.disposables) disposable.dispose();
    for (const entry of this.controllers.values()) {
      entry.selectionDisposable.dispose();
      entry.controller.dispose();
    }
  }
}

function emptyExecResult(chunkId: string, sourceCode = ''): ExecResult {
  return {
    type: 'result',
    chunk_id: chunkId,
    source_code: sourceCode,
    console_segments: [],
    console: '',
    stdout: '',
    stderr: '',
    plots: [],
    plots_html: [],
    dataframes: [],
    output_order: [],
    error: null,
  };
}

function mergeConsoleResult(result: ExecResult, liveResult: ExecResult): ExecResult {
  const liveConsole = liveResult.console ?? '';
  const sourceCode = result.source_code ?? liveResult.source_code;
  const consoleSegments = result.console_segments ?? liveResult.console_segments;
  if (result.console || !liveConsole) {
    if (
      sourceCode && result.source_code !== sourceCode ||
      consoleSegments && result.console_segments !== consoleSegments
    ) {
      return {
        ...result,
        source_code: sourceCode,
        console_segments: consoleSegments,
      };
    }
    return result;
  }
  return {
    ...result,
    console: liveConsole,
    source_code: sourceCode,
    console_segments: consoleSegments,
  };
}

// HTML builders live in notebookOutputHtml.ts (shared with Python controller)


// =============================================================================
// Per-cell figure options — status bar item + quick-pick editor
// =============================================================================

type FigField = {
  key:   'fig_width' | 'fig_height' | 'dpi';
  label: string;
  unit:  string;
  def:   number;
  min:   number;
  max:   number;
};

const FIG_FIELDS: FigField[] = [
  { key: 'fig_width',  label: 'fig.width',  unit: 'in',  def: 7,   min: 1,  max: 20  },
  { key: 'fig_height', label: 'fig.height', unit: 'in',  def: 5,   min: 1,  max: 20  },
  { key: 'dpi',        label: 'dpi',        unit: 'dpi', def: 120, min: 72, max: 600 },
];

/**
 * Status bar provider — shows current fig settings on every R code cell.
 * Clicking opens the quick-pick editor (rNotebook.r.setCellFigureOptions).
 */
export class RCellFigureStatusBar implements vscode.NotebookCellStatusBarItemProvider {

  private readonly _emitter = new vscode.EventEmitter<void>();
  /** Fire to force VS Code to re-query all cell status bar items. */
  readonly onDidChangeCellStatusBarItems = this._emitter.event;

  refresh(): void { this._emitter.fire(); }
  dispose():  void { this._emitter.dispose(); }

  provideCellStatusBarItems(
    cell: vscode.NotebookCell,
  ): vscode.NotebookCellStatusBarItem[] {
    if (cell.kind !== vscode.NotebookCellKind.Code) return [];
    if (!['r', 'R'].includes(cell.document.languageId)) return [];

    const opts = (cell.metadata?.options ?? {}) as Record<string, unknown>;
    const fw  = opts['fig_width']  as number | undefined;
    const fh  = opts['fig_height'] as number | undefined;
    const dpi = opts['dpi']        as number | undefined;

    const parts: string[] = [];
    if (fw !== undefined || fh !== undefined)
      parts.push(`${fw ?? 7}×${fh ?? 5}in`);
    if (dpi !== undefined)
      parts.push(`${dpi}dpi`);

    const text = parts.length ? `⚙ ${parts.join(' · ')}` : '⚙ plot size';

    const item = new vscode.NotebookCellStatusBarItem(
      text,
      vscode.NotebookCellStatusBarAlignment.Right,
    );
    item.command = {
      command:   COMMAND_IDS.rSetCellFigureOptions,
      title:     'Edit Figure Options',
      arguments: [cell],
    };
    item.tooltip = 'fig.width / fig.height / dpi — click to edit';
    return [item];
  }
}

/**
 * Multi-step quick-pick editor for per-cell figure options.
 * Loops until the user picks "Done" or presses Escape.
 * Writes values back to cell.metadata.options via a WorkspaceEdit.
 */
async function editCellFigureOptions(
  cell: vscode.NotebookCell,
  statusBar: RCellFigureStatusBar,
): Promise<void> {

  // Work on a mutable copy so we can show live descriptions while looping
  const opts: Record<string, unknown> = {
    ...((cell.metadata?.options ?? {}) as Record<string, unknown>),
  };

  while (true) {
    type PickItem = vscode.QuickPickItem & { field?: FigField; done?: true };

    const items: PickItem[] = [
      ...FIG_FIELDS.map(f => {
        const cur = opts[f.key] as number | undefined;
        return {
          label:       `$(symbol-ruler) ${f.label}`,
          description: cur !== undefined
            ? `${cur} ${f.unit}`
            : `default (${f.def} ${f.unit})`,
          detail: cur !== undefined
            ? 'Custom value · select to edit or clear'
            : 'Using default · select to override',
          field: f,
        } satisfies PickItem;
      }),
      {
        label:       '$(check) Done',
        description: 'Save and close',
        done:        true,
      },
    ];

    const picked = await vscode.window.showQuickPick(items, {
      title:            'Cell Figure Options',
      placeHolder:      'Select a setting to edit  ·  Escape to cancel without saving',
      matchOnDescription: false,
    });

    if (!picked || picked.done) break;     // Escape or Done → save & exit loop
    if (!picked.field) break;

    const f   = picked.field;
    const cur = opts[f.key] as number | undefined;

    const raw = await vscode.window.showInputBox({
      title:  `Set ${f.label}`,
      prompt: `${f.min}–${f.max} ${f.unit}. Leave empty to reset to default (${f.def}).`,
      value:  cur !== undefined ? String(cur) : '',
      validateInput: v => {
        if (!v.trim()) return null; // empty = reset to default
        const n = parseFloat(v);
        return isNaN(n) || n < f.min || n > f.max
          ? `Enter a number between ${f.min} and ${f.max}`
          : null;
      },
    });

    if (raw === undefined) continue; // user cancelled InputBox → back to QuickPick

    if (!raw.trim()) {
      delete opts[f.key];             // reset to default
    } else {
      opts[f.key] = parseFloat(raw);
    }
  }

  // Persist into cell metadata
  const edit = new vscode.WorkspaceEdit();
  edit.set(cell.notebook.uri, [
    vscode.NotebookEdit.updateCellMetadata(cell.index, {
      ...cell.metadata,
      options: opts,
    }),
  ]);
  await vscode.workspace.applyEdit(edit);

  // Refresh status bar so the new values appear immediately
  statusBar.refresh();
}

/**
 * Register the status bar provider and the setCellFigureOptions command.
 * Call once from extension.activate().
 */
export function registerFigureOptions(ctx: vscode.ExtensionContext): void {
  const provider = new RCellFigureStatusBar();

  ctx.subscriptions.push(
    vscode.notebooks.registerNotebookCellStatusBarItemProvider(
      NOTEBOOK_TYPE,
      provider,
    ),
    provider,
  );

  ctx.subscriptions.push(
    vscode.commands.registerCommand(
      COMMAND_IDS.rSetCellFigureOptions,
      async (cell?: vscode.NotebookCell) => {
        // Called from status bar click: cell is passed as argument.
        // Called from command palette: resolve from active editor selection.
        if (!cell) {
          const editor = vscode.window.activeNotebookEditor;
          if (!editor) {
            vscode.window.showWarningMessage('No active R notebook cell.');
            return;
          }
          const sel = editor.selections[0];
          cell = editor.notebook.cellAt(sel?.start ?? 0);
        }
        await editCellFigureOptions(cell, provider);
      },
    ),
  );
}
