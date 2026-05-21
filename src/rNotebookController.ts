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
  getRAdditionalExecutablePaths,
  getRConfigValue,
  getRFigureDefaults,
  mergeRFigureOptions,
  rememberRExecutablePath,
  updateRConfigValue,
} from './extensionIds';
import { ExecResult, ProgressMessage, StreamMessage, StreamOutputMessage } from './kernelProtocol';
import { discoverRKernelsAsync, RKernelDescriptor } from './kernelDiscovery';
import { getRememberedRNotebookKernel, rememberRNotebookKernel } from './notebookKernelState';
import {
  notebookOutputItemsFromExecResult,
  notebookOutputFromExecResult,
  RAW_EXEC_RESULT_MIME,
  RmdOutputStore,
} from './rmdOutputStore';
import { hasPersistedRmdState } from './rmdPersistedState';
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
  lastRenderStartedAt: number;
  completed: boolean;
};

type RKernelActionDescriptor = {
  id: 'action:manual';
  label: string;
  description?: string;
  detail?: string;
  action: 'manual';
};

type RControllerDescriptor = RKernelDescriptor | RKernelActionDescriptor;

type RControllerEntry = {
  controller: vscode.NotebookController;
  descriptor: RControllerDescriptor;
  selectionDisposable: vscode.Disposable;
};

function isActionDescriptor(descriptor: RControllerDescriptor): descriptor is RKernelActionDescriptor {
  return 'action' in descriptor;
}

const OUTPUT_UPDATE_TIMEOUT_MS = 15_000;
const LIVE_RENDER_INTERVAL_MS = 100;
const LARGE_CONSOLE_CHARS = 200_000;
const HUGE_CONSOLE_CHARS = 1_000_000;
const LARGE_CONSOLE_RENDER_INTERVAL_MS = 1_000;
const HUGE_CONSOLE_RENDER_INTERVAL_MS = 3_000;

class OutputUpdateTimeoutError extends Error {
  constructor(label: string) {
    super(`Notebook ${label} timed out`);
    this.name = 'OutputUpdateTimeoutError';
  }
}

function isOutputUpdateTimeout(err: unknown): boolean {
  return err instanceof OutputUpdateTimeoutError
    || (err instanceof Error && err.name === 'OutputUpdateTimeoutError');
}

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
  private kernelCatalogRefreshSeq = 0;

  constructor(
    ctx: vscode.ExtensionContext,
    private readonly outputStore: RmdOutputStore,
    varNotifier?: VariableNotifier,
  ) {
    this.extensionId = ctx.extension.id;
    this.varNotifier = varNotifier;
    void this.refreshKernelCatalog();
    this.disposables.push(
      vscode.workspace.onDidOpenNotebookDocument((notebook) => {
        if (notebook.notebookType !== NOTEBOOK_TYPE) return;
        void this.prepareOpenedNotebook(notebook);
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
        if (affectsRConfig(event, 'rPath') || affectsRConfig(event, 'additionalRPaths')) {
          void this.refreshKernelCatalog();
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
    void this.refreshKernelCatalog();
    vscode.window.showInformationMessage(`R kernel path set to: ${pick.descriptor.rPath}`);
  }

  private rBinPath(): string {
    return getRConfigValue('rPath', 'Rscript');
  }

  private additionalRPaths(): string[] {
    return getRAdditionalExecutablePaths();
  }

  private execTimeoutMs(): number {
    return getRConfigValue('execTimeoutMs', 0) ?? 0;
  }

  private availableDescriptors(): RKernelDescriptor[] {
    if (this.kernelDescriptors.length === 0) {
      const configuredR = this.rBinPath();
      this.installKernelCatalog([this.fallbackDescriptor(configuredR)]);
      void this.refreshKernelCatalog();
    }
    return this.kernelDescriptors;
  }

  private manualDescriptor(
    rPath: string,
    overrides?: {
      displayName?: string;
      kernelspecName?: string;
    },
  ): RKernelDescriptor {
    const parts = rPath.split(/[\\/]/).filter(Boolean);
    const base = parts[parts.length - 1] ?? rPath;
    const displayName = overrides?.displayName ?? (base === rPath ? `R (${rPath})` : `R (${base})`);
    return {
      id: `executable:${rPath}`,
      label: displayName,
      description: 'R executable',
      detail: rPath,
      rPath,
      displayName,
      kernelspecName: overrides?.kernelspecName,
      source: 'executable',
    };
  }

  private ensureControllerForDescriptor(descriptor: RKernelDescriptor): RControllerEntry {
    const controllerId = this.controllerIdForDescriptor(descriptor);
    const existingDescriptorIndex = this.kernelDescriptors.findIndex((candidate) =>
      candidate.id === descriptor.id,
    );
    if (existingDescriptorIndex >= 0) {
      this.kernelDescriptors[existingDescriptorIndex] = descriptor;
    } else {
      this.kernelDescriptors.push(descriptor);
    }

    const existing = this.controllers.get(controllerId);
    if (existing) {
      existing.descriptor = descriptor;
      existing.controller.label = descriptor.displayName;
      existing.controller.description = descriptor.description ?? 'R kernel';
      existing.controller.detail = descriptor.rPath;
      return existing;
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
      if (isActionDescriptor(selectedDescriptor)) {
        void this.handleActionSelection(event.notebook, selectedDescriptor);
        return;
      }
      void this.handleControllerSelection(event.notebook, selectedDescriptor, controller.id);
    });

    const entry = { controller, descriptor, selectionDisposable };
    this.controllers.set(controllerId, entry);
    return entry;
  }

  private async refreshKernelCatalog(): Promise<void> {
    const refreshSeq = ++this.kernelCatalogRefreshSeq;
    const configuredR = this.rBinPath();
    const additionalRPaths = this.additionalRPaths();
    const fallback = this.fallbackDescriptor(configuredR);
    if (this.kernelDescriptors.length === 0) {
      this.installKernelCatalog([fallback]);
    }

    const discovered = await discoverRKernelsAsync(configuredR, additionalRPaths);
    if (refreshSeq !== this.kernelCatalogRefreshSeq) return;
    this.installKernelCatalog(discovered.length > 0 ? discovered : [fallback]);
  }

  private fallbackDescriptor(configuredR: string): RKernelDescriptor {
    return {
      id: `executable:${configuredR}`,
      label: configuredR,
      description: 'R executable',
      detail: configuredR,
      rPath: configuredR,
      displayName: configuredR,
      source: 'executable',
    };
  }

  private installKernelCatalog(descriptors: RKernelDescriptor[]): void {
    this.kernelDescriptors = descriptors;
    const nextIds = new Set<string>();

    for (const descriptor of this.kernelDescriptors) {
      const controllerId = this.ensureControllerForDescriptor(descriptor).controller.id;
      nextIds.add(controllerId);
    }

    nextIds.add(this.ensureManualPathController().controller.id);

    for (const [controllerId, entry] of this.controllers) {
      if (nextIds.has(controllerId)) continue;
      entry.selectionDisposable.dispose();
      entry.controller.dispose();
      this.controllers.delete(controllerId);
    }

    this.syncAllNotebookAffinities();
  }

  private ensureManualPathController(): RControllerEntry {
    const descriptor: RKernelActionDescriptor = {
      id: 'action:manual',
      label: 'Rscript executable path...',
      description: 'Manual path',
      detail: 'Enter the full path to an Rscript executable',
      action: 'manual',
    };
    const controllerId = this.controllerIdForDescriptor(descriptor);
    const existing = this.controllers.get(controllerId);
    if (existing) {
      existing.descriptor = descriptor;
      existing.controller.label = descriptor.label;
      existing.controller.description = descriptor.description;
      existing.controller.detail = descriptor.detail;
      return existing;
    }

    const controller = vscode.notebooks.createNotebookController(
      controllerId,
      NOTEBOOK_TYPE,
      descriptor.label,
    );
    controller.supportedLanguages = ['r', 'R'];
    controller.supportsExecutionOrder = false;
    controller.description = descriptor.description;
    controller.detail = descriptor.detail;
    controller.executeHandler = async () => undefined;

    const selectionDisposable = controller.onDidChangeSelectedNotebooks((event) => {
      if (!event.selected) return;
      const selectedDescriptor = this.controllers.get(controller.id)?.descriptor ?? descriptor;
      if (!isActionDescriptor(selectedDescriptor)) return;
      void this.handleActionSelection(event.notebook, selectedDescriptor);
    });

    const entry = { controller, descriptor, selectionDisposable };
    this.controllers.set(controllerId, entry);
    return entry;
  }

  private controllerIdForDescriptor(descriptor: RControllerDescriptor): string {
    if (isActionDescriptor(descriptor)) {
      return 'r-notebook-r-kernel-action:manual-path';
    }
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
        !isActionDescriptor(entry.descriptor) && entry.descriptor.id === preferred.id
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
    const matched = this.availableDescriptors().find((descriptor) =>
      this.controllerIdForDescriptor(descriptor) === remembered?.controllerId
      || descriptor.kernelspecName === remembered?.kernelspecName
      || descriptor.displayName === remembered?.displayName
      || descriptor.rPath === remembered?.rPath,
    );
    if (matched) return matched;
    if (remembered?.rPath) {
      const rememberedDescriptor = this.manualDescriptor(remembered.rPath, {
        displayName: remembered.displayName,
        kernelspecName: remembered.kernelspecName,
      });
      const ensuredDescriptor = this.ensureControllerForDescriptor(rememberedDescriptor).descriptor;
      if (!isActionDescriptor(ensuredDescriptor)) return ensuredDescriptor;
    }
    return this.resolveDefaultDescriptor();
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

  private async handleActionSelection(
    notebook: vscode.NotebookDocument,
    descriptor: RKernelActionDescriptor,
  ): Promise<void> {
    if (notebook.notebookType !== NOTEBOOK_TYPE) return;

    const currentDescriptor = this.resolveDescriptorForNotebook(notebook);
    const rPath = await vscode.window.showInputBox({
      title: 'R Kernel Path',
      prompt: 'Full path to Rscript (for example /usr/local/bin/Rscript)',
      value: currentDescriptor.rPath,
    });
    const trimmedRPath = rPath?.trim();

    if (!trimmedRPath) {
      await this.selectDescriptorForNotebook(notebook, currentDescriptor);
      return;
    }

    const pickedDescriptor = this.manualDescriptor(trimmedRPath);
    await this.selectDescriptorForNotebook(notebook, pickedDescriptor);
    await rememberRExecutablePath(pickedDescriptor.rPath);
    vscode.window.showInformationMessage(`R kernel path set to: ${pickedDescriptor.rPath}`);
  }

  private async selectDescriptorForNotebook(
    notebook: vscode.NotebookDocument,
    descriptor: RKernelDescriptor,
  ): Promise<void> {
    const ensuredDescriptor = this.ensureControllerForDescriptor(descriptor).descriptor;
    if (isActionDescriptor(ensuredDescriptor)) return;
    const controllerId = this.controllerIdForDescriptor(ensuredDescriptor);
    rememberRNotebookKernel(notebook.uri.toString(), {
      controllerId,
      displayName: ensuredDescriptor.displayName,
      kernelspecName: ensuredDescriptor.kernelspecName,
      rPath: ensuredDescriptor.rPath,
    });

    const didSelect = await this.selectNotebookController(notebook, controllerId);
    if (!didSelect) {
      await this.handleControllerSelection(notebook, ensuredDescriptor, controllerId);
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
    const selectedDescriptor = this.controllerEntryForController(ctrl)?.descriptor;
    const descriptor = selectedDescriptor && !isActionDescriptor(selectedDescriptor)
      ? selectedDescriptor
      : this.resolveDescriptorForNotebook(notebook);
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
        lastRenderStartedAt: 0,
        completed: false,
      };
      this.liveExecutions.set(liveKey, liveState);
      this.liveCellDocuments.add(cellDocUri);
      const cancelDisp = execution.token.onCancellationRequested(() => session.interrupt());

      try {
        await this.clearExecutionOutputForNewRun(liveState);
        await this.replaceExecutionOutput(liveState, liveState.result, chunkId, { running: true });
        if (startError) throw startError;
        const result = mergeConsoleResult(
          await session.exec(chunkId, code, mergeRFigureOptions(opts)),
          liveState.result,
        );
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
    const hasLiveState = this.hasLiveExecutionState(docUri);
    const interrupted = (session?.isBusy() ?? false) || hasLiveState;
    this.bumpQueueEpoch(docUri);
    session?.interrupt();
    if (!session?.isBusy() && hasLiveState) {
      this.clearLiveExecutionState(docUri);
    }
    return interrupted;
  }

  public isInterruptInProgress(docUri: string): boolean {
    return getSession(docUri)?.isInterruptInProgress() ?? false;
  }

  public interruptRecoveryState(docUri: string): 'current' | 'stale' | 'none' {
    return getSession(docUri)?.recoveryCheckpointState() ?? 'none';
  }

  public async forceInterruptNotebook(docUri: string): Promise<boolean> {
    const session = getSession(docUri);
    const hasLiveState = this.hasLiveExecutionState(docUri);
    const interrupted = (session?.isBusy() ?? false) || hasLiveState;
    if (!interrupted) return false;

    this.bumpQueueEpoch(docUri);
    await session?.forceInterruptAndRecover();
    this.clearLiveExecutionState(docUri);
    return true;
  }

  public hasPendingExecution(docUri: string): boolean {
    if (this.notebookQueues.has(docUri)) return true;
    const session = getSession(docUri);
    if (session?.isBusy()) return true;
    const prefix = `${docUri}::`;
    for (const liveKey of this.liveExecutions.keys()) {
      if (liveKey.startsWith(prefix)) return true;
    }
    return false;
  }

  public async restartNotebook(notebook: vscode.NotebookDocument): Promise<boolean> {
    if (notebook.notebookType !== NOTEBOOK_TYPE) return false;
    const docUri = notebook.uri.toString();
    const descriptor = this.resolveDescriptorForNotebook(notebook);
    const session = getOrCreateSession(docUri, descriptor.rPath, this.execTimeoutMs());
    this.bumpQueueEpoch(docUri);
    this.clearLiveExecutionState(docUri);
    session.setExecutablePath(descriptor.rPath);
    this.bindSession(docUri, session);
    await session.restart();
    this.varNotifier?.notifyChanged(notebook);
    return true;
  }

  private clearLiveExecutionState(docUri: string): void {
    for (const [liveKey, state] of this.liveExecutions) {
      if (!liveKey.startsWith(`${docUri}::`)) continue;
      void this.safeEndExecution(state.execution);
      this.liveExecutions.delete(liveKey);
      this.liveCellDocuments.delete(state.cellDocUri);
      this.deletedLiveCellDocuments.delete(state.cellDocUri);
    }
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
      const now = Date.now();
      if (this.liveExecutions.size === 0) {
        clearInterval(this.renderInterval!);
        this.renderInterval = null;
        return;
      }
      for (const state of this.liveExecutions.values()) {
        if (!state.dirty || state.completed || state.rendering) continue;
        if (this.deletedLiveCellDocuments.has(state.cellDocUri)) continue;
        if (!this.shouldRenderLiveState(state, now)) continue;
        state.dirty = false;
        state.rendering = true;
        state.lastRenderStartedAt = now;
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
    }, LIVE_RENDER_INTERVAL_MS);
  }

  private shouldRenderLiveState(state: LiveExecutionState, now: number): boolean {
    if (!state.output) return true;
    const consoleLength = state.result.console?.length ?? 0;
    const minInterval =
      consoleLength >= HUGE_CONSOLE_CHARS
        ? HUGE_CONSOLE_RENDER_INTERVAL_MS
        : consoleLength >= LARGE_CONSOLE_CHARS
          ? LARGE_CONSOLE_RENDER_INTERVAL_MS
          : LIVE_RENDER_INTERVAL_MS;
    return now - state.lastRenderStartedAt >= minInterval;
  }

  private async waitForPendingRender(state: LiveExecutionState): Promise<void> {
    if (!state.renderPromise) return;
    try {
      await this.withOutputTimeout(state.renderPromise, 'render');
    } catch {}
  }

  private liveKey(docUri: string, chunkId: string): string {
    return `${docUri}::${chunkId}`;
  }

  private hasLiveExecutionState(docUri: string): boolean {
    const prefix = `${docUri}::`;
    for (const liveKey of this.liveExecutions.keys()) {
      if (liveKey.startsWith(prefix)) return true;
    }
    return false;
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
      await this.withOutputTimeout(
        state.execution.clearOutput(),
        'clearOutput',
      ).catch(() => undefined);
      return;
    }
    if (!state.output) {
      await this.replaceExecutionOutputFromScratch(state, items, chunkId, options);
      return;
    }
    state.output.metadata = {
      chunkId,
      running: Boolean(options?.running),
    };
    try {
      await this.withOutputTimeout(
        state.execution.replaceOutputItems(items, state.output),
        'replaceOutputItems',
      );
    } catch (err) {
      if (isOutputUpdateTimeout(err)) return;
      await this.replaceExecutionOutputFromScratch(state, items, chunkId, options);
    }
  }

  private async replaceExecutionOutputFromScratch(
    state: LiveExecutionState,
    items: vscode.NotebookCellOutputItem[],
    chunkId: string,
    options?: { running?: boolean },
  ): Promise<void> {
    state.output = new vscode.NotebookCellOutput(items, {
      chunkId,
      running: Boolean(options?.running),
    });
    await this.withOutputTimeout(
      state.execution.replaceOutput([state.output]),
      'replaceOutput',
    ).catch((err) => {
      if (!isOutputUpdateTimeout(err)) {
        state.output = null;
      }
    });
  }

  private withOutputTimeout<T>(promise: Thenable<T>, label: string): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(
        () => reject(new OutputUpdateTimeoutError(label)),
        OUTPUT_UPDATE_TIMEOUT_MS,
      );
      Promise.resolve(promise).then(
        (value) => {
          clearTimeout(timer);
          resolve(value);
        },
        (error) => {
          clearTimeout(timer);
          reject(error);
        },
      );
    });
  }

  private async safeEndExecution(execution: vscode.NotebookCellExecution): Promise<void> {
    try {
      execution.end(false, Date.now());
    } catch {}
  }

  private async restoreNotebookOutputs(notebook: vscode.NotebookDocument): Promise<void> {
    const docUri = notebook.uri.toString();
    if (this.restoreInFlight.has(docUri)) return;
    const cached = this.outputStore.getForNotebook(notebook);
    if (cached.size === 0) return;

    // Avoid restoring outputs while any notebook is actively running so we do
    // not race live execution updates with a restore pass.
    if (this.liveExecutions.size > 0) return;

    this.restoreInFlight.add(docUri);
    try {
      const edits: vscode.NotebookEdit[] = [];
      for (const cell of notebook.getCells()) {
        if (cell.kind !== vscode.NotebookCellKind.Code) continue;
        if (cell.outputs.length > 0) continue;
        if (this.liveCellDocuments.has(cell.document.uri.toString())) continue;
        const result = cached.get(cell.index);
        if (!result) continue;
        edits.push(
          buildNotebookCellOutputsEdit(
            cell,
            notebookOutputFromExecResult(result, `nb-${cell.index}`),
          ),
        );
      }
      if (edits.length === 0) return;

      const edit = new vscode.WorkspaceEdit();
      edit.set(notebook.uri, edits);
      await vscode.workspace.applyEdit(edit);
    } finally {
      this.restoreInFlight.delete(docUri);
    }
  }

  private shouldRestoreNotebookOutputs(event: vscode.NotebookDocumentChangeEvent): boolean {
    if (event.contentChanges.length > 0) return true;
    return event.cellChanges.some((change) => Boolean(change.document || change.metadata));
  }

  private async prepareOpenedNotebook(notebook: vscode.NotebookDocument): Promise<void> {
    const docUri = notebook.uri.toString();
    try {
      const textDocument = await vscode.workspace.openTextDocument(notebook.uri);
      const hasPersistedState = hasPersistedRmdState(textDocument.getText());
      const hasVisibleOutputs = notebook.getCells().some((cell) =>
        cell.kind === vscode.NotebookCellKind.Code && cell.outputs.length > 0,
      );
      if (!hasPersistedState && !hasVisibleOutputs) {
        this.outputStore.clear(docUri);
      }
    } catch {
      // Best effort only. If the backing text document is unavailable, fall back
      // to the existing in-memory output state.
    }
    this.outputStore.syncNotebook(notebook);
    this.syncNotebookAffinities(notebook);
    await this.restoreNotebookOutputs(notebook);
  }

  private async clearExecutionOutputForNewRun(state: LiveExecutionState): Promise<void> {
    state.output = null;
    await this.withOutputTimeout(
      state.execution.clearOutput(),
      'clearOutput',
    ).catch(() => undefined);
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

function buildNotebookCellOutputsEdit(
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

function figFields(): FigField[] {
  const defaults = getRFigureDefaults();
  return [
    { key: 'fig_width',  label: 'fig.width',  unit: 'in',  def: defaults.fig_width,  min: 1,  max: 20  },
    { key: 'fig_height', label: 'fig.height', unit: 'in',  def: defaults.fig_height, min: 1,  max: 20  },
    { key: 'dpi',        label: 'dpi',        unit: 'dpi', def: defaults.dpi,        min: 72, max: 600 },
  ];
}

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
    const merged = mergeRFigureOptions(opts);
    const hasCustom = opts['fig_width'] !== undefined
      || opts['fig_height'] !== undefined
      || opts['dpi'] !== undefined;
    const text = `⚙ ${merged.fig_width}×${merged.fig_height}in · ${merged.dpi}dpi${hasCustom ? '' : ' default'}`;

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
      ...figFields().map(f => {
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
    vscode.workspace.onDidChangeConfiguration((event) => {
      if (
        affectsRConfig(event, 'defaultFigWidth')
        || affectsRConfig(event, 'defaultFigHeight')
        || affectsRConfig(event, 'defaultDpi')
      ) {
        provider.refresh();
      }
    }),
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
