// =============================================================================
// pyNotebookController.ts — Notebook controller that executes Python cells
// via the R Notebook Python kernel subprocess.
//
// Per-cell figure options (fig_width / fig_height / dpi) mirror the R
// controller's implementation via PyCellFigureStatusBar.
// =============================================================================

import * as vscode from 'vscode';
import { getOrCreatePySession, disposePySession, getPySession } from './pySessionManager';
import {
  COMMAND_IDS,
  EXTENSION_BRAND,
  PY_NOTEBOOK_TYPE as PY_NOTEBOOK_TYPE_ID,
  affectsPythonConfig,
  getNotebookKernelMetadata,
  getPythonConfigValue,
  updatePythonConfigValue,
} from './extensionIds';
import { ExecResult, StreamMessage } from './kernelProtocol';
import { notebookOutputItemsFromExecResult } from './rmdOutputStore';
import {
  discoverPythonKernelsAsync,
  fallbackKernelName,
  PythonKernelDescriptor,
} from './kernelDiscovery';

export const PY_NOTEBOOK_TYPE = PY_NOTEBOOK_TYPE_ID;

// ---------------------------------------------------------------------------
// Minimal interface so controllers can notify the variable provider without
// importing variableProvider.ts (avoids circular deps).
export interface VariableNotifier {
  notifyChanged(notebook: vscode.NotebookDocument): void;
}

type LiveExecutionState = {
  chunkId: string;
  execution: vscode.NotebookCellExecution;
  output: vscode.NotebookCellOutput | null;
  result: ExecResult;
  dirty: boolean;      // has unrendered stream output
  rendering: boolean;  // replaceOutput in-flight
  renderPromise: Promise<void> | null;
  completed: boolean;
};

type NotebookPythonMetadata = {
  kernelspec?: { language?: string; display_name?: string; name?: string };
  language_info?: { name?: string };
  rNotebook?: {
    pythonPath?: string;
    controllerId?: string;
    kernelspecName?: string;
    displayName?: string;
  };
  rnotebook?: {
    pythonPath?: string;
    controllerId?: string;
    kernelspecName?: string;
    displayName?: string;
  };
  [key: string]: unknown;
};

type RememberedPyNotebookKernel = {
  pythonPath: string;
  controllerId: string;
  kernelspecName?: string;
  displayName: string;
};

type PyKernelActionDescriptor = {
  id: 'action:manual';
  label: string;
  description?: string;
  detail?: string;
  action: 'manual';
};

type PyControllerDescriptor = PythonKernelDescriptor | PyKernelActionDescriptor;

type PyControllerEntry = {
  controller: vscode.NotebookController;
  descriptor: PyControllerDescriptor;
  selectionDisposable: vscode.Disposable;
};

function isActionDescriptor(descriptor: PyControllerDescriptor): descriptor is PyKernelActionDescriptor {
  return 'action' in descriptor;
}

const OUTPUT_UPDATE_TIMEOUT_MS = 2_000;
const PY_NOTEBOOK_SELECTIONS_STATE_KEY = 'pythonNotebookSelections';

export class PyNotebookController implements vscode.Disposable {
  private readonly controllers = new Map<string, PyControllerEntry>();
  private readonly liveExecutions = new Map<string, LiveExecutionState>();
  private readonly notebookQueues = new Map<string, Promise<void>>();
  private readonly queueEpochs = new Map<string, number>();
  private readonly rememberedNotebookSelections = new Map<string, RememberedPyNotebookKernel>();
  private readonly boundSessions = new WeakSet<object>();
  private readonly disposables: vscode.Disposable[] = [];
  private readonly extensionId: string;
  private readonly workspaceState: vscode.Memento;
  private kernelDescriptors: PythonKernelDescriptor[] = [];
  private executionOrder = 0;
  private renderInterval: ReturnType<typeof setInterval> | null = null;
  private varNotifier?: VariableNotifier;
  private kernelCatalogRefreshSeq = 0;

  constructor(ctx: vscode.ExtensionContext, varNotifier?: VariableNotifier) {
    this.extensionId = ctx.extension.id;
    this.workspaceState = ctx.workspaceState;
    this.varNotifier = varNotifier;
    this.restoreRememberedNotebookSelections();
    void this.refreshKernelCatalog();
    this.disposables.push(
      vscode.workspace.onDidOpenNotebookDocument((notebook) => {
        this.syncNotebookAffinities(notebook);
      }),
      vscode.workspace.onDidChangeConfiguration((event) => {
        if (affectsPythonConfig(event, 'pythonPath')) {
          void this.refreshKernelCatalog();
        }
      }),
    );
  }

  public async forgetNotebookSelection(docUri: string): Promise<void> {
    if (!this.rememberedNotebookSelections.delete(docUri)) return;
    await this.persistRememberedNotebookSelections();
  }

  public async showKernelPicker(
    notebook?: vscode.NotebookDocument,
    options?: {
      includeManual?: boolean;
      title?: string;
    },
  ): Promise<void> {
    const activeNotebook = vscode.window.activeNotebookEditor?.notebook;
    const targetNotebook = notebook?.notebookType === PY_NOTEBOOK_TYPE
      ? notebook
      : activeNotebook?.notebookType === PY_NOTEBOOK_TYPE
        ? activeNotebook
        : undefined;
    const includeManual = options?.includeManual ?? true;

    const descriptors = this.availableDescriptors();
    const currentDescriptor = targetNotebook
      ? this.resolveDescriptorForNotebook(targetNotebook)
      : this.resolveDefaultDescriptor();

    type PickItem = vscode.QuickPickItem & {
      action?: 'browse' | 'manual';
      descriptor?: PythonKernelDescriptor;
    };

    const items: PickItem[] = [
      ...descriptors.map((descriptor) => ({
        label: descriptor.displayName,
        description: descriptor.id === currentDescriptor?.id
          ? `${descriptor.description ?? 'Python environment'} · current`
          : descriptor.description,
        detail: descriptor.pythonPath,
        descriptor,
      })),
    ];
    if (includeManual) {
      items.push(
        {
          label: '$(folder-opened) Browse…',
          description: 'Pick a Python executable from the file system',
          action: 'browse',
        },
        {
          label: '$(pencil) Enter manually…',
          description: 'Type the full path to a Python executable',
          action: 'manual',
        },
      );
    }

    const pick = await vscode.window.showQuickPick(items, {
      title: options?.title ?? 'Select Python Kernel',
      placeHolder: currentDescriptor?.pythonPath ?? this.pyBinPath(),
      matchOnDescription: true,
      matchOnDetail: true,
    });
    if (!pick) return;

    let descriptor = pick.descriptor;
    if (pick.action === 'browse') {
      const uris = await vscode.window.showOpenDialog({
        title: 'Select Python executable',
        filters: { Executable: ['*'] },
      });
      const pythonPath = uris?.[0]?.fsPath;
      if (!pythonPath) return;
      descriptor = this.manualDescriptor(pythonPath);
    } else if (pick.action === 'manual') {
      const pythonPath = await vscode.window.showInputBox({
        title: 'Python Kernel Path',
        prompt: 'Full path to python3 (for example /usr/local/bin/python3)',
        value: currentDescriptor?.pythonPath ?? this.pyBinPath(),
      });
      if (!pythonPath) return;
      descriptor = this.manualDescriptor(pythonPath);
    }
    if (!descriptor) return;

    if (targetNotebook) {
      await this.selectDescriptorForNotebook(targetNotebook, descriptor);
      vscode.window.showInformationMessage(`Python kernel set to ${descriptor.displayName}.`);
      return;
    }

    await updatePythonConfigValue('pythonPath', descriptor.pythonPath, vscode.ConfigurationTarget.Global);
    void this.refreshKernelCatalog();
    vscode.window.showInformationMessage(`Python path set to: ${descriptor.pythonPath}`);
  }

  private async refreshKernelCatalog(): Promise<void> {
    const refreshSeq = ++this.kernelCatalogRefreshSeq;
    const configuredPython = this.pyBinPath();
    if (this.kernelDescriptors.length === 0) {
      this.installKernelCatalog([this.manualDescriptor(configuredPython)]);
    }

    const discovered = await discoverPythonKernelsAsync(configuredPython);
    if (refreshSeq !== this.kernelCatalogRefreshSeq) return;
    this.installKernelCatalog(discovered.length > 0
      ? discovered
      : [this.manualDescriptor(configuredPython)]);
  }

  private installKernelCatalog(descriptors: PythonKernelDescriptor[]): void {
    this.kernelDescriptors = descriptors;
    const nextIds = new Set<string>();

    for (const descriptor of this.kernelDescriptors) {
      const controllerId = this.ensureControllerForDescriptor(descriptor).controller.id;
      nextIds.add(controllerId);
    }

    for (const notebook of vscode.workspace.notebookDocuments) {
      if (notebook.notebookType !== PY_NOTEBOOK_TYPE) continue;
      const descriptor = this.descriptorFromNotebookMetadata(notebook);
      if (!descriptor) continue;
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

  private pyBinPath(): string {
    return getPythonConfigValue('pythonPath', 'python3');
  }

  private restoreRememberedNotebookSelections(): void {
    const stored = this.workspaceState.get<Record<string, unknown>>(
      PY_NOTEBOOK_SELECTIONS_STATE_KEY,
      {},
    );
    for (const [docUri, value] of Object.entries(stored)) {
      const selection = parseRememberedSelection(value);
      if (!selection) continue;
      this.rememberedNotebookSelections.set(docUri, selection);
    }
  }

  private async persistRememberedNotebookSelections(): Promise<void> {
    await this.workspaceState.update(
      PY_NOTEBOOK_SELECTIONS_STATE_KEY,
      Object.fromEntries(this.rememberedNotebookSelections.entries()),
    );
  }

  private async rememberNotebookSelection(
    docUri: string,
    descriptor: PythonKernelDescriptor,
    controllerId = this.controllerIdForDescriptor(descriptor),
  ): Promise<void> {
    this.rememberedNotebookSelections.set(docUri, {
      pythonPath: descriptor.pythonPath,
      controllerId,
      kernelspecName: descriptor.kernelspecName,
      displayName: descriptor.displayName,
    });
    await this.persistRememberedNotebookSelections();
  }

  private execTimeoutMs(): number {
    return getPythonConfigValue('execTimeoutMs', 3_000_000) ?? 3_000_000;
  }

  private async execute(
    controller: vscode.NotebookController,
    cells: vscode.NotebookCell[],
    notebook: vscode.NotebookDocument,
  ): Promise<void> {
    const selectedDescriptor = this.controllerEntryForController(controller)?.descriptor;
    const descriptor = selectedDescriptor && !isActionDescriptor(selectedDescriptor)
      ? selectedDescriptor
      : this.resolveDescriptorForNotebook(notebook);
    const docUri = notebook.uri.toString();
    return this.enqueueNotebookExecution(docUri, (runEpoch) =>
      this.runQueuedExecution(docUri, runEpoch, descriptor, controller, cells, notebook),
    );
  }

  private async runQueuedExecution(
    docUri: string,
    runEpoch: number,
    descriptor: PythonKernelDescriptor,
    controller: vscode.NotebookController,
    cells: vscode.NotebookCell[],
    notebook: vscode.NotebookDocument,
  ): Promise<void> {
    const session = getOrCreatePySession(
      docUri,
      descriptor.pythonPath,
      descriptor.env,
      this.execTimeoutMs(),
    );
    this.bindSession(docUri, session);

    let startError: Error | undefined;
    try {
      await session.start();
    } catch (err: any) {
      startError = err instanceof Error ? err : new Error(String(err));
    }

    for (const cell of cells) {
      if (this.isQueueInterrupted(docUri, runEpoch)) break;

      const execution = controller.createNotebookCellExecution(cell);
      execution.executionOrder = ++this.executionOrder;
      execution.start(Date.now());

      const chunkId = `py-${cell.index}`;
      const code    = cell.document.getText();
      const opts    = (cell.metadata?.options ?? {}) as Record<string, unknown>;
      const liveKey = this.liveKey(docUri, chunkId);
      const liveState: LiveExecutionState = {
        chunkId,
        execution,
        output: null,
        result: emptyExecResult(chunkId, code),
        dirty: false,
        rendering: false,
        renderPromise: null,
        completed: false,
      };
      this.liveExecutions.set(liveKey, liveState);
      const cancelDisp = execution.token.onCancellationRequested(() => session.interrupt());

      try {
        await this.clearExecutionOutputForNewRun(liveState);
        await this.replaceExecutionOutput(liveState, liveState.result, chunkId, { running: true });
        if (startError) throw startError;
        const result = mergeConsoleResult(await session.exec(chunkId, code, {
          fig_width:  opts['fig_width']  as number | undefined,
          fig_height: opts['fig_height'] as number | undefined,
          dpi:        opts['dpi']        as number | undefined,
        }), liveState.result);
        liveState.completed = true;
        await this.waitForPendingRender(liveState);
        await this.replaceExecutionOutput(liveState, result, chunkId);
        execution.end(!result.error, Date.now());
      } catch (err: any) {
        liveState.completed = true;
        await this.waitForPendingRender(liveState);
        await this.replaceExecutionOutput(liveState, {
          ...liveState.result,
          error: err.message,
        }, chunkId);
        execution.end(false, Date.now());
      } finally {
        cancelDisp.dispose();
        this.liveExecutions.delete(liveKey);
      }

      if (this.isQueueInterrupted(docUri, runEpoch)) break;
    }

    // Notify variable provider so the Variables panel refreshes
    this.varNotifier?.notifyChanged(notebook);
  }

  private async clearExecutionOutputForNewRun(state: LiveExecutionState): Promise<void> {
    state.output = null;
    await this.withOutputTimeout(
      state.execution.clearOutput(),
      'clearOutput',
    ).catch(() => undefined);
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
    const session = getPySession(docUri);
    const interrupted = session?.isBusy() ?? false;
    this.bumpQueueEpoch(docUri);
    session?.interrupt();
    return interrupted;
  }

  public hasPendingExecution(docUri: string): boolean {
    if (this.notebookQueues.has(docUri)) return true;
    const session = getPySession(docUri);
    if (session?.isBusy()) return true;
    const prefix = `${docUri}::`;
    for (const liveKey of this.liveExecutions.keys()) {
      if (liveKey.startsWith(prefix)) return true;
    }
    return false;
  }

  public async restartNotebook(notebook: vscode.NotebookDocument): Promise<boolean> {
    if (notebook.notebookType !== PY_NOTEBOOK_TYPE) return false;
    const docUri = notebook.uri.toString();
    const descriptor = this.resolveDescriptorForNotebook(notebook);
    const session = getOrCreatePySession(
      docUri,
      descriptor.pythonPath,
      descriptor.env,
      this.execTimeoutMs(),
    );
    session.setKernel(descriptor.pythonPath, descriptor.env);
    this.bindSession(docUri, session);
    await session.restart();
    this.varNotifier?.notifyChanged(notebook);
    return true;
  }

  private bindSession(docUri: string, session: ReturnType<typeof getOrCreatePySession>): void {
    if (this.boundSessions.has(session as object)) return;
    this.boundSessions.add(session as object);
    session.on('stream', (msg: StreamMessage) => this.handleStream(docUri, msg));
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

  private handleStream(docUri: string, msg: StreamMessage): void {
    const state = this.liveExecutions.get(this.liveKey(docUri, msg.chunk_id));
    if (!state || state.completed || !msg.text) return;
    if (msg.stream === 'stderr') {
      state.result.stderr += msg.text;
    } else {
      state.result.stdout += msg.text;
    }
    const currentConsole = state.result.console ?? '';
    const needsSeparator =
      currentConsole.length > 0 &&
      !currentConsole.endsWith('\n') &&
      !msg.text.startsWith('\n');
    state.result.console = `${currentConsole}${needsSeparator ? '\n' : ''}${msg.text}`;
    if (state.result.console_segments && state.result.console_segments.length > 0) {
      const segments = state.result.console_segments.map((segment) => ({ ...segment }));
      const lastSegment = segments[segments.length - 1];
      const existingOutput = lastSegment.output ?? '';
      const separator = existingOutput.length > 0 ? '\n' : '';
      lastSegment.output = `${existingOutput}${separator}${msg.text}`;
      state.result.console_segments = segments;
    }
    state.dirty = true;
    this.startRenderLoop();
  }

  /** Shared 100 ms render loop — flushes dirty live states at most 10×/sec. */
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
        state.dirty = false;
        state.rendering = true;
        const renderPromise = this.replaceExecutionOutput(state, state.result, state.chunkId, { running: true });
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

  private availableDescriptors(): PythonKernelDescriptor[] {
    if (this.kernelDescriptors.length === 0) {
      this.installKernelCatalog([this.manualDescriptor(this.pyBinPath())]);
      void this.refreshKernelCatalog();
    }
    return this.kernelDescriptors;
  }

  private manualDescriptor(
    pythonPath: string,
    overrides?: {
      displayName?: string;
      kernelspecName?: string;
    },
  ): PythonKernelDescriptor {
    const displayName = overrides?.displayName ?? simplePythonLabel(pythonPath);
    return {
      id: `executable:${pythonPath}`,
      label: displayName,
      description: 'Python environment',
      detail: pythonPath,
      env: undefined,
      environmentType: undefined,
      pythonPath,
      displayName,
      kernelspecName: overrides?.kernelspecName ?? fallbackKernelName(pythonPath),
      source: 'executable',
    };
  }

  private ensureControllerForDescriptor(descriptor: PythonKernelDescriptor): PyControllerEntry {
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
      existing.controller.description = descriptor.description ?? `${EXTENSION_BRAND} Python Kernel`;
      existing.controller.detail = descriptor.pythonPath;
      return existing;
    }

    const controller = vscode.notebooks.createNotebookController(
      controllerId,
      PY_NOTEBOOK_TYPE,
      descriptor.displayName,
    );
    controller.supportedLanguages = ['python'];
    controller.supportsExecutionOrder = true;
    controller.description = descriptor.description ?? `${EXTENSION_BRAND} Python Kernel`;
    controller.detail = descriptor.pythonPath;
    controller.executeHandler = (cells, notebook, activeController) =>
      this.execute(activeController, cells, notebook);
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

  private ensureManualPathController(): PyControllerEntry {
    const descriptor: PyKernelActionDescriptor = {
      id: 'action:manual',
      label: 'Python executable path...',
      description: 'Manual path',
      detail: 'Enter the full path to a python3 executable',
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
      PY_NOTEBOOK_TYPE,
      descriptor.label,
    );
    controller.supportedLanguages = ['python'];
    controller.supportsExecutionOrder = true;
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

  private resolveDefaultDescriptor(): PythonKernelDescriptor | undefined {
    const configuredPython = this.pyBinPath();
    return this.availableDescriptors().find((descriptor) => descriptor.pythonPath === configuredPython)
      ?? this.availableDescriptors()[0];
  }

  private controllerIdForDescriptor(descriptor: PyControllerDescriptor): string {
    if (isActionDescriptor(descriptor)) {
      return 'r-notebook-py-kernel-action:manual-path';
    }
    return `r-notebook-py-kernel:${descriptor.id}`;
  }

  private controllerEntryForController(
    controller: vscode.NotebookController,
  ): PyControllerEntry | undefined {
    return this.controllers.get(controller.id);
  }

  private syncAllNotebookAffinities(): void {
    for (const notebook of vscode.workspace.notebookDocuments) {
      this.syncNotebookAffinities(notebook);
    }
  }

  private syncNotebookAffinities(notebook: vscode.NotebookDocument): void {
    if (notebook.notebookType !== PY_NOTEBOOK_TYPE) return;
    const preferred = this.resolveDescriptorForNotebook(notebook);
    for (const entry of this.controllers.values()) {
      try {
        entry.controller.updateNotebookAffinity(
          notebook,
          entry.descriptor.id === preferred.id
            ? vscode.NotebookControllerAffinity.Preferred
            : vscode.NotebookControllerAffinity.Default,
        );
      } catch {
        // Cursor can expose affinity APIs while still rejecting extension
        // access on stable builds. Affinity is only a picker hint, so skip it.
      }
    }
  }

  private resolveDescriptorForNotebook(notebook: vscode.NotebookDocument): PythonKernelDescriptor {
    const meta = (notebook.metadata ?? {}) as NotebookPythonMetadata;
    const remembered = this.rememberedNotebookSelections.get(notebook.uri.toString());
    const kernelMeta = getNotebookKernelMetadata(meta);
    const descriptors = this.availableDescriptors();

    const matched = descriptors.find((descriptor) =>
      this.controllerIdForDescriptor(descriptor) === remembered?.controllerId
      || descriptor.kernelspecName === remembered?.kernelspecName
      || descriptor.pythonPath === remembered?.pythonPath
      || descriptor.displayName === remembered?.displayName
      || this.controllerIdForDescriptor(descriptor) === kernelMeta?.controllerId
      || descriptor.kernelspecName === kernelMeta?.kernelspecName
      || descriptor.kernelspecName === meta.kernelspec?.name
      || descriptor.pythonPath === kernelMeta?.pythonPath
      || descriptor.displayName === kernelMeta?.displayName
      || descriptor.displayName === meta.kernelspec?.display_name,
    );
    if (matched) return matched;

    if (remembered?.pythonPath) {
      const rememberedDescriptor = this.manualDescriptor(remembered.pythonPath, {
        displayName: remembered.displayName,
        kernelspecName: remembered.kernelspecName,
      });
      const ensuredDescriptor = this.ensureControllerForDescriptor(rememberedDescriptor).descriptor;
      if (!isActionDescriptor(ensuredDescriptor)) return ensuredDescriptor;
    }

    const persistedDescriptor = this.descriptorFromNotebookMetadata(notebook);
    if (persistedDescriptor) {
      const ensuredDescriptor = this.ensureControllerForDescriptor(persistedDescriptor).descriptor;
      if (!isActionDescriptor(ensuredDescriptor)) return ensuredDescriptor;
    }

    return this.resolveDefaultDescriptor()
      ?? this.manualDescriptor(this.pyBinPath());
  }

  private async handleControllerSelection(
    notebook: vscode.NotebookDocument,
    descriptor: PythonKernelDescriptor,
    controllerId: string,
  ): Promise<void> {
    if (notebook.notebookType !== PY_NOTEBOOK_TYPE) return;
    await this.rememberNotebookSelection(notebook.uri.toString(), descriptor, controllerId);
    const session = getPySession(notebook.uri.toString());
    if (session) {
      session.setKernel(descriptor.pythonPath, descriptor.env);
      await session.restart().catch(() => undefined);
      this.varNotifier?.notifyChanged(notebook);
    }
    this.syncNotebookAffinities(notebook);
  }

  private async handleActionSelection(
    notebook: vscode.NotebookDocument,
    descriptor: PyKernelActionDescriptor,
  ): Promise<void> {
    if (notebook.notebookType !== PY_NOTEBOOK_TYPE) return;

    const currentDescriptor = this.resolveDescriptorForNotebook(notebook);
    const pythonPath = await vscode.window.showInputBox({
      title: 'Python Kernel Path',
      prompt: 'Full path to python3 (for example /usr/local/bin/python3)',
      value: currentDescriptor.pythonPath,
    });

    if (!pythonPath) {
      await this.selectDescriptorForNotebook(notebook, currentDescriptor);
      return;
    }

    const pickedDescriptor = this.manualDescriptor(pythonPath);
    await this.selectDescriptorForNotebook(notebook, pickedDescriptor);
    vscode.window.showInformationMessage(`Python kernel set to ${pickedDescriptor.displayName}.`);
  }

  private async selectDescriptorForNotebook(
    notebook: vscode.NotebookDocument,
    descriptor: PythonKernelDescriptor,
  ): Promise<void> {
    const ensuredDescriptor = this.ensureControllerForDescriptor(descriptor).descriptor;
    if (isActionDescriptor(ensuredDescriptor)) return;
    const controllerId = this.controllerIdForDescriptor(ensuredDescriptor);
    await this.rememberNotebookSelection(notebook.uri.toString(), ensuredDescriptor, controllerId);

    const didSelect = await this.selectNotebookController(notebook, controllerId);
    if (!didSelect) {
      await this.handleControllerSelection(notebook, ensuredDescriptor, controllerId);
      return;
    }

    this.syncNotebookAffinities(notebook);
  }

  private descriptorFromNotebookMetadata(
    notebook: vscode.NotebookDocument,
  ): PythonKernelDescriptor | undefined {
    const meta = (notebook.metadata ?? {}) as NotebookPythonMetadata;
    const kernelMeta = getNotebookKernelMetadata(meta);
    if (!kernelMeta?.pythonPath) return undefined;

    return this.manualDescriptor(kernelMeta.pythonPath, {
      displayName: kernelMeta.displayName ?? meta.kernelspec?.display_name,
      kernelspecName: kernelMeta.kernelspecName ?? meta.kernelspec?.name,
    });
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

  private liveKey(docUri: string, chunkId: string): string {
    return `${docUri}::${chunkId}`;
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
    } catch {
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
    ).catch(() => {
      state.output = null;
    });
  }

  private withOutputTimeout<T>(promise: Thenable<T>, label: string): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(
        () => reject(new Error(`Notebook ${label} timed out`)),
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

/** Tear down the Python session when the notebook editor closes. */
export { disposePySession };

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
    output_order: [],
    error: null,
  };
}

function simplePythonLabel(pythonPath: string): string {
  const parts = pythonPath.split(/[\\/]/).filter(Boolean);
  const base = parts[parts.length - 1] ?? pythonPath;
  return base === pythonPath ? `Python (${pythonPath})` : `Python (${base})`;
}

function mergeConsoleResult(result: ExecResult, liveResult: ExecResult): ExecResult {
  const liveConsole = liveResult.console ?? '';
  const sourceCode = result.source_code ?? liveResult.source_code;
  const consoleSegments = result.console_segments ?? liveResult.console_segments;
  const outputOrder =
    result.output_order && result.output_order.length > 0
      ? result.output_order
      : buildDefaultOutputOrder(result);
  if (result.console || !liveConsole) {
    return {
      ...result,
      source_code: sourceCode,
      console_segments: consoleSegments,
      output_order: outputOrder,
    };
  }
  return {
    ...result,
    console: liveConsole,
    source_code: sourceCode,
    console_segments: consoleSegments,
    output_order: outputOrder,
  };
}

function buildDefaultOutputOrder(result: ExecResult): ExecResult['output_order'] {
  const order: NonNullable<ExecResult['output_order']> = [];
  (result.dataframes ?? []).forEach((df, index) => {
    order.push({ type: 'df', index, name: df?.name });
  });
  (result.plots ?? []).forEach((_, index) => {
    order.push({ type: 'plot', index });
  });
  return order;
}

function parseRememberedSelection(value: unknown): RememberedPyNotebookKernel | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const source = value as Record<string, unknown>;
  if (typeof source.pythonPath !== 'string' || source.pythonPath.length === 0) return undefined;
  if (typeof source.controllerId !== 'string' || source.controllerId.length === 0) return undefined;
  return {
    pythonPath: source.pythonPath,
    controllerId: source.controllerId,
    kernelspecName: typeof source.kernelspecName === 'string' ? source.kernelspecName : undefined,
    displayName: typeof source.displayName === 'string' && source.displayName.length > 0
      ? source.displayName
      : source.pythonPath,
  };
}


// =============================================================================
// Per-cell figure options — status bar item + quick-pick editor
// (mirrors rNotebookController.ts / RCellFigureStatusBar exactly)
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

export class PyCellFigureStatusBar implements vscode.NotebookCellStatusBarItemProvider {

  private readonly _emitter = new vscode.EventEmitter<void>();
  readonly onDidChangeCellStatusBarItems = this._emitter.event;

  refresh(): void { this._emitter.fire(); }
  dispose():  void { this._emitter.dispose(); }

  provideCellStatusBarItems(
    cell: vscode.NotebookCell,
  ): vscode.NotebookCellStatusBarItem[] {
    if (cell.kind !== vscode.NotebookCellKind.Code) return [];
    if (cell.document.languageId !== 'python') return [];

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
      command:   COMMAND_IDS.pySetCellFigureOptions,
      title:     'Edit Figure Options',
      arguments: [cell],
    };
    item.tooltip = 'fig.width / fig.height / dpi — click to edit';
    return [item];
  }
}

async function editPyCellFigureOptions(
  cell: vscode.NotebookCell,
  statusBar: PyCellFigureStatusBar,
): Promise<void> {
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
      title:              'Cell Figure Options',
      placeHolder:        'Select a setting to edit  ·  Escape to cancel without saving',
      matchOnDescription: false,
    });

    if (!picked || picked.done) break;
    if (!picked.field) break;

    const f   = picked.field;
    const cur = opts[f.key] as number | undefined;

    const raw = await vscode.window.showInputBox({
      title:  `Set ${f.label}`,
      prompt: `${f.min}–${f.max} ${f.unit}. Leave empty to reset to default (${f.def}).`,
      value:  cur !== undefined ? String(cur) : '',
      validateInput: v => {
        if (!v.trim()) return null;
        const n = parseFloat(v);
        return isNaN(n) || n < f.min || n > f.max
          ? `Enter a number between ${f.min} and ${f.max}`
          : null;
      },
    });

    if (raw === undefined) continue;

    if (!raw.trim()) {
      delete opts[f.key];
    } else {
      opts[f.key] = parseFloat(raw);
    }
  }

  const edit = new vscode.WorkspaceEdit();
  edit.set(cell.notebook.uri, [
    vscode.NotebookEdit.updateCellMetadata(cell.index, {
      ...cell.metadata,
      options: opts,
    }),
  ]);
  await vscode.workspace.applyEdit(edit);
  statusBar.refresh();
}

export function registerPyFigureOptions(ctx: vscode.ExtensionContext): void {
  const provider = new PyCellFigureStatusBar();

  ctx.subscriptions.push(
    vscode.notebooks.registerNotebookCellStatusBarItemProvider(
      PY_NOTEBOOK_TYPE,
      provider,
    ),
    provider,
  );

  ctx.subscriptions.push(
    vscode.commands.registerCommand(
      COMMAND_IDS.pySetCellFigureOptions,
      async (cell?: vscode.NotebookCell) => {
        if (!cell) {
          const editor = vscode.window.activeNotebookEditor;
          if (!editor) {
            vscode.window.showWarningMessage('No active Python notebook cell.');
            return;
          }
          const sel = editor.selections[0];
          cell = editor.notebook.cellAt(sel?.start ?? 0);
        }
        await editPyCellFigureOptions(cell, provider);
      },
    ),
  );
}
