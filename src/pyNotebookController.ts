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
  discoverPythonKernels,
  fallbackKernelName,
  fallbackPythonDescription,
  fallbackPythonLabel,
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

type PyControllerEntry = {
  controller: vscode.NotebookController;
  descriptor: PythonKernelDescriptor;
  selectionDisposable: vscode.Disposable;
};

export class PyNotebookController implements vscode.Disposable {
  private readonly controllers = new Map<string, PyControllerEntry>();
  private readonly liveExecutions = new Map<string, LiveExecutionState>();
  private readonly notebookQueues = new Map<string, Promise<void>>();
  private readonly queueEpochs = new Map<string, number>();
  private readonly boundSessions = new WeakSet<object>();
  private readonly disposables: vscode.Disposable[] = [];
  private readonly extensionId: string;
  private kernelDescriptors: PythonKernelDescriptor[] = [];
  private executionOrder = 0;
  private renderInterval: ReturnType<typeof setInterval> | null = null;
  private varNotifier?: VariableNotifier;

  constructor(ctx: vscode.ExtensionContext, varNotifier?: VariableNotifier) {
    this.extensionId = ctx.extension.id;
    this.varNotifier = varNotifier;
    this.refreshKernelCatalog();
    this.disposables.push(
      vscode.workspace.onDidOpenNotebookDocument((notebook) => {
        this.syncNotebookAffinities(notebook);
      }),
      vscode.workspace.onDidChangeConfiguration((event) => {
        if (affectsPythonConfig(event, 'pythonPath')) {
          this.refreshKernelCatalog();
        }
      }),
    );
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
    const includeManual = options?.includeManual ?? false;

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
    this.refreshKernelCatalog();
    vscode.window.showInformationMessage(`Python path set to: ${descriptor.pythonPath}`);
  }

  private refreshKernelCatalog(): void {
    this.kernelDescriptors = discoverPythonKernels(this.pyBinPath());
    const nextIds = new Set<string>();

    for (const descriptor of this.kernelDescriptors) {
      const controllerId = this.controllerIdForDescriptor(descriptor);
      nextIds.add(controllerId);

      const existing = this.controllers.get(controllerId);
      if (existing) {
        existing.descriptor = descriptor;
        existing.controller.label = descriptor.displayName;
        existing.controller.description = descriptor.description ?? `${EXTENSION_BRAND} Python Kernel`;
        existing.controller.detail = descriptor.pythonPath;
        continue;
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

  private pyBinPath(): string {
    return getPythonConfigValue('pythonPath', 'python3');
  }

  private execTimeoutMs(): number {
    return getPythonConfigValue('execTimeoutMs', 3_000_000) ?? 3_000_000;
  }

  private async execute(
    controller: vscode.NotebookController,
    cells: vscode.NotebookCell[],
    notebook: vscode.NotebookDocument,
  ): Promise<void> {
    const descriptor = this.controllerEntryForController(controller)?.descriptor
      ?? this.resolveDescriptorForNotebook(notebook);
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
      if (session.cachedVars().vars.length === 0) {
        await this.refreshVariableCache(session);
      }
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
        await this.refreshVariableCache(session);
      } catch (err: any) {
        liveState.completed = true;
        await this.waitForPendingRender(liveState);
        await this.replaceExecutionOutput(liveState, {
          ...liveState.result,
          error: err.message,
        }, chunkId);
        execution.end(false, Date.now());
        await this.refreshVariableCache(session);
      } finally {
        cancelDisp.dispose();
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
    const session = getPySession(docUri);
    const interrupted = session?.isBusy() ?? false;
    this.bumpQueueEpoch(docUri);
    session?.interrupt();
    return interrupted;
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
      this.refreshKernelCatalog();
    }
    return this.kernelDescriptors;
  }

  private manualDescriptor(pythonPath: string): PythonKernelDescriptor {
    return {
      id: `executable:${pythonPath}`,
      label: fallbackPythonLabel(pythonPath),
      description: fallbackPythonDescription(pythonPath),
      detail: pythonPath,
      env: undefined,
      environmentType: undefined,
      pythonPath,
      displayName: fallbackPythonLabel(pythonPath),
      kernelspecName: fallbackKernelName(pythonPath),
      source: 'executable',
    };
  }

  private resolveDefaultDescriptor(): PythonKernelDescriptor | undefined {
    const configuredPython = this.pyBinPath();
    return this.availableDescriptors().find((descriptor) => descriptor.pythonPath === configuredPython)
      ?? this.availableDescriptors()[0];
  }

  private controllerIdForDescriptor(descriptor: PythonKernelDescriptor): string {
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
      entry.controller.updateNotebookAffinity(
        notebook,
        entry.descriptor.id === preferred.id
          ? vscode.NotebookControllerAffinity.Preferred
          : vscode.NotebookControllerAffinity.Default,
      );
    }
  }

  private resolveDescriptorForNotebook(notebook: vscode.NotebookDocument): PythonKernelDescriptor {
    const meta = (notebook.metadata ?? {}) as NotebookPythonMetadata;
    const kernelMeta = getNotebookKernelMetadata(meta);
    const descriptors = this.availableDescriptors();

    return descriptors.find((descriptor) =>
      this.controllerIdForDescriptor(descriptor) === kernelMeta?.controllerId
      || descriptor.kernelspecName === kernelMeta?.kernelspecName
      || descriptor.kernelspecName === meta.kernelspec?.name
      || descriptor.pythonPath === kernelMeta?.pythonPath
      || descriptor.displayName === kernelMeta?.displayName
      || descriptor.displayName === meta.kernelspec?.display_name,
    ) ?? this.resolveDefaultDescriptor()
      ?? this.manualDescriptor(this.pyBinPath());
  }

  private async handleControllerSelection(
    notebook: vscode.NotebookDocument,
    descriptor: PythonKernelDescriptor,
    controllerId: string,
  ): Promise<void> {
    if (notebook.notebookType !== PY_NOTEBOOK_TYPE) return;
    await this.persistSelection(notebook, descriptor, controllerId);
    const session = getPySession(notebook.uri.toString());
    if (session) {
      session.setKernel(descriptor.pythonPath, descriptor.env);
      await session.restart().catch(() => undefined);
      this.varNotifier?.notifyChanged(notebook);
    }
    this.syncNotebookAffinities(notebook);
  }

  private async selectDescriptorForNotebook(
    notebook: vscode.NotebookDocument,
    descriptor: PythonKernelDescriptor,
  ): Promise<void> {
    const controllerId = this.controllerIdForDescriptor(descriptor);
    await this.persistSelection(notebook, descriptor, controllerId);

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

  private async persistSelection(
    notebook: vscode.NotebookDocument,
    descriptor: PythonKernelDescriptor,
    controllerId = this.controllerIdForDescriptor(descriptor),
  ): Promise<void> {
    const meta = (notebook.metadata ?? {}) as NotebookPythonMetadata;
    const nextMeta: NotebookPythonMetadata = {
      ...meta,
      kernelspec: {
        ...(meta.kernelspec ?? {}),
        display_name: descriptor.displayName,
        language: 'python',
        name: descriptor.kernelspecName,
      },
      language_info: {
        ...(meta.language_info ?? {}),
        name: 'python',
      },
      rNotebook: {
        ...(getNotebookKernelMetadata(meta) ?? {}),
        controllerId,
        kernelspecName: descriptor.kernelspecName,
        displayName: descriptor.displayName,
        pythonPath: descriptor.pythonPath,
      },
    };

    if (metadataMatches(meta, nextMeta)) return;

    const edit = new vscode.WorkspaceEdit();
    edit.set(notebook.uri, [vscode.NotebookEdit.updateNotebookMetadata(nextMeta)]);
    await vscode.workspace.applyEdit(edit);
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

  private async refreshVariableCache(
    session: ReturnType<typeof getOrCreatePySession>,
  ): Promise<void> {
    try {
      await session.vars();
    } catch {}
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

function metadataMatches(current: NotebookPythonMetadata, next: NotebookPythonMetadata): boolean {
  const currentKernelMeta = getNotebookKernelMetadata(current);
  const nextKernelMeta = getNotebookKernelMetadata(next);
  return (
    current.kernelspec?.name === next.kernelspec?.name &&
    current.kernelspec?.display_name === next.kernelspec?.display_name &&
    current.language_info?.name === next.language_info?.name &&
    currentKernelMeta?.controllerId === nextKernelMeta?.controllerId &&
    currentKernelMeta?.displayName === nextKernelMeta?.displayName &&
    currentKernelMeta?.kernelspecName === nextKernelMeta?.kernelspecName &&
    currentKernelMeta?.pythonPath === nextKernelMeta?.pythonPath
  );
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
