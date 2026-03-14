import { execFile } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';

import {
  NOTEBOOK_TYPE as R_NOTEBOOK_TYPE,
  PY_NOTEBOOK_TYPE,
  getPythonConfigValue,
  getRConfigValue,
} from './extensionIds';
import { mergeRmdSourceAndState, PersistedRmdState } from './rmdPersistedState';
import { RmdOutputStore } from './rmdOutputStore';

type ExportFormat = 'html' | 'pdf';

const EMPTY_RMD_STATE: PersistedRmdState = {
  version: 2,
  codeCellStates: [],
};

export async function clearNotebookOutputs(
  notebook: vscode.NotebookDocument,
  outputStore: RmdOutputStore,
): Promise<boolean> {
  const edits: vscode.NotebookEdit[] = [];
  let changed = false;

  for (const cell of notebook.getCells()) {
    if (cell.kind !== vscode.NotebookCellKind.Code) continue;

    if (cell.outputs.length > 0) {
      edits.push(vscode.NotebookEdit.updateCellOutputs(cell.index, []));
      changed = true;
    }

    if (notebook.notebookType === PY_NOTEBOOK_TYPE) {
      const clearedMetadata = clearIpynbPersistedOutputs(cell.metadata);
      if (clearedMetadata) {
        edits.push(vscode.NotebookEdit.updateCellMetadata(cell.index, clearedMetadata));
        changed = true;
      }
    }
  }

  if (edits.length > 0) {
    const edit = new vscode.WorkspaceEdit();
    edit.set(notebook.uri, edits);
    await vscode.workspace.applyEdit(edit);
  }

  if (notebook.notebookType === R_NOTEBOOK_TYPE) {
    outputStore.clear(notebook.uri.toString());
  }

  if (changed) {
    await saveNotebookIfSupported(notebook);
  }

  return changed;
}

export async function cleanupClearedNotebookChange(
  event: vscode.NotebookDocumentChangeEvent,
  outputStore: RmdOutputStore,
): Promise<void> {
  if (!event.cellChanges.some((change) => change.outputs !== undefined && change.outputs.length === 0)) {
    return;
  }

  if (event.notebook.notebookType === R_NOTEBOOK_TYPE) {
    const hasRemainingOutputs = event.notebook.getCells().some((cell) =>
      cell.kind === vscode.NotebookCellKind.Code && cell.outputs.length > 0,
    );
    if (!hasRemainingOutputs) {
      outputStore.clear(event.notebook.uri.toString());
      await saveNotebookIfSupported(event.notebook);
    }
    return;
  }

  if (event.notebook.notebookType !== PY_NOTEBOOK_TYPE) return;

  const edits: vscode.NotebookEdit[] = [];
  for (const change of event.cellChanges) {
    if (change.outputs === undefined || change.outputs.length > 0) continue;
    if (change.cell.kind !== vscode.NotebookCellKind.Code) continue;
    const clearedMetadata = clearIpynbPersistedOutputs(change.cell.metadata);
    if (!clearedMetadata) continue;
    edits.push(vscode.NotebookEdit.updateCellMetadata(change.cell.index, clearedMetadata));
  }

  if (edits.length === 0) return;
  const edit = new vscode.WorkspaceEdit();
  edit.set(event.notebook.uri, edits);
  await vscode.workspace.applyEdit(edit);
  await saveNotebookIfSupported(event.notebook);
}

export function clearPersistedRmdOutputs(text: string): string {
  return mergeRmdSourceAndState(text, EMPTY_RMD_STATE);
}

export async function exportDocumentUri(uri: vscode.Uri): Promise<void> {
  const format = await promptExportFormat();
  if (!format) return;

  const fsPath = uri.fsPath;
  const ext = path.extname(fsPath).toLowerCase();
  const parsed = path.parse(fsPath);
  const guessedOutput = vscode.Uri.file(path.join(parsed.dir, `${parsed.name}.${format}`));

  const spec = buildExportSpec(fsPath, ext, format);
  if (!spec) {
    vscode.window.showWarningMessage('Export is supported for .Rmd, .qmd, and .ipynb files.');
    return;
  }

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `Exporting ${path.basename(fsPath)} to ${format.toUpperCase()}...`,
    },
    async () => {
      try {
        await execFileAsync(spec.command, spec.args, path.dirname(fsPath));
      } catch (error) {
        const detail = formatExecError(error);
        vscode.window.showErrorMessage(`Export failed: ${detail}`);
        return;
      }

      const outputExists = await fileExists(guessedOutput);
      const action = await vscode.window.showInformationMessage(
        outputExists
          ? `Exported ${path.basename(guessedOutput.fsPath)}`
          : `Export completed for ${path.basename(fsPath)}`,
        ...(outputExists ? ['Open', 'Reveal'] : ['Reveal Folder']),
      );

      if (action === 'Open' && outputExists) {
        await vscode.env.openExternal(guessedOutput);
      } else if (action === 'Reveal' && outputExists) {
        await vscode.commands.executeCommand('revealFileInOS', guessedOutput);
      } else if (action === 'Reveal Folder') {
        await vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(path.dirname(fsPath)));
      }
    },
  );
}

async function promptExportFormat(): Promise<ExportFormat | undefined> {
  const pick = await vscode.window.showQuickPick<{
    label: string;
    value: ExportFormat;
  }>(
    [
      { label: 'HTML', value: 'html' },
      { label: 'PDF', value: 'pdf' },
    ],
    {
      title: 'Export Notebook',
      placeHolder: 'Choose an export format',
    },
  );
  return pick?.value;
}

function buildExportSpec(
  fsPath: string,
  ext: string,
  format: ExportFormat,
): { command: string; args: string[] } | undefined {
  if (ext === '.qmd') {
    return {
      command: 'quarto',
      args: ['render', fsPath, '--to', format],
    };
  }

  if (ext === '.rmd') {
    const rPath = getRConfigValue('rPath', 'Rscript');
    const outputFormat = format === 'html' ? 'html_document' : 'pdf_document';
    const expr = `rmarkdown::render(${JSON.stringify(fsPath)}, output_format = ${JSON.stringify(outputFormat)})`;
    return {
      command: rPath,
      args: ['-e', expr],
    };
  }

  if (ext === '.ipynb') {
    const pythonPath = getPythonConfigValue('pythonPath', 'python3');
    return {
      command: pythonPath,
      args: ['-m', 'jupyter', 'nbconvert', '--to', format, fsPath],
    };
  }

  return undefined;
}

function clearIpynbPersistedOutputs(
  metadata: vscode.NotebookCellMetadata | undefined,
): vscode.NotebookCellMetadata | null {
  const next = { ...((metadata ?? {}) as Record<string, unknown>) };
  if (!('_ipynb_outputs' in next)) return null;
  delete next._ipynb_outputs;
  return next;
}

async function saveNotebookIfSupported(notebook: vscode.NotebookDocument): Promise<void> {
  const savable = notebook as vscode.NotebookDocument & { save?: () => Thenable<boolean> };
  if (typeof savable.save === 'function') {
    await savable.save();
    return;
  }

  const active = vscode.window.activeNotebookEditor;
  if (active?.notebook.uri.toString() === notebook.uri.toString()) {
    await vscode.commands.executeCommand('workbench.action.files.save');
  }
}

async function execFileAsync(
  command: string,
  args: string[],
  cwd: string,
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    execFile(command, args, { cwd, maxBuffer: 64 * 1024 * 1024 }, (error, _stdout, stderr) => {
      if (!error) {
        resolve();
        return;
      }
      const enriched = error as Error & { stderr?: string };
      enriched.stderr = stderr;
      reject(enriched);
    });
  });
}

function formatExecError(error: unknown): string {
  if (error && typeof error === 'object') {
    const candidate = error as NodeJS.ErrnoException & { stderr?: string };
    const stderr = candidate.stderr?.trim();
    if (stderr) return stderr;
    if (candidate.code === 'ENOENT') {
      return 'Required export tool was not found on PATH.';
    }
    if (candidate.message) return candidate.message;
  }
  return String(error);
}

async function fileExists(uri: vscode.Uri): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch {
    return false;
  }
}
