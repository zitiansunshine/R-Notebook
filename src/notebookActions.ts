import { execFile } from 'child_process';
import { promises as fs } from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as vscode from 'vscode';

import {
  NOTEBOOK_TYPE as R_NOTEBOOK_TYPE,
  PY_NOTEBOOK_TYPE,
  getPythonConfigValue,
} from './extensionIds';
import { mergeRmdSourceAndState, PersistedRmdState } from './rmdPersistedState';
import { RmdOutputStore } from './rmdOutputStore';
import { buildStaticRmdExportHtml } from './staticRmdExport';

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
      edits.push(buildCellOutputsEdit(cell, []));
      changed = true;
    }

    if (notebook.notebookType === PY_NOTEBOOK_TYPE) {
      const clearedMetadata = clearIpynbPersistedOutputs(cell.metadata);
      if (clearedMetadata) {
        edits.push(buildCellMetadataEdit(cell, clearedMetadata));
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
    if (await clearPersistedRmdNotebookState(notebook)) {
      changed = true;
    }
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
    edits.push(buildCellMetadataEdit(change.cell, clearedMetadata));
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

async function clearPersistedRmdNotebookState(
  notebook: vscode.NotebookDocument,
): Promise<boolean> {
  const textDocument = await vscode.workspace.openTextDocument(notebook.uri);
  const currentText = textDocument.getText();
  const nextText = clearPersistedRmdOutputs(currentText);
  if (nextText === currentText) return false;

  const edit = new vscode.WorkspaceEdit();
  edit.replace(textDocument.uri, fullDocumentRange(textDocument), nextText);
  return vscode.workspace.applyEdit(edit);
}

export async function exportDocumentUri(uri: vscode.Uri): Promise<void> {
  const format = await promptExportFormat();
  if (!format) return;

  const fsPath = uri.fsPath;
  const ext = path.extname(fsPath).toLowerCase();
  const parsed = path.parse(fsPath);
  const guessedOutput = vscode.Uri.file(path.join(parsed.dir, `${parsed.name}.${format}`));

  if (!isExportableExtension(ext)) {
    vscode.window.showWarningMessage('Export is supported for .Rmd, .qmd, and .ipynb files.');
    return;
  }

  const exportResult = await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `Exporting ${path.basename(fsPath)} to ${format.toUpperCase()}...`,
    },
    async (): Promise<{
      outputExists: boolean;
      guessedOutput: vscode.Uri;
    } | null> => {
      try {
        if (ext === '.rmd' || ext === '.qmd') {
          await exportStaticRmdLikeDocument(uri, guessedOutput, format);
        } else {
          const spec = buildExportSpec(fsPath, ext, format);
          if (!spec) {
            throw new Error('Export is supported for .Rmd, .qmd, and .ipynb files.');
          }
          await execFileAsync(spec.command, spec.args, path.dirname(fsPath));
        }
      } catch (error) {
        const detail = formatExecError(error);
        vscode.window.showErrorMessage(`Export failed: ${detail}`);
        return null;
      }

      return {
        outputExists: await fileExists(guessedOutput),
        guessedOutput,
      };
    },
  );

  if (!exportResult) return;

  const action = await vscode.window.showInformationMessage(
    exportResult.outputExists
      ? `Exported ${path.basename(exportResult.guessedOutput.fsPath)}`
      : `Export completed for ${path.basename(fsPath)}`,
    ...(exportResult.outputExists ? ['Open', 'Reveal'] : ['Reveal Folder']),
  );

  if (action === 'Open' && exportResult.outputExists) {
    await vscode.env.openExternal(exportResult.guessedOutput);
  } else if (action === 'Reveal' && exportResult.outputExists) {
    await vscode.commands.executeCommand('revealFileInOS', exportResult.guessedOutput);
  } else if (action === 'Reveal Folder') {
    await vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(path.dirname(fsPath)));
  }
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

function isExportableExtension(ext: string): boolean {
  return ext === '.rmd' || ext === '.qmd' || ext === '.ipynb';
}

function buildExportSpec(
  fsPath: string,
  ext: string,
  format: ExportFormat,
): { command: string; args: string[] } | undefined {
  if (ext === '.ipynb') {
    const pythonPath = getPythonConfigValue('pythonPath', 'python3');
    return {
      command: pythonPath,
      args: ['-m', 'jupyter', 'nbconvert', '--to', format, fsPath],
    };
  }

  return undefined;
}

async function exportStaticRmdLikeDocument(
  uri: vscode.Uri,
  outputUri: vscode.Uri,
  format: ExportFormat,
): Promise<void> {
  const document = await vscode.workspace.openTextDocument(uri);
  const html = buildStaticRmdExportHtml(document.getText(), path.basename(uri.fsPath));

  if (format === 'html') {
    await vscode.workspace.fs.writeFile(outputUri, Buffer.from(html, 'utf8'));
    return;
  }

  await exportStaticHtmlToPdf(html, outputUri);
}

async function exportStaticHtmlToPdf(
  html: string,
  outputUri: vscode.Uri,
): Promise<void> {
  const converter = await resolveStaticPdfConverter();
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'r-notebook-export-'));
  const tempHtmlPath = path.join(tempDir, `${path.parse(outputUri.fsPath).name}.html`);

  try {
    await fs.writeFile(tempHtmlPath, html, 'utf8');
    await execFileAsync(
      converter.command,
      [...converter.argsPrefix, '-f', 'html', tempHtmlPath, '-o', outputUri.fsPath],
      path.dirname(outputUri.fsPath),
    );
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
  }
}

async function resolveStaticPdfConverter(): Promise<{ command: string; argsPrefix: string[] }> {
  const candidates = [
    { command: 'pandoc', argsPrefix: [] },
    { command: 'quarto', argsPrefix: ['pandoc'] },
    { command: '/Applications/RStudio.app/Contents/MacOS/pandoc/pandoc', argsPrefix: [] },
  ];

  for (const candidate of candidates) {
    try {
      await execFileAsync(
        candidate.command,
        [...candidate.argsPrefix, '--version'],
        os.tmpdir(),
      );
      return candidate;
    } catch {
      continue;
    }
  }

  throw new Error('Static PDF export requires Pandoc or Quarto to be installed.');
}

function clearIpynbPersistedOutputs(
  metadata: vscode.NotebookCellMetadata | undefined,
): vscode.NotebookCellMetadata | null {
  const next = { ...((metadata ?? {}) as Record<string, unknown>) };
  if (!('_ipynb_outputs' in next)) return null;
  delete next._ipynb_outputs;
  return next;
}

function buildCellOutputsEdit(
  cell: vscode.NotebookCell,
  outputs: readonly vscode.NotebookCellOutput[],
): vscode.NotebookEdit {
  const notebookEdit = vscode.NotebookEdit as typeof vscode.NotebookEdit & {
    updateCellOutputs?: (
      index: number,
      outputs: readonly vscode.NotebookCellOutput[],
    ) => vscode.NotebookEdit;
  };
  if (typeof notebookEdit.updateCellOutputs === 'function') {
    return notebookEdit.updateCellOutputs(cell.index, outputs);
  }
  return buildReplaceCellEdit(cell, { outputs });
}

function buildCellMetadataEdit(
  cell: vscode.NotebookCell,
  metadata: vscode.NotebookCellMetadata,
): vscode.NotebookEdit {
  const notebookEdit = vscode.NotebookEdit as typeof vscode.NotebookEdit & {
    updateCellMetadata?: (
      index: number,
      metadata: vscode.NotebookCellMetadata,
    ) => vscode.NotebookEdit;
  };
  if (typeof notebookEdit.updateCellMetadata === 'function') {
    return notebookEdit.updateCellMetadata(cell.index, metadata);
  }
  return buildReplaceCellEdit(cell, { metadata });
}

function buildReplaceCellEdit(
  cell: vscode.NotebookCell,
  overrides: {
    metadata?: vscode.NotebookCellMetadata;
    outputs?: readonly vscode.NotebookCellOutput[];
  },
): vscode.NotebookEdit {
  const notebookEdit = vscode.NotebookEdit as typeof vscode.NotebookEdit & {
    replaceCells?: (
      range: vscode.NotebookRange,
      newCells: readonly vscode.NotebookCellData[],
    ) => vscode.NotebookEdit;
  };
  if (typeof notebookEdit.replaceCells !== 'function') {
    throw new Error('Notebook edit API is unavailable in this editor build.');
  }

  const replacement = new vscode.NotebookCellData(
    cell.kind,
    cell.document.getText(),
    cell.document.languageId,
  );
  replacement.metadata = (overrides.metadata ?? cell.metadata) as vscode.NotebookCellMetadata;
  replacement.outputs = [...(overrides.outputs ?? cell.outputs)];
  return notebookEdit.replaceCells(
    new vscode.NotebookRange(cell.index, cell.index + 1),
    [replacement],
  );
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

function fullDocumentRange(document: vscode.TextDocument): vscode.Range {
  if (document.lineCount === 0) return new vscode.Range(0, 0, 0, 0);
  return new vscode.Range(0, 0, document.lineCount - 1, document.lineAt(document.lineCount - 1).text.length);
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
