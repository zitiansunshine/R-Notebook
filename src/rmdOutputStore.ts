import * as vscode from 'vscode';
import { ExecResult } from './kernelProtocol';
import { normalizePersistedExecResult } from './rmdPersistedState';
import {
  fingerprintNotebookCodeCell,
  restoreStoredResults,
  StoredExecResultState,
} from './rmdResultMapping';

export const RAW_EXEC_RESULT_MIME = 'application/vnd.rnotebook.exec-result+json';

export class RmdOutputStore {
  private readonly results = new Map<string, Map<string, ExecResult>>();
  private readonly states = new Map<string, StoredExecResultState[]>();
  private readonly hardResetUris = new Set<string>();

  markHardReset(docUri: string): void {
    this.clear(docUri);
    this.hardResetUris.add(docUri);
  }

  applyNotebookChange(event: vscode.NotebookDocumentChangeEvent): void {
    const docUri = event.notebook.uri.toString();
    const next = new Map(this.results.get(docUri) ?? []);
    const suppressTransfer = this.hardResetUris.delete(docUri);

    for (const contentChange of event.contentChanges) {
      const removedStates: StoredExecResultState[] = [];

      for (const cell of contentChange.removedCells) {
        if (cell.kind !== vscode.NotebookCellKind.Code) continue;
        const cellDocUri = cell.document.uri.toString();
        if (!suppressTransfer) {
          removedStates.push({
            fingerprint: fingerprintNotebookCodeCell(
              cell.document.languageId,
              cell.document.getText(),
              ((cell.metadata ?? {}) as { options?: Record<string, unknown> }).options,
            ),
            result: normalizePersistedExecResult(
              next.get(cellDocUri) ?? execResultFromCellOutputs(cell.outputs),
            ),
          });
        }
        next.delete(cellDocUri);
      }

      if (removedStates.length === 0) continue;

      const transferred = restoreStoredResults(
        contentChange.addedCells
          .filter((cell) => cell.kind === vscode.NotebookCellKind.Code)
          .map((cell) => ({
            key: cell.document.uri.toString(),
            fingerprint: fingerprintNotebookCodeCell(
              cell.document.languageId,
              cell.document.getText(),
              ((cell.metadata ?? {}) as { options?: Record<string, unknown> }).options,
            ),
          })),
        removedStates,
      );

      for (const [cellDocUri, result] of transferred) {
        next.set(cellDocUri, result);
      }
    }

    for (const cellChange of event.cellChanges) {
      if (cellChange.cell.kind !== vscode.NotebookCellKind.Code) continue;
      if (cellChange.outputs === undefined) continue;
      const currentResult = normalizePersistedExecResult(execResultFromCellOutputs(cellChange.outputs));
      const cellDocUri = cellChange.cell.document.uri.toString();
      if (currentResult) {
        next.set(cellDocUri, currentResult);
      } else {
        next.delete(cellDocUri);
      }
    }

    this.syncNotebook(event.notebook, next);
    if (next.size === 0) {
      this.results.delete(docUri);
    } else {
      this.results.set(docUri, next);
    }
  }

  getForNotebook(notebook: vscode.NotebookDocument): Map<number, ExecResult> {
    const docUri = notebook.uri.toString();
    const stored = this.results.get(docUri);
    const restored = new Map<number, ExecResult>();

    if (stored) {
      for (const cell of notebook.getCells()) {
        if (cell.kind !== vscode.NotebookCellKind.Code) continue;
        const result = stored.get(cell.document.uri.toString());
        if (result) restored.set(cell.index, result);
      }
    }

    const fallback = restoreStoredResults(
      notebook
        .getCells()
        .filter((cell) => cell.kind === vscode.NotebookCellKind.Code)
        .map((cell) => ({
          key: cell.index,
          fingerprint: fingerprintNotebookCodeCell(
            cell.document.languageId,
            cell.document.getText(),
            ((cell.metadata ?? {}) as { options?: Record<string, unknown> }).options,
          ),
        })),
      this.states.get(docUri) ?? [],
    );

    for (const [cellIndex, result] of fallback) {
      if (!restored.has(cellIndex)) restored.set(cellIndex, result);
    }

    return restored;
  }

  setNotebookCellResult(
    notebook: vscode.NotebookDocument,
    cellIndex: number,
    result: ExecResult | null,
  ): void {
    const docUri = notebook.uri.toString();
    const targetCell = notebook.getCells().find((cell) => cell.index === cellIndex);
    if (!targetCell || targetCell.kind !== vscode.NotebookCellKind.Code) return;

    const next = new Map(this.results.get(docUri) ?? []);
    const normalizedResult = normalizePersistedExecResult(result);
    const cellDocUri = targetCell.document.uri.toString();
    if (normalizedResult) {
      next.set(cellDocUri, normalizedResult);
    } else {
      next.delete(cellDocUri);
    }

    this.syncNotebook(notebook, next);
  }

  syncNotebook(
    notebook: vscode.NotebookDocument,
    existing?: Map<string, ExecResult>,
  ): void {
    const docUri = notebook.uri.toString();
    const next = existing ?? new Map(this.results.get(docUri) ?? []);
    const liveCodeCellUris = new Set(
      notebook
        .getCells()
        .filter((cell) => cell.kind === vscode.NotebookCellKind.Code)
        .map((cell) => cell.document.uri.toString()),
    );

    for (const cellDocUri of next.keys()) {
      if (!liveCodeCellUris.has(cellDocUri)) next.delete(cellDocUri);
    }

    for (const cell of notebook.getCells()) {
      if (cell.kind !== vscode.NotebookCellKind.Code) continue;
      const currentResult = normalizePersistedExecResult(execResultFromCellOutputs(cell.outputs));
      if (!currentResult) continue;
      next.set(cell.document.uri.toString(), currentResult);
    }

    this.writeState(docUri, notebook, next);
    if (existing) return;
  }

  clear(docUri: string): void {
    this.results.delete(docUri);
    this.states.delete(docUri);
  }

  private writeState(
    docUri: string,
    notebook: vscode.NotebookDocument,
    resultsByCellUri: Map<string, ExecResult>,
  ): void {
    if (resultsByCellUri.size === 0) {
      this.results.delete(docUri);
    } else {
      this.results.set(docUri, resultsByCellUri);
    }

    const nextStates: StoredExecResultState[] = [];
    for (const cell of notebook.getCells()) {
      if (cell.kind !== vscode.NotebookCellKind.Code) continue;
      const result = normalizePersistedExecResult(
        resultsByCellUri.get(cell.document.uri.toString()) ?? execResultFromCellOutputs(cell.outputs),
      );
      nextStates.push({
        fingerprint: fingerprintNotebookCodeCell(
          cell.document.languageId,
          cell.document.getText(),
          ((cell.metadata ?? {}) as { options?: Record<string, unknown> }).options,
        ),
        result,
      });
    }
    trimTrailingEmptyStates(nextStates);
    if (nextStates.length === 0) {
      this.states.delete(docUri);
    } else {
      this.states.set(docUri, nextStates);
    }
  }
}

export function notebookOutputFromExecResult(
  result: ExecResult | null,
  chunkId: string,
): vscode.NotebookCellOutput[] {
  const items = notebookOutputItemsFromExecResult(result, chunkId);
  if (!items) return [];
  return [new vscode.NotebookCellOutput(items, { chunkId, running: false })];
}

export function notebookOutputItemsFromExecResult(
  result: ExecResult | null,
  chunkId: string,
  options?: { running?: boolean },
): vscode.NotebookCellOutputItem[] | null {
  void chunkId;
  if (!result) return null;
  if (!options?.running && !hasRenderableExecResult(result)) return null;
  return [vscode.NotebookCellOutputItem.json(result, RAW_EXEC_RESULT_MIME)];
}

export function execResultFromCellOutputs(
  outputs: readonly vscode.NotebookCellOutput[],
): ExecResult | null {
  for (const output of outputs) {
    for (const item of output.items) {
      if (item.mime !== RAW_EXEC_RESULT_MIME) continue;
      try {
        const decoded = new TextDecoder().decode(item.data);
        const parsed = JSON.parse(decoded) as ExecResult;
        return isExecResult(parsed) ? parsed : null;
      } catch {
        return null;
      }
    }
  }
  return null;
}

function isExecResult(value: unknown): value is ExecResult {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Record<string, unknown>;
  return (
    candidate.type === 'result' &&
    typeof candidate.chunk_id === 'string' &&
    (candidate.source_code === undefined || typeof candidate.source_code === 'string') &&
    (candidate.console_segments === undefined || Array.isArray(candidate.console_segments)) &&
    (candidate.console === undefined || typeof candidate.console === 'string') &&
    typeof candidate.stdout === 'string' &&
    typeof candidate.stderr === 'string' &&
    (candidate.error === null || typeof candidate.error === 'string') &&
    Array.isArray(candidate.plots) &&
    Array.isArray(candidate.dataframes) &&
    (candidate.plots_html === undefined || Array.isArray(candidate.plots_html))
  );
}

function hasRenderableExecResult(result: ExecResult): boolean {
  return Boolean(
    (result.console?.trim() ?? '') ||
    result.stdout.trim() ||
    result.stderr.trim() ||
    result.plots.length ||
    (result.plots_html?.length ?? 0) ||
    result.dataframes.length ||
    (result.output_order?.length ?? 0) ||
    result.error,
  );
}

function trimTrailingEmptyStates(values: StoredExecResultState[]): void {
  while (values.length > 0 && !values[values.length - 1].result) values.pop();
}
