// =============================================================================
// rmdNotebookSerializer.ts — VS Code NotebookSerializer for .Rmd files
//
// This is the "Jupyter approach" to Copilot / Cursor-Tab support.
// By registering .Rmd as a notebook type, each R code cell becomes a real
// VS Code text editor (same as Jupyter cells) — so Copilot and Cursor-Tab
// automatically work inside every code cell.
//
// =============================================================================

import * as vscode from 'vscode';
import { ExecResult } from './kernelProtocol';
import { parseRmd } from './rmdParser';
import { execResultFromCellOutputs, notebookOutputFromExecResult } from './rmdOutputStore';
import { fingerprintNotebookCodeCell, StoredExecResultState } from './rmdResultMapping';
import {
  mergeRmdSourceAndState,
  normalizePersistedExecResult,
  restoreChunkResults,
  splitRmdSourceAndState,
} from './rmdPersistedState';

type RmdCellMetadata = {
  kind?: 'yaml_frontmatter';
  options?: Record<string, unknown>;
};

export class RmdNotebookSerializer implements vscode.NotebookSerializer {

  async deserializeNotebook(
    content: Uint8Array,
    _token: vscode.CancellationToken,
  ): Promise<vscode.NotebookData> {
    const text = new TextDecoder().decode(content);
    const { source, state } = splitRmdSourceAndState(text);
    const chunks = parseRmd(source);
    const restoredResults = restoreChunkResults(chunks, state);
    const cells: vscode.NotebookCellData[] = [];

    for (const chunk of chunks) {
      if (chunk.kind === 'yaml_frontmatter') {
        const cell = new vscode.NotebookCellData(
          vscode.NotebookCellKind.Markup,
          '```yaml\n' + chunk.code + '\n```',
          'markdown',
        );
        cell.metadata = { kind: 'yaml_frontmatter' };
        cells.push(cell);

      } else if (chunk.kind === 'prose') {
        cells.push(new vscode.NotebookCellData(
          vscode.NotebookCellKind.Markup,
          chunk.prose,
          'markdown',
        ));

      } else {
        const cell = new vscode.NotebookCellData(
          vscode.NotebookCellKind.Code,
          chunk.code,
          chunk.language || 'r',
        );
        cell.metadata = { options: chunk.options };
        cell.outputs = notebookOutputFromExecResult(restoredResults.get(chunk.id) ?? null, chunk.id);
        cells.push(cell);
      }
    }

    return new vscode.NotebookData(cells);
  }

  async serializeNotebook(
    data: vscode.NotebookData,
    _token: vscode.CancellationToken,
  ): Promise<Uint8Array> {
    const parts: string[] = [];
    const codeCellStates: StoredExecResultState[] = [];

    for (const cell of data.cells) {
      if (cell.kind === vscode.NotebookCellKind.Markup) {
        if (cell.metadata?.kind === 'yaml_frontmatter') {
          const raw = extractYamlFrontmatter(cell.value);
          parts.push(`---\n${raw}\n---`);
        } else {
          parts.push(cell.value);
        }
      } else {
        const meta = (cell.metadata ?? {}) as RmdCellMetadata;
        const lang = cell.languageId || 'r';
        const opts: Record<string, unknown> = meta.options ?? {};
        const optParts: string[] = [];
        if (opts['label']) optParts.push(String(opts['label']));
        for (const [k, v] of Object.entries(opts)) {
          if (k === 'label' || v == null) continue;
          const rKey = k.replace(/_/g, '.');
          if (typeof v === 'boolean') {
            if (!v) optParts.push(`${rKey}=FALSE`);
          } else {
            optParts.push(`${rKey}=${v}`);
          }
        }
        const fence = `\`\`\`{${lang}${optParts.length ? ' ' + optParts.join(', ') : ''}}`;
        parts.push(`${fence}\n${cell.value}\n\`\`\``);
        codeCellStates.push({
          fingerprint: fingerprintNotebookCodeCell(lang, cell.value, opts),
          result: normalizePersistedExecResult(execResultFromCellOutputs(cell.outputs)),
        });
      }
    }

    return new TextEncoder().encode(mergeRmdSourceAndState(parts.join('\n'), {
      version: 2,
      codeCellStates,
    }));
  }
}

function extractYamlFrontmatter(value: string): string {
  const match = value.match(/^```yaml\s*\r?\n([\s\S]*?)\r?\n```[\t ]*$/i);
  return match ? match[1] : value;
}
