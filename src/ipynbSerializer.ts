// =============================================================================
// ipynbSerializer.ts — VS Code NotebookSerializer for Jupyter .ipynb files.
//
// Reads/writes nbformat 4.  Code cells → VS Code Code cells (language from
// notebook metadata).  Markdown/raw cells → Markup cells.
//
// =============================================================================

import * as vscode from 'vscode';
import { RAW_EXEC_RESULT_MIME } from './rmdOutputStore';

// ---------------------------------------------------------------------------
// Minimal nbformat 4 shape

interface JupyterNotebook {
  nbformat:       number;
  nbformat_minor: number;
  metadata:       JupyterMeta;
  cells:          JupyterCell[];
}

interface JupyterMeta {
  kernelspec?:    { language?: string; display_name?: string; name?: string };
  language_info?: { name?: string };
  [k: string]:    unknown;
}

interface JupyterCell {
  cell_type:       'code' | 'markdown' | 'raw';
  metadata:        Record<string, unknown>;
  source:          string | string[];
  outputs?:        unknown[];
  execution_count?: number | null;
}

interface JupyterDisplayOutput {
  output_type: 'display_data' | 'execute_result';
  data?: Record<string, unknown>;
  metadata?: Record<string, unknown>;
  execution_count?: number | null;
}

// ---------------------------------------------------------------------------

function joinSource(src: string | string[]): string {
  return Array.isArray(src) ? src.join('') : src;
}

function splitSource(src: string): string[] {
  // Preserve Jupyter convention: each line ends with \n except possibly the last
  const lines = src.split('\n');
  return lines.map((l, i) => (i < lines.length - 1 ? l + '\n' : l));
}

function detectLanguage(meta: JupyterMeta): string {
  return (
    meta.kernelspec?.language ??
    meta.language_info?.name  ??
    'python'
  );
}

function deserializeCustomOutputs(outputs: unknown[]): vscode.NotebookCellOutput[] {
  const decoded: vscode.NotebookCellOutput[] = [];

  for (const output of outputs) {
    if (!output || typeof output !== 'object') continue;
    const candidate = output as JupyterDisplayOutput;
    if (candidate.output_type !== 'display_data' && candidate.output_type !== 'execute_result') {
      continue;
    }
    const rawExecResult = candidate.data?.[RAW_EXEC_RESULT_MIME];
    if (!rawExecResult) continue;
    decoded.push(new vscode.NotebookCellOutput([
      vscode.NotebookCellOutputItem.json(rawExecResult, RAW_EXEC_RESULT_MIME),
    ]));
  }

  return decoded;
}

function serializeCellOutputs(outputs: readonly vscode.NotebookCellOutput[]): unknown[] {
  const encoded: unknown[] = [];

  for (const output of outputs) {
    for (const item of output.items) {
      if (item.mime !== RAW_EXEC_RESULT_MIME) continue;
      try {
        const raw = new TextDecoder().decode(item.data);
        encoded.push({
          output_type: 'display_data',
          data: {
            [RAW_EXEC_RESULT_MIME]: JSON.parse(raw),
          },
          metadata: {},
        });
      } catch {
        continue;
      }
    }
  }

  return encoded;
}

export class IpynbSerializer implements vscode.NotebookSerializer {

  deserializeNotebook(data: Uint8Array): vscode.NotebookData {
    let nb: JupyterNotebook;
    try {
      nb = JSON.parse(Buffer.from(data).toString('utf-8'));
    } catch {
      return new vscode.NotebookData([]);
    }

    const lang = detectLanguage(nb.metadata ?? {});

    const cells: vscode.NotebookCellData[] = (nb.cells ?? [])
      .filter(c => c.cell_type === 'code' || c.cell_type === 'markdown')
      .map(c => {
        const source = joinSource(c.source ?? '');
        const kind   = c.cell_type === 'markdown'
          ? vscode.NotebookCellKind.Markup
          : vscode.NotebookCellKind.Code;

        const cellLang = kind === vscode.NotebookCellKind.Markup ? 'markdown' : lang;
        const cell = new vscode.NotebookCellData(kind, source, cellLang);
        cell.metadata = { ...c.metadata, _ipynb_outputs: c.outputs ?? [] };
        if (kind === vscode.NotebookCellKind.Code) {
          cell.outputs = deserializeCustomOutputs(c.outputs ?? []);
        }
        return cell;
      });

    const nd = new vscode.NotebookData(cells);
    nd.metadata = nb.metadata ?? {};
    return nd;
  }

  serializeNotebook(data: vscode.NotebookData): Uint8Array {
    const meta = (data.metadata ?? {}) as JupyterMeta;
    const lang = detectLanguage(meta);

    const cells: JupyterCell[] = data.cells.map(cell => {
      const { _ipynb_outputs, ...rest } = (cell.metadata ?? {}) as Record<string, unknown>;
      const renderedOutputs = serializeCellOutputs(cell.outputs ?? []);
      const outputs = renderedOutputs.length > 0
        ? renderedOutputs
        : (Array.isArray(_ipynb_outputs) ? _ipynb_outputs : []);

      return {
        cell_type:       cell.kind === vscode.NotebookCellKind.Markup ? 'markdown' : 'code',
        metadata:        rest,
        source:          splitSource(cell.value),
        outputs,
        execution_count: null,
      };
    });

    const outMeta: JupyterMeta = {
      kernelspec: {
        display_name: lang === 'python' ? 'Python 3' : lang,
        language:     lang,
        name:         lang === 'python' ? 'python3' : lang,
      },
      language_info: { name: lang },
      ...meta,
    };

    const nb: JupyterNotebook = {
      nbformat:       4,
      nbformat_minor: 5,
      metadata:       outMeta,
      cells,
    };

    return Buffer.from(JSON.stringify(nb, null, 1));
  }
}
