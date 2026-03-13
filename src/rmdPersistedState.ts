import { ExecResult } from './kernelProtocol';
import { RmdChunk } from './rmdParser';
import {
  buildStoredStatesFromChunks,
  cloneExecResult,
  normalizeStoredExecResultState,
  restoreChunkResultsWithStoredStates,
  StoredExecResultState,
} from './rmdResultMapping';

const STATE_MARKER = 'RNOTEBOOK_RMD_STATE';
const STATE_BLOCK_RE = new RegExp(
  String.raw`(?:\r?\n)?<!--\s*${STATE_MARKER}\s+([A-Za-z0-9+/=]+)\s*-->\s*$`,
);

export interface PersistedRmdState {
  version: 2;
  codeCellStates: StoredExecResultState[];
  codeCellResults?: (ExecResult | null)[];
}

function emptyState(): PersistedRmdState {
  return { version: 2, codeCellStates: [] };
}

export function splitRmdSourceAndState(text: string): {
  source: string;
  state: PersistedRmdState;
} {
  const match = text.match(STATE_BLOCK_RE);
  if (!match || match.index == null) {
    return { source: text, state: emptyState() };
  }

  const source = text.slice(0, match.index);
  try {
    const decoded = Buffer.from(match[1], 'base64').toString('utf8');
    return { source, state: normalizePersistedState(JSON.parse(decoded)) };
  } catch {
    return { source, state: emptyState() };
  }
}

export function mergeRmdSourceAndState(
  source: string,
  state: PersistedRmdState,
): string {
  const base = splitRmdSourceAndState(source).source;
  const normalized = normalizePersistedState(state);
  if (!normalized.codeCellStates.some((entry) => entry.result)) return base;

  const payload = Buffer.from(JSON.stringify(normalized), 'utf8').toString('base64');
  const separator = base.length === 0 || base.endsWith('\n') ? '' : '\n';
  return `${base}${separator}<!-- ${STATE_MARKER} ${payload} -->\n`;
}

export function buildPersistedRmdStateFromChunks(
  chunks: readonly RmdChunk[],
  resultsByChunkId: ReadonlyMap<string, ExecResult>,
): PersistedRmdState {
  return {
    version: 2,
    codeCellStates: buildStoredStatesFromChunks(chunks, resultsByChunkId),
  };
}

export function restoreChunkResults(
  chunks: readonly RmdChunk[],
  state: PersistedRmdState,
): Map<string, ExecResult> {
  if (state.codeCellStates.length > 0) {
    return restoreChunkResultsWithStoredStates(chunks, state.codeCellStates);
  }

  const results = new Map<string, ExecResult>();
  let codeIndex = 0;
  for (const chunk of chunks) {
    if (chunk.kind !== 'code') continue;
    const result = cloneExecResult(state.codeCellResults?.[codeIndex++] ?? null);
    if (result) results.set(chunk.id, result);
  }
  return results;
}

export function normalizePersistedExecResult(
  value: ExecResult | null | undefined,
): ExecResult | null {
  if (!value || !isExecResult(value) || !hasRenderableExecResult(value)) return null;
  return {
    ...value,
    console: value.console ?? '',
    console_segments: value.console_segments?.map((segment) => ({
      code: segment.code,
      output: segment.output ?? '',
    })),
    plots: [...value.plots],
    plots_html: [...(value.plots_html ?? [])],
    dataframes: [...value.dataframes],
    output_order: value.output_order?.map((item) => ({ ...item })),
  };
}

export function hasRenderableExecResult(
  result: ExecResult | null | undefined,
): result is ExecResult {
  return !!result && Boolean(
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

function normalizePersistedState(value: unknown): PersistedRmdState {
  if (!value || typeof value !== 'object') return emptyState();
  const raw = value as Partial<PersistedRmdState>;
  const codeCellStates = Array.isArray(raw.codeCellStates)
    ? raw.codeCellStates
      .map((entry) => normalizeStoredExecResultState(entry))
      .filter((entry): entry is StoredExecResultState => entry !== null)
    : [];
  const codeCellResults = Array.isArray(raw.codeCellResults)
    ? raw.codeCellResults.map(item => normalizePersistedExecResult(item as ExecResult | null))
    : [];
  return {
    version: 2,
    codeCellStates,
    codeCellResults,
  };
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
