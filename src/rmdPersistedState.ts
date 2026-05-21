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
const STATE_BEGIN_MARKER = `${STATE_MARKER}_BEGIN`;
const STATE_END_MARKER = `${STATE_MARKER}_END`;
const LEGACY_STATE_MARKERS = [
  STATE_MARKER,
  'R_NOTEBOOK_RMD_STATE',
  'NOTA_RMD_STATE',
] as const;
const STATE_BLOCK_RE = new RegExp(
  String.raw`(?:\r?\n)?<!--\s*R Notebook saved outputs: do not edit\s*\r?\n\s*${escapeRegExp(STATE_BEGIN_MARKER)}\s*\r?\n\s*([A-Za-z0-9+/=\s]+?)\s*\r?\n\s*${escapeRegExp(STATE_END_MARKER)}\s*\r?\n\s*-->\s*$`,
);
const LEGACY_STATE_BLOCK_RE = new RegExp(
  String.raw`(?:\r?\n)?<!--\s*(${LEGACY_STATE_MARKERS.map(escapeRegExp).join('|')})(?:\s+([A-Za-z0-9+/=]+))?\s*-->\s*$`,
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
  const { source, payload } = stripTrailingStateBlocks(text);
  if (!payload) return { source, state: emptyState() };

  try {
    const decoded = Buffer.from(payload, 'base64').toString('utf8');
    return { source, state: normalizePersistedState(JSON.parse(decoded)) };
  } catch {
    return { source, state: emptyState() };
  }
}

export function mergeRmdSourceAndState(
  source: string,
  state: PersistedRmdState,
): string {
  const { source: base } = stripTrailingStateBlocks(source);
  const normalized = normalizePersistedState(state);
  if (!normalized.codeCellStates.some((entry) => entry.result)) return base;

  const payload = Buffer.from(JSON.stringify(normalized), 'utf8').toString('base64');
  const separator = base.length === 0 || base.endsWith('\n') ? '' : '\n';
  return `${base}${separator}<!-- R Notebook saved outputs: do not edit
${STATE_BEGIN_MARKER}
${wrapBase64(payload)}
${STATE_END_MARKER}
-->
`;
}

function stripTrailingStateBlocks(text: string): { source: string; payload?: string } {
  let source = text;
  let payload: string | undefined;

  while (true) {
    const match = source.match(STATE_BLOCK_RE);
    if (match && match.index != null) {
      if (!payload && match[1]) payload = match[1].replace(/\s+/g, '');
      source = source.slice(0, match.index);
      continue;
    }

    const legacyMatch = source.match(LEGACY_STATE_BLOCK_RE);
    if (!legacyMatch || legacyMatch.index == null) return { source, payload };
    if (!payload && legacyMatch[2]) payload = legacyMatch[2];
    source = source.slice(0, legacyMatch.index);
  }
}

function wrapBase64(payload: string): string {
  return payload.match(/.{1,100}/g)?.join('\n') ?? payload;
}

export function hasPersistedRmdState(text: string): boolean {
  const { source } = stripTrailingStateBlocks(text);
  return source !== text;
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
    (result.source_code?.trim() ?? '') ||
    (result.console_segments?.some(segment =>
      Boolean(segment?.code?.trim() || segment?.output?.trim()),
    ) ?? false) ||
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
    (candidate.error_trace === undefined || candidate.error_trace === null || typeof candidate.error_trace === 'string') &&
    Array.isArray(candidate.plots) &&
    Array.isArray(candidate.dataframes) &&
    (candidate.plots_html === undefined || Array.isArray(candidate.plots_html))
  );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
