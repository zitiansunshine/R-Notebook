import { ExecResult } from './kernelProtocol';
import { RmdChunk } from './rmdParser';

export interface StoredExecResultState {
  fingerprint: string;
  result: ExecResult | null;
}

export function buildStoredStatesFromChunks(
  chunks: readonly RmdChunk[],
  resultsByChunkId: ReadonlyMap<string, ExecResult>,
): StoredExecResultState[] {
  const states: StoredExecResultState[] = [];
  for (const chunk of chunks) {
    if (chunk.kind !== 'code') continue;
    states.push({
      fingerprint: fingerprintCodeCell(chunk.language, chunk.code, chunk.options),
      result: cloneExecResult(resultsByChunkId.get(chunk.id)),
    });
  }
  trimTrailingEmptyStates(states);
  return states;
}

export function fingerprintNotebookCodeCell(
  languageId: string,
  code: string,
  options?: Record<string, unknown>,
): string {
  return fingerprintCodeCell(languageId, code, options);
}

export function fingerprintChunkCodeCell(chunk: Pick<RmdChunk, 'kind' | 'language' | 'code' | 'options'>): string {
  return chunk.kind === 'code'
    ? fingerprintCodeCell(chunk.language, chunk.code, chunk.options)
    : '';
}

export function restoreStoredResults<K>(
  currentCells: readonly { key: K; fingerprint: string }[],
  storedStates: readonly StoredExecResultState[],
): Map<K, ExecResult> {
  const restored = new Map<K, ExecResult>();
  if (currentCells.length === 0 || storedStates.length === 0) return restored;

  const matches = longestCommonFingerprintMatches(
    storedStates.map((state) => state.fingerprint),
    currentCells.map((cell) => cell.fingerprint),
  );

  let storedStart = 0;
  let currentStart = 0;

  for (const match of [...matches, { storedIndex: storedStates.length, currentIndex: currentCells.length }]) {
    const storedCount = match.storedIndex - storedStart;
    const currentCount = match.currentIndex - currentStart;

    // Preserve outputs through in-place rewrites, but never across insert/delete shifts.
    if (storedCount > 0 && storedCount === currentCount) {
      for (let offset = 0; offset < storedCount; offset += 1) {
        const result = cloneExecResult(storedStates[storedStart + offset]?.result);
        if (result) restored.set(currentCells[currentStart + offset].key, result);
      }
    }

    if (match.storedIndex < storedStates.length && match.currentIndex < currentCells.length) {
      const result = cloneExecResult(storedStates[match.storedIndex]?.result);
      if (result) restored.set(currentCells[match.currentIndex].key, result);
    }

    storedStart = match.storedIndex + 1;
    currentStart = match.currentIndex + 1;
  }

  return restored;
}

export function cloneExecResult(value: ExecResult | null | undefined): ExecResult | null {
  if (!value || !hasRenderableExecResult(value)) return null;
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

export function restoreChunkResultsWithStoredStates(
  chunks: readonly RmdChunk[],
  storedStates: readonly StoredExecResultState[],
): Map<string, ExecResult> {
  const currentCells = chunks
    .filter((chunk) => chunk.kind === 'code')
    .map((chunk) => ({
      key: chunk.id,
      fingerprint: fingerprintChunkCodeCell(chunk),
    }));
  return restoreStoredResults(currentCells, storedStates);
}

export function normalizeStoredExecResultState(value: unknown): StoredExecResultState | null {
  if (!value || typeof value !== 'object') return null;
  const raw = value as Partial<StoredExecResultState>;
  if (typeof raw.fingerprint !== 'string') return null;
  return {
    fingerprint: raw.fingerprint,
    result: cloneExecResult(raw.result as ExecResult | null | undefined),
  };
}

function fingerprintCodeCell(
  language: string | undefined,
  code: string,
  options?: Record<string, unknown>,
): string {
  return JSON.stringify({
    language: (language || 'r').toLowerCase(),
    code,
    options: normalizeValue(options ?? {}),
  });
}

function normalizeValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeValue(entry));
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([, entry]) => entry !== undefined)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, normalizeValue(entry)]),
    );
  }
  return value;
}

function longestCommonFingerprintMatches(
  previous: readonly string[],
  current: readonly string[],
): { storedIndex: number; currentIndex: number }[] {
  const rows = previous.length;
  const cols = current.length;
  const dp: number[][] = Array.from({ length: rows + 1 }, () => Array(cols + 1).fill(0));

  for (let row = rows - 1; row >= 0; row -= 1) {
    for (let col = cols - 1; col >= 0; col -= 1) {
      if (previous[row] === current[col]) {
        dp[row][col] = dp[row + 1][col + 1] + 1;
      } else {
        dp[row][col] = Math.max(dp[row + 1][col], dp[row][col + 1]);
      }
    }
  }

  const matches: { storedIndex: number; currentIndex: number }[] = [];
  let row = 0;
  let col = 0;

  while (row < rows && col < cols) {
    if (previous[row] === current[col]) {
      matches.push({ storedIndex: row, currentIndex: col });
      row += 1;
      col += 1;
      continue;
    }
    if (dp[row + 1][col] >= dp[row][col + 1]) {
      row += 1;
    } else {
      col += 1;
    }
  }

  return matches;
}

function hasRenderableExecResult(result: ExecResult | null | undefined): result is ExecResult {
  return !!result && Boolean(
    (result.console?.trim() ?? '') ||
    result.stdout.trim() ||
    result.stderr.trim() ||
    result.plots.length ||
    (result.plots_html?.length ?? 0) ||
    result.dataframes.length ||
    result.error,
  );
}

function trimTrailingEmptyStates(states: StoredExecResultState[]): void {
  while (states.length > 0 && !states[states.length - 1].result) states.pop();
}
