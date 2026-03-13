// =============================================================================
// rmdParser.ts — Parse .Rmd documents into structured chunks
// =============================================================================

export type ChunkLanguage = 'r' | 'python' | 'bash' | 'sql' | string;
export type ChunkOptionStyle = 'rmd' | 'quarto';

export interface ChunkOptions {
  label?: string;
  eval?: boolean;
  echo?: boolean;
  include?: boolean;
  fig_width?: number;
  fig_height?: number;
  cache?: boolean;
  message?: boolean;
  warning?: boolean;
  [key: string]: unknown;
}

export type ChunkKind = 'code' | 'prose' | 'yaml_frontmatter';

export interface RmdChunk {
  id: string;              // unique stable id: "chunk-<index>"
  kind: ChunkKind;
  language?: ChunkLanguage;
  options: ChunkOptions;
  optionStyle?: ChunkOptionStyle;
  code: string;            // raw code content (empty for prose)
  prose: string;           // markdown text (empty for code chunks)
  /** 0-based line numbers in the original document */
  startLine: number;
  endLine: number;
}

// ---------------------------------------------------------------------------

const FENCE_RE    = /^```\{(\w+)([^}]*)\}\s*$/;
const FENCE_END   = /^```\s*$/;
const YAML_START  = /^---\s*$/;
const YAML_END    = /^(---|\.\.\.)\s*$/;
const QUARTO_OPTION_RE = /^\s*#\|\s*([A-Za-z0-9_.-]+)\s*:\s*(.*?)\s*$/;

function parseOptions(optStr: string): ChunkOptions {
  const opts: ChunkOptions = {};
  if (!optStr.trim()) return opts;

  // label is first positional element (no `=`)
  const parts = optStr.split(',').map(s => s.trim()).filter(Boolean);
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!part.includes('=')) {
      if (i === 0) opts.label = part;
      continue;
    }
    const eq = part.indexOf('=');
    const key = normalizeOptionKey(part.slice(0, eq).trim(), 'rmd');
    const raw = part.slice(eq + 1).trim();
    opts[key] = parseOptionValue(raw);
  }
  return opts;
}

function parseQuartoOptions(codeLines: string[]): { codeLines: string[]; options: ChunkOptions; optionStyle?: ChunkOptionStyle } {
  const options: ChunkOptions = {};
  let index = 0;

  while (index < codeLines.length) {
    const match = QUARTO_OPTION_RE.exec(codeLines[index]);
    if (!match) break;
    const key = normalizeOptionKey(match[1], 'quarto');
    options[key] = parseOptionValue(match[2]);
    index++;
  }

  if (index === 0) {
    return { codeLines, options, optionStyle: undefined };
  }

  if (codeLines[index] === '') {
    index++;
  }

  return {
    codeLines: codeLines.slice(index),
    options,
    optionStyle: 'quarto',
  };
}

function normalizeOptionKey(key: string, style: ChunkOptionStyle): string {
  return style === 'quarto'
    ? key.trim().replace(/[.-]/g, '_')
    : key.trim().replace(/\./g, '_');
}

function parseOptionValue(raw: string): unknown {
  const trimmed = raw.trim();
  if (trimmed === 'TRUE' || trimmed === 'true') return true;
  if (trimmed === 'FALSE' || trimmed === 'false') return false;
  if (trimmed === 'NULL' || trimmed === 'null') return null;
  if (trimmed !== '' && !isNaN(Number(trimmed))) return Number(trimmed);
  return trimmed.replace(/^['"]|['"]$/g, '');
}

function stringifyRmdOptionValue(value: unknown): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value == null) return 'NULL';
  return String(value);
}

function stringifyQuartoOptionValue(value: unknown): string {
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  if (value == null) return 'null';
  return String(value);
}

export function formatCodeChunk(
  language: string,
  options: Record<string, unknown>,
  code: string,
  optionStyle: ChunkOptionStyle = 'rmd',
): string {
  const header = `\`\`\`{${language}}`;
  if (optionStyle === 'quarto') {
    const optionLines: string[] = [];
    if (options.label != null) {
      optionLines.push(`#| label: ${stringifyQuartoOptionValue(options.label)}`);
    }
    for (const [key, value] of Object.entries(options)) {
      if (key === 'label' || value === undefined) continue;
      const quartoKey = key.replace(/_/g, '-');
      optionLines.push(`#| ${quartoKey}: ${stringifyQuartoOptionValue(value)}`);
    }
    const body = [...optionLines, code].filter((line, index, lines) => !(line === '' && index === lines.length - 1)).join('\n');
    return `${header}\n${body}\n\`\`\``;
  }

  const optParts: string[] = [];
  if (options.label != null) optParts.push(String(options.label));
  for (const [key, value] of Object.entries(options)) {
    if (key === 'label' || value === undefined) continue;
    const rKey = key.replace(/_/g, '.');
    optParts.push(`${rKey}=${stringifyRmdOptionValue(value)}`);
  }
  const rmdHeader = `\`\`\`{${language}${optParts.length ? ' ' + optParts.join(', ') : ''}}`;
  return `${rmdHeader}\n${code}\n\`\`\``;
}

export function parseRmd(text: string): RmdChunk[] {
  const lines  = text.split('\n');
  const chunks: RmdChunk[] = [];
  let idx = 0;
  let chunkIndex = 0;

  const makeId = () => `chunk-${chunkIndex++}`;

  // ---- YAML front matter --------------------------------------------------
  if (YAML_START.test(lines[0])) {
    const start = 0;
    let end = 1;
    while (end < lines.length && !YAML_END.test(lines[end])) end++;
    chunks.push({
      id: makeId(),
      kind: 'yaml_frontmatter',
      options: {},
      code: lines.slice(start + 1, end).join('\n'),
      prose: '',
      startLine: start,
      endLine: end,
    });
    idx = end + 1;
  }

  let proseStart = idx;
  let proseLines: string[] = [];

  while (idx < lines.length) {
    const line = lines[idx];
    const fenceMatch = FENCE_RE.exec(line);

    if (fenceMatch) {
      // Flush accumulated prose (skip if entirely whitespace)
      const proseText = proseLines.join('\n');
      if (proseLines.length > 0 && proseText.trim()) {
        chunks.push({
          id: makeId(),
          kind: 'prose',
          options: {},
          code: '',
          prose: proseText,
          startLine: proseStart,
          endLine: idx - 1,
        });
        proseLines = [];
      }

      const lang    = fenceMatch[1].toLowerCase() as ChunkLanguage;
      const optStr  = fenceMatch[2];
      const options = parseOptions(optStr);
      const codeStart = idx;
      idx++;

      const codeLines: string[] = [];
      while (idx < lines.length && !FENCE_END.test(lines[idx])) {
        codeLines.push(lines[idx]);
        idx++;
      }
      const codeEnd = idx;
      const quartoOptions = parseQuartoOptions(codeLines);
      const mergedOptions = {
        ...options,
        ...quartoOptions.options,
      };
      const optionStyle = quartoOptions.optionStyle ?? 'rmd';

      chunks.push({
        id: makeId(),
        kind: 'code',
        language: lang,
        options: mergedOptions,
        optionStyle,
        code: quartoOptions.codeLines.join('\n'),
        prose: '',
        startLine: codeStart,
        endLine: codeEnd,
      });

      idx++;   // skip closing ```
      proseStart = idx;
    } else {
      proseLines.push(line);
      idx++;
    }
  }

  // Remaining prose (skip if entirely whitespace)
  const remainingProse = proseLines.join('\n');
  if (proseLines.length > 0 && remainingProse.trim()) {
    chunks.push({
      id: makeId(),
      kind: 'prose',
      options: {},
      code: '',
      prose: remainingProse,
      startLine: proseStart,
      endLine: lines.length - 1,
    });
  }

  return chunks;
}

/** Reconstruct document text from chunks (round-trip) */
export function chunksToText(chunks: RmdChunk[]): string {
  return chunks.map(c => {
    if (c.kind === 'yaml_frontmatter') {
      return `---\n${c.code}\n---`;
    }
    if (c.kind === 'prose') {
      return c.prose;
    }
    return formatCodeChunk(c.language || 'r', c.options, c.code, c.optionStyle ?? 'rmd');
  }).join('\n');
}
