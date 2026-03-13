// =============================================================================
// rmdParser.ts — Parse .Rmd documents into structured chunks
// =============================================================================

export type ChunkLanguage = 'r' | 'python' | 'bash' | 'sql' | string;

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
    const key = part.slice(0, eq).trim().replace(/\./g, '_');
    const raw = part.slice(eq + 1).trim();
    // coerce common R values
    if (raw === 'TRUE' || raw === 'true')   opts[key] = true;
    else if (raw === 'FALSE' || raw === 'false') opts[key] = false;
    else if (!isNaN(Number(raw)))            opts[key] = Number(raw);
    else                                     opts[key] = raw.replace(/^['"]|['"]$/g, '');
  }
  return opts;
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

      chunks.push({
        id: makeId(),
        kind: 'code',
        language: lang,
        options,
        code: codeLines.join('\n'),
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
    // code chunk
    const optParts: string[] = [];
    if (c.options.label) optParts.push(c.options.label);
    for (const [k, v] of Object.entries(c.options)) {
      if (k === 'label') continue;
      const rKey = k.replace(/_/g, '.');
      if (typeof v === 'boolean') optParts.push(`${rKey}=${v ? 'TRUE' : 'FALSE'}`);
      else optParts.push(`${rKey}=${v}`);
    }
    const header = `\`\`\`{${c.language}${optParts.length ? ' ' + optParts.join(', ') : ''}}`;
    return `${header}\n${c.code}\n\`\`\``;
  }).join('\n');
}
