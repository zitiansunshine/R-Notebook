// =============================================================================
// rmdParser.test.ts — Vitest unit tests for the RMarkdown chunk parser
// Run: npx vitest run  (no VSCode / R needed)
// =============================================================================

import { describe, it, expect } from 'vitest';
import { parseRmd, chunksToText, RmdChunk } from '../src/rmdParser';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const RMD_BASIC = `---
title: "Test"
output: html_document
---

Some **prose** here.

\`\`\`{r setup, message=FALSE}
library(dplyr)
\`\`\`

More prose.

\`\`\`{r plot, fig.width=7, eval=FALSE}
plot(mtcars)
\`\`\`
`;

// ---------------------------------------------------------------------------
// YAML frontmatter
// ---------------------------------------------------------------------------

describe('YAML frontmatter', () => {
  it('parses frontmatter as the first chunk', () => {
    const chunks = parseRmd(RMD_BASIC);
    expect(chunks[0].kind).toBe('yaml_frontmatter');
    expect(chunks[0].code).toContain('title: "Test"');
    expect(chunks[0].startLine).toBe(0);
  });

  it('handles document with no frontmatter', () => {
    const chunks = parseRmd('Just prose\n\n```{r}\n1+1\n```\n');
    expect(chunks[0].kind).toBe('prose');
  });
});

// ---------------------------------------------------------------------------
// Prose chunks
// ---------------------------------------------------------------------------

describe('Prose chunks', () => {
  it('captures prose between code chunks', () => {
    const chunks = parseRmd(RMD_BASIC);
    const prose  = chunks.filter(c => c.kind === 'prose');
    expect(prose.length).toBeGreaterThanOrEqual(2);
    expect(prose.some(p => p.prose.includes('Some **prose**'))).toBe(true);
    expect(prose.some(p => p.prose.includes('More prose'))).toBe(true);
  });

  it('assigns unique IDs to each prose chunk', () => {
    const chunks = parseRmd(RMD_BASIC);
    const ids    = chunks.map(c => c.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

// ---------------------------------------------------------------------------
// Code chunks
// ---------------------------------------------------------------------------

describe('Code chunks', () => {
  it('parses language correctly', () => {
    const chunks = parseRmd(RMD_BASIC);
    const codeChunks = chunks.filter(c => c.kind === 'code');
    expect(codeChunks.every(c => c.language === 'r')).toBe(true);
  });

  it('parses chunk label from first positional option', () => {
    const chunks = parseRmd(RMD_BASIC);
    const setup  = chunks.find(c => c.kind === 'code' && c.options.label === 'setup');
    expect(setup).toBeTruthy();
  });

  it('parses boolean option eval=FALSE', () => {
    const chunks = parseRmd(RMD_BASIC);
    const plot   = chunks.find(c => c.kind === 'code' && c.options.label === 'plot');
    expect(plot?.options.eval).toBe(false);
  });

  it('parses numeric option fig.width', () => {
    const chunks = parseRmd(RMD_BASIC);
    const plot   = chunks.find(c => c.kind === 'code' && c.options.label === 'plot');
    expect(plot?.options.fig_width).toBe(7);
  });

  it('parses boolean option message=FALSE', () => {
    const chunks = parseRmd(RMD_BASIC);
    const setup  = chunks.find(c => c.kind === 'code' && c.options.label === 'setup');
    expect(setup?.options.message).toBe(false);
  });

  it('captures code body correctly', () => {
    const chunks = parseRmd(RMD_BASIC);
    const setup  = chunks.find(c => c.kind === 'code' && c.options.label === 'setup');
    expect(setup?.code.trim()).toBe('library(dplyr)');
  });

  it('records correct start/end line numbers', () => {
    const chunks = parseRmd(RMD_BASIC);
    const setup  = chunks.find(c => c.kind === 'code' && c.options.label === 'setup');
    expect(setup?.startLine).toBeGreaterThan(0);
    expect(setup?.endLine).toBeGreaterThan(setup!.startLine);
  });
});

// ---------------------------------------------------------------------------
// Multiple languages
// ---------------------------------------------------------------------------

describe('Multiple languages', () => {
  const MULTI = `\`\`\`{python}\nprint("hi")\n\`\`\`\n\`\`\`{bash}\nls\n\`\`\`\n`;

  it('parses python chunk language', () => {
    const chunks = parseRmd(MULTI);
    const py = chunks.find(c => c.kind === 'code' && c.language === 'python');
    expect(py).toBeTruthy();
    expect(py?.code.trim()).toBe('print("hi")');
  });

  it('parses bash chunk language', () => {
    const chunks = parseRmd(MULTI);
    const sh = chunks.find(c => c.kind === 'code' && c.language === 'bash');
    expect(sh).toBeTruthy();
    expect(sh?.code.trim()).toBe('ls');
  });
});

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------

describe('Edge cases', () => {
  it('empty document returns empty array', () => {
    expect(parseRmd('')).toEqual([]);
  });

  it('document with only frontmatter returns one chunk', () => {
    const chunks = parseRmd('---\ntitle: "x"\n---\n');
    expect(chunks.length).toBe(1);
    expect(chunks[0].kind).toBe('yaml_frontmatter');
  });

  it('document with only prose returns one prose chunk', () => {
    const chunks = parseRmd('Hello world\n');
    expect(chunks.length).toBe(1);
    expect(chunks[0].kind).toBe('prose');
  });

  it('code chunk with no options parses correctly', () => {
    const chunks = parseRmd('```{r}\nx <- 1\n```\n');
    expect(chunks[0].kind).toBe('code');
    expect(chunks[0].language).toBe('r');
    expect(chunks[0].options.label).toBeUndefined();
  });

  it('handles chunk with only label (no other options)', () => {
    const chunks = parseRmd('```{r myChunk}\nx\n```\n');
    expect(chunks[0].options.label).toBe('myChunk');
  });

  it('handles multiline code blocks correctly', () => {
    const code   = 'x <- 1\ny <- 2\nz <- x + y\nprint(z)';
    const chunks = parseRmd(`\`\`\`{r}\n${code}\n\`\`\`\n`);
    expect(chunks[0].code).toBe(code);
  });

  it('does not treat indented backticks as chunk fence', () => {
    const doc = 'Prose with `inline code` here.\n\n```{r}\n1+1\n```\n';
    const chunks = parseRmd(doc);
    const prose  = chunks.find(c => c.kind === 'prose');
    expect(prose?.prose).toContain('`inline code`');
  });
});

// ---------------------------------------------------------------------------
// Round-trip: parseRmd → chunksToText
// ---------------------------------------------------------------------------

describe('Round-trip reconstruction', () => {
  it('reconstructs a document close to the original', () => {
    const original = RMD_BASIC;
    const chunks   = parseRmd(original);
    const rebuilt  = chunksToText(chunks);

    // Should contain all the key content
    expect(rebuilt).toContain('library(dplyr)');
    expect(rebuilt).toContain('plot(mtcars)');
    expect(rebuilt).toContain('Some **prose**');
  });

  it('round-trips chunk options', () => {
    const doc    = '```{r myLabel, eval=FALSE, fig.width=5}\nplot(x)\n```\n';
    const chunks = parseRmd(doc);
    const text   = chunksToText(chunks);
    expect(text).toContain('myLabel');
    expect(text).toContain('eval=FALSE');
    expect(text).toContain('fig.width=5');
  });

  it('round-trips YAML frontmatter', () => {
    const doc    = '---\ntitle: "Hello"\noutput: html_document\n---\n\nProse here.\n';
    const chunks = parseRmd(doc);
    const text   = chunksToText(chunks);
    expect(text).toContain('title: "Hello"');
    expect(text).toContain('Prose here.');
  });

  it('chunk count is preserved after round-trip', () => {
    const chunks1 = parseRmd(RMD_BASIC);
    const rebuilt = chunksToText(chunks1);
    const chunks2 = parseRmd(rebuilt);
    const codes1  = chunks1.filter(c => c.kind === 'code').map(c => c.code);
    const codes2  = chunks2.filter(c => c.kind === 'code').map(c => c.code);
    expect(codes2).toEqual(codes1);
  });
});

// ---------------------------------------------------------------------------
// ID stability
// ---------------------------------------------------------------------------

describe('Chunk IDs', () => {
  it('all IDs are non-empty strings', () => {
    const chunks = parseRmd(RMD_BASIC);
    for (const c of chunks) {
      expect(typeof c.id).toBe('string');
      expect(c.id.length).toBeGreaterThan(0);
    }
  });

  it('IDs are sequential chunk-N', () => {
    const chunks = parseRmd(RMD_BASIC);
    chunks.forEach((c, i) => {
      expect(c.id).toBe(`chunk-${i}`);
    });
  });
});
