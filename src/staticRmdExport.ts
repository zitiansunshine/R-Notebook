import * as path from 'path';

import { ExecResult } from './kernelProtocol';
import { buildOutputFragment, esc, NOTEBOOK_OUTPUT_CSS } from './notebookOutputHtml';
import { formatCodeChunk, parseRmd, RmdChunk } from './rmdParser';
import { restoreChunkResults, splitRmdSourceAndState } from './rmdPersistedState';

export function buildStaticRmdExportHtml(
  text: string,
  documentName: string,
): string {
  const { source, state } = splitRmdSourceAndState(text);
  const chunks = parseRmd(source);
  const results = restoreChunkResults(chunks, state);
  const title = extractDocumentTitle(chunks, documentName);
  const body = chunks
    .map((chunk) => renderChunk(chunk, results.get(chunk.id) ? chunk.id : undefined, results.get(chunk.id) ?? null))
    .join('\n');

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${esc(title)}</title>
  <style>
    :root {
      --vscode-font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      --vscode-editor-font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
      --vscode-editor-font-size: 13px;
      --vscode-editor-background: #ffffff;
      --vscode-editor-foreground: #1f2328;
      --vscode-editorGroup-border: #d0d7de;
      --vscode-descriptionForeground: #59636e;
    }
    body {
      margin: 0;
      background: #f6f8fa;
      color: #1f2328;
      font-family: var(--vscode-font-family);
      line-height: 1.55;
    }
    .export-doc {
      max-width: 1040px;
      margin: 0 auto;
      padding: 40px 24px 64px;
    }
    .export-title {
      margin: 0 0 28px;
      font-size: 34px;
      line-height: 1.15;
      font-weight: 700;
    }
    .export-prose {
      margin: 0 0 22px;
      font-size: 16px;
    }
    .export-prose p {
      margin: 0 0 16px;
    }
    .export-prose h1,
    .export-prose h2,
    .export-prose h3,
    .export-prose h4,
    .export-prose h5,
    .export-prose h6 {
      margin: 24px 0 12px;
      line-height: 1.2;
    }
    .export-prose blockquote {
      margin: 16px 0;
      padding: 0 0 0 14px;
      border-left: 3px solid #d0d7de;
      color: #59636e;
    }
    .export-cell {
      margin: 0 0 28px;
      border: 1px solid #d0d7de;
      border-radius: 10px;
      background: #ffffff;
      overflow: hidden;
      box-shadow: 0 1px 2px rgba(31, 35, 40, 0.06);
    }
    .export-cell-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 10px 14px;
      border-bottom: 1px solid #d0d7de;
      background: #f6f8fa;
      font-size: 12px;
      color: #59636e;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }
    .export-cell-kind {
      font-weight: 700;
      color: #24292f;
    }
    .export-cell-meta {
      text-transform: none;
      letter-spacing: normal;
      font-weight: 500;
    }
    .export-code {
      margin: 0;
      padding: 14px 16px;
      overflow-x: auto;
      background: #ffffff;
      color: #1f2328;
      font-family: var(--vscode-editor-font-family);
      font-size: 13px;
      line-height: 1.55;
      white-space: pre;
    }
    .export-output {
      padding: 14px 16px 16px;
      border-top: 1px solid #d0d7de;
      background: #fbfcfd;
    }
    .export-empty-note {
      color: #59636e;
      font-size: 13px;
    }
    ${NOTEBOOK_OUTPUT_CSS}
  </style>
</head>
<body>
  <main class="export-doc">
    <h1 class="export-title">${esc(title)}</h1>
    ${body || '<p class="export-empty-note">This document is empty.</p>'}
  </main>
</body>
</html>`;
}

function renderChunk(
  chunk: RmdChunk,
  chunkId: string | undefined,
  result: ExecResult | null,
): string {
  if (chunk.kind === 'prose') {
    return `<section class="export-prose">${renderMarkdown(chunk.prose)}</section>`;
  }

  if (chunk.kind === 'yaml_frontmatter') {
    return buildCodeSection('yaml', 'Front Matter', `---\n${chunk.code}\n---`);
  }

  const codeText = formatCodeChunk(
    chunk.language || 'r',
    chunk.options,
    chunk.code,
    chunk.optionStyle ?? 'rmd',
  );
  const label = typeof chunk.options.label === 'string' && chunk.options.label
    ? chunk.options.label
    : undefined;
  const outputHtml = chunkId && result
    ? buildOutputFragment(result, chunkId) ?? ''
    : '';

  return `<section class="export-cell">
    <div class="export-cell-header">
      <span class="export-cell-kind">${esc((chunk.language || 'r').toUpperCase())} chunk</span>
      <span class="export-cell-meta">${label ? esc(label) : '&nbsp;'}</span>
    </div>
    <pre class="export-code"><code>${esc(codeText)}</code></pre>
    ${outputHtml ? `<div class="export-output">${outputHtml}</div>` : ''}
  </section>`;
}

function buildCodeSection(language: string, label: string, value: string): string {
  return `<section class="export-cell">
    <div class="export-cell-header">
      <span class="export-cell-kind">${esc(language.toUpperCase())}</span>
      <span class="export-cell-meta">${esc(label)}</span>
    </div>
    <pre class="export-code"><code>${esc(value)}</code></pre>
  </section>`;
}

function extractDocumentTitle(chunks: readonly RmdChunk[], documentName: string): string {
  const yaml = chunks.find((chunk) => chunk.kind === 'yaml_frontmatter');
  if (yaml) {
    const match = yaml.code.match(/^\s*title\s*:\s*(.+?)\s*$/mi);
    if (match) {
      return cleanYamlScalar(match[1]);
    }
  }
  return path.parse(documentName).name;
}

function cleanYamlScalar(value: string): string {
  return value.trim().replace(/^['"]|['"]$/g, '');
}

function renderMarkdown(text: string): string {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/^###### (.+)$/gm, '<h6>$1</h6>')
    .replace(/^##### (.+)$/gm, '<h5>$1</h5>')
    .replace(/^#### (.+)$/gm, '<h4>$1</h4>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    .replace(/`(.+?)`/g, '<code>$1</code>')
    .replace(/^> (.+)$/gm, '<blockquote>$1</blockquote>')
    .replace(/^---$/gm, '<hr>')
    .replace(/\n\n/g, '</p><p>')
    .replace(/^/, '<p>')
    .replace(/$/, '</p>');
}
