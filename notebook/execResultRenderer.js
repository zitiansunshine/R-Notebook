const STYLE_ID = 'r-notebook-output-renderer-style';
const TAB_STATE_PREFIX = 'rNotebook.tab.';
const PAGE_STATE_PREFIX = 'rNotebook.page.';
const CONSOLE_STATE_PREFIX = 'rNotebook.console.';
const globalTabState = globalThis.__rNotebookTabState || (globalThis.__rNotebookTabState = new Map());
const globalPageState = globalThis.__rNotebookPageState || (globalThis.__rNotebookPageState = new Map());
const globalConsoleState = globalThis.__rNotebookConsoleState || (globalThis.__rNotebookConsoleState = new Map());
const globalConsoleText = globalThis.__rNotebookConsoleText || (globalThis.__rNotebookConsoleText = new Map());

const STYLES = `
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body { border: none; outline: none; }
.rnb-output-root {
  font-family: var(--vscode-font-family, -apple-system, sans-serif);
  font-size: var(--vscode-editor-font-size, 13px);
  color: var(--vscode-editor-foreground);
  background: var(--vscode-editor-background);
  padding: 4px 0;
}
.output-text {
  background: var(--vscode-editor-background);
  color: var(--vscode-editor-foreground);
  padding: 8px 12px;
  border-radius: 3px;
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-word;
  margin-bottom: 6px;
}
.output-console {
  color: var(--vscode-editor-foreground);
  padding: 10px 12px;
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-word;
  margin: 0;
}
.console-viewer {
  border: 1px solid var(--vscode-editorGroup-border);
  border-radius: 4px;
  overflow: visible;
  margin-bottom: 10px;
  background: var(--vscode-editor-background);
}
.console-title {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 8px;
  padding: 5px 10px;
  border-bottom: 1px solid var(--vscode-editorGroup-border);
  font-size: 12px;
}
.console-meta {
  color: var(--vscode-descriptionForeground);
  font-size: 11px;
}
.console-open-tab-btn {
  background: var(--vscode-button-background, #0078d4);
  border: 1px solid var(--vscode-button-background, #0078d4);
  color: var(--vscode-button-foreground, #fff);
  padding: 2px 8px;
  font-size: 11px;
  cursor: pointer;
  border-radius: 2px;
  margin-left: auto;
}
.console-open-tab-btn:hover,
.console-open-tab-btn:focus,
.console-open-tab-btn:active {
  background: var(--vscode-button-background, #0078d4);
  border-color: var(--vscode-button-background, #0078d4);
  color: var(--vscode-button-foreground, #fff);
}
.console-open-tab-btn:focus {
  outline: 1px solid var(--vscode-focusBorder, #0078d4);
  outline-offset: 1px;
}
.console-scroll {
  max-height: 40em;
  overflow: auto;
}
.console-body {
  min-height: 100%;
}
.console-input {
  background: rgba(0, 120, 212, .08);
  border-bottom: 1px solid rgba(0, 120, 212, .2);
}
.output-console-code {
  color: var(--vscode-textLink-foreground, #4f8cc9);
  padding: 10px 12px 8px;
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-word;
  margin: 0;
}
.output-console.has-source {
  padding-top: 8px;
}
.output-stderr {
  background: var(--vscode-inputValidation-warningBackground);
  color: var(--vscode-inputValidation-warningForeground, #e5c07b);
  padding: 6px 10px;
  border-radius: 3px;
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 12px;
  white-space: pre-wrap;
  margin-bottom: 6px;
}
.output-error {
  background: var(--vscode-inputValidation-errorBackground);
  color: var(--vscode-errorForeground, #e06c75);
  padding: 6px 10px;
  border-radius: 3px;
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 12px;
  white-space: pre-wrap;
  margin-bottom: 6px;
}
.output-html {
  width: 100%;
  overflow: auto;
  margin-bottom: 8px;
}
.plot-wrap {
  position: relative;
  display: inline-block;
  max-width: 100%;
  margin-bottom: 8px;
}
.output-plot {
  display: block;
  max-width: 100%;
  border-radius: 3px;
}
.plot-dl-btn {
  position: absolute;
  top: 6px;
  right: 6px;
  padding: 3px 8px;
  background: rgba(0, 0, 0, .65);
  color: #fff;
  border-radius: 3px;
  font-size: 11px;
  text-decoration: none;
  opacity: 0;
  transition: opacity .15s;
  cursor: pointer;
  white-space: nowrap;
}
.plot-wrap:hover .plot-dl-btn { opacity: 1; }
.df-viewer {
  border: 1px solid var(--vscode-editorGroup-border);
  border-radius: 4px;
  overflow: hidden;
  margin-bottom: 10px;
  font-size: 12px;
  background: var(--vscode-editor-background);
}
.df-title {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  flex-wrap: wrap;
  padding: 5px 10px;
  border-bottom: 1px solid var(--vscode-editorGroup-border);
}
.df-name {
  font-weight: 600;
  overflow-wrap: anywhere;
  word-break: break-word;
}
.df-dims { color: var(--vscode-descriptionForeground); font-size: 11px; }
.df-info { color: var(--vscode-descriptionForeground); font-size: 11px; margin-left: auto; white-space: nowrap; }
.df-note {
  color: var(--vscode-descriptionForeground);
  font-size: 11px;
  margin-left: 10px;
  white-space: nowrap;
}
.df-table-wrap { overflow-x: auto; max-height: 300px; overflow-y: auto; }
.df-table { width: 100%; border-collapse: collapse; white-space: nowrap; }
.df-table thead th {
  position: sticky;
  top: 0;
  background: var(--vscode-editor-background);
  border-bottom: none;
  box-shadow: 0 2px 0 var(--vscode-editorGroup-border, #c0c0c0);
  border-right: 1px solid var(--vscode-editorGroup-border, #c0c0c0);
  padding: 4px 10px;
  text-align: left;
  font-weight: 600;
}
.df-table thead th:last-child { border-right: none; }
.df-table td {
  padding: 3px 10px;
  border-bottom: 1px solid var(--vscode-editorGroup-border);
  border-right: 1px solid var(--vscode-editorGroup-border, #c0c0c0);
}
.df-table td:last-child { border-right: none; }
.row-idx { color: var(--vscode-descriptionForeground); font-size: 11px; }
.na-value { color: var(--vscode-descriptionForeground); font-style: italic; }
.df-pager {
  display: flex;
  align-items: center;
  gap: 3px;
  flex-wrap: wrap;
  padding: 5px 8px;
  border-top: 1px solid var(--vscode-editorGroup-border);
  background: var(--vscode-editor-background);
}
.pg-btn {
  padding: 2px 7px;
  min-width: 28px;
  border: 1px solid var(--vscode-editorGroup-border, #c0c0c0);
  border-radius: 3px;
  background: transparent;
  color: var(--vscode-editor-foreground);
  cursor: pointer;
  font-size: 11px;
  font-family: var(--vscode-font-family, sans-serif);
  line-height: 1.4;
}
.pg-btn:hover:not(:disabled) {
  background: var(--vscode-toolbar-hoverBackground, rgba(128, 128, 128, .15));
}
.pg-btn.pg-active {
  background: var(--vscode-button-background, #0078d4);
  color: var(--vscode-button-foreground, #fff);
  border-color: var(--vscode-button-background, #0078d4);
}
.pg-btn.pg-dis,
.pg-btn:disabled { opacity: .38; cursor: default; }
.pg-ellipsis {
  font-size: 11px;
  padding: 0 2px;
  color: var(--vscode-descriptionForeground);
  line-height: 1.6;
}
.pg-jump-wrap {
  font-size: 11px;
  color: var(--vscode-descriptionForeground);
  margin-left: 6px;
  display: flex;
  align-items: center;
  gap: 4px;
}
.pg-input {
  width: 46px;
  padding: 2px 4px;
  border: 1px solid var(--vscode-input-border, var(--vscode-editorGroup-border));
  border-radius: 3px;
  background: var(--vscode-input-background, transparent);
  color: var(--vscode-input-foreground, var(--vscode-editor-foreground));
  font-size: 11px;
  text-align: center;
}
.pg-input::-webkit-inner-spin-button,
.pg-input::-webkit-outer-spin-button { -webkit-appearance: none; }
.output-thumb-strip {
  display: flex;
  flex-direction: row;
  gap: 8px;
  overflow-x: auto;
  padding: 8px 4px 8px 3px;
  margin-bottom: 8px;
  border-bottom: 1px solid var(--vscode-editorGroup-border, #c0c0c0);
  scrollbar-width: thin;
}
.output-thumb-strip.output-thumb-strip-grid {
  display: grid;
  grid-template-columns: repeat(7, 116px);
  grid-auto-rows: 78px;
  align-content: start;
  overflow: auto;
}
.output-thumb {
  appearance: none;
  flex-shrink: 0;
  width: 116px;
  height: 78px;
  border: 2px solid var(--vscode-editorGroup-border, #c0c0c0);
  border-radius: 4px;
  cursor: pointer;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  background: var(--vscode-editor-background);
  transition: border-color .12s, box-shadow .12s;
  user-select: none;
  padding: 0;
  color: inherit;
  font: inherit;
}
.output-thumb:hover {
  border-color: #0078d4;
  box-shadow: 0 0 0 2px rgba(0, 120, 212, .15);
}
.output-thumb.is-active {
  border-color: #0078d4;
  box-shadow: 0 0 0 3px rgba(0, 120, 212, .22);
}
.output-thumb.thumb-error { border-color: #e06c75; }
.output-thumb img { width: 100%; height: 100%; object-fit: contain; }
.thumb-icon {
  font-size: 22px;
  opacity: .5;
  line-height: 1;
  margin-bottom: 4px;
}
.thumb-err-icon { color: #e06c75; opacity: .8; }
.thumb-label {
  font-size: 10px;
  color: var(--vscode-descriptionForeground);
  text-align: center;
  padding: 0 4px;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.thumb-text-preview {
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 7.5px;
  color: var(--vscode-editor-foreground);
  flex: 1;
  padding: 4px 6px 2px;
  width: 100%;
  overflow: hidden;
  white-space: pre;
  opacity: .65;
}
.thumb-stderr { opacity: .55; color: var(--vscode-inputValidation-warningForeground, #e5c07b); }
.thumb-dims {
  font-size: 9px;
  color: var(--vscode-descriptionForeground);
  opacity: .7;
  margin-top: 1px;
}
.output-tab-main {
  width: 100%;
}
.cf-panel {
  display: none;
}
.cf-panel.is-active {
  display: block;
}
`;

let _rendererCtx = null;

function openConsoleInTab(element, chunkId) {
  const content = globalConsoleText.get(chunkId) || '';
  if (_rendererCtx && typeof _rendererCtx.postMessage === 'function') {
    _rendererCtx.postMessage({ type: 'open_console_in_tab', content, chunkId })
      .then((delivered) => {
        if (!delivered) {
          fallbackOpenConsoleInTab(content, chunkId);
        }
      })
      .catch(() => {
        fallbackOpenConsoleInTab(content, chunkId);
      });
    return;
  }
  fallbackOpenConsoleInTab(content, chunkId);
}

function fallbackOpenConsoleInTab(content, chunkId) {
  if (typeof document === 'undefined') {
    return;
  }
  const blob = new Blob([content || ''], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.target = '_blank';
  anchor.rel = 'noopener noreferrer';
  anchor.download = `${chunkId || 'console'}.txt`;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

export function activate(ctx) {
  _rendererCtx = ctx;
  return {
    renderOutputItem(outputItem, element) {
      Promise.resolve(outputItem.json())
        .then((result) => renderExecResult(outputItem, element, result))
        .catch((err) => {
          ensureStyles(element.ownerDocument);
          element.innerHTML = `<div class="rnb-output-root"><pre class="output-error">Renderer error: ${esc(String(err && err.message ? err.message : err))}</pre></div>`;
        });
    },
  };
}

function renderExecResult(outputItem, element, result) {
  if (!isExecResult(result)) {
    element.innerHTML = '';
    return;
  }

  ensureStyles(element.ownerDocument);

  const chunkId = String(
    result.chunk_id ||
    (outputItem && outputItem.metadata && outputItem.metadata.chunkId) ||
    'chunk',
  );
  const running = Boolean(
    (outputItem && outputItem.metadata && outputItem.metadata.running) ||
    result._rNotebookRunning,
  );
  captureConsoleState(element, chunkId);
  const tabs = buildTabs(result, running);

  // Store console output for the "Open in Tab" feature
  const consoleText = buildConsoleText(result);
  if (consoleText.trim()) {
    globalConsoleText.set(chunkId, consoleText);
  }

  if (tabs.length === 0) {
    element.innerHTML = '';
    return;
  }

  const activeKey = resolveActiveTabKey(element, chunkId, tabs);
  const renderState = { dfById: new Map() };
  const body = tabs.length === 1
    ? renderSinglePanel(tabs[0], renderState)
    : renderTabbedPanels(chunkId, tabs, activeKey, renderState);

  element.innerHTML = `<div class="rnb-output-root">${body}</div>`;

  bindTabInteractions(element, chunkId);
  bindConsoleInteractions(element, chunkId);
  hydrateDataFrames(element, chunkId, renderState.dfById);
  setActiveTab(element, chunkId, activeKey, { forceConsoleBottom: false });
}

function buildTabs(result, running) {
  const tabs = [];
  const consoleText = buildConsoleText(result);
  const consoleSegments = normalizeConsoleSegments(result.console_segments);
  const sourceCode = normalizeSourceCode(result.source_code);
  const mediaTabs = collectMediaTabs(result);
  const hasOtherContent = mediaTabs.length > 0 || Boolean(result.error);

  if (running || hasOtherContent || consoleText.trim()) {
    tabs.push({
      key: 'console',
      type: 'console',
      content: consoleText,
      preview: buildConsolePreview(consoleText, consoleSegments, sourceCode),
      consoleSegments,
      sourceCode,
      live: running,
    });
  }

  tabs.push(...mediaTabs);

  if (result.error) {
    tabs.push({
      key: 'error',
      type: 'error',
      content: result.error,
    });
  }

  return tabs;
}

function collectMediaTabs(result) {
  const tabs = [];
  const referencedPlotIndices = new Set();
  const referencedDfIndices = new Set();

  if (Array.isArray(result.output_order) && result.output_order.length > 0) {
    result.output_order.forEach((item, orderIndex) => {
      if (!item || typeof item.index !== 'number') {
        return;
      }
      if (item.type === 'df') {
        const df = result.dataframes && result.dataframes[item.index];
        if (!df) {
          return;
        }
        referencedDfIndices.add(item.index);
        tabs.push({
          key: `df:${item.index}:${df.name || orderIndex}`,
          type: 'df',
          content: df,
          name: df.name || item.name || 'DataFrame',
        });
        return;
      }

      const plot = result.plots && result.plots[item.index];
      if (!plot) {
        return;
      }
      referencedPlotIndices.add(item.index);
      tabs.push({
        key: `plot:${item.index}:${item.name || item.index}`,
        type: 'plot',
        content: plot,
        name: item.name || `Plot ${item.index + 1}`,
      });
    });
  }

  if (tabs.length === 0) {
    (result.dataframes || []).forEach((df, index) => {
      tabs.push({
        key: `df:${index}:${df && df.name ? df.name : index}`,
        type: 'df',
        content: df,
        name: df && df.name ? df.name : 'DataFrame',
      });
    });
    (result.plots || []).forEach((plot, index) => {
      tabs.push({
        key: `plot:${index}`,
        type: 'plot',
        content: plot,
        name: `Plot ${index + 1}`,
      });
    });
  } else {
    (result.dataframes || []).forEach((df, index) => {
      if (referencedDfIndices.has(index)) {
        return;
      }
      tabs.push({
        key: `df:${index}:${df && df.name ? df.name : index}`,
        type: 'df',
        content: df,
        name: df && df.name ? df.name : 'DataFrame',
      });
    });
    (result.plots || []).forEach((plot, index) => {
      if (referencedPlotIndices.has(index)) {
        return;
      }
      tabs.push({
        key: `plot:${index}`,
        type: 'plot',
        content: plot,
        name: `Plot ${index + 1}`,
      });
    });
  }

  (result.plots_html || []).forEach((html, index) => {
    tabs.push({
      key: `html:${index}`,
      type: 'html',
      content: html,
    });
  });

  return tabs;
}

function renderSinglePanel(tab, renderState) {
  return `<div class="output-tab-main">${renderPanelHtml(tab, renderState)}</div>`;
}

function renderTabbedPanels(chunkId, tabs, activeKey, renderState) {
  const stripClassName = tabs.length >= 8
    ? 'output-thumb-strip output-thumb-strip-grid'
    : 'output-thumb-strip';

  const thumbs = tabs.map((tab, index) => {
    const classes = ['output-thumb'];
    const title = tabHoverTitle(tab, index);
    if (tab.type === 'error') {
      classes.push('thumb-error');
    }
    if (tab.key === activeKey) {
      classes.push('is-active');
    }
    const titleAttr = title ? ` title="${escAttr(title)}"` : '';
    return `<button type="button" class="${classes.join(' ')}" data-tab-key="${escAttr(tab.key)}" data-chunk-id="${escAttr(chunkId)}"${titleAttr}>${buildThumbHtml(tab, index)}</button>`;
  }).join('');

  const panels = tabs.map((tab) => {
    const classes = ['cf-panel'];
    if (tab.key === activeKey) {
      classes.push('is-active');
    }
    return `<div class="${classes.join(' ')}" data-tab-key="${escAttr(tab.key)}">${renderPanelHtml(tab, renderState)}</div>`;
  }).join('');

  return `<div class="rn-output-tabs">
    <div class="${stripClassName}">
      ${thumbs}
    </div>
    <div class="output-tab-main">
      ${panels}
    </div>
  </div>`;
}

function renderPanelHtml(tab, renderState) {
  switch (tab.type) {
    case 'console':
      return buildConsoleViewerHtml(tab.content, {
        consoleSegments: tab.consoleSegments,
        live: tab.live,
        sourceCode: tab.sourceCode,
      });
    case 'plot':
      return buildPlotHtml(tab.content, tab.name);
    case 'html':
      return `<div class="output-html">${tab.content || ''}</div>`;
    case 'df': {
      const dfId = `df_${renderState.dfById.size + 1}`;
      renderState.dfById.set(dfId, tab);
      return `<div class="df-viewer" data-df-id="${escAttr(dfId)}"></div>`;
    }
    case 'stderr':
      return `<pre class="output-stderr">${esc(tab.content)}</pre>`;
    case 'text':
      return `<pre class="output-text">${esc(tab.content)}</pre>`;
    case 'error':
      return `<pre class="output-error">x ${esc(tab.content)}</pre>`;
    default:
      return '';
  }
}

function buildConsoleViewerHtml(content, options = {}) {
  const normalized = normalizeConsoleText(content);
  const consoleSegments = normalizeConsoleSegments(options.consoleSegments);
  const sourceCode = normalizeSourceCode(options.sourceCode);
  const status = options.live
    ? '<span class="console-meta">Streaming</span>'
    : '';
  const liveClass = options.live ? ' console-scroll-live' : '';
  const transcriptHtml = consoleSegments.length > 0
    ? buildConsoleTranscriptHtml(consoleSegments, normalized, { suppressFallback: options.live === true })
    : buildLegacyConsoleHtml(normalized, sourceCode);
  return `<div class="console-viewer">
    <div class="console-title">
      <strong>Console</strong>
      ${status}
      <button class="console-open-tab-btn" title="Open console output in a new tab">Open in Tab</button>
    </div>
    <div class="console-scroll${liveClass}">
      <div class="console-body">${transcriptHtml}</div>
    </div>
  </div>`;
}

function buildPlotHtml(b64, name = 'Plot') {
  return `<div class="plot-wrap">
    <img class="output-plot" src="data:image/png;base64,${b64}" alt="${escAttr(name)}" title="${escAttr(name)}">
    <a class="plot-dl-btn" href="data:image/png;base64,${b64}" download="plot.png">Save PNG</a>
  </div>`;
}

function buildThumbHtml(tab, index) {
  switch (tab.type) {
    case 'console':
      return `<pre class="thumb-text-preview">${esc(previewConsoleText(tab.preview || tab.content))}</pre><div class="thumb-label">Console</div>`;
    case 'plot':
      return `<img src="data:image/png;base64,${tab.content}" alt="${escAttr(tab.name || `Plot ${index + 1}`)}" title="${escAttr(tab.name || `Plot ${index + 1}`)}">`;
    case 'html':
      return '<div class="thumb-icon">&#128202;</div><div class="thumb-label">Interactive</div>';
    case 'df':
      return `<div class="thumb-icon">&#8862;</div><div class="thumb-label">${esc(friendlyName(tab.content.name || 'DataFrame'))}</div><div class="thumb-dims">${formatCount(tab.content.nrow)} x ${formatCount(tab.content.ncol)}</div>`;
    case 'stderr':
      return `<pre class="thumb-text-preview thumb-stderr">${esc(tab.content.split('\n').slice(0, 4).join('\n'))}</pre><div class="thumb-label">Stderr</div>`;
    case 'text':
      return `<pre class="thumb-text-preview">${esc(tab.content.split('\n').slice(0, 4).join('\n'))}</pre><div class="thumb-label">Console</div>`;
    case 'error':
      return '<div class="thumb-icon thumb-err-icon">&times;</div><div class="thumb-label">Error</div>';
    default:
      return '';
  }
}

function bindTabInteractions(element, chunkId) {
  element.querySelectorAll('.output-thumb[data-tab-key]').forEach((button) => {
    button.addEventListener('click', () => {
      const nextKey = button.getAttribute('data-tab-key') || '';
      setActiveTab(element, chunkId, nextKey, {
        forceConsoleBottom: nextKey === 'console',
        resetConsoleHeight: nextKey === 'console',
      });
    });
  });
}

function bindConsoleInteractions(element, chunkId) {
  element.querySelectorAll('.console-title').forEach((title) => {
    title.addEventListener('click', (e) => {
      if (e.target.classList.contains('console-open-tab-btn')) {
        return;
      }
      resetConsoleHeight(element, chunkId, { followBottom: true });
    });

    const btn = title.querySelector('.console-open-tab-btn');
    if (btn) {
      btn.addEventListener('click', () => {
        openConsoleInTab(element, chunkId);
      });
    }
  });

  element.querySelectorAll('.console-scroll').forEach((consoleScroll) => {
    let resizeGuardUntil = 0;
    let lastHeight = consoleScroll.offsetHeight;

    const syncConsoleFollow = () => {
      if (readConsoleState(chunkId).follow !== false) {
        scrollConsoleElementToBottom(consoleScroll);
      }
    };
    const persistConsoleState = (overrides = {}) => {
      const scrollTop = overrides.scrollTop !== undefined ? overrides.scrollTop : consoleScroll.scrollTop;
      const follow = overrides.follow !== undefined ? overrides.follow : isScrolledToBottom(consoleScroll);
      writeConsoleState(chunkId, {
        follow,
        scrollTop,
        height: overrides.height !== undefined ? overrides.height : consoleScroll.offsetHeight,
        viewportBottom: overrides.viewportBottom !== undefined
          ? overrides.viewportBottom
          : scrollTop + consoleScroll.clientHeight,
      });
    };

    const pauseConsoleFollow = () => {
      persistConsoleState({
        follow: false,
        scrollTop: consoleScroll.scrollTop,
        height: consoleScroll.offsetHeight,
        viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
      });
    };

    ['mousedown', 'pointerdown'].forEach((eventName) => {
      consoleScroll.addEventListener(eventName, () => {
        if (readConsoleState(chunkId).follow !== false) {
          resizeGuardUntil = Date.now() + 240;
        }
      });
    });

    ['wheel', 'touchmove'].forEach((eventName) => {
      consoleScroll.addEventListener(eventName, () => {
        pauseConsoleFollow();
      }, { passive: true });
    });

    consoleScroll.addEventListener('scroll', () => {
      const state = readConsoleState(chunkId);
      if (Date.now() < resizeGuardUntil && state.follow !== false) {
        scrollConsoleElementToBottom(consoleScroll);
        persistConsoleState({
          follow: true,
          scrollTop: consoleScroll.scrollTop,
          height: consoleScroll.offsetHeight,
          viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
        });
        return;
      }
      persistConsoleState();
    });

    if (typeof ResizeObserver === 'function') {
      const observer = new ResizeObserver(() => {
        const nextHeight = consoleScroll.offsetHeight;
        const heightChanged = Math.abs(nextHeight - lastHeight) > 0.5;
        lastHeight = nextHeight;

        const state = readConsoleState(chunkId);
        if (heightChanged) {
          resizeGuardUntil = Date.now() + 160;
          if (state.follow === false) {
            const maxScrollTop = Math.max(0, consoleScroll.scrollHeight - consoleScroll.clientHeight);
            const isValidViewportBottom = Number.isFinite(state.viewportBottom) && state.viewportBottom > state.scrollTop;
            const anchoredTop = isValidViewportBottom
              ? state.viewportBottom - consoleScroll.clientHeight
              : state.scrollTop;
            consoleScroll.scrollTop = clamp(anchoredTop, 0, maxScrollTop);
          }
        }
        if (state.follow !== false) {
          syncConsoleFollow();
        }
        persistConsoleState({
          follow: state.follow !== false,
          scrollTop: consoleScroll.scrollTop,
          height: nextHeight,
          viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
        });
      });
      observer.observe(consoleScroll);
      return;
    }

    ['mouseup', 'pointerup'].forEach((eventName) => {
      consoleScroll.addEventListener(eventName, () => {
        const nextHeight = consoleScroll.offsetHeight;
        const heightChanged = Math.abs(nextHeight - lastHeight) > 0.5;
        lastHeight = nextHeight;

        resizeGuardUntil = Date.now() + 160;
        const state = readConsoleState(chunkId);
        if (state.follow === false && heightChanged) {
          const maxScrollTop = Math.max(0, consoleScroll.scrollHeight - consoleScroll.clientHeight);
          const isValidViewportBottom = Number.isFinite(state.viewportBottom) && state.viewportBottom > state.scrollTop;
          const anchoredTop = isValidViewportBottom
            ? state.viewportBottom - consoleScroll.clientHeight
            : state.scrollTop;
          consoleScroll.scrollTop = clamp(anchoredTop, 0, maxScrollTop);
        }
        if (state.follow !== false) {
          syncConsoleFollow();
        }
        persistConsoleState({
          follow: state.follow !== false,
          scrollTop: consoleScroll.scrollTop,
          height: nextHeight,
          viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
        });
      });
    });
  });
}

function setActiveTab(element, chunkId, tabKey, options = {}) {
  const buttons = Array.from(element.querySelectorAll('.output-thumb[data-tab-key]'));
  const panels = Array.from(element.querySelectorAll('.cf-panel[data-tab-key]'));
  const resolvedKey = buttons.length === 0
    ? tabKey
    : buttons.some((button) => button.getAttribute('data-tab-key') === tabKey)
      ? tabKey
      : (buttons[0] && buttons[0].getAttribute('data-tab-key')) || '';

  buttons.forEach((button) => {
    button.classList.toggle('is-active', button.getAttribute('data-tab-key') === resolvedKey);
  });
  panels.forEach((panel) => {
    panel.classList.toggle('is-active', panel.getAttribute('data-tab-key') === resolvedKey);
  });

  if (resolvedKey) {
    storeState(globalTabState, `${TAB_STATE_PREFIX}${chunkId}`, resolvedKey);
  }

  if (resolvedKey === 'console') {
    if (options.resetConsoleHeight) {
      resetConsoleHeight(element, chunkId, { followBottom: options.forceConsoleBottom !== false });
      return;
    }
    if (options.forceConsoleBottom) {
      scrollConsoleToBottom(element, chunkId);
      return;
    }
    restoreConsoleState(element, chunkId);
  }
}

function resolveActiveTabKey(element, chunkId, tabs) {
  const currentElementKey = element.querySelector('.output-thumb.is-active')?.getAttribute('data-tab-key');
  const storedKey = currentElementKey || loadState(globalTabState, `${TAB_STATE_PREFIX}${chunkId}`);
  if (storedKey && tabs.some((tab) => tab.key === storedKey)) {
    return storedKey;
  }
  return tabs[0].key;
}

function captureConsoleState(element, chunkId) {
  const consoleScroll = getActiveConsoleScroll(element);
  const previousState = readConsoleState(chunkId);
  if (!consoleScroll) {
    writeConsoleState(chunkId, { follow: true, scrollTop: 0, height: 0, viewportBottom: 0 });
    return;
  }
  writeConsoleState(chunkId, {
    follow: previousState.follow !== false,
    scrollTop: consoleScroll.scrollTop,
    height: 0,
    viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
  });
}

function restoreConsoleState(element, chunkId) {
  const consoleScroll = getActiveConsoleScroll(element);
  if (!consoleScroll) {
    return;
  }
  const state = applyStoredConsoleHeight(consoleScroll, chunkId);
  if (state.follow !== false) {
    scrollConsoleElementToBottom(consoleScroll);
    writeConsoleState(chunkId, {
      follow: true,
      scrollTop: consoleScroll.scrollTop,
      height: consoleScroll.offsetHeight,
      viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
    });
    return;
  }
  const maxScrollTop = Math.max(0, consoleScroll.scrollHeight - consoleScroll.clientHeight);
  consoleScroll.scrollTop = clamp(state.scrollTop, 0, maxScrollTop);
  writeConsoleState(chunkId, {
    follow: false,
    scrollTop: consoleScroll.scrollTop,
    height: consoleScroll.offsetHeight,
    viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
  });
}

function scrollConsoleToBottom(element, chunkId) {
  const consoleScroll = getActiveConsoleScroll(element);
  if (!consoleScroll) {
    return;
  }
  applyStoredConsoleHeight(consoleScroll, chunkId);
  scrollConsoleElementToBottom(consoleScroll);
  writeConsoleState(chunkId, {
    follow: true,
    scrollTop: consoleScroll.scrollTop,
    height: consoleScroll.offsetHeight,
    viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
  });
}

function resetConsoleHeight(element, chunkId, options = {}) {
  const consoleScroll = getActiveConsoleScroll(element);
  if (!consoleScroll) {
    return;
  }
  consoleScroll.style.height = '';
  if (options.followBottom !== false) {
    scrollConsoleElementToBottom(consoleScroll);
  }
  writeConsoleState(chunkId, {
    follow: options.followBottom !== false,
    scrollTop: options.followBottom !== false ? consoleScroll.scrollTop : consoleScroll.scrollTop,
    height: 0,
    viewportBottom: consoleScroll.scrollTop + consoleScroll.clientHeight,
  });
}

function getActiveConsoleScroll(element) {
  return element.querySelector('.cf-panel.is-active .console-scroll')
    || element.querySelector('.output-tab-main > .console-viewer .console-scroll');
}

function scrollConsoleElementToBottom(consoleScroll) {
  consoleScroll.scrollTop = consoleScroll.scrollHeight;
  if (typeof requestAnimationFrame === 'function') {
    requestAnimationFrame(() => {
      consoleScroll.scrollTop = consoleScroll.scrollHeight;
    });
  }
}

function applyStoredConsoleHeight(consoleScroll, chunkId) {
  const state = readConsoleState(chunkId);
  if (state.height > 0) {
    consoleScroll.style.height = `${state.height}px`;
  }
  return state;
}

function isScrolledToBottom(consoleScroll) {
  return consoleScroll.scrollTop + consoleScroll.clientHeight >= consoleScroll.scrollHeight - 4;
}

function hydrateDataFrames(element, chunkId, dfById) {
  dfById.forEach((tab, dfId) => {
    const container = element.querySelector(`[data-df-id="${cssEscape(dfId)}"]`);
    if (!container) {
      return;
    }
    renderDataFrame(container, chunkId, tab.key, tab.content);
  });
}

function renderDataFrame(container, chunkId, tabKey, df) {
  const fullName = df.name || 'DataFrame';
  const normalizedRowNames = Array.isArray(df.row_names)
    ? df.row_names
    : (df.row_names == null ? [] : [df.row_names]);
  const hasRowNames = normalizedRowNames.length > 0;
  const shownColumnCount = Array.isArray(df.columns) ? df.columns.length : 0;
  const truncatedColumns = Number.isFinite(df.ncol) && shownColumnCount > 0 && df.ncol > shownColumnCount;
  const dfNotice = truncatedColumns
    ? `Showing first ${formatCount(shownColumnCount)} of ${formatCount(df.ncol)} columns`
    : '';
  const headerCells = (df.columns || []).map((column) => {
    const title = escAttr(column && column.type ? column.type : '');
    const name = esc(column && column.name ? column.name : '');
    return `<th title="${title}">${name}</th>`;
  }).join('');

  container.innerHTML = `<div class="df-title">
    <strong class="df-name" title="${escAttr(fullName)}">${esc(fullName)}</strong>
    <span class="df-dims">${formatCount(df.nrow)} x ${formatCount(df.ncol)}</span>
    <span class="df-note"${dfNotice ? ` title="${escAttr(dfNotice)}"` : ''}>${esc(dfNotice)}</span>
    <span class="df-info"></span>
  </div>
  <div class="df-table-wrap">
    <table class="df-table">
      <thead><tr><th class="row-idx"${hasRowNames ? ' title="row names"' : ''}></th>${headerCells}</tr></thead>
      <tbody></tbody>
    </table>
  </div>
  <div class="df-pager"></div>`;

  const totalLoadedPages = Math.max(1, Math.ceil(((df.data && df.data.length) || 0) / 50));
  const storedPage = Number(loadState(globalPageState, `${PAGE_STATE_PREFIX}${chunkId}:${tabKey}`));
  const initialPage = Number.isFinite(storedPage)
    ? clamp(storedPage, 0, Math.max(0, totalLoadedPages - 1))
    : 0;

  const state = { chunkId, tabKey, df, page: initialPage };
  container.addEventListener('click', (event) => {
    const eventTarget = event.target;
    if (!eventTarget || typeof eventTarget.closest !== 'function') {
      return;
    }
    const target = eventTarget.closest('[data-page]');
    if (!target) {
      return;
    }
    const nextPage = Number(target.getAttribute('data-page'));
    if (!Number.isFinite(nextPage)) {
      return;
    }
    state.page = clamp(nextPage, 0, Math.max(0, totalLoadedPages - 1));
    drawDataFramePage(container, state);
  });

  const jumpInput = () => container.querySelector('.pg-input');
  container.addEventListener('change', (event) => {
    const target = event.target;
    if (!target || !target.classList || !target.classList.contains('pg-input')) {
      return;
    }
    const nextPage = clamp((parseInt(target.value, 10) || 1) - 1, 0, Math.max(0, totalLoadedPages - 1));
    state.page = nextPage;
    drawDataFramePage(container, state);
    const input = jumpInput();
    if (input) {
      input.value = String(state.page + 1);
    }
  });

  drawDataFramePage(container, state);
}

function drawDataFramePage(container, state) {
  const rows = Array.isArray(state.df.data) ? state.df.data : [];
  const rowNames = Array.isArray(state.df.row_names)
    ? state.df.row_names
    : (state.df.row_names == null ? [] : [state.df.row_names]);
  const rowsPerPage = 50;
  const totalPages = Math.max(1, Math.ceil(rows.length / rowsPerPage));
  state.page = clamp(state.page, 0, Math.max(0, totalPages - 1));

  const infoEl = container.querySelector('.df-info');
  const bodyEl = container.querySelector('tbody');
  const pagerEl = container.querySelector('.df-pager');
  if (!infoEl || !bodyEl || !pagerEl) {
    return;
  }

  if (rows.length === 0) {
    bodyEl.innerHTML = '<tr><td colspan="100" class="na-value">No data</td></tr>';
    infoEl.textContent = '0 rows';
    pagerEl.innerHTML = '';
    storeState(globalPageState, `${PAGE_STATE_PREFIX}${state.chunkId}:${state.tabKey}`, '0');
    return;
  }

  const start = state.page * rowsPerPage;
  const slice = rows.slice(start, Math.min(start + rowsPerPage, rows.length));
  bodyEl.innerHTML = slice.map((row, rowIndex) => {
    const values = Array.isArray(row) ? row : Object.values(row || {});
    const rowName = rowNames[start + rowIndex] ?? String(start + rowIndex + 1);
    const cells = values.map((value) => {
      if (value == null) {
        return '<td class="na-value">NA</td>';
      }
      return `<td>${esc(String(value))}</td>`;
    }).join('');
    return `<tr><td class="row-idx">${esc(String(rowName))}</td>${cells}</tr>`;
  }).join('');

  const rangeStart = start + 1;
  const rangeEnd = Math.min(start + slice.length, state.df.nrow || rows.length);
  const loadedSuffix = rows.length < state.df.nrow
    ? ` (${formatCount(rows.length)} loaded)`
    : '';
  infoEl.textContent = `Rows ${formatCount(rangeStart)}-${formatCount(rangeEnd)} of ${formatCount(state.df.nrow)}${loadedSuffix}`;
  pagerEl.innerHTML = buildPagerHtml(state.page, totalPages);
  storeState(globalPageState, `${PAGE_STATE_PREFIX}${state.chunkId}:${state.tabKey}`, String(state.page));
}

function buildPagerHtml(currentPage, totalPages) {
  const parts = [];
  parts.push(buildPagerButton('&#9668;&nbsp;Prev', currentPage - 1, currentPage === 0, false));
  buildVisiblePageNumbers(currentPage, totalPages).forEach((page) => {
    if (page < 0) {
      parts.push('<span class="pg-ellipsis">&#8230;</span>');
      return;
    }
    parts.push(buildPagerButton(String(page + 1), page, false, page === currentPage));
  });
  parts.push(buildPagerButton('Next&nbsp;&#9658;', currentPage + 1, currentPage >= totalPages - 1, false));
  parts.push(`<span class="pg-jump-wrap">Go to page <input type="number" class="pg-input" min="1" max="${totalPages}" value="${currentPage + 1}"></span>`);
  return parts.join('');
}

function buildPagerButton(label, page, disabled, active) {
  const classes = ['pg-btn'];
  if (active) {
    classes.push('pg-active');
  }
  if (disabled) {
    classes.push('pg-dis');
  }
  return `<button type="button" class="${classes.join(' ')}" data-page="${page}"${disabled ? ' disabled' : ''}>${label}</button>`;
}

function buildVisiblePageNumbers(currentPage, totalPages) {
  if (totalPages <= 9) {
    return Array.from({ length: totalPages }, (_, index) => index);
  }
  if (currentPage <= 4) {
    return [0, 1, 2, 3, 4, 5, -1, totalPages - 1];
  }
  if (currentPage >= totalPages - 5) {
    return [0, -1, totalPages - 6, totalPages - 5, totalPages - 4, totalPages - 3, totalPages - 2, totalPages - 1];
  }
  return [0, -1, currentPage - 1, currentPage, currentPage + 1, -1, totalPages - 1];
}

function buildConsoleText(result) {
  if (result.console && String(result.console).trim()) {
    return String(result.console);
  }
  return [result.stdout || '', result.stderr || '']
    .filter((part) => String(part).trim().length > 0)
    .join('\n');
}

function previewConsoleText(text) {
  const lines = normalizeConsoleText(text).split('\n');
  return lines.slice(Math.max(0, lines.length - 4)).join('\n');
}

function buildConsolePreview(consoleText, consoleSegments, sourceCode) {
  if (Array.isArray(consoleSegments) && consoleSegments.length > 0) {
    for (let index = consoleSegments.length - 1; index >= 0; index -= 1) {
      const outputText = normalizeConsoleText(consoleSegments[index].output || '');
      if (outputText.trim()) {
        return outputText;
      }
    }
    const lastCode = consoleSegments[consoleSegments.length - 1].code || '';
    return formatSourcePreview(lastCode);
  }
  if (consoleText.trim()) {
    return consoleText;
  }
  return formatSourcePreview(sourceCode);
}

function normalizeConsoleSegments(segments) {
  if (!Array.isArray(segments)) {
    return [];
  }
  return segments
    .filter((segment) => segment && typeof segment.code === 'string')
    .map((segment) => ({
      code: segment.code,
      output: typeof segment.output === 'string' ? segment.output : '',
    }));
}

function buildConsoleTranscriptHtml(consoleSegments, fallbackOutput, options = {}) {
  const blocks = [];
  consoleSegments.forEach((segment) => {
    const code = normalizeSourceCode(segment.code);
    const output = normalizeConsoleText(segment.output || '');
    if (code) {
      blocks.push(`<div class="console-input"><pre class="output-console-code">${esc(formatSourceEcho(code))}</pre></div>`);
    }
    if (output) {
      blocks.push(`<pre class="output-console has-source">${esc(output)}</pre>`);
    }
  });

  const transcriptOutput = consoleSegments
    .map((segment) => normalizeConsoleText(segment.output || ''))
    .filter(Boolean)
    .join('\n');
  const normalizedFallback = normalizeConsoleText(fallbackOutput);
  if (!options.suppressFallback && normalizedFallback && normalizedFallback !== transcriptOutput) {
    blocks.push(`<pre class="output-console">${esc(normalizedFallback)}</pre>`);
  }

  return blocks.join('');
}

function buildLegacyConsoleHtml(normalizedOutput, sourceCode) {
  const sourceHtml = sourceCode
    ? `<div class="console-input"><pre class="output-console-code">${esc(formatSourceEcho(sourceCode))}</pre></div>`
    : '';
  const outputClasses = ['output-console'];
  if (sourceCode) {
    outputClasses.push('has-source');
  }
  const outputHtml = normalizedOutput
    ? `<pre class="${outputClasses.join(' ')}">${esc(normalizedOutput)}</pre>`
    : '';
  return `${sourceHtml}${outputHtml}`;
}

function normalizeSourceCode(text) {
  return String(text || '').replace(/\s+$/g, '');
}

function formatSourceEcho(sourceCode) {
  if (!sourceCode) {
    return '';
  }
  return sourceCode
    .split('\n')
    .map((line) => `> ${line}`)
    .join('\n');
}

function formatSourcePreview(sourceCode) {
  if (!sourceCode) {
    return '';
  }
  return formatSourceEcho(sourceCode).split('\n').slice(0, 4).join('\n');
}

function normalizeConsoleText(text) {
  const chars = Array.from(String(text || '').replace(/\r\n/g, '\n'));
  let currentLine = '';
  const lines = [];
  chars.forEach((char) => {
    if (char === '\r') {
      currentLine = '';
      return;
    }
    if (char === '\n') {
      lines.push(currentLine);
      currentLine = '';
      return;
    }
    currentLine += char;
  });
  lines.push(currentLine);
  return lines.join('\n');
}

function friendlyName(name) {
  if (!name || name.length <= 40) {
    return name || '';
  }
  return `${name.slice(0, 37)}...`;
}

function tabHoverTitle(tab, index) {
  if (tab.type === 'df') {
    return tab.name || (tab.content && tab.content.name) || `DataFrame ${index + 1}`;
  }
  if (tab.type === 'plot') {
    return tab.name || `Plot ${index + 1}`;
  }
  return '';
}

function ensureStyles(doc) {
  if (!doc || doc.getElementById(STYLE_ID)) {
    return;
  }
  const style = doc.createElement('style');
  style.id = STYLE_ID;
  style.textContent = STYLES;
  doc.head.appendChild(style);
}

function storeState(bucket, key, value) {
  bucket.set(key, value);
  try {
    window.sessionStorage.setItem(key, value);
  } catch (_err) {}
  try {
    window.localStorage.setItem(key, value);
  } catch (_err) {}
}

function loadState(bucket, key) {
  if (bucket.has(key)) {
    return bucket.get(key);
  }
  try {
    const sessionValue = window.sessionStorage.getItem(key);
    if (sessionValue) {
      bucket.set(key, sessionValue);
      return sessionValue;
    }
  } catch (_err) {}
  try {
    const localValue = window.localStorage.getItem(key);
    if (localValue) {
      bucket.set(key, localValue);
      return localValue;
    }
  } catch (_err) {}
  return '';
}

function readConsoleState(chunkId) {
  const key = `${CONSOLE_STATE_PREFIX}${chunkId}`;
  if (globalConsoleState.has(key)) {
    return globalConsoleState.get(key);
  }
  try {
    const raw = window.sessionStorage.getItem(key) || window.localStorage.getItem(key);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        const normalized = {
          follow: parsed.follow !== false,
          scrollTop: Number(parsed.scrollTop) || 0,
          height: Number(parsed.height) || 0,
          viewportBottom: Number(parsed.viewportBottom) || 0,
        };
        globalConsoleState.set(key, normalized);
        return normalized;
      }
    }
  } catch (_err) {}
  const fallback = { follow: true, scrollTop: 0, height: 0, viewportBottom: 0 };
  globalConsoleState.set(key, fallback);
  return fallback;
}

function writeConsoleState(chunkId, state) {
  const key = `${CONSOLE_STATE_PREFIX}${chunkId}`;
  const normalized = {
    follow: state.follow !== false,
    scrollTop: Number(state.scrollTop) || 0,
    height: Number(state.height) || 0,
    viewportBottom: Number(state.viewportBottom) || 0,
  };
  globalConsoleState.set(key, normalized);
  const serialized = JSON.stringify(normalized);
  try {
    window.sessionStorage.setItem(key, serialized);
  } catch (_err) {}
  try {
    window.localStorage.setItem(key, serialized);
  } catch (_err) {}
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatCount(value) {
  return Number(value || 0).toLocaleString();
}

function esc(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escAttr(value) {
  return esc(value).replace(/'/g, '&#39;');
}

function cssEscape(value) {
  return String(value).replace(/["\\]/g, '\\$&');
}

function isExecResult(value) {
  return Boolean(
    value &&
    typeof value === 'object' &&
    value.type === 'result' &&
    typeof value.chunk_id === 'string' &&
    (value.source_code === undefined || typeof value.source_code === 'string') &&
    (value.console_segments === undefined || Array.isArray(value.console_segments)) &&
    typeof value.stdout === 'string' &&
    typeof value.stderr === 'string' &&
    Array.isArray(value.plots) &&
    Array.isArray(value.dataframes)
  );
}
