// =============================================================================
// notebookOutputHtml.ts — shared HTML output builders for all R Notebook
// notebook controllers (R and Python).
//
// Each cell execution produces ONE text/html output block whose layout
// matches rmarkdownPanel.css exactly (same CSS classes, same structure).
//   single output → direct render
//   2+ outputs    → radio-button thumbnail strip with minimal JS to persist
//                   the selected tab across streaming rerenders
// =============================================================================

import { DataFrameResult, ExecResult } from './kernelProtocol';

// ---------------------------------------------------------------------------
// Public entry point

export type OutputTab =
  | { type: 'console'; content: string; live?: boolean }
  | { type: 'text';   content: string }
  | { type: 'stderr'; content: string }
  | { type: 'error';  content: string }
  | { type: 'plot';   content: string }          // base64 PNG
  | { type: 'html';   content: string }          // interactive HTML (Plotly, etc.)
  | { type: 'df';     content: DataFrameResult };

type BuildOutputHtmlOptions = {
  running?: boolean;
};

export function buildOutputHtml(
  result: ExecResult,
  chunkId: string,
  options: BuildOutputHtmlOptions = {},
): string | null {
  const body = buildOutputFragment(result, chunkId, options);
  return body ? wrapHtml(body) : null;
}

export function buildOutputFragment(
  result: ExecResult,
  chunkId: string,
  options: BuildOutputHtmlOptions = {},
): string | null {
  // When running but no stream outputs yet, show just the live console.
  if (options.running && !(result.output_order?.length)) {
    return renderLiveConsoleHtml(result);
  }

  const tabs: OutputTab[] = [];
  const consoleText = buildConsoleText(result);

  if (consoleText.trim()) {
    tabs.push({ type: 'console', content: consoleText, live: options.running });
  }

  tabs.push(...collectMediaTabs(result));
  if (result.error)
    tabs.push({ type: 'error', content: result.error });

  if (tabs.length === 0) {
    if (options.running) return renderLiveConsoleHtml(result);
    return null;
  }

  return tabs.length === 1
    ? renderSingleOutputHtml(tabs[0])
    : renderTabsHtml(tabs, chunkId);
}

export function buildErrorHtml(message: string): string {
  return wrapHtml(`<pre class="output-error">✖ ${esc(message)}</pre>`);
}

export function wrapHtml(body: string): string {
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">\
<style>${NOTEBOOK_OUTPUT_CSS}</style></head><body>${body}</body></html>`;
}

// ---- Single output (no thumbnail strip) ------------------------------------

export function renderSingleOutputHtml(tab: OutputTab): string {
  switch (tab.type) {
    case 'console':
      return buildConsoleViewerHtml(tab.content, { live: tab.live });
    case 'text':
      return `<pre class="output-text">${esc(tab.content)}</pre>`;
    case 'stderr':
      return `<pre class="output-stderr">${esc(tab.content)}</pre>`;
    case 'error':
      return `<pre class="output-error">✖ ${esc(tab.content)}</pre>`;
    case 'plot':
      return buildPlotHtml(tab.content);
    case 'html':
      return `<div class="output-html">${tab.content}</div>`;
    case 'df':
      return buildDfViewerHtml(tab.content);
    default:
      return '';
  }
}

function renderLiveConsoleHtml(
  result: Pick<ExecResult, 'console' | 'stdout' | 'stderr'>,
): string {
  const consoleText = buildConsoleText(result);
  return buildConsoleViewerHtml(consoleText, { live: true });
}

function buildConsoleText(
  result: Pick<ExecResult, 'console' | 'stdout' | 'stderr'>,
): string {
  if (result.console && result.console.trim()) return result.console;

  const parts = [result.stdout ?? '', result.stderr ?? '']
    .filter(part => part.trim().length > 0);
  return parts.join(parts.length > 1 ? '\n' : '');
}

function collectMediaTabs(result: ExecResult): OutputTab[] {
  const tabs: OutputTab[] = [];
  const referencedPlotIndices = new Set<number>();
  const referencedDfIndices = new Set<number>();

  if (result.output_order && result.output_order.length > 0) {
    for (const item of result.output_order) {
      if (item.type === 'df') {
        const df = result.dataframes?.[item.index];
        if (!df) continue;
        tabs.push({ type: 'df', content: df });
        referencedDfIndices.add(item.index);
        continue;
      }

      const b64 = result.plots?.[item.index];
      if (!b64) continue;
      tabs.push({ type: 'plot', content: b64 });
      referencedPlotIndices.add(item.index);
    }
  }

  if (tabs.length === 0) {
    for (const df of result.dataframes ?? [])
      tabs.push({ type: 'df', content: df });
    for (const b64 of result.plots ?? [])
      tabs.push({ type: 'plot', content: b64 });
  } else {
    for (let i = 0; i < (result.dataframes?.length ?? 0); i++) {
      if (!referencedDfIndices.has(i))
        tabs.push({ type: 'df', content: result.dataframes[i] });
    }
    for (let i = 0; i < (result.plots?.length ?? 0); i++) {
      if (!referencedPlotIndices.has(i))
        tabs.push({ type: 'plot', content: result.plots[i] });
    }
  }

  for (const html of result.plots_html ?? [])
    tabs.push({ type: 'html', content: html });

  return tabs;
}

let _consoleUid = 0;

function buildConsoleViewerHtml(
  content: string,
  options: { live?: boolean } = {},
): string {
  const normalized = normalizeConsoleText(content);
  const uid = `cf_console_${++_consoleUid}`;
  const status = options.live
    ? '<span class="console-meta">Streaming</span>'
    : '';
  const script = options.live
    ? `<script>(function(){var el=document.getElementById('${uid}');if(el){el.scrollTop=el.scrollHeight;}})()</script>`
    : '';
  return `<div class="console-viewer${options.live ? ' console-viewer-live' : ''}">
  <div class="console-title">
    <strong>Console</strong>
    ${status}
  </div>
  <div class="console-scroll${options.live ? ' console-scroll-live' : ''}" id="${uid}">
    <pre class="output-console">${esc(normalized)}</pre>
  </div>
</div>${script}`;
}

function normalizeConsoleText(text: string): string {
  const chars = Array.from(text.replace(/\r\n/g, '\n'));
  let currentLine = '';
  const lines: string[] = [];

  for (const char of chars) {
    if (char === '\r') {
      currentLine = '';
      continue;
    }
    if (char === '\n') {
      lines.push(currentLine);
      currentLine = '';
      continue;
    }
    currentLine += char;
  }

  lines.push(currentLine);
  return lines.join('\n');
}

// ---- Multiple outputs: CSS-only thumbnail strip + panels -------------------
//
// Structure (all siblings inside .cf-output-tabs):
//   <input type="radio" id="t0" ...>   ← radio inputs first
//   <input type="radio" id="t1" ...>
//   <div class="cf-strip output-thumb-strip">
//     <label for="t0" class="output-thumb">...</label>
//   </div>
//   <div class="cf-panels output-tab-main">
//     <div class="cf-panel">...</div>
//   </div>
//
// CSS sibling combinator: #tN:checked ~ .cf-strip label:nth-child(N+1)
// highlights the active thumbnail; ~ .cf-panels .cf-panel:nth-child(N+1)
// shows the active panel — zero JavaScript required.

function renderTabsHtml(tabs: OutputTab[], id: string): string {
  const radioIds = tabs.map((_, i) => `cf_t${i}_${id}`);
  const tabKeys = tabs.map((tab, i) => buildTabStateKey(tab, i));
  const stripClassName = tabs.length >= 8
    ? 'cf-strip output-thumb-strip output-thumb-strip-grid'
    : 'cf-strip output-thumb-strip';

  const radios = tabs
    .map((_, i) =>
      `<input type="radio" id="${radioIds[i]}" name="cf_tabs_${id}" class="cf-r" data-tab-key="${esc(tabKeys[i])}"${i === 0 ? ' checked' : ''}>`)
    .join('\n  ');

  const tabCss = tabs
    .map((_, i) => `
  #${radioIds[i]}:checked ~ .cf-strip label:nth-child(${i + 1}) {
    border-color: #0078d4;
    box-shadow: 0 0 0 3px rgba(0,120,212,.22);
  }
  #${radioIds[i]}:checked ~ .cf-panels .cf-panel:nth-child(${i + 1}) {
    display: block;
  }`)
    .join('');

  const thumbs = tabs
    .map((tab, i) => {
      const extra = tab.type === 'error' ? ' thumb-error' : '';
      return `<label for="${radioIds[i]}" class="output-thumb${extra}">${buildThumbHtml(tab, i)}</label>`;
    })
    .join('\n    ');

  const panels = tabs
    .map(tab => `<div class="cf-panel">${renderSingleOutputHtml(tab)}</div>`)
    .join('\n    ');

  const stateScript = `<script>(function(){var root=document.currentScript.previousElementSibling;if(!root)return;var key=${JSON.stringify(`cf_tab_state_${id}`)};var radios=Array.from(root.querySelectorAll('.cf-r'));if(radios.length===0)return;var stateBucket='__rNotebookTabState';function readStorage(){var stores=[window.localStorage,window.sessionStorage];for(var i=0;i<stores.length;i++){var store=stores[i];try{var value=store&&store.getItem(key);if(value)return value;}catch(_err){}}return'';}function writeStorage(value){var stores=[window.localStorage,window.sessionStorage];for(var i=0;i<stores.length;i++){var store=stores[i];try{if(store)store.setItem(key,value);}catch(_err){}}}function readWindowNameState(){try{var raw=window.name||'';if(!raw)return{};var parsed=JSON.parse(raw);if(parsed&&typeof parsed==='object'&&parsed[stateBucket]&&typeof parsed[stateBucket]==='object')return parsed[stateBucket];}catch(_err){}return{};}function writeWindowNameState(state){try{var parsed={};var raw=window.name||'';if(raw){var current=JSON.parse(raw);if(current&&typeof current==='object')parsed=current;}parsed[stateBucket]=state;window.name=JSON.stringify(parsed);}catch(_err){}}function readHostState(){try{var host=window.top;if(host&&host!==window){host[stateBucket]=host[stateBucket]||{};return host[stateBucket];}}catch(_err){}return null;}var hostState=readHostState();var state=hostState||readWindowNameState();var stored=readStorage()||state[key]||'';if(stored){var match=radios.find(function(r){return r.getAttribute('data-tab-key')===stored;});if(match)match.checked=true;}function persist(next){writeStorage(next);if(hostState)hostState[key]=next;state[key]=next;writeWindowNameState(state);}radios.forEach(function(r){r.addEventListener('change',function(){if(!r.checked)return;persist(r.getAttribute('data-tab-key')||'');});});if(!stored){var checked=radios.find(function(r){return r.checked;});if(checked)persist(checked.getAttribute('data-tab-key')||'');}})()</script>`;

  return `<style>.cf-r{display:none}.cf-panel{display:none}${tabCss}</style>\
<div class="cf-output-tabs">
  ${radios}
  <div class="${stripClassName}">
    ${thumbs}
  </div>
  <div class="cf-panels output-tab-main">
    ${panels}
  </div>
</div>${stateScript}`;
}

function buildTabStateKey(tab: OutputTab, idx: number): string {
  switch (tab.type) {
    case 'console':
      return 'console';
    case 'text':
      return 'text';
    case 'stderr':
      return 'stderr';
    case 'error':
      return 'error';
    case 'df':
      return `df:${tab.content.name ?? idx}`;
    case 'plot':
      return `plot:${tab.content.length}:${tab.content.slice(0, 32)}`;
    case 'html':
      return `html:${idx}`;
    default:
      return `${idx}`;
  }
}

// ---- Thumbnail card content ------------------------------------------------

function buildThumbHtml(tab: OutputTab, idx: number): string {
  switch (tab.type) {
    case 'console': {
      const preview = previewConsoleText(tab.content);
      return `<pre class="thumb-text-preview">${esc(preview)}</pre>\
<div class="thumb-label">Console</div>`;
    }
    case 'plot':
      return `<img src="data:image/png;base64,${tab.content}" alt="Plot ${idx + 1}">`;
    case 'html':
      return `<div class="thumb-icon">📊</div>\
<div class="thumb-label">Interactive</div>`;
    case 'df': {
      const name = friendlyName(tab.content.name ?? 'DataFrame');
      return `<div class="thumb-icon">⊞</div>\
<div class="thumb-label">${esc(name)}</div>\
<div class="thumb-dims">${tab.content.nrow} × ${tab.content.ncol}</div>`;
    }
    case 'text': {
      const preview = tab.content.split('\n').slice(0, 4).join('\n');
      return `<pre class="thumb-text-preview">${esc(preview)}</pre>\
<div class="thumb-label">Console</div>`;
    }
    case 'stderr': {
      const preview = tab.content.split('\n').slice(0, 4).join('\n');
      return `<pre class="thumb-text-preview thumb-stderr">${esc(preview)}</pre>\
<div class="thumb-label">Stderr</div>`;
    }
    case 'error':
      return `<div class="thumb-icon thumb-err-icon">✖</div>\
<div class="thumb-label">Error</div>`;
    default:
      return '';
  }
}

function previewConsoleText(text: string, lineCount = 4): string {
  const lines = normalizeConsoleText(text).split('\n');
  return lines.slice(Math.max(0, lines.length - lineCount)).join('\n');
}

// ---- Plot ------------------------------------------------------------------

export function buildPlotHtml(b64: string): string {
  return `<div class="plot-wrap">\
<img class="output-plot" src="data:image/png;base64,${b64}" alt="Plot">\
<a class="plot-dl-btn" href="data:image/png;base64,${b64}" download="plot.png">⬇ Save PNG</a>\
</div>`;
}

// ---- DataFrame viewer with RStudio-style JavaScript paginator --------------
// All available rows (up to 2000, sent by the kernel) are embedded as JSON
// and paginated client-side at 50 rows per page.  No round-trip to the kernel
// is required — the paginator works entirely inside the output iframe.

let _dfUid = 0;

export function buildDfViewerHtml(df: DataFrameResult): string {
  const uid  = `dft${++_dfUid}`;
  const name = friendlyName(df.name ?? 'DataFrame');
  const nrow = df.nrow;
  const ncol = df.ncol;
  const rowNames = Array.isArray(df.row_names)
    ? df.row_names
    : (df.row_names == null ? [] : [df.row_names]);
  const shownColumnCount = Array.isArray(df.columns) ? df.columns.length : 0;
  const truncatedColumns = Number.isFinite(ncol) && shownColumnCount > 0 && ncol > shownColumnCount;
  const dfNotice = truncatedColumns
    ? `Showing first ${shownColumnCount.toLocaleString()} of ${ncol.toLocaleString()} columns`
    : '';

  const headerCells = df.columns
    .map(c => `<th title="${esc(c.type ?? '')}"><div class="df-col-head"><span class="df-col-name">${esc(c.name)}</span><span class="df-col-type">(${esc(c.type ?? '?')})</span></div></th>`)
    .join('');

  // Embed all available rows as JSON.  Replace '</' to prevent early
  // </script> tag closure if data contains HTML-like strings.
  const rowsJson = JSON.stringify(df.data).replace(/<\//g, '<\\/');
  const rowNamesJson = JSON.stringify(rowNames).replace(/<\//g, '<\\/');

  // The paginator script uses globally-named callbacks so onclick= attributes
  // (evaluated in window scope) can call them without closures.
  const script = `(function(){
var R=${rowsJson},RN=${rowNamesJson},N=${nrow},Z=50,T=Math.ceil(R.length/Z);
function eh(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function draw(p){
  if(T===0){
    document.getElementById('${uid}-body').innerHTML='<tr><td colspan="100" class="na-value">No data</td></tr>';
    document.getElementById('${uid}-info').textContent='0 rows';
    document.getElementById('${uid}-pg').innerHTML='';
    return;
  }
  var s=p*Z,sl=R.slice(s,Math.min(s+Z,R.length));
  document.getElementById('${uid}-body').innerHTML=sl.map(function(row,i){
    var rowName=RN[s+i] != null ? RN[s+i] : String(s+i+1);
    return'<tr><td class="row-idx">'+eh(String(rowName))+'</td>'+
      Object.values(row).map(function(v){
        return v==null?'<td class="na-value">NA</td>':'<td>'+eh(String(v))+'</td>';
      }).join('')+'</tr>';
  }).join('');
  var r1=p*Z+1,r2=Math.min((p+1)*Z,N);
  document.getElementById('${uid}-info').textContent=
    'Rows '+r1+'\u2013'+r2+' of '+N.toLocaleString()+
    (R.length<N?' ('+R.length.toLocaleString()+' loaded)':'');
  pager(p);
}
function pager(p){
  function btn(lbl,n,dis,act){
    return'<button class="pg-btn'+(act?' pg-active':'')+(dis?' pg-dis':'')+'"'+
      (dis?' disabled':'')+' onclick="window[\\'pg_${uid}\\']('+(+n)+')">'+lbl+'<\\/button>';
  }
  function nums(c,t){
    if(t<=9){var a=[];for(var i=0;i<t;i++)a.push(i);return a;}
    var a=[];
    if(c<=4){for(var i=0;i<6;i++)a.push(i);a.push(-1);a.push(t-1);}
    else if(c>=t-5){a.push(0);a.push(-1);for(var i=t-6;i<t;i++)a.push(i);}
    else{a.push(0);a.push(-1);for(var i=c-1;i<=c+1;i++)a.push(i);a.push(-1);a.push(t-1);}
    return a;
  }
  var h=btn('&#9668;&nbsp;Prev',p-1,p===0,false);
  nums(p,T).forEach(function(n){
    h+=n<0?'<span class="pg-ellipsis">&#8230;<\\/span>':btn(n+1,n,false,n===p);
  });
  h+=btn('Next&nbsp;&#9658;',p+1,p>=T-1,false);
  h+='<span class="pg-jump-wrap">Go to page <input type="number" class="pg-input" min="1" max="'+T+'" value="'+(p+1)+'" onchange="window[\\'jmp_${uid}\\']( this.value )"><\\/span>';
  document.getElementById('${uid}-pg').innerHTML=h;
}
window['pg_${uid}']=function(n){if(n>=0&&n<T)draw(n);};
window['jmp_${uid}']=function(v){var p=parseInt(v,10)-1;draw(Math.max(0,Math.min(p,T-1)));};
draw(0);
})()`;

  return `<div class="df-viewer">
  <div class="df-title">
    <strong>${esc(name)}</strong>
    <span class="df-dims">${nrow.toLocaleString()} × ${ncol}</span>
    <span class="df-note"${dfNotice ? ` title="${esc(dfNotice)}"` : ''}>${esc(dfNotice)}</span>
    <span id="${uid}-info" class="df-info"></span>
  </div>
  <div class="df-table-wrap">
    <table class="df-table">
      <thead><tr><th title="row names"></th>${headerCells}</tr></thead>
      <tbody id="${uid}-body"></tbody>
    </table>
  </div>
  <div class="df-pager" id="${uid}-pg"></div>
</div><script>${script}</script>`;
}

// ---- Helpers ---------------------------------------------------------------

export function esc(s: string): string {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function friendlyName(name: string): string {
  if (!name || name.length <= 40) return name;
  return name.slice(0, 37) + '…';
}

// =============================================================================
// CSS — verbatim from rmarkdownPanel.css (output-relevant rules only).
// VS Code theme variables (--vscode-*) work in notebook cell outputs.
// =============================================================================

export const NOTEBOOK_OUTPUT_CSS = `
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body { border: none; outline: none; }
body {
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
  overflow: hidden;
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
.console-scroll {
  max-height: 24em;
  overflow: auto;
  resize: vertical;
}
.console-scroll-live {
  max-height: 18em;
}
.console-scroll .output-console {
  min-height: 100%;
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
  top: 6px; right: 6px;
  padding: 3px 8px;
  background: rgba(0,0,0,.65);
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
  align-items: baseline;
  gap: 8px;
  padding: 5px 10px;
  border-bottom: 1px solid var(--vscode-editorGroup-border);
}
.df-dims { color: var(--vscode-descriptionForeground); font-size: 11px; }
.df-info { color: var(--vscode-descriptionForeground); font-size: 11px; margin-left: auto; }
.df-note { color: var(--vscode-descriptionForeground); font-size: 11px; margin-left: 10px; white-space: nowrap; }
.df-table-wrap { overflow-x: auto; max-height: 300px; overflow-y: auto; }
.df-table { width: 100%; border-collapse: collapse; white-space: nowrap; }
.df-table thead th {
  position: sticky; top: 0;
  background: var(--vscode-editor-background);
  border-bottom: none;
  box-shadow: 0 2px 0 var(--vscode-editorGroup-border, #c0c0c0);
  border-right: 1px solid var(--vscode-editorGroup-border, #c0c0c0);
  padding: 4px 10px;
  text-align: left; font-weight: 600;
  vertical-align: bottom;
}
.df-table thead th:last-child { border-right: none; }
.df-col-head {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 2px;
}
.df-col-name {
  display: block;
  line-height: 1.25;
}
.df-col-type {
  display: block;
  margin-top: 2px;
  font-size: 10px;
  font-weight: 500;
  line-height: 1.2;
  color: var(--vscode-descriptionForeground);
}
.df-table td {
  padding: 3px 10px;
  border-bottom: 1px solid var(--vscode-editorGroup-border);
  border-right: 1px solid var(--vscode-editorGroup-border, #c0c0c0);
}
.df-table td:last-child { border-right: none; }
.row-idx  { color: var(--vscode-descriptionForeground); font-size: 11px; }
.na-value { color: var(--vscode-descriptionForeground); font-style: italic; }

/* ---- RStudio-style paginator ---- */
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
  background: var(--vscode-toolbar-hoverBackground, rgba(128,128,128,.15));
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
  flex-shrink: 0;
  width: 116px; height: 78px;
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
}
.output-thumb:hover { border-color: #0078d4; box-shadow: 0 0 0 2px rgba(0,120,212,.15); }
.output-thumb.thumb-error { border-color: #e06c75; }
.output-thumb img { width: 100%; height: 100%; object-fit: contain; }
.thumb-icon { font-size: 22px; opacity: .5; line-height: 1; margin-bottom: 4px; }
.thumb-err-icon { color: #e06c75; opacity: .8; }
.thumb-label {
  font-size: 10px;
  color: var(--vscode-descriptionForeground);
  text-align: center;
  padding: 0 4px;
  max-width: 100%;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.thumb-text-preview {
  font-family: var(--vscode-editor-font-family, monospace);
  font-size: 7.5px;
  color: var(--vscode-editor-foreground);
  flex: 1; padding: 4px 6px 2px;
  width: 100%; overflow: hidden;
  white-space: pre; opacity: .65;
}
.thumb-stderr { opacity: .55; color: var(--vscode-inputValidation-warningForeground, #e5c07b); }
.thumb-dims { font-size: 9px; color: var(--vscode-descriptionForeground); opacity: .7; margin-top: 1px; }
`;
