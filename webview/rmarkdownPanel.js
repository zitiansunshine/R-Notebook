// =============================================================================
// rmarkdownPanel.js — RMarkdown editor WebView
// Editable chunks (click code to edit), R syntax highlighting, auto-save.
// =============================================================================

(function () {
  'use strict';

  const vscode      = acquireVsCodeApi();
  const container   = document.getElementById('rmd-container');
  const btnRunAll   = document.getElementById('btn-run-all');
  const btnInterrupt= document.getElementById('btn-interrupt');
  const btnReset    = document.getElementById('btn-reset');
  const btnRPath    = document.getElementById('btn-r-path');
  const btnAddChunk      = document.getElementById('btn-add-chunk');
  const btnAddChunkArrow = document.getElementById('btn-add-chunk-arrow');
  const btnDelChunk      = document.getElementById('btn-del-chunk');
  const btnLineNums      = document.getElementById('btn-line-nums');
  const btnVars          = document.getElementById('btn-vars');
  const btnVarsRefresh   = document.getElementById('btn-vars-refresh');
  const btnVarsClose     = document.getElementById('btn-vars-close');
  const varPanel         = document.getElementById('var-panel');
  const varTableBody     = document.getElementById('var-table-body');
  const statusEl         = document.getElementById('kernel-status');

  let currentChunks     = [];
  let autoSaveTimer     = null;
  let selectedChunkId   = null;
  let showLineNumbers   = true;   // line numbers visible by default
  let pendingCompletion = null;   // { ta, chunkId, cursorPos, codeView, gutterEl }
  let selectedAddLang   = 'r';   // currently chosen type for + Chunk button
  let varPanelOpen      = false;
  const activeOutputTabs = new Map();
  const outputScrollState = new Map(); // Track scroll freeze state per output element

  // ---- Smart Console Scroll Helper Functions --------------------------------

  /**
   * Check if an element is scrolled to the bottom.
   * Returns true if user is at or very near the bottom.
   */
  function isScrolledToBottom(el) {
    const threshold = 5; // Allow 5px tolerance for rounding
    return el.scrollHeight - el.scrollTop - el.clientHeight <= threshold;
  }

  /**
   * Scroll element to the bottom smoothly.
   */
  function scrollToBottom(el) {
    el.scrollTop = el.scrollHeight;
  }

  /**
   * Set up scroll behavior for an output element (console output window).
   * - By default, auto-scrolls to show latest output
   * - When user scrolls, freezes the scroll position
   * - When user clicks, resets to auto-scroll behavior
   */
  function initializeSmartScroll(outputEl) {
    if (!outputEl || outputScrollState.has(outputEl)) return; // Already initialized

    let isUserScrolling = false;
    let scrollFrozen = false;

    const state = { isUserScrolling, scrollFrozen };
    outputScrollState.set(outputEl, state);

    // Listen for user scrolling
    const scrollListener = () => {
      state.isUserScrolling = true;
      // Check if user scrolled away from bottom
      if (!isScrolledToBottom(outputEl)) {
        scrollFrozen = true;
        state.scrollFrozen = true;
      } else {
        // User scrolled back to bottom, resume auto-scroll
        scrollFrozen = false;
        state.scrollFrozen = false;
        state.isUserScrolling = false;
      }
    };

    // Listen for clicks to reset scroll behavior
    const clickListener = () => {
      scrollFrozen = false;
      state.scrollFrozen = false;
      state.isUserScrolling = false;
      // Auto-scroll to bottom when user clicks
      scrollToBottom(outputEl);
    };

    outputEl.addEventListener('scroll', scrollListener, { passive: true });
    outputEl.addEventListener('click', clickListener);

    // Store listeners for potential cleanup
    state.scrollListener = scrollListener;
    state.clickListener = clickListener;
  }

  /**
   * Auto-scroll console output if user hasn't frozen it.
   */
  function autoScrollIfNeeded(outputEl) {
    if (!outputEl) return;

    const state = outputScrollState.get(outputEl);
    if (!state || state.scrollFrozen) return; // Don't scroll if frozen

    scrollToBottom(outputEl);
  }

  /**
   * Remove scroll listeners from an output element.
   */
  function cleanupSmartScroll(outputEl) {
    if (!outputEl) return;

    const state = outputScrollState.get(outputEl);
    if (!state) return;

    outputEl.removeEventListener('scroll', state.scrollListener);
    outputEl.removeEventListener('click', state.clickListener);
    outputScrollState.delete(outputEl);
  }

  // ---- Messages from extension host ----------------------------------------

  window.addEventListener('message', event => {
    const msg = event.data;
    switch (msg.type) {

      case 'init':
        currentChunks = msg.chunks;
        renderAll(msg.chunks);
        syncOutputs(msg.outputs || {});
        break;

      case 'chunks_updated':
        currentChunks = msg.chunks;
        syncChunks(msg.chunks);
        if (msg.outputs) syncOutputs(msg.outputs);
        break;

      case 'chunk_running':
        setChunkState(msg.chunk_id, 'running');
        clearChunkOutput(msg.chunk_id);
        break;

      case 'chunk_progress': {
        const pb = container.querySelector(
          `[data-chunk-id="${msg.chunk_id}"] .chunk-progress-bar`
        );
        if (pb) {
          pb.classList.remove('indeterminate');
          const fill = pb.querySelector('.chunk-progress-bar-fill');
          if (fill) fill.style.height = `${(msg.line / msg.total) * 100}%`;
        }
        break;
      }

      case 'chunk_result': {
        // msg.error = JS exception; msg.result.error = R-level error — both mean error state
        const hasError = msg.error || (msg.result && msg.result.error);
        setChunkState(msg.chunk_id, hasError ? 'error' : 'done');
        if (msg.result) applyResult(msg.chunk_id, msg.result);
        if (msg.error)  applyError(msg.chunk_id, msg.error);
        // Auto-refresh variable inspector if open
        if (varPanelOpen) requestVars();
        break;
      }

      case 'chunk_stream':
        setChunkState(msg.chunk_id, 'running');
        if (msg.result) applyResult(msg.chunk_id, msg.result);
        break;

      case 'df_data':
        renderDfPage(msg);
        break;

      case 'session_reset':
        document.querySelectorAll('.chunk-output').forEach(el => { el.innerHTML = ''; });
        activeOutputTabs.clear();
        setStatus('idle');
        if (varPanelOpen) renderVarTable([]);
        break;

      case 'kernel_exit':
        setStatus('error');
        break;

      case 'kernel_error': {
        setStatus('error');
        statusEl.title = msg.message;
        const banner = document.createElement('div');
        banner.className = 'kernel-error-banner';
        banner.textContent = '⚠ R kernel: ' + msg.message + ' — check your R Path setting';
        container.prepend(banner);
        break;
      }

      case 'kernel_stderr':
        showToast(msg.text, 'warn');
        break;

      case 'completions_result':
        if (pendingCompletion && pendingCompletion.chunkId === msg.chunk_id) {
          showCompletionDropdown(pendingCompletion, msg.completions);
          pendingCompletion = null;
        }
        break;

      case 'vars_result':
        renderVarTable(msg.vars || []);
        // Stop spinning refresh button
        btnVarsRefresh.classList.remove('refreshing');
        break;
    }
  });

  // ---- Toolbar -------------------------------------------------------------

  btnRunAll.addEventListener('click', () => {
    vscode.postMessage({
      type: 'run_all',
      source: reconstructFullText(),
      chunks: collectCurrentChunks(),
    });
    setStatus('running');
  });

  btnInterrupt.addEventListener('click', () => {
    vscode.postMessage({ type: 'interrupt_kernel' });
  });

  btnReset.addEventListener('click', () => {
    vscode.postMessage({ type: 'reset_session' });
  });

  btnRPath.addEventListener('click', () => {
    vscode.postMessage({ type: 'set_r_path' });
  });

  function addChunk(language) {
    const insertAfterIdx = selectedChunkId
      ? currentChunks.findIndex(c => c.id === selectedChunkId)
      : currentChunks.length - 1;
    const insertIdx = insertAfterIdx >= 0 ? insertAfterIdx + 1 : currentChunks.length;
    const tempId = `new-${Date.now()}`;
    const newChunk = {
      id: tempId, kind: 'code', language,
      options: {}, code: '', prose: '', startLine: 0, endLine: 0,
    };
    currentChunks.splice(insertIdx, 0, newChunk);
    const newEl = makeChunkEl(newChunk);
    if (insertAfterIdx >= 0) {
      const refEl = container.querySelector(`[data-chunk-id="${currentChunks[insertAfterIdx].id}"]`);
      if (refEl) { refEl.after(newEl); }
      else container.appendChild(newEl);
    } else {
      container.appendChild(newEl);
    }
    selectChunk(newEl, tempId);
    const ta = newEl.querySelector('.code-textarea');
    if (ta) ta.focus();
    if (autoSaveTimer) clearTimeout(autoSaveTimer);
    postCodeChanged();
  }

  // Update + Chunk button label to reflect the currently selected language
  const LANG_SHORT = { r: 'R', bash: 'Bash', markdown: 'MD' };
  function updateAddBtn() {
    btnAddChunk.textContent = `+ ${LANG_SHORT[selectedAddLang] || selectedAddLang}`;
  }
  updateAddBtn();

  btnAddChunk.addEventListener('click', () => addChunk(selectedAddLang));

  btnAddChunkArrow.addEventListener('click', e => {
    e.stopPropagation();
    // Show selection menu — clicking an item ONLY updates the selection, does not add a chunk
    showChunkTypeMenu(btnAddChunkArrow, lang => {
      selectedAddLang = lang;
      updateAddBtn();
    }, selectedAddLang);
  });

  btnDelChunk.addEventListener('click', () => {
    if (!selectedChunkId) return;
    const idx = currentChunks.findIndex(c => c.id === selectedChunkId);
    if (idx < 0) return;
    currentChunks.splice(idx, 1);
    const el = container.querySelector(`[data-chunk-id="${selectedChunkId}"]`);
    if (el) el.remove();
    selectedChunkId = currentChunks.length > 0
      ? currentChunks[Math.min(idx, currentChunks.length - 1)].id
      : null;
    if (autoSaveTimer) clearTimeout(autoSaveTimer);
    postCodeChanged();
  });

  // ---- Variable inspector --------------------------------------------------

  function toggleVarPanel(open) {
    varPanelOpen = open;
    varPanel.classList.toggle('open', open);
    btnVars.classList.toggle('btn-active', open);
    if (open) requestVars();
  }

  function requestVars() {
    btnVarsRefresh.classList.add('refreshing');
    vscode.postMessage({ type: 'get_vars' });
  }

  btnVars.addEventListener('click', () => toggleVarPanel(!varPanelOpen));
  btnVarsRefresh.addEventListener('click', () => requestVars());
  btnVarsClose.addEventListener('click', () => toggleVarPanel(false));

  function renderVarTable(vars) {
    varTableBody.innerHTML = '';
    if (!vars || vars.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 4;
      td.className = 'var-empty';
      td.textContent = 'Environment is empty.';
      tr.appendChild(td);
      varTableBody.appendChild(tr);
      return;
    }
    for (const v of vars) {
      const tr = document.createElement('tr');

      const tdName = document.createElement('td');
      tdName.textContent = v.name;

      const tdType = document.createElement('td');
      const badge = document.createElement('span');
      badge.className = 'var-type-badge';
      badge.textContent = v.type || '?';
      tdType.appendChild(badge);

      const tdSize = document.createElement('td');
      tdSize.className = 'var-size';
      tdSize.textContent = v.size || '';

      const tdVal = document.createElement('td');
      tdVal.className = 'var-value';
      tdVal.textContent = v.value || '';
      tdVal.title = v.value || '';  // full value on hover

      tr.append(tdName, tdType, tdSize, tdVal);
      varTableBody.appendChild(tr);
    }
  }

  // Reflect default showLineNumbers=true in the toolbar button
  btnLineNums.classList.toggle('btn-active', showLineNumbers);

  btnLineNums.addEventListener('click', () => {
    showLineNumbers = !showLineNumbers;
    btnLineNums.classList.toggle('btn-active', showLineNumbers);
    document.querySelectorAll('.code-gutter').forEach(g => {
      g.classList.toggle('visible', showLineNumbers);
    });
  });

  // Track selected chunk on click
  container.addEventListener('click', e => {
    const chunkEl = e.target.closest('[data-chunk-id]');
    if (chunkEl && chunkEl.dataset.chunkId !== selectedChunkId) {
      selectChunk(chunkEl, chunkEl.dataset.chunkId);
    }
  });

  function selectChunk(el, id) {
    document.querySelectorAll('.chunk-selected').forEach(c => c.classList.remove('chunk-selected'));
    if (el) el.classList.add('chunk-selected');
    selectedChunkId = id;
  }

  // ---- Rendering -----------------------------------------------------------

  function renderAll(chunks) {
    container.innerHTML = '';
    for (const chunk of chunks) {
      container.appendChild(makeChunkEl(chunk));
    }
  }

  function syncChunks(chunks) {
    const existing = new Map(
      [...container.querySelectorAll('[data-chunk-id]')].map(e => [e.dataset.chunkId, e])
    );
    const incoming = new Map(chunks.map(c => [c.id, c]));

    // Remove stale
    for (const [id, el] of existing) {
      if (!incoming.has(id)) el.remove();
    }

    let prev = null;
    for (const chunk of chunks) {
      let el = existing.get(chunk.id);
      if (!el) {
        el = makeChunkEl(chunk);
        if (prev) prev.after(el);
        else container.prepend(el);
      } else {
        // Only update textarea/prose if it is NOT currently being edited
        const ta = el.querySelector('.code-textarea');
        if (ta && document.activeElement !== ta) {
          ta.value = chunk.code;
          const view = el.querySelector('.code-view');
          if (view) {
            view.innerHTML = highlightCode(chunk.language, chunk.code); // grid handles height
          } else {
            autoResizeTextarea(ta); // frontmatter: no grid, needs manual resize
          }
          const gutter = el.querySelector('.code-gutter');
          if (gutter) updateGutter(gutter, chunk.code);
        }
        const prose = el.querySelector('.prose-editor');
        if (prose && document.activeElement !== prose) {
          prose.dataset.plainText = chunk.prose;
          prose.innerHTML = renderMarkdown(chunk.prose);
        }
        // Sync language selector if language changed externally
        const langSel = el.querySelector('.lang-select');
        if (langSel && langSel.value !== (chunk.language || 'r')) {
          langSel.value = chunk.language || 'r';
          langSel.className = `lang-select lang-${chunk.language || 'r'}`;
        }
      }
      prev = el;
    }
  }

  function syncOutputs(outputs) {
    const saved = outputs || {};
    document.querySelectorAll('[data-chunk-id]').forEach(el => {
      const chunkId = el.dataset.chunkId;
      if (!saved[chunkId]) clearChunkOutput(chunkId);
    });
    for (const [id, result] of Object.entries(saved)) {
      setChunkState(id, result && result.error ? 'error' : 'done');
      applyResult(id, result);
    }
  }

  function makeChunkEl(chunk) {
    const wrapper = document.createElement('div');
    wrapper.className = `chunk chunk-${chunk.kind}`;
    wrapper.dataset.chunkId = chunk.id;

    if (chunk.kind === 'yaml_frontmatter') {
      return makeFrontmatterEl(wrapper, chunk);
    }
    if (chunk.kind === 'prose') {
      return makeProseEl(wrapper, chunk);
    }
    return makeCodeEl(wrapper, chunk);
  }

  // ---- Frontmatter (YAML header) -------------------------------------------

  function makeFrontmatterEl(wrapper, chunk) {
    const label = document.createElement('div');
    label.className = 'frontmatter-label';
    label.textContent = 'YAML Front Matter';

    const ta = document.createElement('textarea');
    ta.className = 'code-textarea frontmatter-textarea';
    ta.value = chunk.code;
    ta.spellcheck = false;
    autoResizeTextarea(ta);
    ta.addEventListener('input', () => { autoResizeTextarea(ta); scheduleAutoSave(); });

    wrapper.append(label, ta);
    return wrapper;
  }

  // ---- Prose chunk ---------------------------------------------------------

  function makeProseEl(wrapper, chunk) {
    const div = document.createElement('div');
    div.className = 'prose-editor';
    div.contentEditable = 'true';
    div.spellcheck = true;
    div.dataset.plainText = chunk.prose;
    div.innerHTML = renderMarkdown(chunk.prose);

    div.addEventListener('focus', () => {
      // Switch to plain text for editing
      div.textContent = div.dataset.plainText;
    });
    div.addEventListener('blur', () => {
      div.dataset.plainText = div.innerText;
      div.innerHTML = renderMarkdown(div.innerText);
      scheduleAutoSave();
    });
    div.addEventListener('input', () => scheduleAutoSave());

    wrapper.appendChild(div);
    return wrapper;
  }

  // ---- Chunk options helpers -----------------------------------------------

  /** Read current options from the DOM panel (falls back to chunk.options). */
  function getChunkOptions(el, chunk) {
    if (!el) return chunk.options;
    const opts = Object.assign({}, chunk.options);
    const labelHead = el.querySelector('.chunk-label');
    const labelIn = el.querySelector('.opt-label');
    const figwIn  = el.querySelector('.opt-figw');
    const fighIn  = el.querySelector('.opt-figh');
    const dpiIn   = el.querySelector('.opt-dpi');
    const evalIn  = el.querySelector('.opt-eval');
    const echoIn  = el.querySelector('.opt-echo');
    // Prefer the inline header label (contentEditable) as source of truth
    const labelVal = labelHead ? labelHead.textContent.trim() : (labelIn ? labelIn.value : '');
    if (labelVal) opts.label = labelVal;
    const fw = parseFloat(figwIn && figwIn.value);
    const fh = parseFloat(fighIn && fighIn.value);
    const dp = parseInt(dpiIn   && dpiIn.value, 10);
    if (!isNaN(fw)) opts.fig_width  = fw;
    if (!isNaN(fh)) opts.fig_height = fh;
    if (!isNaN(dp)) opts.dpi        = dp;
    if (evalIn) opts.eval = evalIn.checked;
    if (echoIn) opts.echo = echoIn.checked;
    return opts;
  }

  // ---- Code chunk ----------------------------------------------------------

  function makeCodeEl(wrapper, chunk) {
    // ---- Header
    const header = document.createElement('div');
    header.className = 'chunk-header';

    // Language selector — native <select> avoids accidental lang changes from floating menus
    const langSelect = document.createElement('select');
    langSelect.className = `lang-select lang-${chunk.language || 'r'}`;
    langSelect.title = 'Change block language';
    for (const { lang, label } of CHUNK_TYPE_OPTIONS) {
      const opt = document.createElement('option');
      opt.value = lang;
      opt.textContent = label;
      if (lang === (chunk.language || 'r')) opt.selected = true;
      langSelect.appendChild(opt);
    }
    langSelect.addEventListener('click', e => e.stopPropagation());
    langSelect.addEventListener('change', () => {
      const lang = langSelect.value;
      chunk.language = lang;
      langSelect.className = `lang-select lang-${lang}`;
      codeView.innerHTML = highlightCode(lang, ta.value);
      scheduleAutoSave();
    });

    const labelEl = document.createElement('span');
    labelEl.className = 'chunk-label';
    labelEl.textContent = chunk.options.label || '';
    labelEl.contentEditable = 'true';
    labelEl.spellcheck = false;
    labelEl.title = 'Click to rename chunk';
    labelEl.addEventListener('click', e => e.stopPropagation()); // don't select chunk
    labelEl.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); labelEl.blur(); }
    });
    labelEl.addEventListener('blur', () => {
      const newLabel = labelEl.textContent.trim();
      // Keep optsBar input in sync
      const li = wrapper.querySelector('.opt-label');
      if (li) li.value = newLabel;
      scheduleAutoSave();
    });

    const spinner = document.createElement('span');
    spinner.className = 'spinner hidden';
    spinner.textContent = '⟳';

    const spacer = document.createElement('span');
    spacer.className = 'header-spacer';

    // ⚙ options toggle (stays per-chunk)
    const optsBtn = document.createElement('button');
    optsBtn.className = 'opts-btn';
    optsBtn.title = 'Chunk options (label, fig size, eval…)';
    optsBtn.textContent = '⚙';

    const runBtn = document.createElement('button');
    runBtn.className = 'run-btn';
    runBtn.title = 'Run chunk (Shift+Enter)';
    runBtn.innerHTML = '&#9654; Run';
    runBtn.addEventListener('click', () => {
      const ta = wrapper.querySelector('.code-textarea');
      const code = ta ? ta.value : chunk.code;
      const options = getChunkOptions(wrapper, chunk);
      vscode.postMessage({
        type: 'run_chunk',
        source: reconstructFullText(),
        chunks: collectCurrentChunks(),
        chunk: { ...chunk, code, options },
      });
    });

    header.append(langSelect, labelEl, spinner, spacer, optsBtn, runBtn);

    // ---- Options bar (collapsible, hidden by default) ----------------------
    const optsBar = document.createElement('div');
    optsBar.className = 'chunk-opts-bar';

    function makeOptField(labelTxt, inputEl) {
      const lbl = document.createElement('label');
      lbl.className = 'opt-field';
      lbl.append(document.createTextNode(labelTxt + ' '), inputEl);
      return lbl;
    }

    const labelInput = document.createElement('input');
    labelInput.type = 'text'; labelInput.className = 'opt-label';
    labelInput.placeholder = 'label'; labelInput.value = chunk.options.label || '';
    labelInput.addEventListener('input', () => {
      labelEl.textContent = labelInput.value;
      scheduleAutoSave();
    });

    const figwInput = document.createElement('input');
    figwInput.type = 'number'; figwInput.className = 'opt-figw';
    figwInput.min = '1'; figwInput.max = '20'; figwInput.step = '0.5';
    figwInput.placeholder = '7';
    if (chunk.options.fig_width != null) figwInput.value = chunk.options.fig_width;
    figwInput.addEventListener('input', scheduleAutoSave);

    const fighInput = document.createElement('input');
    fighInput.type = 'number'; fighInput.className = 'opt-figh';
    fighInput.min = '1'; fighInput.max = '20'; fighInput.step = '0.5';
    fighInput.placeholder = '5';
    if (chunk.options.fig_height != null) fighInput.value = chunk.options.fig_height;
    fighInput.addEventListener('input', scheduleAutoSave);

    const dpiInput = document.createElement('input');
    dpiInput.type = 'number'; dpiInput.className = 'opt-dpi';
    dpiInput.min = '72'; dpiInput.max = '300'; dpiInput.step = '10';
    dpiInput.placeholder = '120';
    if (chunk.options.dpi != null) dpiInput.value = chunk.options.dpi;
    dpiInput.addEventListener('input', scheduleAutoSave);

    const evalInput = document.createElement('input');
    evalInput.type = 'checkbox'; evalInput.className = 'opt-eval';
    evalInput.checked = chunk.options.eval !== false;
    evalInput.addEventListener('change', scheduleAutoSave);

    const echoInput = document.createElement('input');
    echoInput.type = 'checkbox'; echoInput.className = 'opt-echo';
    echoInput.checked = chunk.options.echo !== false;
    echoInput.addEventListener('change', scheduleAutoSave);

    const evalLabel = document.createElement('label');
    evalLabel.className = 'opt-field';
    evalLabel.append(evalInput, ' eval');

    const echoLabel = document.createElement('label');
    echoLabel.className = 'opt-field';
    echoLabel.append(echoInput, ' echo');

    optsBar.append(
      makeOptField('Label', labelInput),
      makeOptField('fig.width', figwInput),
      makeOptField('fig.height', fighInput),
      makeOptField('DPI', dpiInput),
      evalLabel,
      echoLabel,
    );

    optsBtn.addEventListener('click', e => {
      e.stopPropagation();
      optsBar.classList.toggle('open');
    });

    // ---- Code editor: absolute-overlay approach for live highlighting ------
    const editorWrap = document.createElement('div');
    editorWrap.className = 'code-editor-wrap';

    // Progress bar (left strip, shown while running)
    const progressBar = document.createElement('div');
    progressBar.className = 'chunk-progress-bar';
    const progressFill = document.createElement('div');
    progressFill.className = 'chunk-progress-bar-fill';
    progressBar.appendChild(progressFill);

    // Line-number gutter
    const gutterEl = document.createElement('div');
    gutterEl.className = 'code-gutter';
    if (showLineNumbers) gutterEl.classList.add('visible');
    updateGutter(gutterEl, chunk.code);

    // Inner container (position: relative so code-view can be absolute)
    const editorMain = document.createElement('div');
    editorMain.className = 'code-editor-main';

    // Highlighted code view (positioned absolutely behind textarea)
    const codeView = document.createElement('pre');
    codeView.className = 'code-view';
    codeView.innerHTML = highlightCode(chunk.language, chunk.code);

    // Editable textarea (transparent text; highlight shows through).
    // The CSS grid layout makes code-view drive the row height automatically —
    // no JS auto-resize needed for code chunks.
    const ta = document.createElement('textarea');
    ta.className = 'code-textarea';
    ta.value = chunk.code;
    ta.rows = 1;          // minimise textarea's own grid-sizing contribution
    ta.spellcheck = false;

    // Live highlighting + gutter on every keystroke
    ta.addEventListener('input', () => {
      codeView.innerHTML = highlightCode(chunk.language, ta.value);
      updateGutter(gutterEl, ta.value);
      scheduleAutoSave();
    });

    // Full editor keyboard shortcuts
    ta.addEventListener('keydown', e => {
      const mod = e.metaKey || e.ctrlKey;

      // ── Escape ──────────────────────────────────────────────────────────
      if (e.key === 'Escape') { removeCompletionDropdown(); return; }

      // ── Completion dropdown navigation ───────────────────────────────────
      const dd = wrapper.querySelector('.completion-dropdown');
      if (dd && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        e.preventDefault();
        navigateDropdown(dd, e.key === 'ArrowDown' ? 1 : -1);
        return;
      }
      if (dd && (e.key === 'Enter' || (e.key === 'Tab' && !e.shiftKey))) {
        const act = dd.querySelector('.completion-item.active');
        if (act) { e.preventDefault(); act.dispatchEvent(new MouseEvent('mousedown')); return; }
      }

      // ── Shift+Enter: run chunk, then jump to next code chunk ────────────
      if (e.key === 'Enter' && e.shiftKey) {
        e.preventDefault(); e.stopImmediatePropagation();
        runBtn.click();
        const curIdx = currentChunks.findIndex(c => c.id === chunk.id);
        const next = currentChunks.slice(curIdx + 1).find(c => c.kind === 'code');
        if (next) {
          const nextEl = container.querySelector(`[data-chunk-id="${next.id}"]`);
          const nextTa = nextEl && nextEl.querySelector('.code-textarea');
          if (nextTa) {
            nextTa.focus();
            nextEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
          }
        }
        return;
      }

      if (mod) {
        // ── Cmd+A / Ctrl+A: select all text in this textarea ──────────────
        if (e.key === 'a') {
          e.preventDefault(); e.stopImmediatePropagation();
          ta.setSelectionRange(0, ta.value.length); return;
        }

        // ── Cmd+C / Ctrl+C: copy ─────────────────────────────────────────
        if (e.key === 'c' && !e.shiftKey) {
          e.preventDefault(); e.stopImmediatePropagation();
          const txt = ta.value.slice(ta.selectionStart, ta.selectionEnd);
          if (txt) navigator.clipboard.writeText(txt).catch(() => {});
          return;
        }

        // ── Cmd+X / Ctrl+X: cut ──────────────────────────────────────────
        if (e.key === 'x') {
          e.preventDefault(); e.stopImmediatePropagation();
          const s = ta.selectionStart, en = ta.selectionEnd;
          const txt = ta.value.slice(s, en);
          if (txt) {
            navigator.clipboard.writeText(txt).catch(() => {});
            // execCommand deletes selection and preserves undo stack
            document.execCommand('insertText', false, '');
            codeView.innerHTML = highlightCode(chunk.language, ta.value);
            updateGutter(gutterEl, ta.value);
            scheduleAutoSave();
          }
          return;
        }

        // ── Cmd+V / Ctrl+V: paste ─────────────────────────────────────────
        if (e.key === 'v') {
          e.preventDefault(); e.stopImmediatePropagation();
          navigator.clipboard.readText().then(txt => {
            insertAtCursor(ta, txt);
            codeView.innerHTML = highlightCode(chunk.language, ta.value);
            updateGutter(gutterEl, ta.value);
            scheduleAutoSave();
          }).catch(() => {}); // clipboard permission denied → silently ignore
          return;
        }

        // ── Cmd+Z / Ctrl+Z: undo  |  Cmd+Shift+Z / Ctrl+Y: redo ─────────
        if (e.key === 'z') {
          e.preventDefault(); e.stopImmediatePropagation();
          document.execCommand(e.shiftKey ? 'redo' : 'undo');
          codeView.innerHTML = highlightCode(chunk.language, ta.value);
          updateGutter(gutterEl, ta.value);
          return;
        }
        if (e.key === 'y') {
          e.preventDefault(); e.stopImmediatePropagation();
          document.execCommand('redo');
          codeView.innerHTML = highlightCode(chunk.language, ta.value);
          updateGutter(gutterEl, ta.value);
          return;
        }

        // ── Cmd+/ or Ctrl+/: toggle line comment ─────────────────────────
        if (e.key === '/') {
          e.preventDefault(); e.stopImmediatePropagation();
          toggleComment(ta);
          codeView.innerHTML = highlightCode(chunk.language, ta.value);
          updateGutter(gutterEl, ta.value);
          scheduleAutoSave(); return;
        }

        // ── Cmd+D / Ctrl+D: duplicate current line ────────────────────────
        if (e.key === 'd' && !e.shiftKey) {
          e.preventDefault(); e.stopImmediatePropagation();
          duplicateLine(ta);
          codeView.innerHTML = highlightCode(chunk.language, ta.value);
          updateGutter(gutterEl, ta.value);
          scheduleAutoSave(); return;
        }

        // For all other Cmd/Ctrl combos, stop propagation so Cursor doesn't
        // intercept them (e.g. Cmd+ArrowLeft for line-start navigation).
        e.stopImmediatePropagation();
        return;
      }

      // ── Tab / Shift+Tab: indent / unindent ────────────────────────────────
      if (e.key === 'Tab') {
        e.preventDefault();
        if (dd && !e.shiftKey) {
          const act = dd.querySelector('.completion-item.active');
          if (act) {
            act.dispatchEvent(new MouseEvent('mousedown'));
            return;
          }
        }
        removeCompletionDropdown();
        if (e.shiftKey) {
          unindentLines(ta);
        } else {
          const cursor = ta.selectionStart;
          const before = ta.value.slice(0, cursor);
          if (ta.selectionEnd > cursor && ta.value.slice(cursor, ta.selectionEnd).includes('\n')) {
            indentLines(ta);
          } else if (/\w$/.test(before)) {
            pendingCompletion = { ta, chunkId: chunk.id, cursorPos: cursor, codeView, gutterEl, wrapper };
            vscode.postMessage({ type: 'get_completions', chunk_id: chunk.id, code: ta.value, cursor_pos: cursor });
            return; // wait for completions_result message
          } else {
            insertAtCursor(ta, '  ');
          }
        }
        codeView.innerHTML = highlightCode(chunk.language, ta.value);
        updateGutter(gutterEl, ta.value);
        scheduleAutoSave(); return;
      }

      // ── Enter: auto-indent (plain Enter without any modifier) ────────────
      if (e.key === 'Enter' && !e.altKey) {
        e.preventDefault();
        const cursor  = ta.selectionStart;
        const before  = ta.value.slice(0, cursor);
        const lastNL  = before.lastIndexOf('\n');
        const curLine = lastNL >= 0 ? before.slice(lastNL + 1) : before;
        const indent  = curLine.match(/^(\s*)/)[1];
        insertAtCursor(ta, '\n' + indent);
        codeView.innerHTML = highlightCode(chunk.language, ta.value);
        updateGutter(gutterEl, ta.value);
        scheduleAutoSave(); return;
      }
    });

    // Dismiss dropdown on blur
    ta.addEventListener('blur', () => setTimeout(removeCompletionDropdown, 150));

    editorMain.append(codeView, ta);
    editorWrap.append(progressBar, gutterEl, editorMain);

    // ---- Outer bordered box (encloses header + options + editor on all sides)
    const codeBlock = document.createElement('div');
    codeBlock.className = 'chunk-code-block';
    codeBlock.append(header, optsBar, editorWrap);

    const output = document.createElement('div');
    output.className = 'chunk-output';

    wrapper.append(codeBlock, output);
    return wrapper;
  }

  // ---- Auto-save -----------------------------------------------------------

  function scheduleAutoSave() {
    if (autoSaveTimer) clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(() => {
      postCodeChanged();
    }, 600);
  }

  function postCodeChanged() {
    vscode.postMessage({
      type: 'code_changed',
      fullText: reconstructFullText(),
      chunks: collectCurrentChunks(),
    });
  }

  function collectCurrentChunks() {
    return currentChunks.map(chunk => {
      const el = container.querySelector(`[data-chunk-id="${chunk.id}"]`);
      const ta = el && el.querySelector('.code-textarea');
      if (chunk.kind === 'code' && ta) {
        return { ...chunk, code: ta.value, options: getChunkOptions(el, chunk) };
      }
      return chunk;
    });
  }

  function reconstructFullText() {
    return currentChunks.map(chunk => {
      const el = container.querySelector(`[data-chunk-id="${chunk.id}"]`);

      if (chunk.kind === 'yaml_frontmatter') {
        const ta = el && el.querySelector('.code-textarea');
        return `---\n${ta ? ta.value : chunk.code}\n---`;
      }

      if (chunk.kind === 'prose') {
        const div = el && el.querySelector('.prose-editor');
        return div ? (div.dataset.plainText || div.innerText) : chunk.prose;
      }

      // code chunk
      const ta = el && el.querySelector('.code-textarea');
      const code = ta ? ta.value : chunk.code;
      const opts = getChunkOptions(el, chunk);
      const optParts = [];
      if (opts.label) optParts.push(opts.label);
      for (const [k, v] of Object.entries(opts)) {
        if (k === 'label') continue;
        if (v == null) continue;
        const rk = k.replace(/_/g, '.');
        if (typeof v === 'boolean') {
          // only write non-default booleans (eval=FALSE, echo=FALSE)
          if (!v) optParts.push(`${rk}=FALSE`);
        } else {
          optParts.push(`${rk}=${v}`);
        }
      }
      const header = `\`\`\`{${chunk.language}${optParts.length ? ' ' + optParts.join(', ') : ''}}`;
      return `${header}\n${code}\n\`\`\``;
    }).join('\n');
  }

  // ---- Tab completion helpers ---------------------------------------------

  function insertAtCursor(ta, text) {
    // execCommand preserves the browser's native undo stack (unlike ta.value=)
    document.execCommand('insertText', false, text);
  }

  function removeCompletionDropdown() {
    document.querySelectorAll('.completion-dropdown').forEach(el => el.remove());
  }

  function navigateDropdown(dd, dir) {
    const items = [...dd.querySelectorAll('.completion-item')];
    const cur = items.findIndex(el => el.classList.contains('active'));
    if (items.length === 0) return;
    items.forEach(el => el.classList.remove('active'));
    const next = (cur + dir + items.length) % items.length;
    items[next].classList.add('active');
    items[next].scrollIntoView({ block: 'nearest' });
  }

  function showCompletionDropdown({ ta, chunkId, cursorPos, codeView, gutterEl, wrapper }, completions) {
    removeCompletionDropdown();
    const language = (wrapper.querySelector('.lang-select') || {}).value || 'r';
    if (!completions || completions.length === 0) {
      // No completions — fall back to inserting 2 spaces
      insertAtCursor(ta, '  ');
      codeView.innerHTML = highlightCode(language, ta.value);
      updateGutter(gutterEl, ta.value);
      return;
    }
    // Find what token was before the cursor to calculate replacement range
    const before = ta.value.slice(0, cursorPos);
    const tokenMatch = before.match(/[\w.]+$/);
    const tokenLen = tokenMatch ? tokenMatch[0].length : 0;
    const tokenStart = cursorPos - tokenLen;

    const dd = document.createElement('div');
    dd.className = 'completion-dropdown';

    completions.slice(0, 12).forEach((comp, i) => {
      const item = document.createElement('div');
      item.className = 'completion-item' + (i === 0 ? ' active' : '');
      item.textContent = comp;
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        // e.preventDefault() keeps focus on textarea so execCommand works
        ta.focus();
        ta.setSelectionRange(tokenStart, cursorPos);
        document.execCommand('insertText', false, comp);
        codeView.innerHTML = highlightCode(language, ta.value);
        updateGutter(gutterEl, ta.value);
        scheduleAutoSave();
        removeCompletionDropdown();
      });
      dd.appendChild(item);
    });

    // Position the dropdown below the code editor
    const editorMain = wrapper.querySelector('.code-editor-main');
    if (editorMain) {
      editorMain.style.position = 'relative';
      editorMain.appendChild(dd);
    }
  }

  // ---- Editor keyboard-shortcut helpers ------------------------------------

  /** Cmd+/ — toggle # comment on each selected line */
  function toggleComment(ta) {
    const val = ta.value;
    const s   = ta.selectionStart;
    const e   = ta.selectionEnd;
    const ls  = val.lastIndexOf('\n', s - 1) + 1;
    const ae  = val.indexOf('\n', e);
    const le  = ae < 0 ? val.length : ae;
    const lines = val.slice(ls, le).split('\n');
    const allCommented = lines.filter(l => l.trim()).every(l => /^[ \t]*#/.test(l));
    let newStart = s, newEnd = e;
    const newLines = lines.map((l, i) => {
      if (allCommented) {
        const stripped = l.replace(/^([ \t]*)# ?/, '$1');
        const delta = stripped.length - l.length;
        if (i === 0) newStart = Math.max(ls, s + delta);
        newEnd += delta;
        return stripped;
      } else {
        if (i === 0) newStart = s + 2;
        newEnd += 2;
        return '# ' + l;
      }
    });
    const newText = newLines.join('\n');
    ta.setSelectionRange(ls, le);
    document.execCommand('insertText', false, newText);
    ta.setSelectionRange(Math.max(ls, newStart), Math.max(newStart, newEnd));
  }

  /** Tab on multi-line selection — indent each line by 2 spaces */
  function indentLines(ta) {
    const val = ta.value;
    const ls  = val.lastIndexOf('\n', ta.selectionStart - 1) + 1;
    const ae  = val.indexOf('\n', ta.selectionEnd);
    const le  = ae < 0 ? val.length : ae;
    const lines    = val.slice(ls, le).split('\n');
    const newLines = lines.map(l => '  ' + l);
    const newText  = newLines.join('\n');
    ta.setSelectionRange(ls, le);
    document.execCommand('insertText', false, newText);
    ta.setSelectionRange(ls, ls + newText.length);
  }

  /** Shift+Tab — remove up to 2 leading spaces from each selected line */
  function unindentLines(ta) {
    const val = ta.value;
    const s   = ta.selectionStart;
    const e   = ta.selectionEnd;
    const ls  = val.lastIndexOf('\n', s - 1) + 1;
    const ae  = val.indexOf('\n', e);
    const le  = ae < 0 ? val.length : ae;
    const lines    = val.slice(ls, le).split('\n');
    let totalDelta = 0;
    const newLines = lines.map((l, i) => {
      const m  = l.match(/^( {1,2})/);
      const removed = m ? m[1].length : 0;
      if (i === 0) totalDelta -= removed;
      return removed ? l.slice(removed) : l;
    });
    const allDelta = newLines.join('\n').length - lines.join('\n').length;
    const newText = newLines.join('\n');
    ta.setSelectionRange(ls, le);
    document.execCommand('insertText', false, newText);
    const newStart = Math.max(ls, s + totalDelta);
    ta.setSelectionRange(newStart, Math.max(newStart, e + allDelta));
  }

  /** Cmd+D — duplicate current line */
  function duplicateLine(ta) {
    const val = ta.value;
    const s   = ta.selectionStart;
    const ls  = val.lastIndexOf('\n', s - 1) + 1;
    const ae  = val.indexOf('\n', s);
    const le  = ae < 0 ? val.length : ae;
    const line = val.slice(ls, le);
    const newCursor = le + 1 + (s - ls);
    ta.setSelectionRange(le, le);
    document.execCommand('insertText', false, '\n' + line);
    ta.setSelectionRange(newCursor, newCursor);
  }

  function updateGutter(gutterEl, code) {
    if (!gutterEl) return;
    const n = (code.match(/\n/g) || []).length + 1;
    let s = '';
    for (let i = 1; i <= n; i++) s += i + '\n';
    gutterEl.textContent = s;
  }

  function autoResizeTextarea(ta) {
    ta.style.height = 'auto';
    ta.style.height = Math.max(60, ta.scrollHeight) + 'px';
  }

  // ---- Output rendering ----------------------------------------------------

  function applyResult(chunkId, result) {
    const outputEl = container.querySelector(`[data-chunk-id="${chunkId}"] .chunk-output`);
    if (!outputEl) return;

    // Preserve scroll freeze state across streaming re-renders (chunk_stream replaces content)
    const oldConsoleEl = outputEl.querySelector('.output-text');
    const wasFrozen = oldConsoleEl
      ? (outputScrollState.get(oldConsoleEl)?.scrollFrozen === true)
      : false;
    if (oldConsoleEl) cleanupSmartScroll(oldConsoleEl);

    outputEl.innerHTML = '';

    const tabs = collectResultTabs(result);

    if (tabs.length === 0) return;

    if (tabs.length === 1) {
      // Single output — render directly, no thumbnail strip
      outputEl.appendChild(makeSingleOutput(chunkId, tabs[0]));
      // Initialize smart scroll for console/text outputs
      const consoleEl = outputEl.querySelector('.output-text');
      if (consoleEl) {
        initializeSmartScroll(consoleEl);
        if (wasFrozen) {
          // User had scrolled up during streaming — restore frozen state
          const st = outputScrollState.get(consoleEl);
          if (st) st.scrollFrozen = true;
        } else {
          // Defer scroll until after browser layout (element may not be sized yet)
          requestAnimationFrame(() => autoScrollIfNeeded(consoleEl));
        }
      }
    } else {
      // Multiple outputs — RStudio-style thumbnail strip + viewer
      const tabsWrap = makeOutputTabs(chunkId, tabs);
      outputEl.appendChild(tabsWrap);
      // Scroll initial console tab to bottom AFTER DOM insertion
      const consoleEl = outputEl.querySelector('.output-tab-main .output-text');
      if (consoleEl) {
        if (!wasFrozen) {
          requestAnimationFrame(() => autoScrollIfNeeded(consoleEl));
        }
      }
    }
  }

  /** Build a thumbnail-strip + main viewer for multiple outputs. */
  function makeOutputTabs(chunkId, tabs) {
    const wrap = document.createElement('div');
    wrap.className = 'output-tabs';

    const strip = document.createElement('div');
    strip.className = 'output-thumb-strip';
    if (tabs.length >= 8) {
      strip.classList.add('output-thumb-strip-grid');
    }

    const mainArea = document.createElement('div');
    mainArea.className = 'output-tab-main';

    function showTab(idx) {
      const tab = tabs[idx];
      if (!tab) return;
      activeOutputTabs.set(chunkId, tabStateKey(tab, idx));
      strip.querySelectorAll('.output-thumb').forEach((th, i) =>
        th.classList.toggle('output-thumb-active', i === idx));

      // Clean up previous console scroll state if switching tabs
      const oldConsoleEl = mainArea.querySelector('.output-text');
      if (oldConsoleEl) cleanupSmartScroll(oldConsoleEl);

      mainArea.innerHTML = '';
      mainArea.appendChild(makeSingleOutput(chunkId, tab));

      // Initialize smart scroll for console/text outputs in tab view.
      // Use requestAnimationFrame so scroll runs after the element is in the DOM
      // (showTab may be called before wrap is appended — scrollHeight would be 0).
      const consoleEl = mainArea.querySelector('.output-text');
      if (consoleEl) {
        initializeSmartScroll(consoleEl);
        requestAnimationFrame(() => autoScrollIfNeeded(consoleEl));
      }
    }

    tabs.forEach((tab, i) => {
      const thumb = makeOutputThumb(tab, i);
      thumb.addEventListener('click', () => showTab(i));
      strip.appendChild(thumb);
    });

    wrap.append(strip, mainArea);
    showTab(initialTabIndex(chunkId, tabs));
    return wrap;
  }

  /** Build a single thumbnail card for the strip. */
  function makeOutputThumb(tab, idx) {
    const card = document.createElement('div');
    card.className = 'output-thumb';

    if (tab.type === 'console') {
      const pre = document.createElement('pre');
      pre.className = 'thumb-text-preview';
      pre.textContent = previewConsoleText(tab.content);
      const label = document.createElement('div');
      label.className = 'thumb-label';
      label.textContent = 'Console';
      card.append(pre, label);
    } else if (tab.type === 'plot') {
      const img = document.createElement('img');
      img.src = `data:image/png;base64,${tab.content}`;
      img.alt = `Plot ${idx + 1}`;
      card.appendChild(img);
    } else if (tab.type === 'df') {
      const icon  = document.createElement('div');
      icon.className  = 'thumb-icon';
      icon.textContent = '⊞';
      const label = document.createElement('div');
      label.className = 'thumb-label';
      label.textContent = friendlyDfName(tab.content.name || 'DataFrame');
      const dims = document.createElement('div');
      dims.className = 'thumb-dims';
      dims.textContent = `${tab.content.nrow} × ${tab.content.ncol}`;
      card.append(icon, label, dims);
    } else if (tab.type === 'text') {
      const pre = document.createElement('pre');
      pre.className = 'thumb-text-preview';
      pre.textContent = tab.content.split('\n').slice(0, 4).join('\n');
      const label = document.createElement('div');
      label.className = 'thumb-label';
      label.textContent = 'Console';
      card.append(pre, label);
    } else if (tab.type === 'stderr') {
      const pre = document.createElement('pre');
      pre.className = 'thumb-text-preview thumb-stderr';
      pre.textContent = tab.content.split('\n').slice(0, 4).join('\n');
      const label = document.createElement('div');
      label.className = 'thumb-label';
      label.textContent = 'Stderr';
      card.append(pre, label);
    } else if (tab.type === 'error') {
      const icon = document.createElement('div');
      icon.className = 'thumb-icon thumb-err-icon';
      icon.textContent = '✖';
      const label = document.createElement('div');
      label.className = 'thumb-label';
      label.textContent = 'Error';
      card.append(icon, label);
      card.classList.add('thumb-error');
    }

    return card;
  }

  /** Render a single output node (fragment) for the main viewer area. */
  function makeSingleOutput(chunkId, tab) {
    const frag = document.createDocumentFragment();
    if (tab.type === 'console' || tab.type === 'text') {
      const pre = document.createElement('pre');
      pre.className = 'output-text';
      pre.textContent = tab.content;
      frag.appendChild(pre);
    } else if (tab.type === 'stderr') {
      const pre = document.createElement('pre');
      pre.className = 'output-stderr';
      pre.textContent = tab.content;
      frag.appendChild(pre);
    } else if (tab.type === 'plot') {
      const plotWrap = document.createElement('div');
      plotWrap.className = 'plot-wrap';
      const img = document.createElement('img');
      img.className = 'output-plot';
      img.src = `data:image/png;base64,${tab.content}`;
      img.alt = 'Plot';
      const dlBtn = document.createElement('a');
      dlBtn.className = 'plot-dl-btn';
      dlBtn.textContent = '⬇ Save PNG';
      dlBtn.href = `data:image/png;base64,${tab.content}`;
      dlBtn.download = `plot-${Date.now()}.png`;
      plotWrap.append(img, dlBtn);
      frag.appendChild(plotWrap);
    } else if (tab.type === 'df') {
      frag.appendChild(makeDfViewer(chunkId, tab.content));
    } else if (tab.type === 'error') {
      const pre = document.createElement('pre');
      pre.className = 'output-error';
      pre.textContent = '✖ ' + tab.content;
      frag.appendChild(pre);
    }
    return frag;
  }

  function applyError(chunkId, errorText) {
    const outputEl = container.querySelector(`[data-chunk-id="${chunkId}"] .chunk-output`);
    if (!outputEl) return;

    // Clean up previous scroll state
    const oldConsoleEl = outputEl.querySelector('.output-text');
    if (oldConsoleEl) cleanupSmartScroll(oldConsoleEl);

    outputEl.innerHTML = '';
    const pre = document.createElement('pre');
    pre.className = 'output-error';
    pre.textContent = '✖ ' + errorText;
    outputEl.appendChild(pre);
  }

  function clearChunkOutput(chunkId) {
    const outputEl = container.querySelector(`[data-chunk-id="${chunkId}"] .chunk-output`);
    if (outputEl) {
      // Clean up scroll state
      const consoleEl = outputEl.querySelector('.output-text');
      if (consoleEl) cleanupSmartScroll(consoleEl);
      outputEl.innerHTML = '';
    }
    activeOutputTabs.delete(chunkId);
  }

  function initialTabIndex(chunkId, tabs) {
    const activeKey = activeOutputTabs.get(chunkId);
    if (!activeKey) return 0;
    const index = tabs.findIndex((tab, idx) => tabStateKey(tab, idx) === activeKey);
    return index >= 0 ? index : 0;
  }

  function tabStateKey(tab, idx) {
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
        return `df:${tab.content.name || idx}`;
      case 'plot':
        return `plot:${tab.content.length}:${tab.content.slice(0, 32)}`;
      default:
        return `${tab.type}:${idx}`;
    }
  }

  function collectResultTabs(result) {
    const tabs = [];
    const consoleText = getConsoleText(result);
    if (consoleText.trim()) {
      tabs.push({ type: 'console', content: consoleText });
    }

    const plots = Array.isArray(result.plots) ? result.plots : [];
    const dataframes = Array.isArray(result.dataframes) ? result.dataframes : [];
    const outputOrder = Array.isArray(result.output_order) ? result.output_order : [];
    const usedPlots = new Set();
    const usedDataframes = new Set();

    if (outputOrder.length > 0) {
      outputOrder.forEach(item => {
        if (item.type === 'plot') {
          const plot = plots[item.index];
          if (plot === undefined) return;
          tabs.push({ type: 'plot', content: plot });
          usedPlots.add(item.index);
          return;
        }

        const df = dataframes[item.index];
        if (!df) return;
        tabs.push({ type: 'df', content: df });
        usedDataframes.add(item.index);
      });

      dataframes.forEach((df, index) => {
        if (!usedDataframes.has(index)) {
          tabs.push({ type: 'df', content: df });
        }
      });
      plots.forEach((plot, index) => {
        if (!usedPlots.has(index)) {
          tabs.push({ type: 'plot', content: plot });
        }
      });
    } else {
      dataframes.forEach(df => tabs.push({ type: 'df', content: df }));
      plots.forEach(plot => tabs.push({ type: 'plot', content: plot }));
    }

    if (result.error) {
      tabs.push({ type: 'error', content: result.error });
    }

    return tabs;
  }

  function getConsoleText(result) {
    if (result.console && result.console.trim()) {
      return result.console;
    }

    const parts = [result.stdout || '', result.stderr || '']
      .filter(part => part.trim().length > 0);
    return parts.join(parts.length > 1 ? '\n' : '');
  }

  function previewConsoleText(text) {
    const lines = text.split(/\r?\n/);
    return lines.slice(Math.max(0, lines.length - 4)).join('\n');
  }

  // ---- DataFrame viewer ---------------------------------------------------

  /** Return a readable display name for a data frame.
   *  For function-call expressions, show only the outermost function name
   *  and the first argument, e.g. "head(mtcars)" stays as-is, but a deeply
   *  nested chain is trimmed to ≤ 40 chars with "…".
   */
  function friendlyDfName(name) {
    if (!name || name.length <= 40) return name;
    return name.slice(0, 37) + '…';
  }

  function makeDfViewer(chunkId, df) {
    const wrap = document.createElement('div');
    wrap.className = 'df-viewer';
    wrap.dataset.dfName    = df.name;
    wrap.dataset.chunkId   = chunkId;
    wrap.dataset.currentPage = '0';
    wrap.dataset.totalPages  = String(df.pages);

    const titleBar = document.createElement('div');
    titleBar.className = 'df-title';
    const displayName = friendlyDfName(df.name);
    titleBar.innerHTML = `<strong>${esc(displayName)}</strong> <span class="df-dims">${df.nrow} × ${df.ncol}</span>`;

    const tableWrap = document.createElement('div');
    tableWrap.className = 'df-table-wrap';
    tableWrap.appendChild(buildDfTable(df));

    const paginator = document.createElement('div');
    paginator.className = 'df-paginator';
    updatePaginator(paginator, df);

    paginator.addEventListener('click', e => {
      const btn = e.target.closest('button[data-page]');
      if (!btn) return;
      const page = parseInt(btn.dataset.page, 10);
      vscode.postMessage({ type: 'df_page', chunk_id: chunkId, name: df.name, page, page_size: 50 });
      wrap.dataset.currentPage = String(page);
    });

    wrap.append(titleBar, tableWrap, paginator);
    return wrap;
  }

  function buildDfTable(df) {
    const table = document.createElement('table');
    table.className = 'df-table';
    const thead = table.createTHead();
    const hrow  = thead.insertRow();
    hrow.insertCell().textContent = '';
    for (const col of df.columns) {
      const th = document.createElement('th');
      th.textContent = col.name;
      th.title = col.type;
      hrow.appendChild(th);
    }
    const tbody = table.createTBody();
    const startIdx = df.page * 50 + 1;
    df.data.forEach((row, i) => {
      const tr = tbody.insertRow();
      const idx = tr.insertCell();
      idx.className = 'row-idx';
      idx.textContent = String(startIdx + i);
      for (const val of Object.values(row)) {
        const td = tr.insertCell();
        td.textContent = val === null ? 'NA' : String(val);
        if (val === null) td.className = 'na-value';
      }
    });
    return table;
  }

  function renderDfPage(msg) {
    const wrap = container.querySelector(
      `[data-chunk-id="${msg.chunk_id}"][data-df-name="${msg.name}"]`
    ) || container.querySelector(`[data-df-name="${msg.name}"]`);
    if (!wrap) return;
    wrap.dataset.currentPage = String(msg.page);
    wrap.dataset.totalPages  = String(msg.pages);
    const tableWrap = wrap.querySelector('.df-table-wrap');
    tableWrap.innerHTML = '';
    tableWrap.appendChild(buildDfTable(msg));
    updatePaginator(wrap.querySelector('.df-paginator'), msg);
  }

  function updatePaginator(el, df) {
    el.innerHTML = '';
    const current = df.page, total = df.pages;
    const info = document.createElement('span');
    info.className = 'page-info';
    info.textContent = `Page ${current + 1} of ${total}`;
    const prev = document.createElement('button');
    prev.textContent = '‹ Prev';
    prev.dataset.page = String(Math.max(0, current - 1));
    prev.disabled = current === 0;
    const next = document.createElement('button');
    next.textContent = 'Next ›';
    next.dataset.page = String(Math.min(total - 1, current + 1));
    next.disabled = current >= total - 1;
    el.append(prev, next, info);
  }

  // ---- Chunk state ---------------------------------------------------------

  function setChunkState(id, state) {
    const el = container.querySelector(`[data-chunk-id="${id}"]`);
    if (!el) return;
    el.classList.remove('state-running', 'state-done', 'state-error');
    el.classList.add(`state-${state}`);
    const sp = el.querySelector('.spinner');
    if (sp) sp.classList.toggle('hidden', state !== 'running');

    const pb   = el.querySelector('.chunk-progress-bar');
    const fill = pb && pb.querySelector('.chunk-progress-bar-fill');
    if (pb) {
      if (state === 'running') {
        if (fill) fill.style.height = '0%';
        pb.classList.add('active', 'indeterminate');
      } else {
        // Animate to 100% then hide
        pb.classList.remove('indeterminate');
        if (fill) fill.style.height = '100%';
        setTimeout(() => pb.classList.remove('active'), 350);
      }
    }
  }

  function setStatus(state) {
    statusEl.className = `status-${state}`;
    statusEl.textContent = state === 'running' ? '● Running'
                         : state === 'error'   ? '● Error'
                         :                       '● Idle';
  }

  // ---- Chunk-type menu (toolbar ▾ + lang badge) ----------------------------

  const CHUNK_TYPE_OPTIONS = [
    { lang: 'r',        icon: '{ }', label: 'R' },
    { lang: 'bash',     icon: '$',   label: 'Bash' },
    { lang: 'markdown', icon: '#',   label: 'Markdown' },
  ];

  /**
   * Show a floating dropdown to pick a chunk language.
   * @param {HTMLElement} anchor    - element to position below
   * @param {Function}    onPick    - called with lang string when selected
   * @param {string}      [current] - currently selected lang (shows checkmark); omit for lang-badge inline switcher
   */
  function showChunkTypeMenu(anchor, onPick, current) {
    // Dismiss any existing menu first
    document.querySelectorAll('.chunk-type-menu').forEach(el => el.remove());

    const menu = document.createElement('div');
    menu.className = 'chunk-type-menu';

    // Declare dismiss before items so item handlers can remove it on pick
    let dismiss;

    CHUNK_TYPE_OPTIONS.forEach(({ lang, icon, label }) => {
      const item = document.createElement('div');
      item.className = 'chunk-type-item';

      const badge = document.createElement('span');
      badge.className = `cti-badge lang-badge lang-${lang}`;
      badge.textContent = icon;

      // Checkmark column — only shown when a current selection is provided (toolbar dropdown)
      const check = document.createElement('span');
      check.className = 'cti-check';
      if (current !== undefined) {
        check.textContent = lang === current ? '✓' : '';
      }

      item.append(badge, document.createTextNode(' ' + label), check);
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        menu.remove();
        document.removeEventListener('mousedown', dismiss, true);
        onPick(lang);
      });
      menu.appendChild(item);
    });

    // Position below the anchor element
    const rect = anchor.getBoundingClientRect();
    menu.style.top  = `${rect.bottom + 4}px`;
    menu.style.left = `${rect.left}px`;
    document.body.appendChild(menu);

    // Dismiss on any outside mousedown
    dismiss = e => {
      if (!menu.contains(e.target)) {
        menu.remove();
        document.removeEventListener('mousedown', dismiss, true);
      }
    };
    // Use capture so we see the event before stopPropagation() in buttons
    setTimeout(() => document.addEventListener('mousedown', dismiss, true), 0);
  }

  // ---- Syntax highlighting dispatcher --------------------------------------

  /** Dispatch to the right highlighter based on language. */
  function highlightCode(language, src) {
    if (!language || language === 'r') return highlightR(src);
    // Bash: highlight comments (#) and strings; everything else plain
    if (language === 'bash') return highlightBash(src);
    // Markdown and all others: plain escaped text
    return esc(src);
  }

  // ---- Bash Syntax Highlighter (minimal) -----------------------------------

  function highlightBash(src) {
    const out = [];
    let i = 0;
    const n = src.length;
    const BASH_KW = new Set([
      'if','then','else','elif','fi','for','do','done','while','until',
      'case','esac','in','function','return','exit','export','local',
      'echo','cd','ls','grep','awk','sed','cat','rm','cp','mv','mkdir',
      'source','set','unset','shift','read','true','false',
    ]);
    while (i < n) {
      const ch = src[i];
      // Comment: # to end of line
      if (ch === '#') {
        const end = src.indexOf('\n', i);
        const tok = end < 0 ? src.slice(i) : src.slice(i, end);
        out.push(`<span class="r-comment">${esc(tok)}</span>`);
        i += tok.length; continue;
      }
      // String: "..." or '...'
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        while (j < n && src[j] !== ch) { if (src[j] === '\\') j++; j++; }
        out.push(`<span class="r-string">${esc(src.slice(i, j + 1))}</span>`);
        i = j + 1; continue;
      }
      // Variable: $VAR or ${VAR}
      if (ch === '$') {
        let j = i + 1;
        if (j < n && src[j] === '{') {
          const end = src.indexOf('}', j);
          j = end < 0 ? n : end + 1;
        } else {
          while (j < n && /[\w]/.test(src[j])) j++;
        }
        out.push(`<span class="r-op">${esc(src.slice(i, j))}</span>`);
        i = j; continue;
      }
      // Word / keyword
      if (/[a-zA-Z_]/.test(ch)) {
        let j = i;
        while (j < n && /[a-zA-Z0-9_\-]/.test(src[j])) j++;
        const word = src.slice(i, j);
        if (BASH_KW.has(word)) {
          out.push(`<span class="r-keyword">${esc(word)}</span>`);
        } else {
          // Check for function call (word followed by optional space + '(')
          let k = j; while (k < n && src[k] === ' ') k++;
          if (src[k] === '(') out.push(`<span class="r-func">${esc(word)}</span>`);
          else out.push(esc(word));
        }
        i = j; continue;
      }
      out.push(esc(ch)); i++;
    }
    return out.join('');
  }

  // ---- R Syntax Highlighter ------------------------------------------------

  const R_KEYWORDS = new Set([
    'function','if','else','for','while','repeat','break','next','return','in',
    'TRUE','FALSE','T','F','NULL','NA','NA_integer_','NA_real_',
    'NA_complex_','NA_character_','Inf','NaN',
  ]);

  function highlightR(src) {
    const out = [];
    let i = 0;
    const n = src.length;

    while (i < n) {
      const ch = src[i];

      // Comment: # to end of line
      if (ch === '#') {
        const end = src.indexOf('\n', i);
        const tok = end < 0 ? src.slice(i) : src.slice(i, end);
        out.push(`<span class="r-comment">${esc(tok)}</span>`);
        i += tok.length;
        continue;
      }

      // String: "..." or '...'
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        while (j < n && src[j] !== ch) { if (src[j] === '\\') j++; j++; }
        out.push(`<span class="r-string">${esc(src.slice(i, j + 1))}</span>`);
        i = j + 1;
        continue;
      }

      // Backtick identifier: `...`
      if (ch === '`') {
        const end = src.indexOf('`', i + 1);
        const tok = end < 0 ? src.slice(i) : src.slice(i, end + 1);
        out.push(`<span class="r-string">${esc(tok)}</span>`);
        i += tok.length;
        continue;
      }

      // Percent operator: %...%
      if (ch === '%') {
        const end = src.indexOf('%', i + 1);
        const tok = end < 0 ? src.slice(i) : src.slice(i, end + 1);
        out.push(`<span class="r-op">${esc(tok)}</span>`);
        i += tok.length;
        continue;
      }

      // Number
      if (/[0-9]/.test(ch) || (ch === '.' && i + 1 < n && /[0-9]/.test(src[i + 1]))) {
        let j = i;
        while (j < n && /[0-9.]/.test(src[j])) j++;
        if (j < n && /[eE]/.test(src[j])) {
          j++;
          if (j < n && /[+\-]/.test(src[j])) j++;
          while (j < n && /[0-9]/.test(src[j])) j++;
        }
        if (j < n && /[Li]/.test(src[j])) j++;
        out.push(`<span class="r-number">${esc(src.slice(i, j))}</span>`);
        i = j;
        continue;
      }

      // Identifier / keyword / function call
      if (/[a-zA-Z_.]/.test(ch)) {
        let j = i;
        while (j < n && /[a-zA-Z0-9_.]/.test(src[j])) j++;
        const word = src.slice(i, j);
        // Skip trailing whitespace to check for '('
        let k = j;
        while (k < n && src[k] === ' ') k++;
        if (R_KEYWORDS.has(word)) {
          out.push(`<span class="r-keyword">${esc(word)}</span>`);
        } else if (src[k] === '(') {
          out.push(`<span class="r-func">${esc(word)}</span>`);
        } else {
          out.push(esc(word));
        }
        i = j;
        continue;
      }

      // Multi-char operators: <-, ->, <<-, ->>, ==, !=, <=, >=, |>
      if (/[<>\-=!|&]/.test(ch)) {
        let tok = ch;
        const next = src[i + 1] || '';
        if ((ch === '<' && next === '-') || (ch === '-' && next === '>') ||
            (ch === '<' && next === '<') || (ch === '-' && next === '-') ||
            (ch === '=' && next === '=') || (ch === '!' && next === '=') ||
            (ch === '<' && next === '=') || (ch === '>' && next === '=') ||
            (ch === '|' && next === '>') || (ch === '&' && next === '&') ||
            (ch === '|' && next === '|')) {
          tok = ch + next;
        }
        out.push(`<span class="r-op">${esc(tok)}</span>`);
        i += tok.length;
        continue;
      }

      out.push(esc(ch));
      i++;
    }
    return out.join('');
  }

  function esc(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // ---- Markdown renderer (prose) -------------------------------------------

  function renderMarkdown(text) {
    if (!text) return '';
    return text
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/^###### (.+)$/gm, '<h6>$1</h6>')
      .replace(/^##### (.+)$/gm,  '<h5>$1</h5>')
      .replace(/^#### (.+)$/gm,   '<h4>$1</h4>')
      .replace(/^### (.+)$/gm,    '<h3>$1</h3>')
      .replace(/^## (.+)$/gm,     '<h2>$1</h2>')
      .replace(/^# (.+)$/gm,      '<h1>$1</h1>')
      .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
      .replace(/\*\*(.+?)\*\*/g,     '<strong>$1</strong>')
      .replace(/\*(.+?)\*/g,         '<em>$1</em>')
      .replace(/`(.+?)`/g,           '<code>$1</code>')
      .replace(/^> (.+)$/gm,         '<blockquote>$1</blockquote>')
      .replace(/^---$/gm,            '<hr>')
      .replace(/\n\n/g, '</p><p>')
      .replace(/^/, '<p>').replace(/$/, '</p>');
  }

  // ---- Toast ---------------------------------------------------------------

  function showToast(text, level = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${level}`;
    toast.textContent = String(text).slice(0, 300);
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 5000);
  }

})();
