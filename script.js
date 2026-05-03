// ── Elements ────────────────────────────────────────
  const editor    = document.getElementById('editor');
  const sbWords   = document.getElementById('sb-words');
  const sbChars   = document.getElementById('sb-chars');
  const sbPara    = document.getElementById('sb-para');
  const sbDoc     = document.getElementById('sb-doc');
  const titleInput= document.getElementById('doc-title-input');
  const savedBadge= document.getElementById('saved-badge');

  // ── Saved state ─────────────────────────────────────
  let savedTimer  = null;
  let findMatches = [];
  let findIdx     = -1;
  let savedRange  = null; // saved selection for dialogs
  let tableHover = { rows: 0, cols: 0 };

  // ── Update stats ─────────────────────────────────────
  function updateStats() {
    const text  = editor.innerText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    const paras = (editor.innerHTML.match(/<(p|h[1-6]|blockquote|pre|li)/gi) || []).length;
    sbWords.textContent = 'Words: ' + words.toLocaleString();
    sbChars.textContent = 'Chars: ' + chars.toLocaleString();
    sbPara.textContent  = 'Paragraphs: ' + paras;
    markUnsaved();
  }

  function markUnsaved() {
    savedBadge.innerHTML = '<span style="width:6px;height:6px;border-radius:50%;background:#f5a623;display:inline-block;"></span> Unsaved';
    clearTimeout(savedTimer);
    savedTimer = setTimeout(autoSave, 3000);
  }

  function autoSave() {
    localStorage.setItem('writeflow_content', editor.innerHTML);
    localStorage.setItem('writeflow_title',   titleInput.value);
    savedBadge.innerHTML = '<span style="width:6px;height:6px;border-radius:50%;background:#6bcf7f;display:inline-block;"></span> Saved';
  }

  // ── Format commands ──────────────────────────────────
  function fmt(cmd) {
    editor.focus();
    document.execCommand(cmd, false, null);
    updateActiveButtons();
  }

  function applyFont(val) {
    editor.focus();
    document.execCommand('fontName', false, val);
  }

  function applyFontSize(pt) {
    editor.focus();
    const sizePt = String(pt).trim();
    if (!sizePt) return;

    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;

    // If no selection, set editor default size (so next typed text matches)
    if (sel.isCollapsed) {
      editor.style.fontSize = `${sizePt}pt`;
      toast(`Font size set to ${sizePt}pt`);
      return;
    }

    // Reliable approach: use execCommand('fontSize') then replace <font> tags with spans.
    // This handles complex selections better than Range.surroundContents.
    document.execCommand('fontSize', false, '7');
    const fonts = editor.querySelectorAll('font[size="7"]');
    fonts.forEach((font) => {
      const span = document.createElement('span');
      span.style.fontSize = `${sizePt}pt`;
      span.innerHTML = font.innerHTML;
      font.replaceWith(span);
    });
  }

  async function pasteCmd() {
    editor.focus();

    // Best effort: clipboard API (requires user gesture + permissions).
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const text = await navigator.clipboard.readText();
        if (text) {
          document.execCommand('insertText', false, text);
          toast('Pasted');
          return;
        }
      }
    } catch (_) {
      // ignore; we'll fall back below
    }

    // Fallback: let the browser handle Ctrl+V / context-menu paste.
    toast('Press Ctrl+V to paste (browser security)');
  }

  function applyColor(cmd, color) {
    editor.focus();
    document.execCommand(cmd, false, color);
  }

  function applyHighlight(color) {
    editor.focus();
    document.execCommand('hiliteColor', false, color);
  }

  function applyBlock(tag) {
    editor.focus();
    document.execCommand('formatBlock', false, '<' + tag + '>');
    setTimeout(() => {
      document.getElementById('block-format').value = tag;
    }, 50);
  }

  function applyLineSpacing(val) {
    editor.style.setProperty('--editor-lh', String(val));
  }

  function setParaSpacing(spaceEm) {
    editor.style.setProperty('--para-space', `${spaceEm}em`);
  }

  function applyStylePreset(name) {
    editor.focus();

    if (name === 'normal') {
      applyBlock('p');
      applyLineSpacing('1.75');
      setParaSpacing(0.4);
    } else if (name === 'nospace') {
      applyBlock('p');
      applyLineSpacing('1.2');
      setParaSpacing(0.0);
    } else if (name === 'heading') {
      applyBlock('h2');
      applyLineSpacing('1.3');
      setParaSpacing(0.2);
    }

    updateStyleTiles(name);
    updateActiveButtons();
    updateStats();
  }

  function updateStyleTiles(force) {
    const tiles = {
      normal: document.getElementById('style-normal'),
      nospace: document.getElementById('style-nospace'),
      heading: document.getElementById('style-heading'),
    };
    Object.values(tiles).forEach(t => t && t.classList.remove('active'));

    if (force && tiles[force]) {
      tiles[force].classList.add('active');
      return;
    }

    let block = '';
    try {
      block = String(document.queryCommandValue('formatBlock') || '')
        .toLowerCase()
        .replace(/[<>]/g, '');
    } catch (_) {
      block = '';
    }

    const lh = (getComputedStyle(editor).getPropertyValue('--editor-lh') || '').trim();
    const ps = (getComputedStyle(editor).getPropertyValue('--para-space') || '').trim();

    if (block === 'h2' || block === 'h1' || block === 'h3' || block === 'h4') {
      tiles.heading && tiles.heading.classList.add('active');
      return;
    }

    if ((ps === '0em' || ps === '0') && (lh === '1.2' || lh === '1.15')) {
      tiles.nospace && tiles.nospace.classList.add('active');
      return;
    }

    tiles.normal && tiles.normal.classList.add('active');
  }

  // ── Track active states ──────────────────────────────
  function updateActiveButtons() {
    const cmds = {
      'btn-bold': 'bold', 'btn-italic': 'italic', 'btn-underline': 'underline',
      'btn-strikethrough': 'strikeThrough', 'btn-superscript': 'superscript',
      'btn-subscript': 'subscript',
      'btn-left': 'justifyLeft', 'btn-center': 'justifyCenter',
      'btn-right': 'justifyRight', 'btn-justify': 'justifyFull',
      'btn-ul': 'insertUnorderedList', 'btn-ol': 'insertOrderedList',
    };
    for (const [id, cmd] of Object.entries(cmds)) {
      const el = document.getElementById(id);
      if (el) el.classList.toggle('active', document.queryCommandState(cmd));
    }
    // Update block format select
    const block = document.queryCommandValue('formatBlock').toLowerCase().replace(/[<>]/g,'');
    const bf = document.getElementById('block-format');
    if (bf && block) bf.value = block;
    updateStyleTiles();
  }

  editor.addEventListener('keyup',   updateStats);
  editor.addEventListener('input',   updateStats);
  editor.addEventListener('mouseup', updateActiveButtons);
  editor.addEventListener('keyup',   updateActiveButtons);

  // ── Title sync ───────────────────────────────────────
  titleInput.addEventListener('input', () => {
    sbDoc.textContent = titleInput.value || 'Untitled Document';
    markUnsaved();
  });

  // ── Ribbon tabs ──────────────────────────────────────
  function showTab(name, el) {
    document.querySelectorAll('.rtab').forEach(t => t.classList.remove('active'));
    el.classList.add('active');
    const panels = {
      file:'tab-file',
      home:'tab-home',
      insert:'tab-insert',
      draw:'tab-draw',
      design:'tab-design',
      layout:'tab-layout',
      references:'tab-references',
      mailings:'tab-mailings',
      review:'tab-review',
      view:'tab-view',
      help:'tab-help',
    };
    Object.values(panels).forEach(id => {
      const p = document.getElementById(id);
      if (p) p.style.display = 'none';
    });
    const target = document.getElementById(panels[name]);
    if (target) target.style.display = 'flex';
  }
  // Show File tab by default (Word-like)
  document.getElementById('tab-file').style.display = 'flex';

  // ── Keyboard shortcuts ───────────────────────────────
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'f') { e.preventDefault(); toggleFind(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); saveHTML(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'p') { e.preventDefault(); window.print(); }
    if (e.key === 'Escape') { closeLinkDialog(); closeTableDialog(); if (findOpen) toggleFind(); }
  });

  // ── File ops ─────────────────────────────────────────
  function newDoc() {
    if (!confirm('Create a new document? Unsaved changes will be lost.')) return;
    editor.innerHTML = '<p><br></p>';
    titleInput.value = 'Untitled Document';
    sbDoc.textContent= 'Untitled Document';
    autoSave(); updateStats();
    toast('New document created');
  }

  function openFile() {
    document.getElementById('file-input').click();
  }

  document.getElementById('file-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const content = ev.target.result;
      if (file.name.endsWith('.html') || file.name.endsWith('.htm')) {
        const parser = new DOMParser();
        const doc    = parser.parseFromString(content, 'text/html');
        editor.innerHTML = doc.body.innerHTML;
      } else {
        editor.innerText = content;
      }
      titleInput.value  = file.name.replace(/\.[^.]+$/, '');
      sbDoc.textContent = titleInput.value;
      updateStats();
      toast('Opened: ' + file.name);
    };
    reader.readAsText(file);
    this.value = '';
  });

  function saveHTML() {
    const title = titleInput.value || 'document';
    const html  = `<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8" />\n<title>${title}</title>\n<style>body{font-family:'Lora',Georgia,serif;max-width:800px;margin:60px auto;padding:0 24px;line-height:1.75;color:#2c2825;font-size:13pt}h1,h2,h3,h4{font-weight:700;margin:.5em 0 .2em}blockquote{border-left:3px solid #c0692a;padding:.4em 1.2em;color:#6b6460;font-style:italic;background:#fdf8f3}table{border-collapse:collapse;width:100%}td,th{border:1px solid #e0dbd2;padding:7px 12px}th{background:#f5f0e8;font-weight:600}</style>\n</head>\n<body>\n${editor.innerHTML}\n</body>\n</html>`;
    const blob  = new Blob([html], { type: 'text/html' });
    const a     = document.createElement('a');
    a.href      = URL.createObjectURL(blob);
    a.download  = title + '.html';
    a.click();
    URL.revokeObjectURL(a.href);
    autoSave();
    toast('✅ Saved as ' + title + '.html');
  }

  function clearDoc() {
    if (!confirm('Clear all content?')) return;
    editor.innerHTML = '<p><br></p>';
    updateStats();
    toast('Document cleared');
  }

  // ── Insert helpers ───────────────────────────────────
  function insertHR() {
    editor.focus();
    document.execCommand('insertHTML', false, '<hr>');
    toast('Divider inserted');
  }

  function insertImage() {
    const url = prompt('Enter image URL:');
    if (!url) return;
    const alt = prompt('Alt text (optional):') || '';
    editor.focus();
    document.execCommand('insertHTML', false, `<img src="${url}" alt="${alt}" style="max-width:100%;height:auto;border-radius:4px;margin:8px 0;" />`);
    toast('Image inserted');
  }

  function insertDateTime() {
    const now = new Date();
    const str = now.toLocaleString('en-IN', { dateStyle:'long', timeStyle:'short' });
    editor.focus();
    document.execCommand('insertText', false, str);
  }

  function insertSpecialChar() {
    const chars = ['©','®','™','€','£','¥','°','±','×','÷','≤','≥','≠','∞','√','π','α','β','γ','δ','→','←','↑','↓','★','♥','♦','♣','♠','•','–','—','"','"','\'','\''];
    const ch = prompt('Choose symbol:\n' + chars.join('  '));
    if (ch) { editor.focus(); document.execCommand('insertText', false, ch); }
  }

  function showWordCount() {
    const text  = editor.innerText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    const charsNoSpace = text.replace(/\s/g,'').length;
    const paras = (editor.innerHTML.match(/<(p|h[1-6]|blockquote|pre)/gi) || []).length;
    alert(`📊 Document Statistics\n\nWords: ${words}\nCharacters (with spaces): ${chars}\nCharacters (no spaces): ${charsNoSpace}\nParagraphs: ${paras}`);
  }

  // ── Find & Replace ───────────────────────────────────
  let findOpen = false;

  function toggleFind() {
    findOpen = !findOpen;
    const panel = document.getElementById('find-panel');
    panel.classList.toggle('open', findOpen);
    if (findOpen) { document.getElementById('find-input').focus(); doFind(); }
    else clearHighlights();
  }

  function doFind() {
    clearHighlights();
    const term = document.getElementById('find-input').value.trim();
    const countEl = document.getElementById('find-count');
    if (!term) { countEl.textContent = ''; return; }
    const body = editor;
    const regex = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'gi');
    let html = body.innerHTML;
    let count = 0;
    html = html.replace(regex, m => { count++; return `<mark class="find-hl" style="background:#fff176;color:inherit;">${m}</mark>`; });
    body.innerHTML = html;
    countEl.textContent = count ? `${count} found` : 'Not found';
    findIdx = -1;
  }

  function findNext() {
    const marks = editor.querySelectorAll('.find-hl');
    if (!marks.length) return;
    if (findIdx >= 0 && findIdx < marks.length) marks[findIdx].style.background = '#fff176';
    findIdx = (findIdx + 1) % marks.length;
    marks[findIdx].style.background = '#ffb347';
    marks[findIdx].scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function clearHighlights() {
    editor.querySelectorAll('.find-hl').forEach(el => {
      el.replaceWith(document.createTextNode(el.textContent));
    });
    editor.normalize();
  }

  function doReplace() {
    const term = document.getElementById('find-input').value;
    const repl = document.getElementById('replace-input').value;
    if (!term) return;
    // Find first highlighted
    const mark = editor.querySelector('.find-hl');
    if (mark) { mark.replaceWith(document.createTextNode(repl)); editor.normalize(); }
    doFind();
  }

  function doReplaceAll() {
    const term = document.getElementById('find-input').value.trim();
    const repl = document.getElementById('replace-input').value;
    if (!term) return;
    clearHighlights();
    const regex = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'gi');
    let count = 0;
    const walk = node => {
      if (node.nodeType === 3) {
        const matches = node.textContent.match(regex);
        if (matches) {
          count += matches.length;
          const span = document.createElement('span');
          span.innerHTML = node.textContent.replace(regex, repl);
          node.replaceWith(span);
        }
      } else {
        [...node.childNodes].forEach(walk);
      }
    };
    walk(editor);
    editor.normalize();
    toast(`Replaced ${count} occurrence(s)`);
    document.getElementById('find-count').textContent = `${count} replaced`;
  }

  // ── Link dialog ──────────────────────────────────────
  function openLinkDialog() {
    // Save current selection
    const sel = window.getSelection();
    if (sel.rangeCount) savedRange = sel.getRangeAt(0).cloneRange();
    document.getElementById('link-text').value = sel.toString();
    document.getElementById('link-url').value  = '';
    document.getElementById('link-dialog').classList.add('open');
    setTimeout(() => document.getElementById('link-url').focus(), 50);
  }

  function closeLinkDialog() {
    document.getElementById('link-dialog').classList.remove('open');
    savedRange = null;
  }

  function applyLink() {
    const url  = document.getElementById('link-url').value.trim();
    const text = document.getElementById('link-text').value.trim();
    if (!url) { toast('Please enter a URL'); return; }
    const href = url.startsWith('http') ? url : 'https://' + url;
    editor.focus();
    if (savedRange) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
    if (text && !(window.getSelection().toString())) {
      document.execCommand('insertHTML', false, `<a href="${href}" target="_blank">${text}</a>`);
    } else {
      document.execCommand('createLink', false, href);
      const links = editor.querySelectorAll('a[href="' + href + '"]');
      links.forEach(l => l.setAttribute('target', '_blank'));
    }
    closeLinkDialog();
    toast('Link inserted');
  }

  // ── Table dialog ─────────────────────────────────────
  function openTableDialog() {
    editor.focus();
    const sel = window.getSelection();
    if (sel.rangeCount) savedRange = sel.getRangeAt(0).cloneRange();
    document.getElementById('table-dialog').classList.add('open');
    setTimeout(() => document.getElementById('tbl-rows').focus(), 50);
  }

  function closeTableDialog() {
    document.getElementById('table-dialog').classList.remove('open');
    savedRange = null;
  }

  function applyTable() {
    const rows = parseInt(document.getElementById('tbl-rows').value) || 3;
    const cols = parseInt(document.getElementById('tbl-cols').value) || 3;
    let html = '<table><thead><tr>';
    for (let c = 0; c < cols; c++) html += `<th>Header ${c+1}</th>`;
    html += '</tr></thead><tbody>';
    for (let r = 0; r < rows - 1; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) html += '<td>&nbsp;</td>';
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';
    editor.focus();
    if (savedRange) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
    document.execCommand('insertHTML', false, html);
    closeTableDialog();
    toast('Table inserted');
  }

  // ── View controls ────────────────────────────────────
  function toggleRuler(show) {
    document.getElementById('page').classList.toggle('no-guide', !show);
  }

  function setPageZoom(val) {
    const px = Math.round(816 * val / 100);
    document.getElementById('page').style.width = px + 'px';
    document.getElementById('zoom-label').textContent = val + '%';
  }

  // ── Toast ─────────────────────────────────────────────
  function toast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(el._t);
    el._t = setTimeout(() => el.classList.remove('show'), 2200);
  }

  // ── Auto-load saved ───────────────────────────────────
  function loadSaved() {
    const content = localStorage.getItem('writeflow_content');
    const title   = localStorage.getItem('writeflow_title');
    if (content) { editor.innerHTML = content; toast('Restored last session'); }
    if (title)   { titleInput.value = title; sbDoc.textContent = title; }
    updateStats();
  }

  // ── Init ──────────────────────────────────────────────
  loadSaved();
  updateStats();
  updateStyleTiles('normal');

  // Make dialog links work
  document.getElementById('editor').addEventListener('click', e => {
    if (e.target.tagName === 'A' && e.ctrlKey) {
      window.open(e.target.href, '_blank');
    }
  });

  // ── Table grid dropdown (Insert → Table) ──────────────
  function ensureTableGrid() {
    const grid = document.getElementById('table-grid');
    if (!grid || grid.childElementCount) return;
    const maxCols = 10;
    const maxRows = 8;
    for (let r = 1; r <= maxRows; r++) {
      for (let c = 1; c <= maxCols; c++) {
        const cell = document.createElement('div');
        cell.className = 'tbl-cell';
        cell.dataset.r = String(r);
        cell.dataset.c = String(c);
        cell.addEventListener('mouseenter', () => setTableHover(r, c));
        cell.addEventListener('click', (e) => {
          e.preventDefault();
          insertTableGrid(r, c);
          closeAllDropdowns();
        });
        grid.appendChild(cell);
      }
    }
    setTableHover(0, 0);
  }

  function setTableHover(rows, cols) {
    tableHover = { rows, cols };
    const label = document.getElementById('table-size-label');
    if (label) label.textContent = `${rows} × ${cols}`;
    const cells = document.querySelectorAll('#table-grid .tbl-cell');
    cells.forEach((cell) => {
      const r = parseInt(cell.dataset.r, 10);
      const c = parseInt(cell.dataset.c, 10);
      cell.classList.toggle('on', r <= rows && c <= cols && rows > 0 && cols > 0);
    });
  }

  function insertTableGrid(rows, cols) {
    if (!rows || !cols) return;
    let html = '<table><tbody>';
    for (let r = 0; r < rows; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) html += '<td>&nbsp;</td>';
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';
    editor.focus();
    document.execCommand('insertHTML', false, html);
    toast(`Inserted ${rows}×${cols} table`);
  }

  function closeAllDropdowns() {
    document.querySelectorAll('.rb-dd.open').forEach((dd) => dd.classList.remove('open'));
  }

  function toggleTableDropdown(ev) {
    if (ev) ev.preventDefault();
    ensureTableGrid();
    const dd = document.getElementById('table-dd');
    if (!dd) return;
    const isOpen = dd.classList.contains('open');
    closeAllDropdowns();
    dd.classList.toggle('open', !isOpen);
    if (!isOpen) setTableHover(0, 0);
  }

  window.toggleTableDropdown = toggleTableDropdown;
  window.closeAllDropdowns = closeAllDropdowns;
  window.pasteCmd = pasteCmd;
  window.applyStylePreset = applyStylePreset;

  document.addEventListener('click', (e) => {
    const dd = document.getElementById('table-dd');
    if (!dd) return;
    if (!dd.contains(e.target)) closeAllDropdowns();
  });
