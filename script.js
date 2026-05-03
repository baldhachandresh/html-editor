// ── Elements ────────────────────────────────────────
  const pagesWrap = document.getElementById('pages');
  const pageArea  = document.getElementById('page-area');
  function getEditors() {
    return Array.from(document.querySelectorAll('.editor-page'));
  }
  function getActiveEditor() {
    const el = document.activeElement;
    if (el && el.classList && el.classList.contains('editor-page')) return el;
    return getEditors()[0];
  }
  function ensureAtLeastOnePage() {
    if (!pagesWrap) return;
    if (getEditors().length) return;
    const page = document.createElement('div');
    page.className = 'page';
    page.dataset.page = '1';
    const ed = document.createElement('div');
    ed.className = 'editor-page';
    ed.contentEditable = 'true';
    ed.spellcheck = true;
    ed.innerHTML = '<p><br></p>';
    page.appendChild(ed);
    pagesWrap.appendChild(page);
  }
  ensureAtLeastOnePage();
  let editor = getActiveEditor();
  const sbWords   = document.getElementById('sb-words');
  const sbChars   = document.getElementById('sb-chars');
  const sbPara    = document.getElementById('sb-para');
  const sbDoc     = document.getElementById('sb-doc');
  const sbPages   = document.getElementById('sb-pages');
  const sbPaper   = document.getElementById('sb-paper');
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
    const allText = getEditors().map(e => e.innerText || '').join('\n');
    const allHTML = getEditors().map(e => e.innerHTML || '').join('\n');
    const text  = allText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    const paras = (allHTML.match(/<(p|h[1-6]|blockquote|pre|li)/gi) || []).length;
    sbWords.textContent = 'Words: ' + words.toLocaleString();
    sbChars.textContent = 'Chars: ' + chars.toLocaleString();
    sbPara.textContent  = 'Paragraphs: ' + paras;
    updatePagesStatus();
    markUnsaved();
  }

  function markUnsaved() {
    savedBadge.innerHTML = '<span style="width:6px;height:6px;border-radius:50%;background:#f5a623;display:inline-block;"></span> Unsaved';
    clearTimeout(savedTimer);
    savedTimer = setTimeout(autoSave, 3000);
  }

  function autoSave() {
    const html = getEditors().map(e => e.innerHTML).join('\n<!-- page -->\n');
    localStorage.setItem('writeflow_content', html);
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

  // Editor events (delegated for multi-page)
  function bindEditorEvents(ed) {
    ed.addEventListener('keyup', updateStats);
    ed.addEventListener('input', () => { schedulePaginate(); updateStats(); });
    ed.addEventListener('mouseup', updateActiveButtons);
    ed.addEventListener('keyup', updateActiveButtons);
    ed.addEventListener('focus', () => { editor = ed; });
  }
  getEditors().forEach(bindEditorEvents);

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

  // Pictures (local upload)
  const pictureInput = document.getElementById('picture-input');
  if (pictureInput) {
    pictureInput.addEventListener('change', function (e) {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      insertPictureFile(file);
      this.value = '';
    });
  }

  function saveHTML() {
    const title = titleInput.value || 'document';
    const bodyHtml = getEditors()
      .map((e) => e.innerHTML)
      .filter(Boolean)
      .join('\n<div style="page-break-after:always"></div>\n');
    const html  = `<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8" />\n<title>${title}</title>\n<style>body{font-family:'Lora',Georgia,serif;max-width:800px;margin:60px auto;padding:0 24px;line-height:1.75;color:#2c2825;font-size:13pt}h1,h2,h3,h4{font-weight:700;margin:.5em 0 .2em}blockquote{border-left:3px solid #c0692a;padding:.4em 1.2em;color:#6b6460;font-style:italic;background:#fdf8f3}table{border-collapse:collapse;width:100%}td,th{border:1px solid #e0dbd2;padding:7px 12px}th{background:#f5f0e8;font-weight:600}</style>\n</head>\n<body>\n${bodyHtml}\n</body>\n</html>`;
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

  function openPictureFile() {
    const input = document.getElementById('picture-input');
    if (!input) return;
    input.click();
  }

  function insertPictureFile(file) {
    if (!file) return;
    if (!file.type || !file.type.startsWith('image/')) {
      toast('Please choose an image file');
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      const src = ev.target && ev.target.result;
      if (!src) return;
      editor.focus();
      const safeName = (file.name || 'image').replace(/"/g, '');
      document.execCommand(
        'insertHTML',
        false,
        `<img src="${src}" alt="${safeName}" style="max-width:100%;height:auto;border-radius:6px;margin:10px 0;border:1px solid rgba(0,0,0,.06);" />`
      );
      toast('Picture inserted');
      updateStats();
    };
    reader.readAsDataURL(file);
  }

  function insertImage() {
    const url = prompt('Enter image URL:');
    if (!url) return;
    const alt = prompt('Alt text (optional):') || '';
    editor.focus();
    document.execCommand('insertHTML', false, `<img src="${url}" alt="${alt}" style="max-width:100%;height:auto;border-radius:4px;margin:8px 0;" />`);
    toast('Image inserted');
  }

  function normalizeVideoUrl(url) {
    try {
      const u = new URL(url);
      if (u.hostname.includes('youtube.com') || u.hostname.includes('youtu.be')) {
        let id = '';
        if (u.hostname.includes('youtu.be')) id = u.pathname.replace('/', '');
        else id = u.searchParams.get('v') || '';
        if (id) return `https://www.youtube.com/embed/${id}`;
      }
      return url;
    } catch (_) {
      return url;
    }
  }

  function insertOnlineVideo() {
    const url = prompt('Paste a video URL (YouTube supported):');
    if (!url) return;
    const src = normalizeVideoUrl(url.trim());
    editor.focus();
    document.execCommand(
      'insertHTML',
      false,
      `<div class="embed-video" contenteditable="false"><iframe src="${src}" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe></div><p><br></p>`
    );
    toast('Video embed inserted');
    updateStats();
  }

  function insertDateTime() {
    const now = new Date();
    const str = now.toLocaleString('en-IN', { dateStyle:'long', timeStyle:'short' });
    editor.focus();
    document.execCommand('insertText', false, str);
  }

  function getUniqueId(base) {
    let id = base;
    let i = 2;
    while (document.getElementById(id)) id = `${base}-${i++}`;
    return id;
  }

  function insertBookmark() {
    const name = (prompt('Bookmark name:') || '').trim();
    if (!name) return;
    const slug = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'bookmark';
    const id = getUniqueId(`bm-${slug}`);
    editor.focus();
    document.execCommand(
      'insertHTML',
      false,
      `<span id="${id}" data-bookmark="${name.replace(/"/g, '&quot;')}" style="display:inline-block;width:0;height:0;line-height:0;"></span>`
    );
    toast(`Bookmark added: ${name}`);
    updateStats();
  }

  function listBookmarks() {
    const nodes = editor.querySelectorAll('[data-bookmark]');
    const items = [];
    nodes.forEach((n) => {
      const label = n.getAttribute('data-bookmark') || n.id || 'Bookmark';
      items.push({ id: n.id, label });
    });
    const seen = new Set();
    return items.filter((x) => x.id && !seen.has(x.id) && seen.add(x.id));
  }

  function insertCrossReference() {
    const bms = listBookmarks();
    if (!bms.length) {
      toast('No bookmarks found');
      return;
    }
    const names = bms.map((b) => b.label).join('\n');
    const chosen = (prompt(`Type a bookmark name to link:\n\n${names}`) || '').trim();
    if (!chosen) return;
    const target =
      bms.find((b) => b.label.toLowerCase() === chosen.toLowerCase()) ||
      bms.find((b) => b.id.toLowerCase() === chosen.toLowerCase()) ||
      bms[0];
    editor.focus();
    document.execCommand('insertHTML', false, `<a href="#${target.id}">${target.label}</a>`);
    toast('Cross-reference inserted');
    updateStats();
  }

  function convertTextToTable() {
    editor.focus();
    const sel = window.getSelection();
    const text = sel && sel.toString ? sel.toString() : '';
    const raw = text && text.trim() ? text : (prompt('Paste/Type the text to convert into a table:') || '');
    if (!raw.trim()) return;
    const delim = (prompt('Column separator (default: tab). Use "," for CSV:', '\t') ?? '\t');
    const lines = raw.split(/\r?\n/).filter((l) => l.trim().length);
    const rows = lines.map((l) => l.split(delim).map((c) => c.trim()));
    const maxCols = Math.max(...rows.map((r) => r.length), 1);

    let html = '<table><tbody>';
    rows.forEach((r) => {
      html += '<tr>';
      for (let c = 0; c < maxCols; c++) {
        const v = (r[c] || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        html += `<td>${v || '&nbsp;'}</td>`;
      }
      html += '</tr>';
    });
    html += '</tbody></table><p><br></p>';
    document.execCommand('insertHTML', false, html);
    toast('Converted text to table');
    updateStats();
  }

  function insertQuickTable(type) {
    if (type === '2x2') return insertTableGrid(2, 2);
    toast('Quick table not available');
  }

  // ── Pagination (Word-like pages) ──────────────────────
  let paginateTimer = null;
  function schedulePaginate() {
    clearTimeout(paginateTimer);
    paginateTimer = setTimeout(paginateAll, 120);
  }

  function createPage(afterPageEl) {
    if (!pagesWrap) return null;
    const page = document.createElement('div');
    page.className = 'page';
    const ed = document.createElement('div');
    ed.className = 'editor-page';
    ed.contentEditable = 'true';
    ed.spellcheck = true;
    ed.innerHTML = '<p><br></p>';
    page.appendChild(ed);
    if (afterPageEl && afterPageEl.parentNode === pagesWrap) afterPageEl.after(page);
    else pagesWrap.appendChild(page);
    bindEditorEvents(ed);
    renumberPages();
    return page;
  }

  function renumberPages() {
    document.querySelectorAll('.page').forEach((p, idx) => {
      p.dataset.page = String(idx + 1);
    });
    updatePagesStatus();
  }

  function pageOverflows(ed) {
    const max = getPageInnerHeightPx();
    return ed.scrollHeight > max + 1;
  }

  function pageHasSpace(ed) {
    const max = getPageInnerHeightPx();
    return ed.scrollHeight < max - 60;
  }

  function moveLastNodeToNext(fromEd, toEd) {
    const node = fromEd.lastChild;
    if (!node) return false;
    // Keep at least one paragraph in the page
    if (fromEd.childNodes.length === 1) return false;
    toEd.insertBefore(node, toEd.firstChild);
    return true;
  }

  function moveFirstNodeToPrev(fromEd, toEd) {
    const node = fromEd.firstChild;
    if (!node) return false;
    toEd.appendChild(node);
    return true;
  }

  function cleanupEmptyPages() {
    const pages = Array.from(document.querySelectorAll('.page'));
    // Keep at least one page
    for (let i = pages.length - 1; i >= 1; i--) {
      const ed = pages[i].querySelector('.editor-page');
      const html = (ed && ed.innerHTML || '').replace(/\s|&nbsp;|<br\s*\/?>/gi, '');
      if (!html) pages[i].remove();
      else break;
    }
    renumberPages();
  }

  function paginateAll() {
    if (!pagesWrap) return;
    const pages = Array.from(document.querySelectorAll('.page'));
    if (!pages.length) return;

    // Forward pass: push overflow to next pages
    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      const ed = page.querySelector('.editor-page');
      if (!ed) continue;
      while (pageOverflows(ed)) {
        const nextPage = pages[i + 1] || createPage(page);
        const nextEd = nextPage.querySelector('.editor-page');
        if (!nextEd) break;
        // Ensure next has at least a blank paragraph at the end
        if (!moveLastNodeToNext(ed, nextEd)) break;
      }
    }

    // Backward pass: pull content back if there is space
    const pages2 = Array.from(document.querySelectorAll('.page'));
    for (let i = pages2.length - 1; i > 0; i--) {
      const page = pages2[i];
      const prev = pages2[i - 1];
      const ed = page.querySelector('.editor-page');
      const prevEd = prev.querySelector('.editor-page');
      if (!ed || !prevEd) continue;
      while (pageHasSpace(prevEd) && ed.childNodes.length) {
        if (!moveFirstNodeToPrev(ed, prevEd)) break;
        if (pageOverflows(prevEd)) {
          // Undo if we overfilled
          moveLastNodeToNext(prevEd, ed);
          break;
        }
      }
    }

    cleanupEmptyPages();
  }

  function toggleShapesDropdown(ev) {
    if (ev) ev.preventDefault();
    const dd = document.getElementById('shapes-dd');
    if (!dd) return;
    const isOpen = dd.classList.contains('open');
    closeAllDropdowns();
    dd.classList.toggle('open', !isOpen);
  }

  function insertShape(kind) {
    const stroke = '#b9b2a6';
    const fill = '#ffffff';
    const accent = '#c0692a';
    let svg = '';
    if (kind === 'rect') svg = `<svg width="220" height="90" viewBox="0 0 220 90" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="216" height="86" rx="0" fill="${fill}" stroke="${stroke}" stroke-width="2"/></svg>`;
    else if (kind === 'round') svg = `<svg width="220" height="90" viewBox="0 0 220 90" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="216" height="86" rx="14" fill="${fill}" stroke="${stroke}" stroke-width="2"/></svg>`;
    else if (kind === 'circle') svg = `<svg width="120" height="120" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><circle cx="60" cy="60" r="54" fill="${fill}" stroke="${stroke}" stroke-width="2"/></svg>`;
    else if (kind === 'arrow') svg = `<svg width="240" height="80" viewBox="0 0 240 80" xmlns="http://www.w3.org/2000/svg"><path d="M10 40h160" stroke="${accent}" stroke-width="6" stroke-linecap="round"/><path d="M150 18l40 22-40 22" fill="none" stroke="${accent}" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
    else if (kind === 'line') svg = `<svg width="240" height="24" viewBox="0 0 240 24" xmlns="http://www.w3.org/2000/svg"><line x1="8" y1="12" x2="232" y2="12" stroke="${accent}" stroke-width="4" stroke-linecap="round"/></svg>`;
    else return;

    editor.focus();
    document.execCommand('insertHTML', false, `<span class="shape-wrap" contenteditable="false">${svg}</span><p><br></p>`);
    toast('Shape inserted');
    updateStats();
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
    document.querySelectorAll('.page').forEach(p => p.classList.toggle('no-guide', !show));
  }

  function setPageZoom(val) {
    const baseW = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--page-w'), 10) || 816;
    const px = Math.round(baseW * val / 100);
    document.getElementById('pages').style.width = px + 'px';
    document.getElementById('zoom-label').textContent = val + '%';
    updatePagesStatus();
  }

  function setPaperSize(size) {
    const root = document.documentElement;
    const s = String(size || 'letter').toLowerCase();
    if (s === 'a4') {
      root.style.setProperty('--page-w', '794px');  // ~8.27in @96dpi
      root.style.setProperty('--page-h', '1123px'); // ~11.69in @96dpi
      if (sbPaper) sbPaper.textContent = 'A4 (210 × 297 mm)';
      toast('Paper set to A4');
    } else {
      root.style.setProperty('--page-w', '816px');
      root.style.setProperty('--page-h', '1056px');
      if (sbPaper) sbPaper.textContent = 'Letter (8.5 × 11 in)';
      toast('Paper set to Letter');
    }
    // Reset zoom slider to 100% baseline
    const zoom = document.getElementById('page-zoom');
    if (zoom) zoom.value = '100';
    const pages = document.getElementById('pages');
    if (pages) pages.style.width = '';
    const zlabel = document.getElementById('zoom-label');
    if (zlabel) zlabel.textContent = '100%';
    updatePagesStatus();
    schedulePaginate();
  }

  function getPageInnerHeightPx() {
    // Approx: page height minus editor vertical padding (72px top + 72px bottom)
    const h = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--page-h'), 10) || 1056;
    return Math.max(200, h - 144);
  }

  function updatePagesStatus() {
    if (!sbPages) return;
    const pageH = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--page-h'), 10) || 1056;
    const total = Math.max(1, document.querySelectorAll('.page').length);

    // Current page based on scroll position + page height + gap
    const gap = 28;
    const current = pageArea
      ? Math.min(total, Math.max(1, Math.floor(pageArea.scrollTop / (pageH + gap)) + 1))
      : 1;
    sbPages.textContent = `Page: ${current} / ${total}`;
  }

  // ── Toast ─────────────────────────────────────────────
  function toast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(el._t);
    el._t = setTimeout(() => el.classList.remove('show'), 2200);
  }

  // ── Draw pad ──────────────────────────────────────────
  const drawOverlay = document.getElementById('draw-overlay');
  const drawCanvas = document.getElementById('draw-canvas');
  const drawCtx = drawCanvas ? drawCanvas.getContext('2d') : null;

  let drawOpen = false;
  let drawTool = 'pen'; // pen | highlighter | eraser
  let drawColor = '#c0692a';
  let drawSize = 6;
  let drawing = false;
  let lastPt = null;

  function setDrawTool(tool) {
    drawTool = tool;
    if (drawTool === 'eraser') toast('Eraser');
    else if (drawTool === 'highlighter') toast('Highlighter');
    else toast('Pen');
  }

  function setDrawColor(color) {
    drawColor = color;
  }

  function setDrawSize(val) {
    drawSize = Math.max(1, parseInt(val, 10) || 6);
    const label = document.getElementById('draw-size-label');
    if (label) label.textContent = String(drawSize);
  }

  function openDrawPad(tool) {
    if (!drawOverlay || !drawCanvas || !drawCtx) return;
    drawOpen = true;
    drawOverlay.classList.add('open');
    drawOverlay.setAttribute('aria-hidden', 'false');
    setDrawTool(tool || 'pen');
    syncDrawSettingsFromUI();
    resizeDrawCanvasToDisplay();
    toast('Draw pad opened');
  }

  function closeDrawPad() {
    if (!drawOverlay) return;
    drawOpen = false;
    drawOverlay.classList.remove('open');
    drawOverlay.setAttribute('aria-hidden', 'true');
  }

  function clearDrawPad() {
    if (!drawCtx || !drawCanvas) return;
    drawCtx.save();
    drawCtx.setTransform(1, 0, 0, 1, 0, 0);
    drawCtx.clearRect(0, 0, drawCanvas.width, drawCanvas.height);
    drawCtx.restore();
    toast('Drawing cleared');
  }

  function saveDrawPad() {
    if (!drawCanvas) return;
    const dataUrl = drawCanvas.toDataURL('image/png');
    editor.focus();
    document.execCommand(
      'insertHTML',
      false,
      `<img src="${dataUrl}" alt="Drawing" style="max-width:100%;height:auto;border-radius:10px;margin:10px 0;border:1px solid rgba(0,0,0,.08);" />`
    );
    toast('Drawing inserted');
    updateStats();
    closeDrawPad();
  }

  function syncDrawSettingsFromUI() {
    const c = document.getElementById('draw-color');
    const s = document.getElementById('draw-size');
    if (c && c.value) drawColor = c.value;
    if (s && s.value) setDrawSize(s.value);
  }

  function resizeDrawCanvasToDisplay() {
    if (!drawCanvas || !drawCtx) return;
    // Preserve current drawing when resizing to fit container
    const old = document.createElement('canvas');
    old.width = drawCanvas.width;
    old.height = drawCanvas.height;
    const octx = old.getContext('2d');
    if (octx) octx.drawImage(drawCanvas, 0, 0);

    const rect = drawCanvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const w = Math.max(1, Math.floor(rect.width * dpr));
    const h = Math.max(1, Math.floor(rect.height * dpr));
    drawCanvas.width = w;
    drawCanvas.height = h;
    drawCtx.setTransform(1, 0, 0, 1, 0, 0);
    drawCtx.imageSmoothingEnabled = true;
    drawCtx.drawImage(old, 0, 0, w, h);
  }

  function canvasPointFromEvent(e) {
    const rect = drawCanvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const x = (e.clientX - rect.left) * dpr;
    const y = (e.clientY - rect.top) * dpr;
    return { x, y };
  }

  function beginStroke(e) {
    if (!drawOpen || !drawCanvas || !drawCtx) return;
    drawing = true;
    lastPt = canvasPointFromEvent(e);
    drawCanvas.setPointerCapture && drawCanvas.setPointerCapture(e.pointerId);
  }

  function moveStroke(e) {
    if (!drawing || !drawOpen || !drawCanvas || !drawCtx || !lastPt) return;
    const pt = canvasPointFromEvent(e);

    drawCtx.lineCap = 'round';
    drawCtx.lineJoin = 'round';

    if (drawTool === 'eraser') {
      drawCtx.globalCompositeOperation = 'destination-out';
      drawCtx.strokeStyle = 'rgba(0,0,0,1)';
      drawCtx.lineWidth = drawSize * (window.devicePixelRatio || 1);
    } else if (drawTool === 'highlighter') {
      drawCtx.globalCompositeOperation = 'source-over';
      drawCtx.strokeStyle = drawColor;
      drawCtx.globalAlpha = 0.35;
      drawCtx.lineWidth = (drawSize * 2) * (window.devicePixelRatio || 1);
    } else {
      drawCtx.globalCompositeOperation = 'source-over';
      drawCtx.strokeStyle = drawColor;
      drawCtx.globalAlpha = 1;
      drawCtx.lineWidth = drawSize * (window.devicePixelRatio || 1);
    }

    drawCtx.beginPath();
    drawCtx.moveTo(lastPt.x, lastPt.y);
    drawCtx.lineTo(pt.x, pt.y);
    drawCtx.stroke();
    drawCtx.globalAlpha = 1;

    lastPt = pt;
  }

  function endStroke(e) {
    if (!drawOpen) return;
    drawing = false;
    lastPt = null;
    try { drawCanvas.releasePointerCapture && drawCanvas.releasePointerCapture(e.pointerId); } catch (_) {}
  }

  if (drawCanvas) {
    drawCanvas.addEventListener('pointerdown', beginStroke);
    drawCanvas.addEventListener('pointermove', moveStroke);
    drawCanvas.addEventListener('pointerup', endStroke);
    drawCanvas.addEventListener('pointercancel', endStroke);
  }

  window.addEventListener('resize', () => {
    if (drawOpen) resizeDrawCanvasToDisplay();
  });

  window.openDrawPad = openDrawPad;
  window.closeDrawPad = closeDrawPad;
  window.clearDrawPad = clearDrawPad;
  window.saveDrawPad = saveDrawPad;
  window.setDrawTool = setDrawTool;
  window.setDrawColor = setDrawColor;
  window.setDrawSize = setDrawSize;

  // ── Auto-load saved ───────────────────────────────────
  function loadSaved() {
    const content = localStorage.getItem('writeflow_content');
    const title   = localStorage.getItem('writeflow_title');
    if (content) {
      // Rebuild pages from stored content split markers
      const parts = String(content).split('<!-- page -->');
      if (pagesWrap) pagesWrap.innerHTML = '';
      parts.forEach((html, idx) => {
        const page = document.createElement('div');
        page.className = 'page';
        page.dataset.page = String(idx + 1);
        const ed = document.createElement('div');
        ed.className = 'editor-page';
        ed.contentEditable = 'true';
        ed.spellcheck = true;
        ed.innerHTML = html.trim() || '<p><br></p>';
        page.appendChild(ed);
        pagesWrap.appendChild(page);
      });
      getEditors().forEach(bindEditorEvents);
      editor = getActiveEditor();
      toast('Restored last session');
    }
    if (title)   { titleInput.value = title; sbDoc.textContent = title; }
    updateStats();
  }

  // ── Init ──────────────────────────────────────────────
  loadSaved();
  updateStats();
  updateStyleTiles('normal');
  updatePagesStatus();

  // Make links work (Ctrl+Click) across pages
  document.addEventListener('click', e => {
    if (e.target && e.target.tagName === 'A' && e.ctrlKey) {
      window.open(e.target.href, '_blank');
    }
  });

  // Update current page indicator while scrolling
  const pageAreaEl = document.getElementById('page-area');
  if (pageAreaEl) pageAreaEl.addEventListener('scroll', updatePagesStatus, { passive: true });

  window.setPaperSize = setPaperSize;

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
    const openAny = document.querySelector('.rb-dd.open');
    if (!openAny) return;
    const anyDd = e.target && e.target.closest ? e.target.closest('.rb-dd') : null;
    if (!anyDd) closeAllDropdowns();
  });

  window.openPictureFile = openPictureFile;
  window.insertOnlineVideo = insertOnlineVideo;
  window.toggleShapesDropdown = toggleShapesDropdown;
  window.insertShape = insertShape;
  window.insertBookmark = insertBookmark;
  window.insertCrossReference = insertCrossReference;
  window.convertTextToTable = convertTextToTable;
  window.insertQuickTable = insertQuickTable;
