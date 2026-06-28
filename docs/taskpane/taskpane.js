/**
 * Mind Map Studio — Task Pane Controller
 * Wires the MindMap engine to the UI and handles Office.js / OneNote API integration.
 */

/* ══════════════════════════════════════
   STATE
══════════════════════════════════════ */
let mindMap = null;
let officeReady = false;
let pendingAction = null; // { type: 'addRoot' | 'addChild', parentId? }
let selectedPageForLink = null;

/* ══════════════════════════════════════
   INIT
══════════════════════════════════════ */

function initApp() {
  const svg     = document.getElementById('mindmap-svg');
  const tooltip = document.getElementById('node-tooltip');

  mindMap = new MindMap(svg, tooltip);

  bindToolbar();
  bindPanel();
  bindModalClose();
  bindAddNodeModal();
  bindNoteModal();
  bindLinkModal();
  bindDeleteModal();
  bindKeyboard();
  bindEmptyState();

  // MindMap event handlers
  mindMap.on('nodeSelect',       onNodeSelect);
  mindMap.on('nodeDeselect',     onNodeDeselect);
  mindMap.on('requestAddChild',  (parentId) => openAddNodeModal('addChild', parentId));
  mindMap.on('requestEditNode',  (nodeId)   => openAddNodeModalForEdit(nodeId));
  mindMap.on('requestNavigateLink', (nodeId, link) => navigateToOneNoteLink(link));
  mindMap.on('nodeAdd',          updateEmptyState);
  mindMap.on('nodeDelete',       updateEmptyState);
  mindMap.on('themeChange',      onThemeChange);

  // Zoom badge live update
  const svgEl = document.getElementById('mindmap-svg');
  const badge  = document.getElementById('zoom-badge');
  const origApply = mindMap._applyTransform.bind(mindMap);
  mindMap._applyTransform = function() {
    origApply();
    if (badge) badge.textContent = Math.round(mindMap._scale * 100) + '%';
  };

  updateEmptyState();
  centerMapOnResize();

  // Enable horizontal drag-scroll on both toolbar rows
  document.querySelectorAll('.tb-row').forEach(enableHorizontalScroll);

  // Arrow buttons for toolbar rows
  const tbRows = document.querySelectorAll('.tb-row');
  document.querySelectorAll('.tb-arrow').forEach(btn => {
    btn.addEventListener('click', () => {
      const row = tbRows[parseInt(btn.dataset.row)];
      if (row) row.scrollLeft += parseInt(btn.dataset.dir) * 120;
    });
    // Hold to keep scrolling
    let held;
    btn.addEventListener('mousedown', () => {
      held = setInterval(() => {
        const row = tbRows[parseInt(btn.dataset.row)];
        if (row) row.scrollLeft += parseInt(btn.dataset.dir) * 40;
      }, 80);
    });
    window.addEventListener('mouseup', () => clearInterval(held));
  });
}

// Office.js bootstrap — falls back gracefully if not in Office
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    officeReady = true;
    initApp();
  });
} else {
  // Running outside Office (browser dev mode)
  document.addEventListener('DOMContentLoaded', initApp);
}

/* ══════════════════════════════════════
   TOOLBAR BINDINGS
══════════════════════════════════════ */

function bindToolbar() {
  // Add Root Topic
  document.getElementById('tb-add-root').addEventListener('click', () => {
    openAddNodeModal('addRoot');
  });

  // Shape selector (toolbar, for NEW nodes default)
  document.querySelectorAll('.tb-shape-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tb-shape-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  // Font family — apply to selected node
  document.getElementById('tb-font-family').addEventListener('change', (e) => {
    if (mindMap.selectedId) applyPropertyToSelected({ fontFamily: e.target.value });
  });

  // Font size
  document.getElementById('tb-font-size').addEventListener('change', (e) => {
    const v = parseInt(e.target.value);
    if (mindMap.selectedId && v >= 8 && v <= 48) applyPropertyToSelected({ fontSize: v });
  });

  // Bold
  const boldBtn = document.getElementById('tb-bold');
  boldBtn.addEventListener('click', () => {
    if (!mindMap.selectedId) return;
    const n = mindMap.nodes[mindMap.selectedId];
    const next = !n.bold;
    boldBtn.classList.toggle('active', next);
    applyPropertyToSelected({ bold: next });
  });

  // Italic
  const italicBtn = document.getElementById('tb-italic');
  italicBtn.addEventListener('click', () => {
    if (!mindMap.selectedId) return;
    const n = mindMap.nodes[mindMap.selectedId];
    const next = !n.italic;
    italicBtn.classList.toggle('active', next);
    applyPropertyToSelected({ italic: next });
  });

  // Colors
  bindColorInput('tb-bg-color',     'tb-bg-swatch',     (v) => applyPropertyToSelected({ color: v, manualColor: true }));
  bindColorInput('tb-text-color',   'tb-text-swatch',   (v) => applyPropertyToSelected({ textColor: v, manualColor: true }));
  bindColorInput('tb-border-color', 'tb-border-swatch', (v) => applyPropertyToSelected({ borderColor: v, manualColor: true }));

  // Themes
  document.querySelectorAll('.tb-theme-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tb-theme-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      mindMap.applyTheme(btn.dataset.theme);
    });
  });

  // Zoom
  document.getElementById('tb-zoom-in').addEventListener('click', () => mindMap.zoomIn());
  document.getElementById('tb-zoom-out').addEventListener('click', () => mindMap.zoomOut());
  document.getElementById('tb-fit').addEventListener('click', () => mindMap.fitToScreen());

  // Save
  document.getElementById('tb-save').addEventListener('click', saveToFile);

  // Load
  document.getElementById('tb-load').addEventListener('click', () => {
    document.getElementById('tb-load-input').click();
  });
  document.getElementById('tb-load-input').addEventListener('change', loadFromFile);

  // Insert into OneNote
  document.getElementById('tb-insert-onenote').addEventListener('click', insertIntoOneNote);
}

function bindColorInput(inputId, swatchId, onChange) {
  const input  = document.getElementById(inputId);
  const swatch = document.getElementById(swatchId);
  if (!input) return;
  input.addEventListener('input', (e) => {
    swatch.style.background = e.target.value;
    onChange(e.target.value);
  });
}

/* ══════════════════════════════════════
   PANEL BINDINGS
══════════════════════════════════════ */

function bindPanel() {
  document.getElementById('panel-close').addEventListener('click', () => {
    mindMap.deselectAll();
  });

  // Live text update
  document.getElementById('pn-text').addEventListener('input', debounce((e) => {
    if (mindMap.selectedId) applyPropertyToSelected({ text: e.target.value });
  }, 300));

  // Shape buttons in panel
  document.querySelectorAll('#node-panel .shape-option').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#node-panel .shape-option').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (mindMap.selectedId) applyPropertyToSelected({ shape: btn.dataset.shape });
    });
  });

  // Panel color pickers
  document.getElementById('pn-bg-color').addEventListener('input', (e) => {
    applyPropertyToSelected({ color: e.target.value, manualColor: true });
  });
  document.getElementById('pn-text-color').addEventListener('input', (e) => {
    applyPropertyToSelected({ textColor: e.target.value, manualColor: true });
  });
  document.getElementById('pn-border-color').addEventListener('input', (e) => {
    applyPropertyToSelected({ borderColor: e.target.value, manualColor: true });
  });

  // Font
  document.getElementById('pn-font-family').addEventListener('change', (e) => {
    applyPropertyToSelected({ fontFamily: e.target.value });
  });
  document.getElementById('pn-font-size').addEventListener('change', (e) => {
    const v = parseInt(e.target.value);
    if (v >= 8 && v <= 48) applyPropertyToSelected({ fontSize: v });
  });

  // Bold / Italic toggles
  document.getElementById('pn-bold').addEventListener('click', () => {
    if (!mindMap.selectedId) return;
    const n = mindMap.nodes[mindMap.selectedId];
    applyPropertyToSelected({ bold: !n.bold });
    document.getElementById('pn-bold').classList.toggle('active', !n.bold);
  });
  document.getElementById('pn-italic').addEventListener('click', () => {
    if (!mindMap.selectedId) return;
    const n = mindMap.nodes[mindMap.selectedId];
    applyPropertyToSelected({ italic: !n.italic });
    document.getElementById('pn-italic').classList.toggle('active', !n.italic);
  });

  // Add Child Branch
  document.getElementById('pn-add-child').addEventListener('click', () => {
    if (mindMap.selectedId) openAddNodeModal('addChild', mindMap.selectedId);
  });

  // Add Note
  document.getElementById('pn-add-note').addEventListener('click', () => {
    if (mindMap.selectedId) openNoteModal(mindMap.selectedId);
  });

  // Link to Page
  document.getElementById('pn-link-page').addEventListener('click', () => {
    if (mindMap.selectedId) openLinkModal(mindMap.selectedId);
  });

  // Reset to theme colors
  document.getElementById('pn-reset-colors').addEventListener('click', () => {
    if (!mindMap.selectedId) return;
    const n = mindMap.nodes[mindMap.selectedId];
    const colors = getThemeColors(mindMap.themeId, n.depth);
    applyPropertyToSelected({
      color: colors.bg,
      textColor: colors.text,
      borderColor: colors.border,
      manualColor: false
    });
    populatePanelForNode(mindMap.selectedId);
    showToast('Theme colors restored');
  });

  // Delete node
  document.getElementById('pn-delete').addEventListener('click', () => {
    if (!mindMap.selectedId) return;
    openConfirmDelete(mindMap.selectedId);
  });
}

/* ══════════════════════════════════════
   NODE SELECT / DESELECT
══════════════════════════════════════ */

function onNodeSelect(nodeId, node) {
  document.getElementById('node-panel').classList.remove('hidden');
  populatePanelForNode(nodeId);
}

function onNodeDeselect() {
  document.getElementById('node-panel').classList.add('hidden');
}

function populatePanelForNode(nodeId) {
  const n = mindMap.nodes[nodeId];
  if (!n) return;

  document.getElementById('pn-text').value = n.text || '';

  // Shape
  document.querySelectorAll('#node-panel .shape-option').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.shape === n.shape);
  });

  // Colors
  document.getElementById('pn-bg-color').value     = n.color || '#4a90d9';
  document.getElementById('pn-text-color').value   = n.textColor || '#ffffff';
  document.getElementById('pn-border-color').value = n.borderColor || '#2171c2';

  // Font
  document.getElementById('pn-font-family').value = n.fontFamily || 'Segoe UI, Arial, sans-serif';
  document.getElementById('pn-font-size').value   = n.fontSize || 14;
  document.getElementById('pn-bold').classList.toggle('active', !!n.bold);
  document.getElementById('pn-italic').classList.toggle('active', !!n.italic);

  // Sync toolbar inputs
  syncToolbarToNode(n);
}

function syncToolbarToNode(n) {
  document.getElementById('tb-font-family').value = n.fontFamily || 'Segoe UI, Arial, sans-serif';
  document.getElementById('tb-font-size').value   = n.fontSize || 14;
  document.getElementById('tb-bold').classList.toggle('active', !!n.bold);
  document.getElementById('tb-italic').classList.toggle('active', !!n.italic);
  const bgSwatch     = document.getElementById('tb-bg-swatch');
  const textSwatch   = document.getElementById('tb-text-swatch');
  const borderSwatch = document.getElementById('tb-border-swatch');
  if (bgSwatch)     bgSwatch.style.background     = n.color || '#4a90d9';
  if (textSwatch)   textSwatch.style.background   = n.textColor || '#ffffff';
  if (borderSwatch) borderSwatch.style.background = n.borderColor || '#2171c2';
}

function applyPropertyToSelected(updates) {
  if (!mindMap.selectedId) return;
  mindMap.updateNode(mindMap.selectedId, updates);
  // Keep panel in sync if colors changed
  if (updates.color)       document.getElementById('pn-bg-color').value     = updates.color;
  if (updates.textColor)   document.getElementById('pn-text-color').value   = updates.textColor;
  if (updates.borderColor) document.getElementById('pn-border-color').value = updates.borderColor;
}

function onThemeChange(themeId) {
  if (mindMap.selectedId) populatePanelForNode(mindMap.selectedId);
}

/* ══════════════════════════════════════
   ADD NODE MODAL
══════════════════════════════════════ */

function openAddNodeModal(type, parentId = null) {
  pendingAction = { type, parentId };

  const title = type === 'addRoot' ? 'Add Root Topic' : 'Add Child Branch';
  document.getElementById('modal-add-title').textContent = title;
  document.getElementById('modal-add-confirm').textContent = type === 'addRoot' ? 'Add Topic' : 'Add Branch';

  // Reset form
  document.getElementById('modal-node-text').value = '';
  document.getElementById('modal-use-theme').checked = true;
  document.getElementById('modal-font-size').value = 14;

  // Set default shape from toolbar
  const activeShape = document.querySelector('.tb-shape-btn.active')?.dataset.shape || 'rect';
  document.querySelectorAll('#modal-add-node .shape-card').forEach(c => {
    c.classList.toggle('selected', c.dataset.shape === activeShape);
  });

  // Set default colors from current toolbar swatches
  document.getElementById('modal-bg-color').value     = document.getElementById('tb-bg-color').value;
  document.getElementById('modal-text-color').value   = document.getElementById('tb-text-color').value;
  document.getElementById('modal-border-color').value = document.getElementById('tb-border-color').value;

  openModal('modal-add-node');
  setTimeout(() => document.getElementById('modal-node-text').focus(), 100);
}

function openAddNodeModalForEdit(nodeId) {
  const n = mindMap.nodes[nodeId];
  if (!n) return;
  // Just select the node — editing is done via the panel
  mindMap.selectNode(nodeId);
}

function bindAddNodeModal() {
  // Shape card selection
  document.querySelectorAll('#modal-add-node .shape-card').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelectorAll('#modal-add-node .shape-card').forEach(c => c.classList.remove('selected'));
      card.classList.add('selected');
    });
  });

  // Use-theme toggle
  document.getElementById('modal-use-theme').addEventListener('change', (e) => {
    const colorRow = e.target.closest('.field-group').querySelector('.color-row');
    colorRow.style.opacity = e.target.checked ? '0.4' : '1';
    colorRow.style.pointerEvents = e.target.checked ? 'none' : 'all';
  });
  // Apply initial state
  const colorRow = document.querySelector('#modal-add-node .color-row');
  colorRow.style.opacity = '0.4';
  colorRow.style.pointerEvents = 'none';

  // Confirm
  document.getElementById('modal-add-confirm').addEventListener('click', () => {
    const text = document.getElementById('modal-node-text').value.trim();
    if (!text) {
      document.getElementById('modal-node-text').focus();
      return;
    }
    const shape      = document.querySelector('#modal-add-node .shape-card.selected')?.dataset.shape || 'rect';
    const useTheme   = document.getElementById('modal-use-theme').checked;
    const fontFamily = document.getElementById('modal-font-family').value;
    const fontSize   = parseInt(document.getElementById('modal-font-size').value) || 14;

    const opts = { shape, fontFamily, fontSize };
    if (!useTheme) {
      opts.color       = document.getElementById('modal-bg-color').value;
      opts.textColor   = document.getElementById('modal-text-color').value;
      opts.borderColor = document.getElementById('modal-border-color').value;
    }

    if (pendingAction.type === 'addRoot') {
      mindMap.addRoot(text, opts);
      mindMap.centerRoot();
    } else {
      mindMap.addChild(pendingAction.parentId, text, opts);
    }
    closeModal('modal-add-node');
    updateEmptyState();
  });

  // Enter key in text field
  document.getElementById('modal-node-text').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') document.getElementById('modal-add-confirm').click();
  });
}

/* ══════════════════════════════════════
   NOTE MODAL
══════════════════════════════════════ */

let noteTargetId = null;

function openNoteModal(nodeId) {
  noteTargetId = nodeId;
  const n = mindMap.nodes[nodeId];
  document.getElementById('modal-note-text').value = n?.note || '';
  openModal('modal-note');
  setTimeout(() => document.getElementById('modal-note-text').focus(), 100);
}

function bindNoteModal() {
  document.getElementById('modal-note-confirm').addEventListener('click', () => {
    if (!noteTargetId) return;
    const text = document.getElementById('modal-note-text').value.trim();
    mindMap.setNodeNote(noteTargetId, text);
    closeModal('modal-note');
    showToast(text ? '✎ Note saved' : 'Note removed');
  });
}

/* ══════════════════════════════════════
   LINK MODAL
══════════════════════════════════════ */

let linkTargetId = null;

function openLinkModal(nodeId) {
  linkTargetId = nodeId;
  const n = mindMap.nodes[nodeId];
  selectedPageForLink = n?.oneNoteLink || null;

  // Show current link if any
  const currentDiv  = document.getElementById('link-current');
  const currentName = document.getElementById('link-current-name');
  if (n?.oneNoteLink) {
    currentDiv.style.display = 'block';
    currentName.textContent  = n.oneNoteLink.name || n.oneNoteLink.id || '—';
  } else {
    currentDiv.style.display = 'none';
  }

  document.getElementById('link-manual-name').value = n?.oneNoteLink?.name || '';
  loadOneNotePages();
  openModal('modal-link');
}

function bindLinkModal() {
  document.getElementById('link-refresh-pages').addEventListener('click', loadOneNotePages);

  document.getElementById('link-remove').addEventListener('click', () => {
    selectedPageForLink = null;
    document.getElementById('link-current').style.display = 'none';
    document.getElementById('link-manual-name').value = '';
    document.querySelectorAll('.page-item').forEach(i => i.classList.remove('selected'));
  });

  document.getElementById('modal-link-confirm').addEventListener('click', () => {
    if (!linkTargetId) return;
    const manualName = document.getElementById('link-manual-name').value.trim();

    let link = selectedPageForLink;
    if (!link && manualName) {
      link = { id: null, name: manualName };
    }

    mindMap.setNodeLink(linkTargetId, link || null);
    closeModal('modal-link');
    showToast(link ? `🔗 Linked to "${link.name || link.id}"` : '🔗 Link removed');
    populatePanelForNode(linkTargetId);
  });
}

async function loadOneNotePages() {
  const list = document.getElementById('link-page-list');
  list.innerHTML = '<div style="text-align:center;padding:16px;color:var(--text-muted);"><div class="spinner" style="margin:0 auto 8px;"></div>Loading pages…</div>';

  try {
    const pages = await getOneNotePages();
    if (!pages.length) {
      list.innerHTML = '<div style="text-align:center;padding:12px;color:var(--text-muted);font-size:12px;">No pages found. Enter a name manually below.</div>';
      return;
    }
    list.innerHTML = '';
    pages.forEach(page => {
      const item = document.createElement('div');
      item.className = 'page-item';
      item.innerHTML = `
        <span class="page-icon">📄</span>
        <div>
          <div class="page-name">${escapeHtml(page.name)}</div>
          ${page.section ? `<div class="page-section">${escapeHtml(page.section)}</div>` : ''}
        </div>`;
      item.addEventListener('click', () => {
        document.querySelectorAll('.page-item').forEach(i => i.classList.remove('selected'));
        item.classList.add('selected');
        selectedPageForLink = { id: page.id, name: page.name, section: page.section, url: page.url };
        document.getElementById('link-manual-name').value = page.name;
      });
      list.appendChild(item);
    });
    // Re-select existing link if any
    if (linkTargetId) {
      const cur = mindMap.nodes[linkTargetId]?.oneNoteLink;
      if (cur?.id) {
        list.querySelectorAll('.page-item').forEach((item, i) => {
          if (pages[i]?.id === cur.id) item.classList.add('selected');
        });
      }
    }
  } catch (err) {
    list.innerHTML = `<div style="padding:12px;color:var(--text-muted);font-size:12px;">Could not load pages. ${officeReady ? 'Make sure OneNote is open.' : 'Running outside Office.'}<br/>Enter a name manually below.</div>`;
  }
}

/* ══════════════════════════════════════
   DELETE MODAL
══════════════════════════════════════ */

let deleteTargetId = null;

function openConfirmDelete(nodeId) {
  deleteTargetId = nodeId;
  const n = mindMap.nodes[nodeId];
  document.getElementById('confirm-delete-name').textContent = `"${n?.text || 'this node'}"`;
  openModal('modal-confirm-delete');
}

function bindDeleteModal() {
  document.getElementById('confirm-delete-btn').addEventListener('click', () => {
    if (!deleteTargetId) return;
    mindMap.deleteNode(deleteTargetId);
    deleteTargetId = null;
    closeModal('modal-confirm-delete');
    showToast('Node deleted');
  });
}

/* ══════════════════════════════════════
   MODAL HELPERS
══════════════════════════════════════ */

function openModal(id) {
  document.getElementById(id).classList.add('visible');
}

function closeModal(id) {
  document.getElementById(id).classList.remove('visible');
}

function bindModalClose() {
  // Close buttons
  document.querySelectorAll('[data-close]').forEach(btn => {
    btn.addEventListener('click', () => closeModal(btn.dataset.close));
  });
  // Click outside modal
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal(overlay.id);
    });
  });
  // Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.modal-overlay.visible').forEach(o => closeModal(o.id));
    }
  });
}

/* ══════════════════════════════════════
   KEYBOARD SHORTCUTS
══════════════════════════════════════ */

function bindKeyboard() {
  document.addEventListener('keydown', (e) => {
    if (isModalOpen()) return;

    const ctrl = e.ctrlKey || e.metaKey;

    // Ctrl+Enter — add root
    if (ctrl && e.key === 'Enter') {
      e.preventDefault();
      openAddNodeModal('addRoot');
    }
    // Tab — add child to selected
    if (e.key === 'Tab' && mindMap.selectedId) {
      e.preventDefault();
      openAddNodeModal('addChild', mindMap.selectedId);
    }
    // Delete / Backspace — delete selected
    if ((e.key === 'Delete' || e.key === 'Backspace') && mindMap.selectedId && !isInputFocused()) {
      openConfirmDelete(mindMap.selectedId);
    }
    // F2 — focus text field in panel
    if (e.key === 'F2' && mindMap.selectedId) {
      const t = document.getElementById('pn-text');
      if (t) { t.focus(); t.select(); }
    }
    // +/= — zoom in
    if ((e.key === '+' || e.key === '=') && ctrl) {
      e.preventDefault();
      mindMap.zoomIn();
    }
    // - — zoom out
    if (e.key === '-' && ctrl) {
      e.preventDefault();
      mindMap.zoomOut();
    }
    // 0 — fit to screen
    if (e.key === '0' && ctrl) {
      e.preventDefault();
      mindMap.fitToScreen();
    }
    // Escape — deselect
    if (e.key === 'Escape') {
      mindMap.deselectAll();
    }
  });
}

function isModalOpen() {
  return !!document.querySelector('.modal-overlay.visible');
}

function isInputFocused() {
  const el = document.activeElement;
  return el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT');
}

/* ══════════════════════════════════════
   EMPTY STATE
══════════════════════════════════════ */

function bindEmptyState() {
  document.getElementById('es-add-btn').addEventListener('click', () => {
    openAddNodeModal('addRoot');
  });
}

function updateEmptyState() {
  const hasNodes = Object.keys(mindMap.nodes).length > 0;
  document.getElementById('empty-state').classList.toggle('hidden', hasNodes);
}

/* ══════════════════════════════════════
   SAVE / LOAD
══════════════════════════════════════ */

function saveToFile() {
  const json = mindMap.serialize();
  const blob = new Blob([json], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'mindmap-' + new Date().toISOString().slice(0, 10) + '.json';
  a.click();
  URL.revokeObjectURL(url);
  showToast('💾 Mind map saved');
}

function loadFromFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      mindMap.deserialize(ev.target.result);
      // Sync theme button UI
      document.querySelectorAll('.tb-theme-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.theme === mindMap.themeId);
      });
      mindMap.fitToScreen();
      updateEmptyState();
      showToast('📂 Mind map loaded');
    } catch {
      showToast('❌ Failed to load file');
    }
  };
  reader.readAsText(file);
  e.target.value = '';
}

/* ══════════════════════════════════════
   INSERT INTO ONENOTE
══════════════════════════════════════ */

async function insertIntoOneNote() {
  if (!officeReady || typeof OneNote === 'undefined') {
    showToast('⚠️ Not running inside OneNote');
    return;
  }
  if (!mindMap.rootId) {
    showToast('⚠️ Nothing to insert — build a mind map first');
    return;
  }

  showToast('⏳ Preparing image…');

  try {
    const pngDataUrl = await exportMapToPng();

    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
      // addOutline(left, top, html) — inserts HTML content at position on page
      page.addOutline(20, 120, `<img src="${pngDataUrl}" />`);
      await context.sync();
    });

    showToast('✅ Mind map added to the OneNote page!');
  } catch (err) {
    console.error('Insert error:', err);
    showToast('❌ Could not insert: ' + (err.message || err));
  }
}

/** Render the mind map SVG to a 2× resolution PNG and return a base64 data URL. */
function exportMapToPng() {
  return new Promise((resolve, reject) => {
    const nodes = Object.values(mindMap.nodes);
    if (!nodes.length) { reject(new Error('No nodes')); return; }

    // Compute tight bounding box around all node shapes
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    nodes.forEach(n => {
      minX = Math.min(minX, n.x - n.w / 2);
      minY = Math.min(minY, n.y - n.h / 2);
      maxX = Math.max(maxX, n.x + n.w / 2);
      maxY = Math.max(maxY, n.y + n.h / 2);
    });

    const pad  = 40;
    const vbX  = minX - pad,  vbY = minY - pad;
    const vbW  = maxX - minX + pad * 2;
    const vbH  = maxY - minY + pad * 2;
    const scale = 2; // retina / crisp export

    // Clone the SVG and set a fixed viewBox for export
    const svgEl = document.getElementById('mindmap-svg');
    const clone = svgEl.cloneNode(true);
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    clone.setAttribute('width',   vbW * scale);
    clone.setAttribute('height',  vbH * scale);
    clone.setAttribute('viewBox', `${vbX} ${vbY} ${vbW} ${vbH}`);

    // Reset canvas pan/zoom transform — the viewBox handles framing
    const mainG = clone.querySelector('#mm-main');
    if (mainG) mainG.setAttribute('transform', 'scale(1)');

    // Fill background
    const bg = clone.querySelector('.mm-bg');
    if (bg) { bg.setAttribute('width', '100%'); bg.setAttribute('height', '100%'); }

    const svgStr  = new XMLSerializer().serializeToString(clone);
    const svgB64  = btoa(unescape(encodeURIComponent(svgStr)));
    const imgSrc  = 'data:image/svg+xml;base64,' + svgB64;

    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width  = vbW * scale;
      canvas.height = vbH * scale;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => reject(new Error('SVG → canvas render failed'));
    img.src = imgSrc;
  });
}

/* ══════════════════════════════════════
   ONENOTE API — Load Pages
══════════════════════════════════════ */

async function getOneNotePages() {
  if (!officeReady || typeof OneNote === 'undefined') {
    // Return dummy data for dev/browser mode
    return [
      { id: 'p1', name: 'Chemistry Notes',   section: 'Science' },
      { id: 'p2', name: 'Math Summary',       section: 'Math' },
      { id: 'p3', name: 'History Chapter 3',  section: 'History' },
      { id: 'p4', name: 'Biology — Cells',    section: 'Science' },
      { id: 'p5', name: 'Vocabulary List',    section: 'Languages' }
    ];
  }

  return new Promise((resolve, reject) => {
    OneNote.run(async (context) => {
      try {
        const notebooks = context.application.getNotebooks();
        notebooks.load('name');
        await context.sync();

        const pages = [];
        for (const nb of notebooks.items) {
          const sections = nb.sections;
          sections.load('name');
          await context.sync();
          for (const sec of sections.items) {
            const ps = sec.pages;
            ps.load('title,id,links');
            await context.sync();
            ps.items.forEach(p => pages.push({
              id:      p.id,
              name:    p.title || '(Untitled)',
              section: `${nb.name} / ${sec.name}`,
              url:     p.links?.oneNoteClientUrl?.href || null
            }));
          }
        }
        resolve(pages);
      } catch (e) {
        reject(e);
      }
    });
  });
}

async function navigateToOneNoteLink(link) {
  if (!link) return;
  if (!officeReady || typeof OneNote === 'undefined') {
    showToast(`Would navigate to: "${link.name || link.id}"`);
    return;
  }
  try {
    if (link.url) {
      // Try to open the page via its client URL
      window.open(link.url, '_blank');
    } else if (link.id) {
      await OneNote.run(async (context) => {
        const page = context.application.getPageOrNullObject(link.id);
        page.load('id');
        await context.sync();
        if (!page.isNullObject) {
          context.application.navigateTo(page, null);
          await context.sync();
        }
      });
    }
    showToast(`📄 Navigating to "${link.name}"`);
  } catch (err) {
    showToast('❌ Could not navigate to page');
    console.error(err);
  }
}

/* ══════════════════════════════════════
   RESIZE HANDLER
══════════════════════════════════════ */

function centerMapOnResize() {
  window.addEventListener('resize', debounce(() => {
    if (mindMap && !mindMap.rootId) return;
  }, 400));
}

/** Make an element horizontally scrollable via mouse drag and scroll wheel. */
function enableHorizontalScroll(el) {
  if (!el) return;
  let down = false, startX = 0, scrollLeft = 0;

  el.addEventListener('mousedown', (e) => {
    down = true;
    startX = e.pageX - el.getBoundingClientRect().left;
    scrollLeft = el.scrollLeft;
    el.style.cursor = 'grabbing';
    e.preventDefault();
  });
  window.addEventListener('mouseup',   () => { down = false; el.style.cursor = ''; });
  window.addEventListener('mousemove', (e) => {
    if (!down) return;
    const x = e.pageX - el.getBoundingClientRect().left;
    el.scrollLeft = scrollLeft - (x - startX);
  });

  // Mouse wheel scrolls horizontally
  el.addEventListener('wheel', (e) => {
    e.preventDefault();
    el.scrollLeft += e.deltaY || e.deltaX;
  }, { passive: false });

  // Touch scroll
  let touchStartX = 0, touchScrollLeft = 0;
  el.addEventListener('touchstart', (e) => {
    touchStartX    = e.touches[0].clientX;
    touchScrollLeft = el.scrollLeft;
  }, { passive: true });
  el.addEventListener('touchmove', (e) => {
    const dx = touchStartX - e.touches[0].clientX;
    el.scrollLeft = touchScrollLeft + dx;
  }, { passive: true });
}

/* ══════════════════════════════════════
   UTILITIES
══════════════════════════════════════ */

function showToast(msg, duration = 2500) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), duration);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function debounce(fn, delay) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), delay);
  };
}
