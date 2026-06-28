/**
 * MindMap Engine
 * SVG-based mind map with auto-layout, pan/zoom, collapse/expand,
 * per-node styling, themes, note tooltips, and OneNote linking.
 */

const MM_H_GAP = 230;   // horizontal distance between parent & child centers
const MM_V_GAP = 28;    // vertical gap between sibling subtrees
const MM_NODE_H = 44;   // default node height
const MM_NODE_MIN_W = 120;
const MM_NODE_PAD_X = 20;
const MM_NODE_PAD_Y = 12;
const MM_BTN_R = 10;    // radius of +/- buttons

function mmGenId() {
  return 'n' + Math.random().toString(36).slice(2, 9) + Date.now().toString(36);
}

class MindMap {
  constructor(svgEl, tooltipEl) {
    this.svg = svgEl;
    this.tooltipEl = tooltipEl;
    this.nodes = {};       // id -> node object
    this.rootId = null;
    this.selectedId = null;
    this.themeId = 'classic';
    this._handlers = {};   // event name -> [callbacks]

    this._panX = 0;
    this._panY = 0;
    this._scale = 1;
    this._panning = false;
    this._panStart = null;

    this._buildSVGStructure();
    this._attachCanvasEvents();
  }

  // ─────────────────────────────── SVG INIT ────────────────────────────────

  _buildSVGStructure() {
    this.svg.innerHTML = '';
    this.svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

    // Defs: arrowhead markers
    const defs = this._el('defs');
    defs.innerHTML = `
      <marker id="arrow-end" markerWidth="10" markerHeight="7"
              refX="9" refY="3.5" orient="auto">
        <polygon points="0 0, 10 3.5, 0 7" fill="#95a5a6" class="arrow-marker"/>
      </marker>
      <filter id="node-shadow" x="-20%" y="-20%" width="140%" height="140%">
        <feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="rgba(0,0,0,0.15)"/>
      </filter>`;
    this.svg.appendChild(defs);

    // Background rect to catch pan events
    this._bgRect = this._el('rect');
    this._bgRect.setAttribute('width', '100%');
    this._bgRect.setAttribute('height', '100%');
    this._bgRect.setAttribute('fill', '#f7f9fc');
    this._bgRect.setAttribute('class', 'mm-bg');
    this.svg.appendChild(this._bgRect);

    // Main transform group
    this._mainG = this._el('g');
    this._mainG.setAttribute('id', 'mm-main');
    this.svg.appendChild(this._mainG);

    this._edgesG = this._el('g', { id: 'mm-edges' });
    this._nodesG = this._el('g', { id: 'mm-nodes' });
    this._ctrlsG = this._el('g', { id: 'mm-ctrls' });
    this._mainG.appendChild(this._edgesG);
    this._mainG.appendChild(this._nodesG);
    this._mainG.appendChild(this._ctrlsG);

    this._applyTransform();
  }

  _applyTransform() {
    this._mainG.setAttribute('transform',
      `translate(${this._panX}, ${this._panY}) scale(${this._scale})`);
  }

  // ──────────────────────────── PUBLIC API ─────────────────────────────────

  /** Add a root node. Returns node id. */
  addRoot(text, opts = {}) {
    if (this.rootId && this.nodes[this.rootId]) {
      // Already has a root – add as sibling (another root means additional independent root)
      // For simplicity, treat subsequent roots as children of an implicit root
      // Or just add at the same level – let's allow multiple roots by making a virtual root
      // We'll just add it as another top-level node connected to nothing
      return this._createNode(text, null, opts, 0);
    }
    const id = this._createNode(text, null, opts, 0);
    this.rootId = id;
    this._layout();
    this._render();
    this._emit('nodeAdd', id);
    return id;
  }

  /** Add a child node to parentId. Returns node id. */
  addChild(parentId, text, opts = {}) {
    if (!this.nodes[parentId]) return null;
    const parent = this.nodes[parentId];
    const depth = parent.depth + 1;
    const id = this._createNode(text, parentId, opts, depth);
    parent.children.push(id);
    this._layout();
    this._render();
    this._emit('nodeAdd', id);
    this.selectNode(id);
    return id;
  }

  /** Delete a node and all its descendants. */
  deleteNode(nodeId) {
    if (!nodeId || nodeId === this.rootId) return;
    const node = this.nodes[nodeId];
    if (!node) return;

    // Remove from parent's children list
    if (node.parent && this.nodes[node.parent]) {
      this.nodes[node.parent].children =
        this.nodes[node.parent].children.filter(id => id !== nodeId);
    }
    // Recursively delete descendants
    this._deleteSubtree(nodeId);

    if (this.selectedId === nodeId) {
      this.selectedId = null;
      this._emit('nodeDeselect');
    }
    this._layout();
    this._render();
    this._emit('nodeDelete', nodeId);
  }

  /** Update node properties (text, color, shape, font, etc.) */
  updateNode(nodeId, updates) {
    const node = this.nodes[nodeId];
    if (!node) return;
    Object.assign(node, updates);
    this._layout();
    this._render();
    this._emit('nodeUpdate', nodeId);
  }

  /** Toggle collapse/expand of node's children. */
  toggleCollapse(nodeId) {
    const node = this.nodes[nodeId];
    if (!node || !node.children.length) return;
    node.collapsed = !node.collapsed;
    this._layout();
    this._render();
  }

  /** Select a node (highlight + emit event). */
  selectNode(nodeId) {
    const prev = this.selectedId;
    this.selectedId = nodeId;
    if (prev) this._updateNodeVisual(prev);
    if (nodeId) this._updateNodeVisual(nodeId);
    this._emit('nodeSelect', nodeId, this.nodes[nodeId]);
  }

  /** Deselect current node. */
  deselectAll() {
    const prev = this.selectedId;
    this.selectedId = null;
    if (prev) this._updateNodeVisual(prev);
    this._emit('nodeDeselect');
  }

  /** Apply a named theme to the entire map. */
  applyTheme(themeId) {
    this.themeId = themeId;
    // Update canvas background
    const theme = (typeof THEMES !== 'undefined') ? THEMES[themeId] : null;
    if (theme) this._bgRect.setAttribute('fill', theme.canvas);
    // Recompute colors for all nodes that haven't been manually overridden
    Object.values(this.nodes).forEach(n => {
      if (!n.manualColor) {
        const colors = getThemeColors(themeId, n.depth);
        n.color = colors.bg;
        n.textColor = colors.text;
        n.borderColor = colors.border;
      }
    });
    this._render();
    this._emit('themeChange', themeId);
  }

  /** Set a hover note/association for a node. */
  setNodeNote(nodeId, note) {
    if (this.nodes[nodeId]) {
      this.nodes[nodeId].note = note;
      this._render();
    }
  }

  /** Set a OneNote page link for a node. */
  setNodeLink(nodeId, link) {
    if (this.nodes[nodeId]) {
      this.nodes[nodeId].oneNoteLink = link;
      this._render();
    }
  }

  /** Zoom in. */
  zoomIn() { this._zoom(1.2, this._svgCenterX(), this._svgCenterY()); }

  /** Zoom out. */
  zoomOut() { this._zoom(0.8, this._svgCenterX(), this._svgCenterY()); }

  /** Reset zoom and center the map. */
  fitToScreen() {
    if (!this.rootId) return;
    // Find bounding box of all nodes
    const nodes = Object.values(this.nodes);
    if (!nodes.length) return;
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    nodes.forEach(n => {
      minX = Math.min(minX, n.x - n.w / 2);
      minY = Math.min(minY, n.y - n.h / 2);
      maxX = Math.max(maxX, n.x + n.w / 2);
      maxY = Math.max(maxY, n.y + n.h / 2);
    });
    const svgW = this.svg.clientWidth || 800;
    const svgH = this.svg.clientHeight || 600;
    const padding = 60;
    const mapW = maxX - minX + padding * 2;
    const mapH = maxY - minY + padding * 2;
    this._scale = Math.min(svgW / mapW, svgH / mapH, 1.5);
    this._panX = svgW / 2 - (minX + (maxX - minX) / 2) * this._scale;
    this._panY = svgH / 2 - (minY + (maxY - minY) / 2) * this._scale;
    this._applyTransform();
  }

  /** Center the root node in the SVG. */
  centerRoot() {
    this._panX = (this.svg.clientWidth || 800) / 2;
    this._panY = (this.svg.clientHeight || 600) / 2;
    this._scale = 1;
    this._applyTransform();
  }

  /** Serialize map to JSON. */
  serialize() {
    return JSON.stringify({
      rootId: this.rootId,
      themeId: this.themeId,
      nodes: this.nodes
    });
  }

  /** Restore map from JSON. */
  deserialize(json) {
    try {
      const data = JSON.parse(json);
      this.nodes = data.nodes || {};
      this.rootId = data.rootId || null;
      this.themeId = data.themeId || 'classic';
      this.selectedId = null;
      this._layout();
      this._render();
      const theme = (typeof THEMES !== 'undefined') ? THEMES[this.themeId] : null;
      if (theme) this._bgRect.setAttribute('fill', theme.canvas);
    } catch (e) {
      console.error('MindMap.deserialize error:', e);
    }
  }

  /** Register an event callback. */
  on(event, cb) {
    if (!this._handlers[event]) this._handlers[event] = [];
    this._handlers[event].push(cb);
    return () => {
      this._handlers[event] = this._handlers[event].filter(f => f !== cb);
    };
  }

  // ─────────────────────────────── LAYOUT ──────────────────────────────────

  _layout() {
    if (!this.rootId || !this.nodes[this.rootId]) return;
    const root = this.nodes[this.rootId];

    // Assign node sizes first
    Object.values(this.nodes).forEach(n => this._measureNode(n));

    // Place root at origin
    root.x = 0;
    root.y = 0;

    // Split direct children: even-index -> right, odd-index -> left
    const visChildren = this._visibleChildren(this.rootId);
    const rightKids = visChildren.filter((_, i) => i % 2 === 0);
    const leftKids  = visChildren.filter((_, i) => i % 2 === 1);

    this._layoutSide(rightKids, root, 'right');
    this._layoutSide(leftKids,  root, 'left');

    // Handle extra roots (multi-root support)
    Object.values(this.nodes)
      .filter(n => !n.parent && n.id !== this.rootId)
      .forEach((n, i) => {
        n.x = 0;
        n.y = root.y + (i + 1) * 300;
        this._layoutSide(this._visibleChildren(n.id).filter((_, j) => j % 2 === 0), n, 'right');
        this._layoutSide(this._visibleChildren(n.id).filter((_, j) => j % 2 === 1), n, 'left');
      });
  }

  _layoutSide(childIds, parent, direction) {
    if (!childIds.length) return;
    const totalH = childIds.reduce((s, id) => s + this._subtreeH(id), 0);
    const childX = parent.x + (direction === 'right' ? MM_H_GAP : -MM_H_GAP);
    let curY = parent.y - totalH / 2;
    childIds.forEach(id => {
      const h = this._subtreeH(id);
      this._layoutSubtree(id, childX, curY + h / 2, direction);
      curY += h;
    });
  }

  _layoutSubtree(nodeId, x, y, direction) {
    const node = this.nodes[nodeId];
    if (!node) return;
    node.x = x;
    node.y = y;
    node.direction = direction;
    if (node.collapsed) return;
    const kids = this._visibleChildren(nodeId);
    if (!kids.length) return;
    const totalH = kids.reduce((s, id) => s + this._subtreeH(id), 0);
    const childX = x + (direction === 'right' ? MM_H_GAP : -MM_H_GAP);
    let curY = y - totalH / 2;
    kids.forEach(id => {
      const h = this._subtreeH(id);
      this._layoutSubtree(id, childX, curY + h / 2, direction);
      curY += h;
    });
  }

  _subtreeH(nodeId) {
    const node = this.nodes[nodeId];
    if (!node) return MM_NODE_H + MM_V_GAP;
    if (node.collapsed || !node.children.length) return (node.h || MM_NODE_H) + MM_V_GAP;
    const kids = this._visibleChildren(nodeId);
    if (!kids.length) return (node.h || MM_NODE_H) + MM_V_GAP;
    return Math.max(
      (node.h || MM_NODE_H) + MM_V_GAP,
      kids.reduce((s, id) => s + this._subtreeH(id), 0)
    );
  }

  _measureNode(node) {
    const fontSize = node.fontSize || 14;
    const text = node.text || '';
    const charW = fontSize * 0.58;
    node.w = Math.max(MM_NODE_MIN_W, text.length * charW + MM_NODE_PAD_X * 2);
    node.h = fontSize + MM_NODE_PAD_Y * 2;
  }

  _visibleChildren(nodeId) {
    const node = this.nodes[nodeId];
    if (!node) return [];
    return node.children.filter(id => this.nodes[id]);
  }

  // ─────────────────────────────── RENDER ──────────────────────────────────

  _render() {
    this._edgesG.innerHTML = '';
    this._nodesG.innerHTML = '';
    this._ctrlsG.innerHTML = '';

    if (!this.rootId) return;

    // Update arrow marker color from theme
    const theme = (typeof THEMES !== 'undefined') ? THEMES[this.themeId] : null;
    const arrowColor = theme ? theme.edge : '#95a5a6';
    const markers = this.svg.querySelectorAll('.arrow-marker');
    markers.forEach(m => m.setAttribute('fill', arrowColor));

    // Render all edges first (so they appear below nodes)
    Object.values(this.nodes).forEach(node => {
      if (node.parent && this.nodes[node.parent]) {
        this._renderEdge(this.nodes[node.parent], node, arrowColor);
        this._renderCollapseBtn(this.nodes[node.parent], node);
      }
    });

    // Render nodes
    Object.values(this.nodes).forEach(node => {
      this._renderNode(node);
    });
  }

  _renderNode(node) {
    const g = this._el('g');
    g.setAttribute('transform', `translate(${node.x}, ${node.y})`);
    g.setAttribute('class', 'mm-node' + (node.id === this.selectedId ? ' mm-selected' : ''));
    g.setAttribute('data-id', node.id);
    g.style.cursor = 'pointer';

    const hw = node.w / 2, hh = node.h / 2;
    const isSelected = node.id === this.selectedId;
    const stroke = isSelected ? '#f39c12' : (node.borderColor || '#666');
    const strokeW = isSelected ? 3 : 2;

    // Shape
    const fill   = node.color || '#4a90d9';
    const shadow = 'url(#node-shadow)';
    const baseAttrs = { fill, stroke, 'stroke-width': strokeW, filter: shadow };

    if (node.shape === 'ellipse') {
      g.appendChild(this._el('ellipse', { cx: 0, cy: 0, rx: hw, ry: hh, ...baseAttrs }));

    } else if (node.shape === 'rect-sharp') {
      g.appendChild(this._el('rect', { x: -hw, y: -hh, width: node.w, height: node.h, rx: 2, ...baseAttrs }));

    } else if (node.shape === 'diamond') {
      // Expand bounding box so text fits inside the diamond
      const dw = hw * 1.4, dh = hh * 1.5;
      node._shapeW = dw; node._shapeH = dh;
      const pts = `0,${-dh} ${dw},0 0,${dh} ${-dw},0`;
      g.appendChild(this._el('polygon', { points: pts, ...baseAttrs }));

    } else if (node.shape === 'hexagon') {
      // Flat-top hexagon
      const r2 = hw + 10;
      const pts = [
        [r2 * 0.5, -hh], [r2, 0], [r2 * 0.5, hh],
        [-r2 * 0.5, hh], [-r2, 0], [-r2 * 0.5, -hh]
      ].map(p => p.join(',')).join(' ');
      g.appendChild(this._el('polygon', { points: pts, ...baseAttrs }));

    } else if (node.shape === 'parallelogram') {
      const sk = hh * 0.5; // skew offset
      const pts = `${-hw + sk},${-hh} ${hw + sk},${-hh} ${hw - sk},${hh} ${-hw - sk},${hh}`;
      g.appendChild(this._el('polygon', { points: pts, ...baseAttrs }));

    } else {
      // Default: rounded rectangle
      g.appendChild(this._el('rect', { x: -hw, y: -hh, width: node.w, height: node.h, rx: 8, ...baseAttrs }));
    }

    // Text
    const textLines = this._wrapText(node.text || '', node.w - MM_NODE_PAD_X * 2, node.fontSize || 14);
    const lineH = (node.fontSize || 14) * 1.3;
    const totalTH = lineH * textLines.length;
    textLines.forEach((line, i) => {
      const t = this._el('text', {
        x: 0,
        y: -totalTH / 2 + lineH * i + lineH * 0.5,
        'text-anchor': 'middle',
        'dominant-baseline': 'middle',
        fill: node.textColor || '#fff',
        'font-family': node.fontFamily || 'Segoe UI, Arial, sans-serif',
        'font-size': node.fontSize || 14,
        'font-weight': node.bold ? 'bold' : 'normal',
        'font-style': node.italic ? 'italic' : 'normal',
        'pointer-events': 'none'
      });
      t.textContent = line;
      g.appendChild(t);
    });

    // Note indicator (📝 small dot, top-right)
    if (node.note) {
      const noteBtn = this._el('circle', {
        cx: hw - 6, cy: -hh + 6, r: 6,
        fill: '#f39c12', stroke: '#fff', 'stroke-width': 1.5,
        class: 'mm-note-indicator', cursor: 'help'
      });
      noteBtn.addEventListener('mouseenter', (e) => this._showTooltip(e, node.note));
      noteBtn.addEventListener('mouseleave', () => this._hideTooltip());
      g.appendChild(noteBtn);

      const noteT = this._el('text', {
        x: hw - 6, y: -hh + 6,
        'text-anchor': 'middle', 'dominant-baseline': 'middle',
        'font-size': 7, fill: '#fff', 'pointer-events': 'none'
      });
      noteT.textContent = '✎';
      g.appendChild(noteT);
    }

    // OneNote link indicator (chain icon, top-left)
    if (node.oneNoteLink) {
      const linkBtn = this._el('circle', {
        cx: -hw + 6, cy: -hh + 6, r: 6,
        fill: '#8e44ad', stroke: '#fff', 'stroke-width': 1.5,
        class: 'mm-link-indicator', cursor: 'pointer'
      });
      linkBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this._emit('requestNavigateLink', node.id, node.oneNoteLink);
      });
      linkBtn.title = 'Navigate to OneNote page';
      const linkT = this._el('text', {
        x: -hw + 6, y: -hh + 6,
        'text-anchor': 'middle', 'dominant-baseline': 'middle',
        'font-size': 7, fill: '#fff', 'pointer-events': 'none'
      });
      linkT.textContent = '⛓';
      g.appendChild(linkBtn);
      g.appendChild(linkT);
    }

    // Add-child button (+) on outgoing side
    const addBtnX = (node.direction === 'left' ? -hw - MM_BTN_R - 4 : hw + MM_BTN_R + 4);
    const addBtnCircle = this._el('circle', {
      cx: addBtnX, cy: 0, r: MM_BTN_R,
      fill: '#fff', stroke: node.borderColor || '#666',
      'stroke-width': 1.5, class: 'mm-add-btn', cursor: 'pointer',
      opacity: 0.85
    });
    const addBtnT = this._el('text', {
      x: addBtnX, y: 0,
      'text-anchor': 'middle', 'dominant-baseline': 'middle',
      'font-size': 14, 'font-weight': 'bold',
      fill: node.borderColor || '#666', 'pointer-events': 'none'
    });
    addBtnT.textContent = '+';
    addBtnCircle.addEventListener('click', (e) => {
      e.stopPropagation();
      this._emit('requestAddChild', node.id);
    });
    g.appendChild(addBtnCircle);
    g.appendChild(addBtnT);

    // Node click → select
    g.addEventListener('click', (e) => {
      e.stopPropagation();
      this.selectNode(node.id);
    });

    // Double-click → request edit
    g.addEventListener('dblclick', (e) => {
      e.stopPropagation();
      this._emit('requestEditNode', node.id);
    });

    this._nodesG.appendChild(g);
  }

  _renderEdge(parent, child, arrowColor) {
    const isRight = child.direction === 'right' || child.x >= parent.x;
    const x1 = isRight ? parent.x + parent.w / 2 : parent.x - parent.w / 2;
    const y1 = parent.y;
    const x2 = isRight ? child.x - child.w / 2 : child.x + child.w / 2;
    const y2 = child.y;

    const mx = (x1 + x2) / 2;
    const path = this._el('path', {
      d: `M ${x1} ${y1} C ${mx} ${y1}, ${mx} ${y2}, ${x2} ${y2}`,
      fill: 'none',
      stroke: arrowColor || '#95a5a6',
      'stroke-width': 2,
      'marker-end': 'url(#arrow-end)',
      class: 'mm-edge'
    });
    this._edgesG.appendChild(path);
  }

  _renderCollapseBtn(parent, child) {
    if (!child.children.length) return;

    const isRight = child.direction === 'right' || child.x >= parent.x;
    const cx = isRight ? child.x - child.w / 2 - 14 : child.x + child.w / 2 + 14;
    const cy = child.y;

    const g = this._el('g');
    g.style.cursor = 'pointer';

    const circle = this._el('circle', {
      cx, cy, r: MM_BTN_R,
      fill: '#fff', stroke: child.borderColor || '#666', 'stroke-width': 1.5
    });
    const sign = this._el('text', {
      x: cx, y: cy,
      'text-anchor': 'middle', 'dominant-baseline': 'middle',
      'font-size': 14, 'font-weight': 'bold',
      fill: child.borderColor || '#666', 'pointer-events': 'none'
    });
    sign.textContent = child.collapsed ? '+' : '−';

    g.appendChild(circle);
    g.appendChild(sign);
    g.addEventListener('click', (e) => {
      e.stopPropagation();
      this.toggleCollapse(child.id);
    });

    this._ctrlsG.appendChild(g);
  }

  /** Re-render only a single node's visual (for selection state updates). */
  _updateNodeVisual(nodeId) {
    const existing = this._nodesG.querySelector(`[data-id="${nodeId}"]`);
    if (existing) existing.remove();
    const node = this.nodes[nodeId];
    if (node) this._renderNode(node);
  }

  // ────────────────────────────── TOOLTIP ──────────────────────────────────

  _showTooltip(e, text) {
    if (!this.tooltipEl) return;
    this.tooltipEl.textContent = text;
    this.tooltipEl.style.display = 'block';
    this._moveTooltip(e);
  }

  _moveTooltip(e) {
    if (!this.tooltipEl) return;
    const rect = this.svg.getBoundingClientRect();
    this.tooltipEl.style.left = (e.clientX - rect.left + 14) + 'px';
    this.tooltipEl.style.top  = (e.clientY - rect.top  - 10) + 'px';
  }

  _hideTooltip() {
    if (this.tooltipEl) this.tooltipEl.style.display = 'none';
  }

  // ─────────────────────────── CANVAS EVENTS ───────────────────────────────

  _attachCanvasEvents() {
    // Pan on background drag
    this.svg.addEventListener('mousedown', (e) => {
      if (e.target === this.svg || e.target === this._bgRect) {
        this._panning = true;
        this._panStart = { x: e.clientX - this._panX, y: e.clientY - this._panY };
        this.svg.style.cursor = 'grabbing';
      }
    });

    window.addEventListener('mousemove', (e) => {
      if (this._panning) {
        this._panX = e.clientX - this._panStart.x;
        this._panY = e.clientY - this._panStart.y;
        this._applyTransform();
      }
    });

    window.addEventListener('mouseup', () => {
      if (this._panning) {
        this._panning = false;
        this.svg.style.cursor = '';
      }
    });

    // Click background → deselect
    this.svg.addEventListener('click', (e) => {
      if (e.target === this.svg || e.target === this._bgRect) {
        this.deselectAll();
      }
    });

    // Zoom with wheel
    this.svg.addEventListener('wheel', (e) => {
      e.preventDefault();
      const delta = e.deltaY < 0 ? 1.1 : 0.9;
      const rect = this.svg.getBoundingClientRect();
      const mx = e.clientX - rect.left;
      const my = e.clientY - rect.top;
      this._zoom(delta, mx, my);
    }, { passive: false });

    // Touch pan (single finger)
    let lastTouch = null;
    this.svg.addEventListener('touchstart', (e) => {
      if (e.touches.length === 1) lastTouch = { x: e.touches[0].clientX, y: e.touches[0].clientY };
    });
    this.svg.addEventListener('touchmove', (e) => {
      if (e.touches.length === 1 && lastTouch) {
        const dx = e.touches[0].clientX - lastTouch.x;
        const dy = e.touches[0].clientY - lastTouch.y;
        this._panX += dx; this._panY += dy;
        lastTouch = { x: e.touches[0].clientX, y: e.touches[0].clientY };
        this._applyTransform();
        e.preventDefault();
      }
    }, { passive: false });
    this.svg.addEventListener('touchend', () => { lastTouch = null; });
  }

  // ──────────────────────────── INTERNALS ──────────────────────────────────

  _createNode(text, parentId, opts, depth) {
    const id = mmGenId();
    const colors = (typeof getThemeColors !== 'undefined')
      ? getThemeColors(this.themeId, depth)
      : { bg: '#4a90d9', text: '#ffffff', border: '#2171c2' };

    const node = {
      id,
      text: text || 'Topic',
      parent: parentId,
      children: [],
      depth,
      shape: opts.shape || 'rect',
      color: opts.color || colors.bg,
      textColor: opts.textColor || colors.text,
      borderColor: opts.borderColor || colors.border,
      fontSize: opts.fontSize || 14,
      fontFamily: opts.fontFamily || 'Segoe UI, Arial, sans-serif',
      bold: opts.bold || false,
      italic: opts.italic || false,
      manualColor: !!(opts.color),
      collapsed: false,
      note: opts.note || '',
      oneNoteLink: opts.oneNoteLink || null,
      direction: opts.direction || (parentId ? (this.nodes[parentId]?.direction || 'right') : 'right'),
      x: 0, y: 0, w: MM_NODE_MIN_W, h: MM_NODE_H
    };

    this.nodes[id] = node;
    return id;
  }

  _deleteSubtree(nodeId) {
    const node = this.nodes[nodeId];
    if (!node) return;
    [...node.children].forEach(cid => this._deleteSubtree(cid));
    delete this.nodes[nodeId];
  }

  _zoom(factor, cx, cy) {
    const newScale = Math.max(0.15, Math.min(4, this._scale * factor));
    this._panX = cx - (cx - this._panX) * (newScale / this._scale);
    this._panY = cy - (cy - this._panY) * (newScale / this._scale);
    this._scale = newScale;
    this._applyTransform();
  }

  _svgCenterX() { return (this.svg.clientWidth || 800) / 2; }
  _svgCenterY() { return (this.svg.clientHeight || 600) / 2; }

  _wrapText(text, maxWidth, fontSize) {
    const maxChars = Math.floor(maxWidth / (fontSize * 0.58));
    if (text.length <= maxChars) return [text];
    const words = text.split(' ');
    const lines = [];
    let line = '';
    words.forEach(w => {
      if ((line + ' ' + w).trim().length <= maxChars) {
        line = (line + ' ' + w).trim();
      } else {
        if (line) lines.push(line);
        line = w;
      }
    });
    if (line) lines.push(line);
    return lines.length ? lines : [text.slice(0, maxChars)];
  }

  _emit(event, ...args) {
    (this._handlers[event] || []).forEach(cb => cb(...args));
  }

  /** Create SVG element with optional attribute map. */
  _el(tag, attrs = {}) {
    const el = document.createElementNS('http://www.w3.org/2000/svg', tag);
    Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
    return el;
  }
}

if (typeof module !== 'undefined') module.exports = { MindMap };
