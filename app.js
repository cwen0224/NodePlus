const editor = document.getElementById("editor");
const nodesLayer = document.getElementById("nodes-layer");
const svg = document.getElementById("connections");
const helpFab = document.getElementById("help-fab");
const shortcutPanelEl = document.getElementById("shortcut-panel");
const shortcutCloseBtn = document.getElementById("shortcut-close-btn");
const lineDeleteBtn = document.getElementById("line-delete-btn");
const lassoBox = document.getElementById("lasso-box");
const contextMenuEl = document.getElementById("context-menu");
const navigatorEl = document.getElementById("navigator");
const navigatorMap = document.getElementById("navigator-map");
const navigatorView = document.getElementById("navigator-view");
const nodeTemplate = document.getElementById("node-template");

const ZOOM_MIN = 0.35;
const ZOOM_MAX = 2.4;
const HISTORY_LIMIT = 120;
const LINK_SNAP_RADIUS = 60;
const HANDLE_SIDES = ["top", "right", "bottom", "left"];
const NODE_DEFAULT_WIDTH = 170;
const NODE_DEFAULT_HEIGHT = 95;
const CLIPBOARD_SCHEMA = "node-editor.clipboard";
const CLIPBOARD_VERSION = 1;

const state = {
  nextNodeId: 1,
  nextConnectionId: 1,
  nodes: new Map(),
  connections: [],
  viewport: { x: 0, y: 0 },
  zoom: 1,
  draggingNode: null,
  panning: null,
  lasso: null,
  navigatorDrag: null,
  navigatorTransform: null,
  linking: null,
  selectedNodeIds: new Set(),
  clipboard: null,
  clipboardSignature: "",
  pasteSerial: 0,
  history: {
    undo: [],
    redo: [],
    lastHash: "",
  },
  draggedNodeMoved: false,
  pointerScreen: { x: 0, y: 0 },
  pointerWorld: { x: 0, y: 0 },
  hoverConnectionId: null,
  contextMenu: null,
};

bootstrap();

function bootstrap() {
  createNode(80, 80);
  createNode(360, 220);
  createNode(610, 110);

  editor.addEventListener("pointerdown", onPointerDown);
  editor.addEventListener("pointermove", onPointerMove);
  editor.addEventListener("pointerup", onPointerUp);
  editor.addEventListener("pointercancel", onPointerUp);
  editor.addEventListener("pointerleave", onPointerLeave);
  editor.addEventListener("contextmenu", onEditorContextMenu);
  editor.addEventListener("wheel", onWheel, { passive: false });
  helpFab.addEventListener("click", () => toggleShortcutPanel());
  shortcutCloseBtn.addEventListener("click", hideShortcutPanel);
  shortcutPanelEl.addEventListener("pointerdown", onShortcutPanelPointerDown);
  contextMenuEl.addEventListener("click", onContextMenuClick);
  contextMenuEl.addEventListener("pointerdown", (event) => event.stopPropagation());
  navigatorEl.addEventListener("pointerdown", onNavigatorPointerDown);
  navigatorEl.addEventListener("pointermove", onNavigatorPointerMove);
  navigatorEl.addEventListener("pointerup", onNavigatorPointerUp);
  navigatorEl.addEventListener("pointercancel", onNavigatorPointerUp);
  window.addEventListener("pointerdown", onWindowPointerDown, true);
  window.addEventListener("resize", onResize);
  window.addEventListener("keydown", onKeyDown);
  window.addEventListener("paste", onWindowPaste);

  lineDeleteBtn.addEventListener("click", () => {
    if (state.hoverConnectionId == null) {
      return;
    }
    deleteConnection(state.hoverConnectionId);
    state.hoverConnectionId = null;
    lineDeleteBtn.classList.remove("visible");
  });

  renderConnections();
  initializeHistory();
}

function createNode(x, y, options = {}) {
  const nodeId = options.id ?? String(state.nextNodeId++);
  const fragment = nodeTemplate.content.cloneNode(true);
  const nodeEl = fragment.querySelector(".node");
  const title = normalizeNodeTitle(options.title ?? `節點 ${nodeId}`, nodeId);
  const content = normalizeNodeContent(options.content);
  const titleEl = nodeEl.querySelector(".node-title");
  const contentEl = nodeEl.querySelector(".node-content");

  nodeEl.dataset.nodeId = nodeId;
  titleEl.textContent = title;
  contentEl.textContent = content;
  titleEl.setAttribute("contenteditable", "true");
  titleEl.setAttribute("spellcheck", "false");
  contentEl.setAttribute("contenteditable", "true");
  contentEl.setAttribute("spellcheck", "false");

  const nodeModel = {
    id: nodeId,
    x,
    y,
    width: NODE_DEFAULT_WIDTH,
    height: NODE_DEFAULT_HEIGHT,
    title,
    content,
    el: nodeEl,
  };
  state.nodes.set(nodeId, nodeModel);

  nodeEl.addEventListener("pointerdown", (event) => onNodePointerDown(event, nodeId));

  titleEl.addEventListener("pointerdown", (event) => {
    event.stopPropagation();
  });
  titleEl.addEventListener("focus", () => {
    titleEl.dataset.beforeEdit = nodeModel.title;
  });
  titleEl.addEventListener("input", () => {
    nodeModel.title = normalizeNodeTitle(titleEl.textContent, nodeId);
  });
  titleEl.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      titleEl.blur();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      titleEl.textContent = titleEl.dataset.beforeEdit ?? nodeModel.title;
      titleEl.blur();
    }
  });
  titleEl.addEventListener("blur", () => {
    const beforeEdit = titleEl.dataset.beforeEdit ?? nodeModel.title;
    const normalized = normalizeNodeTitle(titleEl.textContent, nodeId);
    titleEl.textContent = normalized;
    nodeModel.title = normalized;
    if (normalized !== beforeEdit) {
      commitHistory();
    }
  });

  contentEl.addEventListener("pointerdown", (event) => {
    event.stopPropagation();
  });
  contentEl.addEventListener("focus", () => {
    contentEl.dataset.beforeEdit = nodeModel.content;
  });
  contentEl.addEventListener("input", () => {
    nodeModel.content = normalizeNodeContent(contentEl.textContent);
  });
  contentEl.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      contentEl.blur();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      contentEl.textContent = contentEl.dataset.beforeEdit ?? nodeModel.content;
      contentEl.blur();
    }
  });
  contentEl.addEventListener("blur", () => {
    const beforeEdit = contentEl.dataset.beforeEdit ?? nodeModel.content;
    const normalized = normalizeNodeContent(contentEl.textContent);
    contentEl.textContent = normalized;
    nodeModel.content = normalized;
    if (normalized !== beforeEdit) {
      commitHistory();
    }
  });

  nodeEl.querySelectorAll(".handle").forEach((handle) => {
    handle.addEventListener("pointerdown", (event) => startLink(event, nodeId, handle.dataset.side));
    handle.addEventListener("pointerup", (event) => completeLink(event, nodeId, handle.dataset.side));
  });

  nodeEl.querySelector(".node-delete-btn").addEventListener("click", (event) => {
    event.stopPropagation();
    deleteNode(nodeId);
  });

  nodesLayer.appendChild(fragment);
  applyNodePosition(nodeModel);
  syncNodeSelectionClasses();
  syncNextNodeId(nodeId);
  updateNavigator();
  return nodeModel;
}

function onNodePointerDown(event, nodeId) {
  if (event.button !== 0) {
    return;
  }
  if (
    event.target.closest(".handle") ||
    event.target.closest(".node-delete-btn") ||
    event.target.closest(".node-title[contenteditable=\"true\"]") ||
    event.target.closest(".node-content[contenteditable=\"true\"]")
  ) {
    return;
  }

  event.stopPropagation();

  if (event.ctrlKey || event.metaKey) {
    toggleNodeSelection(nodeId);
    return;
  }

  if (!state.selectedNodeIds.has(nodeId)) {
    setSelection([nodeId]);
  }

  startNodeDrag(event, nodeId);
}

function startNodeDrag(event, nodeId) {
  if (!state.nodes.has(nodeId)) {
    return;
  }

  const pointerWorld = pointerToWorld(event.clientX, event.clientY);
  const dragTargets = state.selectedNodeIds.has(nodeId)
    ? Array.from(state.selectedNodeIds)
    : [nodeId];

  state.draggingNode = {
    pointerStartX: pointerWorld.x,
    pointerStartY: pointerWorld.y,
    anchors: dragTargets
      .map((id) => state.nodes.get(id))
      .filter(Boolean)
      .map((node) => ({ nodeId: node.id, x: node.x, y: node.y })),
  };
  state.draggedNodeMoved = false;

  const activeNode = state.nodes.get(nodeId);
  activeNode?.el.setPointerCapture(event.pointerId);
}

function startLink(event, nodeId, side) {
  event.stopPropagation();

  state.linking = {
    fromNodeId: nodeId,
    fromSide: side,
    previewPoint: pointerToWorld(event.clientX, event.clientY),
  };

  ensurePreviewPath();
  updatePreviewPath();
}

function completeLink(event, nodeId, side) {
  event.stopPropagation();

  finishLinkTo(nodeId, side);
}

function finishLinkTo(nodeId, side) {
  if (!state.linking) {
    return false;
  }

  const { fromNodeId, fromSide } = state.linking;
  if (fromNodeId === nodeId && fromSide === side) {
    return false;
  }

  const created = createConnection({
    fromNodeId,
    fromSide,
    toNodeId: nodeId,
    toSide: side,
  });

  state.linking = null;
  clearPreviewPath();
  renderConnections();
  if (created) {
    commitHistory();
  }
  return created;
}

function findClosestHandle(worldPoint, excludeNodeId, excludeSide) {
  const snapRadiusWorld = LINK_SNAP_RADIUS / state.zoom;
  let closest = null;

  state.nodes.forEach((node) => {
    HANDLE_SIDES.forEach((side) => {
      if (node.id === excludeNodeId && side === excludeSide) {
        return;
      }

      const handlePoint = getHandlePointWorld(node, side);
      const distance = Math.hypot(worldPoint.x - handlePoint.x, worldPoint.y - handlePoint.y);
      if (distance > snapRadiusWorld) {
        return;
      }

      if (!closest || distance < closest.distance) {
        closest = { nodeId: node.id, side, distance };
      }
    });
  });

  return closest;
}

function createConnection(link) {
  if (link.fromNodeId === link.toNodeId) {
    return false;
  }

  const duplicated = state.connections.some((connection) => {
    const sameForward =
      connection.fromNodeId === link.fromNodeId &&
      connection.fromSide === link.fromSide &&
      connection.toNodeId === link.toNodeId &&
      connection.toSide === link.toSide;

    const sameReverse =
      connection.fromNodeId === link.toNodeId &&
      connection.fromSide === link.toSide &&
      connection.toNodeId === link.fromNodeId &&
      connection.toSide === link.fromSide;

    return sameForward || sameReverse;
  });

  if (duplicated) {
    return false;
  }

  state.connections.push({
    id: String(state.nextConnectionId++),
    ...link,
  });
  return true;
}

function deleteNode(nodeId) {
  deleteNodes([nodeId]);
}

function deleteConnection(connectionId) {
  const before = state.connections.length;
  state.connections = state.connections.filter((connection) => connection.id !== connectionId);
  if (before === state.connections.length) {
    return;
  }
  renderConnections();
  commitHistory();
}

function onPointerDown(event) {
  hideContextMenu();

  if (event.button !== 0) {
    return;
  }

  if (
    event.target.closest("#navigator") ||
    event.target.closest(".node") ||
    event.target.closest(".line-delete-btn") ||
    event.target.closest(".handle")
  ) {
    return;
  }

  if (event.ctrlKey || event.metaKey) {
    startLasso(event);
    return;
  }

  if (!event.ctrlKey && !event.metaKey && !event.shiftKey) {
    clearSelection();
  }

  state.panning = {
    startClientX: event.clientX,
    startClientY: event.clientY,
    originX: state.viewport.x,
    originY: state.viewport.y,
  };

  editor.setPointerCapture(event.pointerId);
  editor.classList.add("panning");
}

function onPointerMove(event) {
  state.pointerScreen = pointerToEditor(event.clientX, event.clientY);
  state.pointerWorld = editorToWorld(state.pointerScreen.x, state.pointerScreen.y);

  if (state.lasso) {
    updateLasso(event);
    updateNodeProximity();
    updateConnectionProximity();
    return;
  }

  if (state.panning) {
    state.viewport.x = state.panning.originX + (event.clientX - state.panning.startClientX);
    state.viewport.y = state.panning.originY + (event.clientY - state.panning.startClientY);

    state.nodes.forEach((node) => applyNodePosition(node));
    renderConnections();
    state.pointerWorld = editorToWorld(state.pointerScreen.x, state.pointerScreen.y);
  }

  if (state.draggingNode) {
    const dx = state.pointerWorld.x - state.draggingNode.pointerStartX;
    const dy = state.pointerWorld.y - state.draggingNode.pointerStartY;

    state.draggingNode.anchors.forEach((anchor) => {
      const node = state.nodes.get(anchor.nodeId);
      if (!node) {
        return;
      }
      node.x = anchor.x + dx;
      node.y = anchor.y + dy;
      applyNodePosition(node);
    });

    if (Math.abs(dx) > 0.01 || Math.abs(dy) > 0.01) {
      state.draggedNodeMoved = true;
    }
    renderConnections();
  }

  if (state.linking) {
    state.linking.previewPoint = state.pointerWorld;
    updatePreviewPath();
  }

  updateNodeProximity();
  updateConnectionProximity();
}

function onEditorContextMenu(event) {
  event.preventDefault();
  const worldPoint = pointerToWorld(event.clientX, event.clientY);
  const targetNodeEl = event.target.closest(".node");
  const targetNodeId = targetNodeEl?.dataset.nodeId ?? null;

  if (targetNodeId && !state.selectedNodeIds.has(targetNodeId)) {
    setSelection([targetNodeId]);
  }

  showContextMenu(event.clientX, event.clientY, worldPoint, targetNodeId);
}

function onWindowPointerDown(event) {
  if (!state.contextMenu) {
    return;
  }
  if (event.target.closest("#context-menu")) {
    return;
  }
  hideContextMenu();
}

function onWindowPaste(event) {
  if (isTypingTarget(event.target)) {
    return;
  }
  const text = event.clipboardData?.getData("text/plain");
  if (!text) {
    return;
  }

  const payload = parseClipboardPayloadText(text);
  if (!payload) {
    return;
  }

  event.preventDefault();
  stashClipboardPayload(payload, text);
  importClipboardPayload(payload);
}

function onShortcutPanelPointerDown(event) {
  if (event.target === shortcutPanelEl) {
    hideShortcutPanel();
  }
}

function onContextMenuClick(event) {
  const button = event.target.closest("button[data-action]");
  if (!button || button.disabled || !state.contextMenu) {
    return;
  }

  const { action } = button.dataset;
  const contextNodeId = state.contextMenu.targetNodeId;
  const contextWorldPoint = state.contextMenu.worldPoint;

  const ensureTargetSelected = () => {
    if (contextNodeId && !state.selectedNodeIds.has(contextNodeId)) {
      setSelection([contextNodeId]);
      return true;
    }
    return false;
  };

  switch (action) {
    case "add-node":
      createNodeAtWorld(contextWorldPoint);
      break;
    case "copy":
      ensureTargetSelected();
      copySelectionToClipboard(false);
      break;
    case "cut":
      ensureTargetSelected();
      copySelectionToClipboard(true);
      break;
    case "paste":
      pasteClipboard(contextWorldPoint);
      break;
    case "duplicate":
      ensureTargetSelected();
      duplicateSelection(contextWorldPoint);
      break;
    case "delete":
      ensureTargetSelected();
      deleteSelectionOrHover();
      break;
    case "select-all":
      setSelection(Array.from(state.nodes.keys()));
      break;
    case "fit-view":
      fitViewToNodes();
      break;
    case "undo":
      undoHistory();
      break;
    case "redo":
      redoHistory();
      break;
    default:
      break;
  }

  hideContextMenu();
}

function showContextMenu(clientX, clientY, worldPoint, targetNodeId) {
  state.contextMenu = {
    clientX,
    clientY,
    worldPoint,
    targetNodeId,
  };

  contextMenuEl.classList.add("visible");
  updateContextMenuAvailability();
  positionContextMenu(clientX, clientY);
}

function positionContextMenu(clientX, clientY) {
  const margin = 8;
  const menuWidth = contextMenuEl.offsetWidth;
  const menuHeight = contextMenuEl.offsetHeight;
  const left = clamp(clientX, margin, window.innerWidth - menuWidth - margin);
  const top = clamp(clientY, margin, window.innerHeight - menuHeight - margin);
  contextMenuEl.style.left = `${left}px`;
  contextMenuEl.style.top = `${top}px`;
}

function updateContextMenuAvailability() {
  if (!state.contextMenu) {
    return;
  }

  const selectionSize = state.selectedNodeIds.size;
  const hasTargetNode = Boolean(state.contextMenu.targetNodeId);
  const hasNodeForAction = selectionSize > 0 || hasTargetNode;
  const hasClipboard = Boolean(state.clipboard && state.clipboard.nodes.length > 0);
  const canReadSystemClipboard = Boolean(navigator.clipboard && typeof navigator.clipboard.readText === "function");

  setContextMenuItemDisabled("copy", !hasNodeForAction);
  setContextMenuItemDisabled("cut", !hasNodeForAction);
  setContextMenuItemDisabled("duplicate", !hasNodeForAction);
  setContextMenuItemDisabled("delete", !hasNodeForAction && state.hoverConnectionId == null);
  setContextMenuItemDisabled("paste", !hasClipboard && !canReadSystemClipboard);
  setContextMenuItemDisabled("undo", state.history.undo.length <= 1);
  setContextMenuItemDisabled("redo", state.history.redo.length === 0);
}

function setContextMenuItemDisabled(action, disabled) {
  const item = contextMenuEl.querySelector(`button[data-action=\"${action}\"]`);
  if (!item) {
    return;
  }
  item.disabled = disabled;
}

function hideContextMenu() {
  if (!state.contextMenu) {
    return;
  }
  state.contextMenu = null;
  contextMenuEl.classList.remove("visible");
}

function toggleShortcutPanel() {
  if (shortcutPanelEl.classList.contains("visible")) {
    hideShortcutPanel();
    return;
  }
  showShortcutPanel();
}

function showShortcutPanel() {
  hideContextMenu();
  shortcutPanelEl.classList.add("visible");
  shortcutPanelEl.setAttribute("aria-hidden", "false");
}

function hideShortcutPanel() {
  shortcutPanelEl.classList.remove("visible");
  shortcutPanelEl.setAttribute("aria-hidden", "true");
}

function startLasso(event) {
  const startWorld = pointerToWorld(event.clientX, event.clientY);
  state.lasso = {
    pointerId: event.pointerId,
    startWorld,
    currentWorld: startWorld,
    baseSelection: new Set(state.selectedNodeIds),
  };

  editor.setPointerCapture(event.pointerId);
  editor.classList.add("lassoing");
  lassoBox.classList.add("visible");
  updateLassoBox(startWorld, startWorld);
}

function updateLasso(event) {
  if (!state.lasso || state.lasso.pointerId !== event.pointerId) {
    return;
  }

  const currentWorld = pointerToWorld(event.clientX, event.clientY);
  state.lasso.currentWorld = currentWorld;
  updateLassoBox(state.lasso.startWorld, currentWorld);

  const worldRect = getWorldRect(state.lasso.startWorld, currentWorld);
  const selectedIds = new Set(state.lasso.baseSelection);
  state.nodes.forEach((node) => {
    if (doesWorldRectIntersectNode(worldRect, node)) {
      selectedIds.add(node.id);
    }
  });
  setSelection(Array.from(selectedIds));
}

function finishLasso() {
  if (!state.lasso) {
    return;
  }

  const pointerId = state.lasso.pointerId;
  if (editor.hasPointerCapture(pointerId)) {
    editor.releasePointerCapture(pointerId);
  }

  state.lasso = null;
  editor.classList.remove("lassoing");
  lassoBox.classList.remove("visible");
}

function onNavigatorPointerDown(event) {
  if (event.button !== 0) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  state.navigatorDrag = { pointerId: event.pointerId };
  navigatorEl.setPointerCapture(event.pointerId);
  navigatorEl.classList.add("navigating");
  moveViewportFromNavigatorEvent(event);
}

function onNavigatorPointerMove(event) {
  if (!state.navigatorDrag || state.navigatorDrag.pointerId !== event.pointerId) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  moveViewportFromNavigatorEvent(event);
}

function onNavigatorPointerUp(event) {
  if (!state.navigatorDrag || state.navigatorDrag.pointerId !== event.pointerId) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  if (navigatorEl.hasPointerCapture(event.pointerId)) {
    navigatorEl.releasePointerCapture(event.pointerId);
  }

  state.navigatorDrag = null;
  navigatorEl.classList.remove("navigating");
}

function onWheel(event) {
  event.preventDefault();

  const anchorScreen = pointerToEditor(event.clientX, event.clientY);
  const anchorWorld = editorToWorld(anchorScreen.x, anchorScreen.y);
  const zoomFactor = Math.exp(-event.deltaY * 0.0015);
  const nextZoom = clamp(state.zoom * zoomFactor, ZOOM_MIN, ZOOM_MAX);
  applyZoomAt(nextZoom, anchorScreen, anchorWorld);
  state.pointerScreen = anchorScreen;
  state.pointerWorld = anchorWorld;
  updateConnectionProximity();
}

function onResize() {
  hideContextMenu();
  state.nodes.forEach((node) => applyNodePosition(node));
  renderConnections();

  if (state.linking) {
    updatePreviewPath();
  }
}

function onKeyDown(event) {
  if (isTypingTarget(event.target)) {
    return;
  }

  if (event.key === "Escape" && shortcutPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideShortcutPanel();
    return;
  }

  if (event.key === "?" && !event.ctrlKey && !event.metaKey) {
    event.preventDefault();
    toggleShortcutPanel();
    return;
  }

  if (shortcutPanelEl.classList.contains("visible")) {
    return;
  }

  if (event.key === "Escape" && state.contextMenu) {
    event.preventDefault();
    hideContextMenu();
    return;
  }

  const mod = event.ctrlKey || event.metaKey;
  const key = event.key.toLowerCase();

  if (mod && key === "n") {
    event.preventDefault();
    createNodeAtViewportCenter();
    return;
  }

  if (mod && key === "a") {
    event.preventDefault();
    setSelection(Array.from(state.nodes.keys()));
    return;
  }

  if (mod && key === "c") {
    event.preventDefault();
    copySelectionToClipboard(false);
    return;
  }

  if (mod && key === "x") {
    event.preventDefault();
    copySelectionToClipboard(true);
    return;
  }

  if (mod && key === "v") {
    const canReadSystemClipboard = Boolean(navigator.clipboard && typeof navigator.clipboard.readText === "function");
    if (!canReadSystemClipboard && state.clipboard) {
      event.preventDefault();
      importClipboardPayload(state.clipboard);
      return;
    }
  }

  if (mod && key === "d") {
    event.preventDefault();
    duplicateSelection();
    return;
  }

  if (mod && key === "z" && event.shiftKey) {
    event.preventDefault();
    redoHistory();
    return;
  }

  if (mod && key === "z") {
    event.preventDefault();
    undoHistory();
    return;
  }

  if (mod && key === "y") {
    event.preventDefault();
    redoHistory();
    return;
  }

  if (mod && (key === "=" || key === "+")) {
    event.preventDefault();
    zoomByFactor(1.12);
    return;
  }

  if (mod && key === "-") {
    event.preventDefault();
    zoomByFactor(1 / 1.12);
    return;
  }

  if (mod && key === "0") {
    event.preventDefault();
    resetZoom();
    return;
  }

  if (!mod && key === "f") {
    event.preventDefault();
    fitViewToNodes();
    return;
  }

  if (!mod && (event.key === "Delete" || event.key === "Backspace")) {
    event.preventDefault();
    deleteSelectionOrHover();
    return;
  }

  if (!mod && event.key === "Escape") {
    event.preventDefault();
    cancelCurrentInteraction();
    return;
  }

  if (!mod && event.key.startsWith("Arrow")) {
    event.preventDefault();
    const step = event.shiftKey ? 30 : 10;
    switch (event.key) {
      case "ArrowUp":
        moveSelectionBy(0, -step);
        break;
      case "ArrowDown":
        moveSelectionBy(0, step);
        break;
      case "ArrowLeft":
        moveSelectionBy(-step, 0);
        break;
      case "ArrowRight":
        moveSelectionBy(step, 0);
        break;
      default:
        break;
    }
  }
}

function onPointerUp(event) {
  const moved = state.draggedNodeMoved;
  if (state.lasso) {
    finishLasso();
  }
  state.draggingNode = null;
  state.draggedNodeMoved = false;
  state.panning = null;
  editor.classList.remove("panning");

  if (state.linking) {
    const pointerWorld = event ? pointerToWorld(event.clientX, event.clientY) : state.pointerWorld;
    const target = findClosestHandle(pointerWorld, state.linking.fromNodeId, state.linking.fromSide);
    if (target) {
      finishLinkTo(target.nodeId, target.side);
    } else {
      state.linking = null;
      clearPreviewPath();
    }
  }

  if (moved) {
    commitHistory();
  }
}

function onPointerLeave() {
  if (state.draggingNode || state.panning || state.lasso) {
    return;
  }

  if (state.linking) {
    state.linking = null;
    clearPreviewPath();
  }

  state.hoverConnectionId = null;
  lineDeleteBtn.classList.remove("visible");
  state.nodes.forEach((node) => node.el.classList.remove("near"));
}

function renderConnections() {
  const preview = svg.querySelector(".connection.preview");
  svg.innerHTML = "";

  state.connections.forEach((connection) => {
    const fromNode = state.nodes.get(connection.fromNodeId);
    const toNode = state.nodes.get(connection.toNodeId);

    if (!fromNode || !toNode) {
      return;
    }

    const fromPoint = worldToEditor(getHandlePointWorld(fromNode, connection.fromSide));
    const toPoint = worldToEditor(getHandlePointWorld(toNode, connection.toSide));

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.classList.add("connection");
    path.dataset.connectionId = connection.id;
    path.setAttribute("d", buildPath(fromPoint, toPoint));
    svg.appendChild(path);
  });

  if (preview) {
    svg.appendChild(preview);
  }

  updateConnectionProximity();
  updateNavigator();
}

function ensurePreviewPath() {
  if (svg.querySelector(".connection.preview")) {
    return;
  }

  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.classList.add("connection", "preview");
  svg.appendChild(path);
}

function clearPreviewPath() {
  const preview = svg.querySelector(".connection.preview");
  preview?.remove();
}

function updatePreviewPath() {
  if (!state.linking) {
    return;
  }

  const sourceNode = state.nodes.get(state.linking.fromNodeId);
  if (!sourceNode) {
    return;
  }

  const fromPoint = worldToEditor(getHandlePointWorld(sourceNode, state.linking.fromSide));
  const toPoint = worldToEditor(state.linking.previewPoint);
  const preview = svg.querySelector(".connection.preview");

  if (!preview) {
    return;
  }

  preview.setAttribute("d", buildPath(fromPoint, toPoint));
}

function updateNodeProximity() {
  const threshold = 22;

  state.nodes.forEach((node) => {
    const dx = distanceToRange(state.pointerWorld.x, node.x, node.x + node.el.offsetWidth);
    const dy = distanceToRange(state.pointerWorld.y, node.y, node.y + node.el.offsetHeight);
    const distance = Math.hypot(dx, dy);

    node.el.classList.toggle("near", distance <= threshold);
  });
}

function updateConnectionProximity() {
  const threshold = 14;
  let nearest = null;

  state.connections.forEach((connection) => {
    const fromNode = state.nodes.get(connection.fromNodeId);
    const toNode = state.nodes.get(connection.toNodeId);

    if (!fromNode || !toNode) {
      return;
    }

    const from = getHandlePointWorld(fromNode, connection.fromSide);
    const to = getHandlePointWorld(toNode, connection.toSide);
    const distance = pointToSegmentDistance(state.pointerWorld, from, to);

    if (!nearest || distance < nearest.distance) {
      nearest = {
        connectionId: connection.id,
        distance,
        midpoint: {
          x: (from.x + to.x) / 2,
          y: (from.y + to.y) / 2,
        },
      };
    }
  });

  if (!nearest || nearest.distance > threshold) {
    state.hoverConnectionId = null;
    lineDeleteBtn.classList.remove("visible");
    return;
  }

  state.hoverConnectionId = nearest.connectionId;
  const midpoint = worldToEditor(nearest.midpoint);
  lineDeleteBtn.style.left = `${midpoint.x}px`;
  lineDeleteBtn.style.top = `${midpoint.y}px`;
  lineDeleteBtn.classList.add("visible");
}

function getHandlePointWorld(node, side) {
  const width = node.el.offsetWidth;
  const height = node.el.offsetHeight;

  switch (side) {
    case "top":
      return { x: node.x + width / 2, y: node.y };
    case "right":
      return { x: node.x + width, y: node.y + height / 2 };
    case "bottom":
      return { x: node.x + width / 2, y: node.y + height };
    case "left":
      return { x: node.x, y: node.y + height / 2 };
    default:
      return { x: node.x + width / 2, y: node.y + height / 2 };
  }
}

function buildPath(start, end) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const c1 = { x: start.x + dx * 0.35, y: start.y + dy * 0.05 };
  const c2 = { x: start.x + dx * 0.65, y: start.y + dy * 0.95 };
  return `M ${start.x} ${start.y} C ${c1.x} ${c1.y}, ${c2.x} ${c2.y}, ${end.x} ${end.y}`;
}

function pointToSegmentDistance(point, a, b) {
  const vx = b.x - a.x;
  const vy = b.y - a.y;

  if (vx === 0 && vy === 0) {
    return Math.hypot(point.x - a.x, point.y - a.y);
  }

  const wx = point.x - a.x;
  const wy = point.y - a.y;
  const t = clamp((wx * vx + wy * vy) / (vx * vx + vy * vy), 0, 1);
  const projX = a.x + t * vx;
  const projY = a.y + t * vy;
  return Math.hypot(point.x - projX, point.y - projY);
}

function pointerToEditor(clientX, clientY) {
  const rect = editor.getBoundingClientRect();
  return {
    x: clientX - rect.left,
    y: clientY - rect.top,
  };
}

function pointerToWorld(clientX, clientY) {
  const pointer = pointerToEditor(clientX, clientY);
  return editorToWorld(pointer.x, pointer.y);
}

function editorToWorld(x, y) {
  return {
    x: (x - state.viewport.x) / state.zoom,
    y: (y - state.viewport.y) / state.zoom,
  };
}

function worldToEditor(point) {
  return {
    x: point.x * state.zoom + state.viewport.x,
    y: point.y * state.zoom + state.viewport.y,
  };
}

function updateLassoBox(aWorld, bWorld) {
  const topLeft = worldToEditor({
    x: Math.min(aWorld.x, bWorld.x),
    y: Math.min(aWorld.y, bWorld.y),
  });
  const bottomRight = worldToEditor({
    x: Math.max(aWorld.x, bWorld.x),
    y: Math.max(aWorld.y, bWorld.y),
  });

  lassoBox.style.left = `${topLeft.x}px`;
  lassoBox.style.top = `${topLeft.y}px`;
  lassoBox.style.width = `${Math.max(1, bottomRight.x - topLeft.x)}px`;
  lassoBox.style.height = `${Math.max(1, bottomRight.y - topLeft.y)}px`;
}

function getWorldRect(a, b) {
  return {
    minX: Math.min(a.x, b.x),
    minY: Math.min(a.y, b.y),
    maxX: Math.max(a.x, b.x),
    maxY: Math.max(a.y, b.y),
  };
}

function doesWorldRectIntersectNode(rect, node) {
  const nodeRect = {
    minX: node.x,
    minY: node.y,
    maxX: node.x + node.el.offsetWidth,
    maxY: node.y + node.el.offsetHeight,
  };

  return !(
    rect.maxX < nodeRect.minX ||
    rect.minX > nodeRect.maxX ||
    rect.maxY < nodeRect.minY ||
    rect.minY > nodeRect.maxY
  );
}

function applyNodePosition(node) {
  const point = worldToEditor(node);
  node.el.style.left = `${point.x}px`;
  node.el.style.top = `${point.y}px`;
  node.el.style.transform = `scale(${state.zoom})`;
  node.el.style.transformOrigin = "top left";
}

function moveViewportFromNavigatorEvent(event) {
  if (!state.navigatorTransform) {
    return;
  }

  const rect = navigatorEl.getBoundingClientRect();
  const mapX = clamp(event.clientX - rect.left, 0, rect.width);
  const mapY = clamp(event.clientY - rect.top, 0, rect.height);
  const worldPoint = navigatorToWorld(mapX, mapY);

  state.viewport.x = editor.clientWidth * 0.5 - worldPoint.x * state.zoom;
  state.viewport.y = editor.clientHeight * 0.5 - worldPoint.y * state.zoom;
  state.pointerWorld = worldPoint;
  state.pointerScreen = worldToEditor(worldPoint);

  state.nodes.forEach((node) => applyNodePosition(node));
  renderConnections();

  if (state.linking) {
    updatePreviewPath();
  }

  updateNodeProximity();
}

function updateNavigator() {
  const mapWidth = navigatorEl.clientWidth;
  const mapHeight = navigatorEl.clientHeight;

  if (mapWidth <= 0 || mapHeight <= 0) {
    return;
  }

  const bounds = getNavigatorWorldBounds();
  const contentWidth = Math.max(1, bounds.maxX - bounds.minX);
  const contentHeight = Math.max(1, bounds.maxY - bounds.minY);
  const scale = Math.min(mapWidth / contentWidth, mapHeight / contentHeight);
  const offsetX = (mapWidth - contentWidth * scale) * 0.5;
  const offsetY = (mapHeight - contentHeight * scale) * 0.5;

  state.navigatorTransform = {
    minX: bounds.minX,
    minY: bounds.minY,
    scale,
    offsetX,
    offsetY,
  };

  navigatorMap.setAttribute("viewBox", `0 0 ${mapWidth} ${mapHeight}`);
  navigatorMap.innerHTML = "";

  const worldTopLeft = editorToWorld(0, 0);
  const worldBottomRight = editorToWorld(editor.clientWidth, editor.clientHeight);
  const viewWorld = {
    minX: Math.min(worldTopLeft.x, worldBottomRight.x),
    minY: Math.min(worldTopLeft.y, worldBottomRight.y),
    maxX: Math.max(worldTopLeft.x, worldBottomRight.x),
    maxY: Math.max(worldTopLeft.y, worldBottomRight.y),
  };

  state.connections.forEach((connection) => {
    const fromNode = state.nodes.get(connection.fromNodeId);
    const toNode = state.nodes.get(connection.toNodeId);
    if (!fromNode || !toNode) {
      return;
    }

    const from = getHandlePointWorld(fromNode, connection.fromSide);
    const to = getHandlePointWorld(toNode, connection.toSide);
    const fromMap = worldToNavigator(from);
    const toMap = worldToNavigator(to);

    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.classList.add("navigator-connection");
    line.setAttribute("x1", `${fromMap.x}`);
    line.setAttribute("y1", `${fromMap.y}`);
    line.setAttribute("x2", `${toMap.x}`);
    line.setAttribute("y2", `${toMap.y}`);
    navigatorMap.appendChild(line);
  });

  state.nodes.forEach((node) => {
    const topLeft = worldToNavigator({ x: node.x, y: node.y });
    const nodeRight = node.x + node.el.offsetWidth;
    const nodeBottom = node.y + node.el.offsetHeight;
    const inView =
      nodeRight >= viewWorld.minX &&
      node.x <= viewWorld.maxX &&
      nodeBottom >= viewWorld.minY &&
      node.y <= viewWorld.maxY;
    const size = {
      w: Math.max(8, node.el.offsetWidth * state.navigatorTransform.scale),
      h: Math.max(6, node.el.offsetHeight * state.navigatorTransform.scale),
    };

    const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.classList.add("navigator-node");
    if (inView) {
      rect.classList.add("in-view");
    }
    rect.setAttribute("x", `${topLeft.x}`);
    rect.setAttribute("y", `${topLeft.y}`);
    rect.setAttribute("width", `${size.w}`);
    rect.setAttribute("height", `${size.h}`);
    rect.setAttribute("rx", "2");
    navigatorMap.appendChild(rect);
  });

  const viewTopLeft = worldToNavigator(worldTopLeft);
  const viewBottomRight = worldToNavigator(worldBottomRight);
  const viewLeft = Math.min(viewTopLeft.x, viewBottomRight.x);
  const viewTop = Math.min(viewTopLeft.y, viewBottomRight.y);
  const viewWidth = Math.max(8, Math.abs(viewBottomRight.x - viewTopLeft.x));
  const viewHeight = Math.max(8, Math.abs(viewBottomRight.y - viewTopLeft.y));

  navigatorView.style.left = `${viewLeft}px`;
  navigatorView.style.top = `${viewTop}px`;
  navigatorView.style.width = `${viewWidth}px`;
  navigatorView.style.height = `${viewHeight}px`;
}

function getNavigatorWorldBounds() {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  const viewTopLeft = editorToWorld(0, 0);
  const viewBottomRight = editorToWorld(editor.clientWidth, editor.clientHeight);
  minX = Math.min(minX, viewTopLeft.x, viewBottomRight.x);
  minY = Math.min(minY, viewTopLeft.y, viewBottomRight.y);
  maxX = Math.max(maxX, viewTopLeft.x, viewBottomRight.x);
  maxY = Math.max(maxY, viewTopLeft.y, viewBottomRight.y);

  state.nodes.forEach((node) => {
    minX = Math.min(minX, node.x);
    minY = Math.min(minY, node.y);
    maxX = Math.max(maxX, node.x + node.el.offsetWidth);
    maxY = Math.max(maxY, node.y + node.el.offsetHeight);
  });

  if (!Number.isFinite(minX)) {
    minX = -300;
    minY = -300;
    maxX = 300;
    maxY = 300;
  }

  const padding = 120;
  return {
    minX: minX - padding,
    minY: minY - padding,
    maxX: maxX + padding,
    maxY: maxY + padding,
  };
}

function worldToNavigator(point) {
  if (!state.navigatorTransform) {
    return { x: 0, y: 0 };
  }

  return {
    x: state.navigatorTransform.offsetX + (point.x - state.navigatorTransform.minX) * state.navigatorTransform.scale,
    y: state.navigatorTransform.offsetY + (point.y - state.navigatorTransform.minY) * state.navigatorTransform.scale,
  };
}

function navigatorToWorld(x, y) {
  if (!state.navigatorTransform) {
    return { x: 0, y: 0 };
  }

  return {
    x: state.navigatorTransform.minX + (x - state.navigatorTransform.offsetX) / state.navigatorTransform.scale,
    y: state.navigatorTransform.minY + (y - state.navigatorTransform.offsetY) / state.navigatorTransform.scale,
  };
}

function isTypingTarget(target) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  const tag = target.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || target.isContentEditable;
}

function normalizeNodeTitle(value, nodeId) {
  const singleLine = String(value ?? "").replace(/\r?\n/g, " ");
  const normalized = singleLine.trim();
  return normalized || `節點 ${nodeId}`;
}

function normalizeNodeContent(value) {
  const singleLine = String(value ?? "").replace(/\r?\n/g, " ");
  const normalized = singleLine.trim();
  return normalized || "拖曳節點或拉線連接";
}

function createNodeAtViewportCenter() {
  const center = getViewportCenterWorld();
  createNodeAtWorld(center);
}

function createNodeAtWorld(worldPoint) {
  const node = createNode(worldPoint.x - NODE_DEFAULT_WIDTH / 2, worldPoint.y - NODE_DEFAULT_HEIGHT / 2);
  setSelection([node.id]);
  commitHistory();
  return node;
}

function getViewportCenterWorld() {
  return editorToWorld(editor.clientWidth * 0.5, editor.clientHeight * 0.5);
}

function setSelection(nodeIds) {
  state.selectedNodeIds.clear();
  nodeIds.forEach((id) => {
    if (state.nodes.has(id)) {
      state.selectedNodeIds.add(id);
    }
  });
  syncNodeSelectionClasses();
}

function clearSelection() {
  if (state.selectedNodeIds.size === 0) {
    return;
  }
  state.selectedNodeIds.clear();
  syncNodeSelectionClasses();
}

function toggleNodeSelection(nodeId) {
  if (state.selectedNodeIds.has(nodeId)) {
    state.selectedNodeIds.delete(nodeId);
  } else if (state.nodes.has(nodeId)) {
    state.selectedNodeIds.add(nodeId);
  }
  syncNodeSelectionClasses();
}

function syncNodeSelectionClasses() {
  state.nodes.forEach((node) => {
    node.el.classList.toggle("selected", state.selectedNodeIds.has(node.id));
  });
}

function deleteNodes(nodeIds, options = {}) {
  const shouldRecord = options.recordHistory !== false;
  const uniqueIds = Array.from(new Set(nodeIds)).filter((id) => state.nodes.has(id));
  if (uniqueIds.length === 0) {
    return false;
  }

  const removedSet = new Set(uniqueIds);
  uniqueIds.forEach((id) => {
    const node = state.nodes.get(id);
    node?.el.remove();
    state.nodes.delete(id);
    state.selectedNodeIds.delete(id);
  });

  state.connections = state.connections.filter((connection) => {
    return !removedSet.has(connection.fromNodeId) && !removedSet.has(connection.toNodeId);
  });

  state.hoverConnectionId = null;
  lineDeleteBtn.classList.remove("visible");
  syncNodeSelectionClasses();
  renderConnections();
  if (shouldRecord) {
    commitHistory();
  }
  return true;
}

function deleteSelectionOrHover() {
  if (state.selectedNodeIds.size > 0) {
    return deleteNodes(Array.from(state.selectedNodeIds));
  }
  if (state.hoverConnectionId != null) {
    deleteConnection(state.hoverConnectionId);
    return true;
  }
  return false;
}

function cancelCurrentInteraction() {
  if (state.linking) {
    state.linking = null;
    clearPreviewPath();
  }
  if (state.lasso) {
    finishLasso();
  }
  state.draggingNode = null;
  state.panning = null;
  editor.classList.remove("panning");
  clearSelection();
}

function moveSelectionBy(dx, dy) {
  if (state.selectedNodeIds.size === 0) {
    return false;
  }

  state.selectedNodeIds.forEach((id) => {
    const node = state.nodes.get(id);
    if (!node) {
      return;
    }
    node.x += dx;
    node.y += dy;
    applyNodePosition(node);
  });

  renderConnections();
  commitHistory();
  return true;
}

function copySelectionToClipboard(cutSelection, options = {}) {
  const shouldWriteSystem = options.writeSystem !== false;
  const selected = Array.from(state.selectedNodeIds).filter((id) => state.nodes.has(id));
  if (selected.length === 0) {
    return false;
  }

  const ordered = selected.sort((a, b) => Number(a) - Number(b));
  const payload = buildClipboardPayloadFromNodeIds(ordered);
  const payloadText = JSON.stringify(payload);
  stashClipboardPayload(payload, payloadText);
  if (shouldWriteSystem) {
    writeClipboardPayloadToSystem(payloadText);
  }

  if (cutSelection) {
    deleteNodes(ordered);
  }
  return true;
}

async function pasteClipboard(anchorWorld = null) {
  const systemPayload = await readClipboardPayloadFromSystem();
  if (systemPayload) {
    stashClipboardPayload(systemPayload.payload, systemPayload.text);
    return importClipboardPayload(systemPayload.payload, anchorWorld);
  }

  if (!state.clipboard) {
    return false;
  }
  return importClipboardPayload(state.clipboard, anchorWorld);
}

function importClipboardPayload(payload, anchorWorld = null) {
  const normalized = normalizeClipboardPayload(payload);
  if (!normalized || normalized.nodes.length === 0) {
    return false;
  }

  const pointerInsideEditor =
    state.pointerScreen.x >= 0 &&
    state.pointerScreen.y >= 0 &&
    state.pointerScreen.x <= editor.clientWidth &&
    state.pointerScreen.y <= editor.clientHeight;
  const anchor = anchorWorld ?? (pointerInsideEditor ? state.pointerWorld : getViewportCenterWorld());
  const offsetUnit = 24;
  const offset = offsetUnit * (state.pasteSerial + 1);
  const originX = anchor.x - normalized.width * 0.5 + offset;
  const originY = anchor.y - normalized.height * 0.5 + offset;
  const idMap = new Map();
  const newNodeIds = [];

  normalized.nodes.forEach((nodeData) => {
    const newNode = createNode(originX + nodeData.offsetX, originY + nodeData.offsetY, {
      title: nodeData.title,
      content: nodeData.content,
    });
    idMap.set(nodeData.sourceId, newNode.id);
    newNodeIds.push(newNode.id);
  });

  normalized.connections.forEach((connection) => {
    const fromNodeId = idMap.get(connection.fromSourceId);
    const toNodeId = idMap.get(connection.toSourceId);
    if (!fromNodeId || !toNodeId) {
      return;
    }
    createConnection({
      fromNodeId,
      fromSide: connection.fromSide,
      toNodeId,
      toSide: connection.toSide,
    });
  });

  state.pasteSerial += 1;
  setSelection(newNodeIds);
  renderConnections();
  commitHistory();
  return true;
}

function duplicateSelection(anchorWorld = null) {
  const copied = copySelectionToClipboard(false, { writeSystem: false });
  if (!copied) {
    return false;
  }
  if (!state.clipboard) {
    return false;
  }
  return importClipboardPayload(state.clipboard, anchorWorld);
}

function buildClipboardPayloadFromNodeIds(nodeIds) {
  const idSet = new Set(nodeIds);
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  const nodes = nodeIds.map((id) => {
    const node = state.nodes.get(id);
    minX = Math.min(minX, node.x);
    minY = Math.min(minY, node.y);
    maxX = Math.max(maxX, node.x + node.el.offsetWidth);
    maxY = Math.max(maxY, node.y + node.el.offsetHeight);
    return {
      id: String(id),
      offsetX: node.x - minX,
      offsetY: node.y - minY,
      title: node.title,
      content: node.content,
    };
  });

  // offsets depend on final minX/minY, normalize after bounds are finalized.
  nodes.forEach((node) => {
    const source = state.nodes.get(node.id);
    node.offsetX = source.x - minX;
    node.offsetY = source.y - minY;
  });

  const connections = state.connections
    .filter((connection) => idSet.has(connection.fromNodeId) && idSet.has(connection.toNodeId))
    .map((connection) => ({
      fromId: String(connection.fromNodeId),
      fromSide: connection.fromSide,
      toId: String(connection.toNodeId),
      toSide: connection.toSide,
    }));

  return {
    type: CLIPBOARD_SCHEMA,
    version: CLIPBOARD_VERSION,
    nodes,
    connections,
    width: Math.max(1, maxX - minX),
    height: Math.max(1, maxY - minY),
  };
}

function parseClipboardPayloadText(text) {
  if (typeof text !== "string" || text.trim() === "") {
    return null;
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    return null;
  }

  if (Array.isArray(parsed)) {
    parsed = { nodes: parsed };
  }

  const normalized = normalizeClipboardPayload(parsed);
  if (!normalized) {
    return null;
  }

  return {
    type: CLIPBOARD_SCHEMA,
    version: CLIPBOARD_VERSION,
    nodes: normalized.nodes.map((node) => ({
      id: node.sourceId,
      offsetX: node.offsetX,
      offsetY: node.offsetY,
      title: node.title,
      content: node.content,
    })),
    connections: normalized.connections.map((connection) => ({
      fromId: connection.fromSourceId,
      fromSide: connection.fromSide,
      toId: connection.toSourceId,
      toSide: connection.toSide,
    })),
    width: normalized.width,
    height: normalized.height,
  };
}

function normalizeClipboardPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const rawNodes = Array.isArray(payload.nodes) ? payload.nodes : [];
  if (rawNodes.length === 0) {
    return null;
  }

  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  const prelimNodes = [];

  rawNodes.forEach((rawNode, index) => {
    if (!rawNode || typeof rawNode !== "object") {
      return;
    }
    const sourceId = String(rawNode.id ?? rawNode.localId ?? rawNode.sourceId ?? index);
    const title = typeof rawNode.title === "string" ? rawNode.title : `節點 ${sourceId}`;
    const content = typeof rawNode.content === "string" ? rawNode.content : "拖曳節點或拉線連接";
    const hasOffset = Number.isFinite(rawNode.offsetX) && Number.isFinite(rawNode.offsetY);
    const hasPosition = Number.isFinite(rawNode.x) && Number.isFinite(rawNode.y);

    if (!hasOffset && !hasPosition) {
      return;
    }

    const node = {
      sourceId,
      title,
      content,
      offsetX: hasOffset ? rawNode.offsetX : null,
      offsetY: hasOffset ? rawNode.offsetY : null,
      x: hasPosition ? rawNode.x : null,
      y: hasPosition ? rawNode.y : null,
    };
    prelimNodes.push(node);

    if (!hasOffset) {
      minX = Math.min(minX, rawNode.x);
      minY = Math.min(minY, rawNode.y);
    }
  });

  if (prelimNodes.length === 0) {
    return null;
  }

  if (!Number.isFinite(minX)) {
    minX = 0;
    minY = 0;
  }

  const nodes = prelimNodes.map((node) => ({
    sourceId: node.sourceId,
    title: node.title,
    content: node.content,
    offsetX: node.offsetX ?? node.x - minX,
    offsetY: node.offsetY ?? node.y - minY,
  }));

  const nodeIdSet = new Set(nodes.map((node) => node.sourceId));
  const rawConnections = Array.isArray(payload.connections) ? payload.connections : [];
  const connections = rawConnections
    .map((rawConnection) => {
      if (!rawConnection || typeof rawConnection !== "object") {
        return null;
      }
      const fromSourceId = String(
        rawConnection.fromId ?? rawConnection.fromLocalId ?? rawConnection.fromNodeId ?? rawConnection.from
      );
      const toSourceId = String(rawConnection.toId ?? rawConnection.toLocalId ?? rawConnection.toNodeId ?? rawConnection.to);
      const fromSide = rawConnection.fromSide;
      const toSide = rawConnection.toSide;
      if (
        !nodeIdSet.has(fromSourceId) ||
        !nodeIdSet.has(toSourceId) ||
        !HANDLE_SIDES.includes(fromSide) ||
        !HANDLE_SIDES.includes(toSide)
      ) {
        return null;
      }
      return { fromSourceId, fromSide, toSourceId, toSide };
    })
    .filter(Boolean);

  let computedWidth = Number(payload.width);
  let computedHeight = Number(payload.height);
  if (!(Number.isFinite(computedWidth) && computedWidth > 0)) {
    const maxOffsetX = nodes.reduce((max, node) => Math.max(max, node.offsetX), 0);
    computedWidth = maxOffsetX + NODE_DEFAULT_WIDTH;
  }
  if (!(Number.isFinite(computedHeight) && computedHeight > 0)) {
    const maxOffsetY = nodes.reduce((max, node) => Math.max(max, node.offsetY), 0);
    computedHeight = maxOffsetY + NODE_DEFAULT_HEIGHT;
  }

  return {
    nodes,
    connections,
    width: Math.max(1, computedWidth),
    height: Math.max(1, computedHeight),
  };
}

function stashClipboardPayload(payload, signatureText = null) {
  if (!payload) {
    return;
  }
  const signature = signatureText ?? JSON.stringify(payload);
  if (signature !== state.clipboardSignature) {
    state.pasteSerial = 0;
  }
  state.clipboard = payload;
  state.clipboardSignature = signature;
}

function writeClipboardPayloadToSystem(payloadText) {
  if (!navigator.clipboard || typeof navigator.clipboard.writeText !== "function") {
    return;
  }
  navigator.clipboard.writeText(payloadText).catch(() => {});
}

async function readClipboardPayloadFromSystem() {
  if (!navigator.clipboard || typeof navigator.clipboard.readText !== "function") {
    return null;
  }

  try {
    const text = await navigator.clipboard.readText();
    const payload = parseClipboardPayloadText(text);
    if (!payload) {
      return null;
    }
    return { payload, text };
  } catch {
    return null;
  }
}

function zoomByFactor(factor) {
  const centerScreen = { x: editor.clientWidth * 0.5, y: editor.clientHeight * 0.5 };
  const centerWorld = editorToWorld(centerScreen.x, centerScreen.y);
  const nextZoom = clamp(state.zoom * factor, ZOOM_MIN, ZOOM_MAX);
  applyZoomAt(nextZoom, centerScreen, centerWorld);
}

function resetZoom() {
  const centerScreen = { x: editor.clientWidth * 0.5, y: editor.clientHeight * 0.5 };
  const centerWorld = editorToWorld(centerScreen.x, centerScreen.y);
  applyZoomAt(1, centerScreen, centerWorld);
}

function applyZoomAt(nextZoom, anchorScreen, anchorWorld) {
  if (nextZoom === state.zoom) {
    return;
  }
  state.zoom = nextZoom;
  state.viewport.x = anchorScreen.x - anchorWorld.x * state.zoom;
  state.viewport.y = anchorScreen.y - anchorWorld.y * state.zoom;

  state.nodes.forEach((node) => applyNodePosition(node));
  renderConnections();
  if (state.linking) {
    updatePreviewPath();
  }
  updateNodeProximity();
}

function fitViewToNodes() {
  const targets =
    state.selectedNodeIds.size > 0
      ? Array.from(state.selectedNodeIds).map((id) => state.nodes.get(id)).filter(Boolean)
      : Array.from(state.nodes.values());
  if (targets.length === 0) {
    return;
  }

  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  targets.forEach((node) => {
    minX = Math.min(minX, node.x);
    minY = Math.min(minY, node.y);
    maxX = Math.max(maxX, node.x + node.el.offsetWidth);
    maxY = Math.max(maxY, node.y + node.el.offsetHeight);
  });

  const padding = 80;
  const width = Math.max(1, maxX - minX + padding * 2);
  const height = Math.max(1, maxY - minY + padding * 2);
  const nextZoom = clamp(Math.min(editor.clientWidth / width, editor.clientHeight / height), ZOOM_MIN, ZOOM_MAX);
  const centerWorld = { x: (minX + maxX) * 0.5, y: (minY + maxY) * 0.5 };
  const centerScreen = { x: editor.clientWidth * 0.5, y: editor.clientHeight * 0.5 };
  applyZoomAt(nextZoom, centerScreen, centerWorld);
}

function initializeHistory() {
  const snapshot = captureSnapshot();
  state.history.undo = [snapshot];
  state.history.redo = [];
  state.history.lastHash = hashSnapshot(snapshot);
}

function captureSnapshot() {
  const nodes = Array.from(state.nodes.values())
    .map((node) => ({
      id: node.id,
      x: node.x,
      y: node.y,
      title: node.title,
      content: node.content,
    }))
    .sort((a, b) => Number(a.id) - Number(b.id));

  const connections = state.connections
    .map((connection) => ({ ...connection }))
    .sort((a, b) => Number(a.id) - Number(b.id));

  const selectedNodeIds = Array.from(state.selectedNodeIds).sort((a, b) => Number(a) - Number(b));

  return {
    nextNodeId: state.nextNodeId,
    nextConnectionId: state.nextConnectionId,
    nodes,
    connections,
    selectedNodeIds,
  };
}

function hashSnapshot(snapshot) {
  return JSON.stringify(snapshot);
}

function restoreSnapshot(snapshot) {
  state.draggingNode = null;
  state.panning = null;
  state.linking = null;
  clearPreviewPath();

  state.nodes.forEach((node) => node.el.remove());
  state.nodes.clear();
  state.connections = [];
  state.selectedNodeIds.clear();

  snapshot.nodes.forEach((nodeData) => {
    createNode(nodeData.x, nodeData.y, {
      id: nodeData.id,
      title: nodeData.title,
      content: nodeData.content,
    });
  });

  state.connections = snapshot.connections.map((connection) => ({ ...connection }));
  state.nextNodeId = snapshot.nextNodeId;
  state.nextConnectionId = snapshot.nextConnectionId;
  setSelection(snapshot.selectedNodeIds);

  state.hoverConnectionId = null;
  lineDeleteBtn.classList.remove("visible");
  editor.classList.remove("panning");
  renderConnections();
  updateNodeProximity();
}

function commitHistory() {
  const snapshot = captureSnapshot();
  const hash = hashSnapshot(snapshot);
  if (hash === state.history.lastHash) {
    return;
  }
  state.history.undo.push(snapshot);
  if (state.history.undo.length > HISTORY_LIMIT) {
    state.history.undo.shift();
  }
  state.history.redo = [];
  state.history.lastHash = hash;
}

function undoHistory() {
  if (state.history.undo.length <= 1) {
    return;
  }
  const current = state.history.undo.pop();
  state.history.redo.push(current);
  const previous = state.history.undo[state.history.undo.length - 1];
  restoreSnapshot(previous);
  state.history.lastHash = hashSnapshot(previous);
}

function redoHistory() {
  if (state.history.redo.length === 0) {
    return;
  }
  const snapshot = state.history.redo.pop();
  state.history.undo.push(snapshot);
  restoreSnapshot(snapshot);
  state.history.lastHash = hashSnapshot(snapshot);
}

function syncNextNodeId(nodeId) {
  const numericId = Number(nodeId);
  if (!Number.isNaN(numericId)) {
    state.nextNodeId = Math.max(state.nextNodeId, numericId + 1);
  }
}

function distanceToRange(value, min, max) {
  if (value < min) {
    return min - value;
  }
  if (value > max) {
    return value - max;
  }
  return 0;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}
