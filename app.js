const editor = document.getElementById("editor");
const nodesLayer = document.getElementById("nodes-layer");
const svg = document.getElementById("connections");
const undoBtn = document.getElementById("undo-btn");
const redoBtn = document.getElementById("redo-btn");
const gitSyncBtn = document.getElementById("git-sync-btn");
const setMediaPathBtn = document.getElementById("set-media-path-btn");
const gitSyncStatusEl = document.getElementById("git-sync-status");
const gitSyncPanelEl = document.getElementById("git-sync-panel");
const gitSyncCloseBtn = document.getElementById("git-sync-close-btn");
const gitSyncCloseFooterBtn = document.getElementById("git-sync-close-footer-btn");
const gitSyncOwnerInput = document.getElementById("git-sync-owner-input");
const gitSyncRepoInput = document.getElementById("git-sync-repo-input");
const gitSyncBranchInput = document.getElementById("git-sync-branch-input");
const gitSyncPathInput = document.getElementById("git-sync-path-input");
const gitSyncTokenInput = document.getElementById("git-sync-token-input");
const gitSyncAutosyncInput = document.getElementById("git-sync-autosync-checkbox");
const gitSyncStartupInput = document.getElementById("git-sync-startup-checkbox");
const gitSyncMessageEl = document.getElementById("git-sync-message");
const gitSyncLoadBtn = document.getElementById("git-sync-load-btn");
const gitSyncNowBtn = document.getElementById("git-sync-now-btn");
const gitSyncSaveBtn = document.getElementById("git-sync-save-btn");
const gitSyncClearBtn = document.getElementById("git-sync-clear-btn");
const definedNodeMenuEl = document.getElementById("defined-node-menu");
const definedNodeMenuBtn = document.getElementById("defined-node-menu-btn");
const definedNodePanelEl = document.getElementById("defined-node-panel");
const definedNodeCountEl = document.getElementById("defined-node-count");
const definedNodeListEl = document.getElementById("defined-node-list");
const helpFab = document.getElementById("help-fab");
const shortcutPanelEl = document.getElementById("shortcut-panel");
const shortcutCloseBtn = document.getElementById("shortcut-close-btn");
const shortcutEditorListEl = document.getElementById("shortcut-editor-list");
const shortcutResetBtn = document.getElementById("shortcut-reset-btn");
const lineDeleteBtn = document.getElementById("line-delete-btn");
const lassoBox = document.getElementById("lasso-box");
const contextMenuEl = document.getElementById("context-menu");
const jsonEditorPanelEl = document.getElementById("json-editor-panel");
const jsonEditorCloseBtn = document.getElementById("json-editor-close-btn");
const jsonEditorTitleEl = document.getElementById("json-editor-title");
const jsonEditorTextarea = document.getElementById("json-editor-textarea");
const jsonEditorFunctionalControlsEl = document.getElementById("json-editor-functional-controls");
const jsonEditorCancelBtn = document.getElementById("json-editor-cancel-btn");
const jsonEditorSaveBtn = document.getElementById("json-editor-save-btn");
const nodeTextEditorPanelEl = document.getElementById("node-text-editor-panel");
const nodeTextEditorCloseBtn = document.getElementById("node-text-editor-close-btn");
const nodeTextEditorTitleInput = document.getElementById("node-text-editor-title-input");
const nodeTextEditorContentTextarea = document.getElementById("node-text-editor-content-textarea");
const nodeTextEditorEditJsonBtn = document.getElementById("node-text-editor-edit-json-btn");
const nodeTextEditorDefineJsonBtn = document.getElementById("node-text-editor-define-json-btn");
const nodeTextEditorCancelBtn = document.getElementById("node-text-editor-cancel-btn");
const nodeTextEditorSaveBtn = document.getElementById("node-text-editor-save-btn");
const mediaViewerPanelEl = document.getElementById("media-viewer-panel");
const mediaViewerCloseBtn = document.getElementById("media-viewer-close-btn");
const mediaViewerTitleEl = document.getElementById("media-viewer-title");
const mediaViewerMetaEl = document.getElementById("media-viewer-meta");
const mediaViewerBodyEl = document.getElementById("media-viewer-body");
const navigatorEl = document.getElementById("navigator");
const navigatorMap = document.getElementById("navigator-map");
const navigatorView = document.getElementById("navigator-view");
const nodeTemplate = document.getElementById("node-template");

const ZOOM_MIN = 0.35;
const ZOOM_MAX = 2.4;
const HISTORY_LIMIT = 120;
const LINK_SNAP_RADIUS = 60;
const HANDLE_SIDES = ["top", "right", "bottom", "left"];
const GRID_WORLD_SIZE = 28;
const NODE_DEFAULT_WIDTH = 170;
const NODE_DEFAULT_HEIGHT = 95;
const CLIPBOARD_SCHEMA = "node-editor.clipboard";
const CLIPBOARD_VERSION = 1;
const SELECTION_JSON_EDIT_SCHEMA = "node-editor.selection-json-edit";
const SELECTION_JSON_EDIT_VERSION = 1;
const MEDIA_NODE_SCHEMA = "node-editor.media-record";
const BINDING_NODE_SCHEMA = "node-editor.binding-record";
const BINDING_NODE_VERSION = 1;
const NODE_DEFINITION_VERSION = 1;
const NODE_DEFINITION_TITLE_KEY = "__nodeTitle";
const NODE_DEFINITION_DESCRIPTION_KEY = "__nodeDescription";
const MEDIA_MANIFEST_FILE = "media-manifest.json";
const MEDIA_HANDLE_DB_NAME = "node-editor-settings";
const MEDIA_HANDLE_STORE = "kv";
const PROJECT_ROOT_HANDLE_KEY = "project-root-handle";
const LEGACY_MEDIA_HANDLE_KEY = "media-directory-handle";
const PROJECT_FILE_HANDLE_KEY = "project-state-file-handle";
const PROJECT_STATE_SCHEMA = "node-editor.project";
const PROJECT_STATE_VERSION = 1;
const PROJECT_STATE_FILE = "project-state.json";
const PROJECT_DRAFT_STORAGE_KEY = "node-editor.project-draft.v1";
const GIT_SYNC_STORAGE_KEY = "node-editor.git-sync.v1";
const GIT_SYNC_DEBOUNCE_MS = 1800;
const GIT_SYNC_DEFAULT_COMMIT_MESSAGE = "chore(editor): autosave project-state.json";
const GITHUB_API_BASE = "https://api.github.com";
const GITHUB_API_VERSION = "2026-03-10";
const GITHUB_CONTENTS_ACCEPT = "application/vnd.github+json";
const MEDIA_THUMBNAIL_TOKEN = "[embedded-thumbnail-data-url]";
const THUMBNAIL_MAX_WIDTH = 560;
const THUMBNAIL_MAX_HEIGHT = 360;
const NODE_VISUAL_TYPES = ["default", "media", "image", "video", "audio", "text", "json", "binary"];
const NODE_VISUAL_CLASS_PREFIX = "node-type-";
const NAVIGATOR_NODE_VISUAL_CLASS_PREFIX = "navigator-node-type-";
const BINDING_DIRECTIVE_PATTERN = /^node_(top|right|bottom|left)(?:_(\d+))?$/i;
const BINDING_DIRECTIVE_LEGACY_PATTERN = /^n([trbl])(\d*)$/i;
const BINDING_CHECK_DIRECTIVE_PATTERN = /^check(?:_(\d+)|(\d*))$/i;
const BINDING_VALUE_DATE_FIELD_PATTERN = /^date_yyyy_mm_dd_(\d+)$/i;
const BINDING_VALUE_TEXT_FIELD_PATTERN = /^text_(\d+)$/i;
const DEFINED_NODE_TITLE = "筆記模板";
const GAMEPAD_REPEAT_INITIAL_MS = 260;
const GAMEPAD_REPEAT_INTERVAL_MS = 92;
const DEFAULT_TIMECODE_FPS = 30;
const SHORTCUT_STORAGE_KEY = "node-editor.shortcuts.v1";
const SHORTCUT_MODIFIER_ORDER = ["Mod", "Alt", "Shift"];
const SHORTCUT_MODIFIER_SET = new Set(["Mod", "Alt", "Shift"]);
const IS_MAC_PLATFORM = /mac|iphone|ipad|ipod/i.test(navigator.platform ?? navigator.userAgent ?? "");
const GENERIC_JSON_PREVIEW_MAX_CHARS = 80;

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
  nodeDefinitions: [],
  history: {
    undo: [],
    redo: [],
    lastHash: "",
  },
  mediaDragDepth: 0,
  projectRootHandle: null,
  projectStateFileHandle: null,
  mediaStoreHandle: null,
  mediaStoreMode: null,
  mediaManifest: [],
  projectSaveQueued: null,
  projectSaveInFlight: false,
  projectLastSavedHash: "",
  projectLastSavedSnapshot: null,
  gitSync: {
    settings: null,
    queuedSnapshot: null,
    queuedHash: "",
    timerId: null,
    inFlight: false,
    lastRemoteSha: "",
    lastSyncedHash: "",
    lastSyncedAt: null,
    lastMessage: "",
    lastError: "",
  },
  jsonEditorNodeId: null,
  jsonEditorOriginalThumbnailDataUrl: null,
  jsonEditorMode: null,
  jsonEditorSelectionNodeIds: [],
  jsonEditorFunctionalConfig: null,
  nodeTextEditorNodeId: null,
  mediaViewerNodeId: null,
  mediaViewerObjectUrl: null,
  draggedNodeMoved: false,
  pointerScreen: { x: 0, y: 0 },
  pointerWorld: { x: 0, y: 0 },
  hoverConnectionId: null,
  contextMenu: null,
  shortcuts: {},
  shortcutEditingActionId: null,
  gamepadRafId: null,
  gamepadButtonStates: new Map(),
  gamepadRepeatDueAt: {},
};

const SHORTCUT_ACTION_DEFS = [
  { id: "toggleShortcutPanel", label: "開關快捷鍵編輯", description: "開啟或關閉快捷鍵編輯面板", defaultCombo: "Shift+Slash" },
  { id: "newNode", label: "新增節點", description: "在視窗中心新增節點", defaultCombo: "Mod+KeyN" },
  { id: "selectAll", label: "全選節點", description: "選取目前所有節點", defaultCombo: "Mod+KeyA" },
  { id: "copy", label: "複製", description: "複製目前選取節點", defaultCombo: "Mod+KeyC" },
  { id: "cut", label: "剪下", description: "剪下目前選取節點", defaultCombo: "Mod+KeyX" },
  { id: "paste", label: "貼上", description: "貼上剪貼簿節點", defaultCombo: "Mod+KeyV" },
  { id: "duplicate", label: "複製一份", description: "複製目前選取節點到新位置", defaultCombo: "Mod+KeyD" },
  { id: "saveProject", label: "儲存筆記", description: "儲存到 project-state.json", defaultCombo: "Mod+KeyS" },
  { id: "loadProject", label: "讀取筆記", description: "讀取 project-state.json", defaultCombo: "Mod+KeyO" },
  { id: "loadLegacyProject", label: "開啟舊筆記", description: "讀取舊版 JSON 存檔", defaultCombo: "Mod+Shift+KeyO" },
  { id: "undo", label: "復原", description: "回到上一個操作", defaultCombo: "Mod+KeyZ" },
  { id: "redo", label: "重做", description: "重新套用操作", defaultCombo: "Mod+Shift+KeyZ" },
  { id: "redoAlt", label: "重做（替代）", description: "重做替代鍵", defaultCombo: "Mod+KeyY" },
  { id: "zoomIn", label: "放大", description: "放大白板", defaultCombo: "Mod+Equal" },
  { id: "zoomInAlt", label: "放大（替代）", description: "放大替代鍵", defaultCombo: "Mod+Shift+Equal" },
  { id: "zoomOut", label: "縮小", description: "縮小白板", defaultCombo: "Mod+Minus" },
  { id: "resetZoom", label: "重設縮放", description: "縮放回到 100%", defaultCombo: "Mod+Digit0" },
  { id: "fitView", label: "聚焦內容", description: "聚焦目前選取或全部內容", defaultCombo: "KeyF" },
  { id: "deleteSelection", label: "刪除選取", description: "刪除節點或 hover 連線", defaultCombo: "Delete" },
  { id: "deleteSelectionAlt", label: "刪除（替代）", description: "刪除替代鍵", defaultCombo: "Backspace" },
  { id: "cancel", label: "取消操作", description: "取消目前操作", defaultCombo: "Escape" },
  { id: "moveUp", label: "向上移動", description: "移動選取節點", defaultCombo: "ArrowUp" },
  { id: "moveDown", label: "向下移動", description: "移動選取節點", defaultCombo: "ArrowDown" },
  { id: "moveLeft", label: "向左移動", description: "移動選取節點", defaultCombo: "ArrowLeft" },
  { id: "moveRight", label: "向右移動", description: "移動選取節點", defaultCombo: "ArrowRight" },
];
const SHORTCUT_ACTION_IDS = new Set(SHORTCUT_ACTION_DEFS.map((item) => item.id));

void bootstrap();

async function bootstrap() {
  editor.addEventListener("pointerdown", onPointerDown);
  editor.addEventListener("pointermove", onPointerMove);
  editor.addEventListener("pointerup", onPointerUp);
  editor.addEventListener("pointercancel", onPointerUp);
  editor.addEventListener("pointerleave", onPointerLeave);
  editor.addEventListener("contextmenu", onEditorContextMenu);
  editor.addEventListener("wheel", onWheel, { passive: false });
  editor.addEventListener("dragenter", onEditorDragEnter);
  editor.addEventListener("dragover", onEditorDragOver);
  editor.addEventListener("dragleave", onEditorDragLeave);
  editor.addEventListener("drop", onEditorDrop);
  undoBtn?.addEventListener("click", undoHistory);
  redoBtn?.addEventListener("click", redoHistory);
  definedNodeMenuBtn?.addEventListener("click", onDefinedNodeMenuButtonClick);
  definedNodePanelEl?.addEventListener("click", onDefinedNodePanelClick);
  definedNodePanelEl?.addEventListener("pointerdown", (event) => event.stopPropagation());
  jsonEditorPanelEl.addEventListener("pointerdown", onJsonEditorPanelPointerDown);
  jsonEditorCloseBtn.addEventListener("click", hideJsonEditor);
  jsonEditorCancelBtn.addEventListener("click", hideJsonEditor);
  jsonEditorSaveBtn.addEventListener("click", onJsonEditorSaveClick);
  jsonEditorTextarea.addEventListener("keydown", onJsonEditorTextareaKeyDown);
  jsonEditorTextarea.addEventListener("input", onJsonEditorTextareaInput);
  nodeTextEditorPanelEl.addEventListener("pointerdown", onNodeTextEditorPanelPointerDown);
  nodeTextEditorCloseBtn.addEventListener("click", hideNodeTextEditor);
  nodeTextEditorCancelBtn.addEventListener("click", hideNodeTextEditor);
  nodeTextEditorEditJsonBtn?.addEventListener("click", onNodeTextEditorEditJsonClick);
  nodeTextEditorDefineJsonBtn?.addEventListener("click", onNodeTextEditorDefineJsonClick);
  nodeTextEditorSaveBtn.addEventListener("click", onNodeTextEditorSaveClick);
  nodeTextEditorTitleInput.addEventListener("input", onNodeTextEditorInput);
  nodeTextEditorContentTextarea.addEventListener("input", onNodeTextEditorInput);
  nodeTextEditorContentTextarea.addEventListener("keydown", onNodeTextEditorKeyDown);
  nodeTextEditorTitleInput.addEventListener("keydown", onNodeTextEditorKeyDown);
  gitSyncBtn?.addEventListener("click", () => toggleGitSyncPanel());
  gitSyncCloseBtn?.addEventListener("click", hideGitSyncPanel);
  gitSyncCloseFooterBtn?.addEventListener("click", hideGitSyncPanel);
  gitSyncPanelEl?.addEventListener("pointerdown", onGitSyncPanelPointerDown);
  [gitSyncOwnerInput, gitSyncRepoInput, gitSyncBranchInput, gitSyncPathInput, gitSyncTokenInput].forEach((input) =>
    input?.addEventListener("input", onGitSyncSettingsInputChange)
  );
  gitSyncAutosyncInput?.addEventListener("change", onGitSyncSettingsInputChange);
  gitSyncStartupInput?.addEventListener("change", onGitSyncSettingsInputChange);
  gitSyncLoadBtn?.addEventListener("click", onGitSyncLoadClick);
  gitSyncNowBtn?.addEventListener("click", onGitSyncNowClick);
  gitSyncSaveBtn?.addEventListener("click", onGitSyncSaveClick);
  gitSyncClearBtn?.addEventListener("click", onGitSyncClearTokenClick);
  mediaViewerPanelEl.addEventListener("pointerdown", onMediaViewerPanelPointerDown);
  mediaViewerCloseBtn.addEventListener("click", hideMediaViewer);
  helpFab.addEventListener("click", () => toggleShortcutPanel());
  shortcutCloseBtn.addEventListener("click", hideShortcutPanel);
  shortcutPanelEl.addEventListener("pointerdown", onShortcutPanelPointerDown);
  shortcutEditorListEl?.addEventListener("click", onShortcutEditorListClick);
  shortcutResetBtn?.addEventListener("click", onShortcutResetClick);
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
  window.addEventListener("gamepadconnected", onGamepadConnected);
  window.addEventListener("gamepaddisconnected", onGamepadDisconnected);

  lineDeleteBtn.addEventListener("click", () => {
    if (state.hoverConnectionId == null) {
      return;
    }
    deleteConnection(state.hoverConnectionId);
    state.hoverConnectionId = null;
    lineDeleteBtn.classList.remove("visible");
  });

  initializeShortcutBindings();
  await initializeGitSyncFromStorage();
  initializeMediaPathFromStorage();
  renderDefinedNodeTypeList();
  updateEditorGridBackground();
  const autoLoaded = await autoLoadLastProjectOnStartup();
  if (!autoLoaded) {
    createNode(80, 80, { metadata: createDefaultBindingDirectiveSeed() });
    createNode(360, 220, { metadata: createDefaultBindingDirectiveSeed() });
    createNode(610, 110, { metadata: createDefaultBindingDirectiveSeed() });
    renderConnections();
    initializeHistory();
  }
  startGamepadPolling();
}

function createNode(x, y, options = {}) {
  const nodeId = options.id ?? String(state.nextNodeId++);
  const fragment = nodeTemplate.content.cloneNode(true);
  const nodeEl = fragment.querySelector(".node");
  const metadata = normalizeNodeMetadata(options.metadata ?? null);
  const isMediaNode = isMediaMetadata(metadata);
  const isBindingNode = isBindingMetadata(metadata);
  const isGenericJsonNode = isPlainRecord(metadata) && !isMediaNode && !isBindingNode;
  const title = normalizeNodeTitle(options.title ?? `節點 ${nodeId}`, nodeId);
  const normalizedOptionContent = normalizeNodeContent(options.content);
  const shouldUseJsonSummary =
    isGenericJsonNode &&
    (title === "未定義結點" ||
      parseJsonObjectTextSafely(normalizedOptionContent) != null ||
      looksLikeVerboseJsonText(normalizedOptionContent));
  const fallbackContent = shouldUseJsonSummary
    ? normalizeNodeContent(buildJsonNodeSummaryText(metadata))
    : title === "未定義結點"
      ? buildUndefinedNodeCompactContent(metadata, options.content)
      : normalizedOptionContent;
  const content = isMediaNode
    ? buildMediaNodeContent(metadata)
    : isBindingNode
      ? buildBindingNodeContent(metadata)
      : fallbackContent;
  const titleEl = nodeEl.querySelector(".node-title");
  const contentEl = nodeEl.querySelector(".node-content");

  nodeEl.dataset.nodeId = nodeId;
  titleEl.textContent = title;
  contentEl.textContent = content;
  titleEl.removeAttribute("contenteditable");
  contentEl.removeAttribute("contenteditable");
  titleEl.setAttribute("title", "雙擊編輯文字");
  contentEl.setAttribute(
    "title",
    isMediaNode ? "雙擊編輯 JSON" : isBindingNode ? "雙擊編輯綁定 JSON" : "雙擊編輯文字"
  );

  const nodeModel = {
    id: nodeId,
    x,
    y,
    width: NODE_DEFAULT_WIDTH,
    height: NODE_DEFAULT_HEIGHT,
    title,
    content,
    metadata,
    el: nodeEl,
  };
  state.nodes.set(nodeId, nodeModel);
  syncUndefinedNodeClass(nodeModel);
  renderNodeHandles(nodeModel);
  renderNodeMediaPreview(nodeModel);
  renderBindingToggleControls(nodeModel);

  nodeEl.addEventListener("pointerdown", (event) => onNodePointerDown(event, nodeId));
  nodeEl.addEventListener("dblclick", (event) => onNodeDoubleClick(event, nodeId));

  titleEl.addEventListener("dblclick", (event) => onNodeTextDoubleClick(event, nodeId));

  contentEl.addEventListener("dblclick", (event) => onNodeContentDoubleClick(event, nodeId));

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

function syncUndefinedNodeClass(nodeModel) {
  if (!nodeModel?.el) {
    return;
  }
  const isUndefinedNode =
    nodeModel.title === "未定義結點" && !isMediaMetadata(nodeModel.metadata) && !isBindingMetadata(nodeModel.metadata);
  nodeModel.el.classList.toggle("undefined-node", isUndefinedNode);
}

function renderNodeMediaPreview(nodeModel) {
  if (!nodeModel?.el) {
    return;
  }
  const nodeEl = nodeModel.el;
  syncUndefinedNodeClass(nodeModel);
  applyNodeVisualType(nodeModel);
  nodeEl.classList.remove("media-node");
  const existing = nodeEl.querySelector(".media-thumbnail");
  existing?.remove();

  const preview = resolveMediaPreview(nodeModel.metadata);
  if (!preview) {
    return;
  }

  const thumb = document.createElement("div");
  thumb.className = "media-thumbnail";
  thumb.setAttribute("contenteditable", "false");
  thumb.setAttribute("role", "button");
  thumb.setAttribute("tabindex", "0");
  thumb.setAttribute("title", "點擊開啟媒體預覽");
  thumb.addEventListener("pointerdown", (event) => event.stopPropagation());
  thumb.addEventListener("click", (event) => onMediaThumbnailClick(event, nodeModel.id));
  thumb.addEventListener("keydown", (event) => onMediaThumbnailKeyDown(event, nodeModel.id));

  if (preview.thumbnailDataUrl) {
    const image = document.createElement("img");
    image.src = preview.thumbnailDataUrl;
    image.alt = preview.altText ?? "media thumbnail";
    image.loading = "lazy";
    image.draggable = false;
    thumb.appendChild(image);
  } else {
    thumb.classList.add("fallback");
    thumb.textContent = preview.label ?? "MEDIA";
  }

  const contentEl = nodeEl.querySelector(".node-content");
  if (!contentEl) {
    return;
  }
  nodeEl.insertBefore(thumb, contentEl);
  nodeEl.classList.add("media-node");
}

function resolveMediaPreview(metadata) {
  if (!isMediaMetadata(metadata)) {
    return null;
  }

  const mediaKind = metadata.file?.kind ?? null;
  const thumbnailDataUrl = metadata.preview?.thumbnailDataUrl;
  if (typeof thumbnailDataUrl === "string" && thumbnailDataUrl.startsWith("data:image/")) {
    return {
      thumbnailDataUrl,
      altText: metadata.file?.name || "media thumbnail",
    };
  }

  const kindLabelMap = {
    image: "IMAGE",
    video: "VIDEO",
    audio: "AUDIO",
    text: "TEXT",
    json: "JSON",
    binary: "FILE",
  };
  if (mediaKind && kindLabelMap[mediaKind]) {
    return {
      label: kindLabelMap[mediaKind],
    };
  }
  return null;
}

function isMediaMetadata(metadata) {
  return Boolean(metadata && typeof metadata === "object" && metadata.schema === MEDIA_NODE_SCHEMA);
}

function isBindingMetadata(metadata) {
  return Boolean(metadata && typeof metadata === "object" && metadata.schema === BINDING_NODE_SCHEMA);
}

function isPlainRecord(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function deepCloneJsonValue(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return null;
  }
}

function buildCompactTextPreview(value, maxChars = GENERIC_JSON_PREVIEW_MAX_CHARS) {
  const normalized = String(value ?? "").replace(/\s+/g, " ").trim();
  if (normalized.length <= maxChars) {
    return normalized;
  }
  return `${normalized.slice(0, Math.max(1, maxChars - 3))}...`;
}

function buildGenericJsonPreview(value, maxChars = GENERIC_JSON_PREVIEW_MAX_CHARS) {
  let text = "";
  try {
    text = JSON.stringify(value);
  } catch {
    text = String(value ?? "");
  }
  return buildCompactTextPreview(text, maxChars);
}

function buildJsonNodeSummaryText(value) {
  if (Array.isArray(value)) {
    return `JSON 陣列 · ${value.length} 筆`;
  }
  if (isPlainRecord(value)) {
    return "JSON 物件 · 雙擊編輯";
  }
  return "JSON 內容 · 雙擊編輯";
}

function looksLikeVerboseJsonText(value) {
  if (typeof value !== "string") {
    return false;
  }
  const normalized = value.replace(/\s+/g, " ").trim();
  if (normalized.length < 32) {
    return false;
  }
  if (!(normalized.startsWith("{") || normalized.startsWith("["))) {
    return false;
  }
  return normalized.includes(":") || normalized.includes(",");
}

function buildUndefinedNodeCompactContent(metadata, rawContent) {
  if (isPlainRecord(metadata)) {
    return normalizeNodeContent(buildJsonNodeSummaryText(metadata));
  }

  const parsedObject = parseJsonObjectTextSafely(rawContent);
  if (parsedObject) {
    return normalizeNodeContent(buildJsonNodeSummaryText(parsedObject));
  }

  return normalizeNodeContent(buildCompactTextPreview(rawContent));
}

function buildHandleSideKey(side, index = 1) {
  if (!HANDLE_SIDES.includes(side)) {
    return null;
  }
  const safeIndex = Number.isFinite(index) ? Math.max(1, Math.floor(index)) : 1;
  return safeIndex <= 1 ? side : `${side}:${safeIndex}`;
}

function parseHandleSideReference(value) {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim().toLowerCase();
  const match = /^(top|right|bottom|left)(?::(\d+))?$/.exec(trimmed);
  if (!match) {
    return null;
  }
  const side = match[1];
  const index = match[2] ? Math.max(1, Number.parseInt(match[2], 10)) : 1;
  const key = buildHandleSideKey(side, index);
  if (!key) {
    return null;
  }
  return { side, index, key };
}

function normalizeHandleSideKey(value) {
  const parsed = parseHandleSideReference(value);
  return parsed ? parsed.key : null;
}

function isValidHandleSideReference(value) {
  return normalizeHandleSideKey(value) != null;
}

function parseBindingDirectiveKey(key) {
  if (typeof key !== "string") {
    return null;
  }
  const trimmed = key.trim();
  const modernMatch = BINDING_DIRECTIVE_PATTERN.exec(trimmed);
  if (modernMatch) {
    const side = modernMatch[1].toLowerCase();
    const index = modernMatch[2] ? Math.max(1, Number.parseInt(modernMatch[2], 10)) : 1;
    const handleKey = buildHandleSideKey(side, index);
    return handleKey ? { side, index, handleKey } : null;
  }

  const legacyMatch = BINDING_DIRECTIVE_LEGACY_PATTERN.exec(trimmed);
  if (!legacyMatch) {
    return null;
  }

  const sideMap = {
    t: "top",
    r: "right",
    b: "bottom",
    l: "left",
  };
  const side = sideMap[legacyMatch[1].toLowerCase()];
  if (!side) {
    return null;
  }
  const index = legacyMatch[2] ? Math.max(1, Number.parseInt(legacyMatch[2], 10)) : 1;
  const handleKey = buildHandleSideKey(side, index);
  return handleKey ? { side, index, handleKey } : null;
}

function parseBindingCheckDirectiveKey(key) {
  if (typeof key !== "string") {
    return null;
  }
  const match = BINDING_CHECK_DIRECTIVE_PATTERN.exec(key.trim());
  if (!match) {
    return null;
  }
  const rawIndex = match[1] || match[2];
  const index = rawIndex ? Math.max(1, Number.parseInt(rawIndex, 10)) : 1;
  return {
    index,
    checkKey: `check_${index}`,
  };
}

function coerceBindingToggleValue(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "" || normalized === "false" || normalized === "0" || normalized === "off" || normalized === "no") {
      return false;
    }
    if (normalized === "true" || normalized === "1" || normalized === "on" || normalized === "yes") {
      return true;
    }
  }
  return Boolean(value);
}

function ensureBindingValueDefaultsFromChecks(values, checks, options = {}) {
  if (!isPlainRecord(values) || !isPlainRecord(checks)) {
    return;
  }
  const coerceExisting = options.coerceExisting === true;
  Object.values(checks).forEach((fieldName) => {
    if (typeof fieldName !== "string" || fieldName.trim() === "") {
      return;
    }
    const normalizedField = normalizeBindingTargetExpression(fieldName.trim()) ?? fieldName.trim();
    const previous = getBindingValueByTarget(values, normalizedField);
    if (previous === undefined) {
      setBindingValueByTarget(values, normalizedField, false);
      return;
    }
    if (coerceExisting) {
      setBindingValueByTarget(values, normalizedField, coerceBindingToggleValue(previous));
    }
  });
}

function resolveBindingCheckEntries(metadata) {
  if (!isBindingMetadata(metadata)) {
    return [];
  }
  const checksSource = isPlainRecord(metadata.checks) ? metadata.checks : {};
  const entries = [];
  Object.entries(checksSource).forEach(([rawKey, rawField]) => {
    const parsed = parseBindingCheckDirectiveKey(rawKey);
    if (!parsed) {
      return;
    }
    if (typeof rawField !== "string" || rawField.trim() === "") {
      return;
    }
    const trimmedField = rawField.trim();
    const normalizedField = normalizeBindingTargetExpression(trimmedField) ?? trimmedField;
    entries.push({
      ...parsed,
      fieldName: normalizedField,
    });
  });
  entries.sort((a, b) => a.index - b.index);
  return entries;
}

function compareBindingHandleKeys(handleA, handleB) {
  const parsedA = parseHandleSideReference(handleA);
  const parsedB = parseHandleSideReference(handleB);
  if (!parsedA || !parsedB) {
    return String(handleA).localeCompare(String(handleB));
  }
  const sideOrder = HANDLE_SIDES.indexOf(parsedA.side) - HANDLE_SIDES.indexOf(parsedB.side);
  if (sideOrder !== 0) {
    return sideOrder;
  }
  return parsedA.index - parsedB.index;
}

function resolveBindingReferenceEntries(metadata) {
  if (!isBindingMetadata(metadata)) {
    return [];
  }

  const refs = isPlainRecord(metadata.refs) ? metadata.refs : {};
  const counts = resolveNodeHandleCounts(metadata);
  const handleSet = new Set(buildHandleSideKeysFromCounts(counts));
  Object.keys(refs).forEach((key) => {
    const normalized = normalizeHandleSideKey(key);
    if (normalized) {
      handleSet.add(normalized);
    }
  });

  return Array.from(handleSet)
    .sort(compareBindingHandleKeys)
    .map((handleKey) => {
      const rawField = refs[handleKey];
      const fieldName =
        typeof rawField === "string" && rawField.trim() !== ""
          ? normalizeBindingTargetExpression(rawField.trim()) ?? rawField.trim()
          : "";
      return {
        handleKey,
        fieldName,
      };
    });
}

function cloneHandleCounts(counts) {
  const base = createDefaultHandleCounts();
  if (!isPlainRecord(counts)) {
    return base;
  }
  HANDLE_SIDES.forEach((side) => {
    const raw = Number(counts[side]);
    if (Number.isFinite(raw)) {
      base[side] = Math.max(0, Math.floor(raw));
    }
  });
  return base;
}

function sanitizeJsonEditorFunctionalConfig(config, options = {}) {
  if (!isPlainRecord(config)) {
    return null;
  }

  const handleCounts = cloneHandleCounts(config.handleCounts);
  const refs = {};
  const checks = {};
  const valueControls = {};

  const refsSource = isPlainRecord(config.refs) ? config.refs : {};
  Object.entries(refsSource).forEach(([rawHandle, rawTarget]) => {
    const normalizedHandle = normalizeHandleSideKey(rawHandle) ?? parseBindingDirectiveKey(rawHandle)?.handleKey ?? null;
    if (!normalizedHandle) {
      return;
    }
    const parsedHandle = parseHandleSideReference(normalizedHandle);
    if (!parsedHandle) {
      return;
    }
    handleCounts[parsedHandle.side] = Math.max(handleCounts[parsedHandle.side], parsedHandle.index);
    if (typeof rawTarget !== "string" || rawTarget.trim() === "") {
      refs[normalizedHandle] = "";
      return;
    }
    refs[normalizedHandle] = normalizeBindingTargetExpression(rawTarget.trim()) ?? rawTarget.trim();
  });

  const checksSource = isPlainRecord(config.checks) ? config.checks : {};
  Object.entries(checksSource).forEach(([rawCheckKey, rawTarget]) => {
    const parsedCheck = parseBindingCheckDirectiveKey(rawCheckKey);
    if (!parsedCheck) {
      return;
    }
    if (typeof rawTarget !== "string" || rawTarget.trim() === "") {
      checks[parsedCheck.checkKey] = "";
      return;
    }
    checks[parsedCheck.checkKey] = normalizeBindingTargetExpression(rawTarget.trim()) ?? rawTarget.trim();
  });

  const valueControlsSource = isPlainRecord(config.valueControls) ? config.valueControls : {};
  Object.entries(valueControlsSource).forEach(([rawFieldKey, rawValue]) => {
    const parsedField = parseBindingValueControlField(rawFieldKey);
    if (!parsedField) {
      return;
    }
    if (parsedField.kind === "date") {
      valueControls[parsedField.fieldName] = normalizeBindingDateValue(rawValue);
      return;
    }
    valueControls[parsedField.fieldName] = normalizeBindingTextValue(rawValue);
  });

  if (options.ensureDefaultChecks === true) {
    if (!Object.prototype.hasOwnProperty.call(checks, "check_1")) {
      checks.check_1 = "";
    }
    if (!Object.prototype.hasOwnProperty.call(checks, "check_2")) {
      checks.check_2 = "";
    }
  }

  const hasHandles = HANDLE_SIDES.some((side) => handleCounts[side] > 0);
  const hasRefs = Object.keys(refs).length > 0;
  const hasChecks = Object.keys(checks).length > 0;
  const hasValueControls = Object.keys(valueControls).length > 0;
  if (!hasHandles && !hasRefs && !hasChecks && !hasValueControls) {
    return null;
  }

  return { handleCounts, refs, checks, valueControls };
}

function createJsonEditorFunctionalConfigFromBindingMetadata(metadata, options = {}) {
  const normalized = normalizeBindingMetadata(metadata);
  if (!normalized) {
    return null;
  }

  const refs = {};
  const checks = {};
  const valueControls = {};
  const refsSource = isPlainRecord(normalized.refs) ? normalized.refs : {};
  Object.entries(refsSource).forEach(([handleKey, fieldName]) => {
    const normalizedHandle = normalizeHandleSideKey(handleKey);
    if (!normalizedHandle) {
      return;
    }
    if (typeof fieldName !== "string" || fieldName.trim() === "") {
      refs[normalizedHandle] = "";
      return;
    }
    refs[normalizedHandle] = normalizeBindingTargetExpression(fieldName.trim()) ?? fieldName.trim();
  });

  const checksSource = isPlainRecord(normalized.checks) ? normalized.checks : {};
  Object.entries(checksSource).forEach(([checkKey, fieldName]) => {
    const parsed = parseBindingCheckDirectiveKey(checkKey);
    if (!parsed) {
      return;
    }
    if (typeof fieldName !== "string" || fieldName.trim() === "") {
      checks[parsed.checkKey] = "";
      return;
    }
    checks[parsed.checkKey] = normalizeBindingTargetExpression(fieldName.trim()) ?? fieldName.trim();
  });

  const valuesSource = isPlainRecord(normalized.values) ? normalized.values : {};
  Object.entries(valuesSource).forEach(([fieldKey, fieldValue]) => {
    const parsedField = parseBindingValueControlField(fieldKey);
    if (!parsedField) {
      return;
    }
    if (parsedField.kind === "date") {
      valueControls[parsedField.fieldName] = normalizeBindingDateValue(fieldValue);
      return;
    }
    valueControls[parsedField.fieldName] = normalizeBindingTextValue(fieldValue);
  });

  return sanitizeJsonEditorFunctionalConfig(
    {
      handleCounts: resolveNodeHandleCounts(normalized),
      refs,
      checks,
      valueControls,
    },
    options
  );
}

function extractJsonEditorFunctionalConfigFromObject(sourceObject, options = {}) {
  const cleanObject = {};
  const extractedConfig = {
    handleCounts: createDefaultHandleCounts(),
    refs: {},
    checks: {},
    valueControls: {},
  };

  if (isPlainRecord(sourceObject)) {
    Object.entries(sourceObject).forEach(([key, value]) => {
      const directive = parseBindingDirectiveKey(key);
      if (directive) {
        extractedConfig.handleCounts[directive.side] = Math.max(extractedConfig.handleCounts[directive.side], directive.index);
        if (typeof value === "string" && value.trim() !== "") {
          extractedConfig.refs[directive.handleKey] = normalizeBindingTargetExpression(value.trim()) ?? value.trim();
        } else {
          extractedConfig.refs[directive.handleKey] = "";
        }
        return;
      }

      const check = parseBindingCheckDirectiveKey(key);
      if (check) {
        if (typeof value === "string" && value.trim() !== "") {
          extractedConfig.checks[check.checkKey] = normalizeBindingTargetExpression(value.trim()) ?? value.trim();
        } else {
          extractedConfig.checks[check.checkKey] = "";
        }
        return;
      }

      const valueControl = parseBindingValueControlField(key);
      if (valueControl) {
        if (valueControl.kind === "date") {
          extractedConfig.valueControls[valueControl.fieldName] = normalizeBindingDateValue(value);
        } else {
          extractedConfig.valueControls[valueControl.fieldName] = normalizeBindingTextValue(value);
        }
        return;
      }

      cleanObject[key] = deepCloneJsonValue(value);
    });
  }

  if (options.seedDefaults === true) {
    Object.keys(createDefaultBindingDirectiveSeed()).forEach((key) => {
      const directive = parseBindingDirectiveKey(key);
      if (directive) {
        extractedConfig.handleCounts[directive.side] = Math.max(extractedConfig.handleCounts[directive.side], directive.index);
      }
      const check = parseBindingCheckDirectiveKey(key);
      if (check && !Object.prototype.hasOwnProperty.call(extractedConfig.checks, check.checkKey)) {
        extractedConfig.checks[check.checkKey] = "";
      }
    });
  }

  const functionalConfig = sanitizeJsonEditorFunctionalConfig(extractedConfig, options);
  return {
    cleanObject: reorderNodeMetaEditorKeys(cleanObject),
    functionalConfig,
  };
}

function mergeJsonEditorFunctionalConfigIntoObject(sourceObject, functionalConfig, options = {}) {
  const output = ensureDefinitionDirectiveDefaultsObject(sourceObject);
  Object.keys(output).forEach((key) => {
    if (parseBindingDirectiveKey(key) || parseBindingCheckDirectiveKey(key) || parseBindingValueControlField(key)) {
      delete output[key];
    }
  });

  const normalizedConfig = sanitizeJsonEditorFunctionalConfig(functionalConfig, options);
  if (!normalizedConfig) {
    return output;
  }

  const includeEmptyRefs = options.includeEmptyRefs !== false;
  const includeEmptyChecks = options.includeEmptyChecks !== false;
  const includeValueControls = options.includeValueControls !== false;

  HANDLE_SIDES.forEach((side) => {
    const total = Math.max(0, Number(normalizedConfig.handleCounts?.[side]) || 0);
    for (let index = 1; index <= total; index += 1) {
      const handleKey = buildHandleSideKey(side, index);
      if (!handleKey) {
        continue;
      }
      const directiveKey = formatBindingDirectiveFromHandleKey(handleKey);
      const target = typeof normalizedConfig.refs?.[handleKey] === "string" ? normalizedConfig.refs[handleKey].trim() : "";
      if (target) {
        output[directiveKey] = target;
      } else if (includeEmptyRefs) {
        output[directiveKey] = null;
      }
    }
  });

  Object.keys(normalizedConfig.checks)
    .sort((a, b) => (parseBindingCheckDirectiveKey(a)?.index ?? 1) - (parseBindingCheckDirectiveKey(b)?.index ?? 1))
    .forEach((checkKey) => {
      const target = typeof normalizedConfig.checks[checkKey] === "string" ? normalizedConfig.checks[checkKey].trim() : "";
      if (target) {
        output[checkKey] = target;
      } else if (includeEmptyChecks) {
        output[checkKey] = null;
      }
    });

  if (includeValueControls) {
    Object.entries(normalizedConfig.valueControls ?? {})
      .sort((a, b) => {
        const pa = parseBindingValueControlField(a[0]);
        const pb = parseBindingValueControlField(b[0]);
        if (!pa || !pb) {
          return String(a[0]).localeCompare(String(b[0]));
        }
        const kindOrder = { text: 0, date: 1 };
        const kindDiff = (kindOrder[pa.kind] ?? 9) - (kindOrder[pb.kind] ?? 9);
        if (kindDiff !== 0) {
          return kindDiff;
        }
        return pa.index - pb.index;
      })
      .forEach(([fieldKey, fieldValue]) => {
        const parsedField = parseBindingValueControlField(fieldKey);
        if (!parsedField) {
          return;
        }
        output[parsedField.fieldName] = parsedField.kind === "date"
          ? normalizeBindingDateValue(fieldValue)
          : normalizeBindingTextValue(fieldValue);
      });
  }

  return output;
}

function isFunctionalInternalTargetExpression(expression) {
  const normalized = normalizeBindingTargetExpression(expression);
  if (!normalized) {
    return false;
  }

  if (!normalized.startsWith("$")) {
    return (
      normalized === NODE_DEFINITION_TITLE_KEY ||
      normalized === NODE_DEFINITION_DESCRIPTION_KEY ||
      parseBindingDirectiveKey(normalized) != null ||
      parseBindingCheckDirectiveKey(normalized) != null ||
      parseBindingValueControlField(normalized) != null
    );
  }

  const tokens = parseBindingTargetPathTokens(normalized);
  if (!tokens || tokens.length !== 1 || tokens[0].type !== "key") {
    return false;
  }
  const key = tokens[0].value;
  return (
    key === NODE_DEFINITION_TITLE_KEY ||
    key === NODE_DEFINITION_DESCRIPTION_KEY ||
    parseBindingDirectiveKey(key) != null ||
    parseBindingCheckDirectiveKey(key) != null ||
    parseBindingValueControlField(key) != null
  );
}

function collectJsonEditorFunctionalTargetOptions() {
  const parsed = parseJsonObjectTextSafely(jsonEditorTextarea.value);
  if (!parsed) {
    return [];
  }
  const valuesRoot = deepCloneJsonValue(parsed) ?? {};
  if (!isPlainRecord(valuesRoot)) {
    return [];
  }
  delete valuesRoot[NODE_DEFINITION_TITLE_KEY];
  delete valuesRoot[NODE_DEFINITION_DESCRIPTION_KEY];
  Object.keys(valuesRoot).forEach((key) => {
    if (parseBindingDirectiveKey(key) || parseBindingCheckDirectiveKey(key) || parseBindingValueControlField(key)) {
      delete valuesRoot[key];
    }
  });

  return collectBindingTargetOptions(valuesRoot).filter((option) => !isFunctionalInternalTargetExpression(option.value));
}

function escapeRegExp(value) {
  return String(value ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function countOrderedKeyMatches(text, keyList) {
  if (!Array.isArray(keyList) || keyList.length === 0) {
    return 0;
  }
  let cursor = 0;
  let matched = 0;
  for (const key of keyList) {
    const needle = `"${String(key).replace(/"/g, '\\"')}"`;
    const foundAt = text.indexOf(needle, cursor);
    if (foundAt < 0) {
      continue;
    }
    matched += 1;
    cursor = foundAt + needle.length;
  }
  return matched;
}

function collectAllNeedleMatches(text, needle) {
  if (typeof text !== "string" || typeof needle !== "string" || needle === "") {
    return [];
  }
  const matches = [];
  const regex = new RegExp(escapeRegExp(needle), "g");
  let match;
  while ((match = regex.exec(text)) != null) {
    matches.push({ index: match.index, length: needle.length });
    if (regex.lastIndex === match.index) {
      regex.lastIndex += 1;
    }
  }
  return matches;
}

function findJsonTextRangeByBindingTarget(text, targetExpression) {
  if (typeof text !== "string" || text.trim() === "") {
    return null;
  }

  const normalizedTarget = normalizeBindingTargetExpression(targetExpression);
  if (!normalizedTarget) {
    return null;
  }

  let tokens = [];
  if (normalizedTarget.startsWith("$")) {
    tokens = parseBindingTargetPathTokens(normalizedTarget) ?? [];
  }

  const keyTokens = tokens.filter((token) => token.type === "key").map((token) => token.value);
  const contextKeys = keyTokens.slice(0, -1);
  const lastToken = tokens.length > 0 ? tokens[tokens.length - 1] : null;
  const lastKeyToken = [...tokens].reverse().find((token) => token.type === "key") ?? null;
  const primaryKey = lastToken?.type === "key" ? lastToken.value : lastKeyToken?.value ?? null;

  const needleCandidates = [];
  if (primaryKey) {
    needleCandidates.push(`"${String(primaryKey).replace(/"/g, '\\"')}"`);
  }
  if (!normalizedTarget.startsWith("$")) {
    needleCandidates.push(`"${String(normalizedTarget).replace(/"/g, '\\"')}"`);
  }

  for (const needle of needleCandidates) {
    const matches = collectAllNeedleMatches(text, needle);
    if (matches.length === 0) {
      continue;
    }
    let best = null;
    matches.forEach((item) => {
      const windowStart = Math.max(0, item.index - 12000);
      const prefix = text.slice(windowStart, item.index);
      const contextMatched = countOrderedKeyMatches(prefix, contextKeys);
      const nearObjectKey = text[item.index + item.length] === ":" ? 2 : 0;
      const score = contextMatched * 10 + nearObjectKey;
      if (!best || score > best.score || (score === best.score && item.index < best.index)) {
        best = { ...item, score };
      }
    });
    if (best) {
      return { start: best.index, end: best.index + best.length };
    }
  }

  if (!normalizedTarget.startsWith("$")) {
    const fallbackIndex = text.indexOf(normalizedTarget);
    if (fallbackIndex >= 0) {
      return { start: fallbackIndex, end: fallbackIndex + normalizedTarget.length };
    }
  }

  return null;
}

function navigateJsonEditorToBindingTarget(targetExpression, options = {}) {
  if (!jsonEditorTextarea) {
    return;
  }
  const range = findJsonTextRangeByBindingTarget(jsonEditorTextarea.value, targetExpression);
  if (!range) {
    return;
  }

  const lineHeightRaw = Number.parseFloat(window.getComputedStyle(jsonEditorTextarea).lineHeight || "");
  const lineHeight = Number.isFinite(lineHeightRaw) && lineHeightRaw > 0 ? lineHeightRaw : 18;
  const line = jsonEditorTextarea.value.slice(0, range.start).split("\n").length - 1;
  const targetScrollTop = Math.max(0, line * lineHeight - jsonEditorTextarea.clientHeight * 0.35);
  jsonEditorTextarea.scrollTop = targetScrollTop;
  jsonEditorTextarea.setSelectionRange(range.start, range.end);
  if (options.focus === true) {
    jsonEditorTextarea.focus();
  }
}

function findNextHandleIndexForSide(config, side) {
  const base = Math.max(0, Number(config?.handleCounts?.[side]) || 0);
  const refs = isPlainRecord(config?.refs) ? config.refs : {};
  let maxIndex = base;
  Object.keys(refs).forEach((handleKey) => {
    const parsed = parseHandleSideReference(handleKey);
    if (!parsed || parsed.side !== side) {
      return;
    }
    maxIndex = Math.max(maxIndex, parsed.index);
  });
  return maxIndex + 1;
}

function removeHandleFromFunctionalConfig(config, handleKey) {
  if (!isPlainRecord(config) || !isPlainRecord(config.handleCounts) || !isPlainRecord(config.refs)) {
    return false;
  }
  const parsed = parseHandleSideReference(handleKey);
  if (!parsed) {
    return false;
  }

  const total = Math.max(0, Number(config.handleCounts[parsed.side]) || 0);
  if (total <= 0) {
    return false;
  }

  const nextRefs = {};
  Object.entries(config.refs).forEach(([currentHandleKey, target]) => {
    const currentParsed = parseHandleSideReference(currentHandleKey);
    if (!currentParsed || currentParsed.side !== parsed.side) {
      nextRefs[currentHandleKey] = target;
      return;
    }
    if (currentParsed.index === parsed.index) {
      return;
    }
    if (currentParsed.index > parsed.index) {
      const shiftedKey = buildHandleSideKey(parsed.side, currentParsed.index - 1);
      if (shiftedKey) {
        nextRefs[shiftedKey] = target;
      }
      return;
    }
    nextRefs[currentHandleKey] = target;
  });

  config.refs = nextRefs;
  config.handleCounts[parsed.side] = Math.max(0, total - 1);
  return true;
}

function findNextCheckIndex(config) {
  const checks = isPlainRecord(config?.checks) ? config.checks : {};
  let maxIndex = 0;
  Object.keys(checks).forEach((checkKey) => {
    const parsed = parseBindingCheckDirectiveKey(checkKey);
    if (!parsed) {
      return;
    }
    maxIndex = Math.max(maxIndex, parsed.index);
  });
  return maxIndex + 1;
}

function removeCheckFromFunctionalConfig(config, checkKey) {
  if (!isPlainRecord(config) || !isPlainRecord(config.checks)) {
    return false;
  }
  const parsed = parseBindingCheckDirectiveKey(checkKey);
  if (!parsed) {
    return false;
  }

  const nextChecks = {};
  Object.entries(config.checks).forEach(([currentCheckKey, target]) => {
    const currentParsed = parseBindingCheckDirectiveKey(currentCheckKey);
    if (!currentParsed) {
      return;
    }
    if (currentParsed.index === parsed.index) {
      return;
    }
    const nextIndex = currentParsed.index > parsed.index ? currentParsed.index - 1 : currentParsed.index;
    nextChecks[`check_${nextIndex}`] = target;
  });
  config.checks = nextChecks;
  return true;
}

function findNextValueControlIndex(config, kind) {
  const controls = isPlainRecord(config?.valueControls) ? config.valueControls : {};
  let maxIndex = 0;
  Object.keys(controls).forEach((fieldKey) => {
    const parsed = parseBindingValueControlField(fieldKey);
    if (!parsed || parsed.kind !== kind) {
      return;
    }
    maxIndex = Math.max(maxIndex, parsed.index);
  });
  return maxIndex + 1;
}

function removeValueControlFromFunctionalConfig(config, fieldKey) {
  if (!isPlainRecord(config) || !isPlainRecord(config.valueControls)) {
    return false;
  }
  const parsed = parseBindingValueControlField(fieldKey);
  if (!parsed) {
    return false;
  }

  const nextControls = {};
  Object.entries(config.valueControls).forEach(([currentFieldKey, value]) => {
    const currentParsed = parseBindingValueControlField(currentFieldKey);
    if (!currentParsed || currentParsed.kind !== parsed.kind) {
      nextControls[currentFieldKey] = value;
      return;
    }
    if (currentParsed.index === parsed.index) {
      return;
    }
    const nextIndex = currentParsed.index > parsed.index ? currentParsed.index - 1 : currentParsed.index;
    const renamedKey =
      parsed.kind === "date" ? `date_yyyy_mm_dd_${nextIndex}` : `text_${nextIndex}`;
    nextControls[renamedKey] = value;
  });
  config.valueControls = nextControls;
  return true;
}

function renderJsonEditorFunctionalControls() {
  if (!jsonEditorFunctionalControlsEl) {
    return;
  }

  const activeMode = state.jsonEditorMode === "binding" || state.jsonEditorMode === "node-definition";
  const config = sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig);
  if (!activeMode || !config) {
    state.jsonEditorFunctionalConfig = null;
    jsonEditorFunctionalControlsEl.hidden = true;
    jsonEditorFunctionalControlsEl.textContent = "";
    return;
  }

  state.jsonEditorFunctionalConfig = config;
  jsonEditorFunctionalControlsEl.hidden = false;
  jsonEditorFunctionalControlsEl.textContent = "";

  const options = collectJsonEditorFunctionalTargetOptions();
  const optionValues = new Set(options.map((item) => item.value));

  const noteEl = document.createElement("p");
  noteEl.className = "json-editor-functional-note";
  noteEl.textContent = "功能性參數改在這裡配對；上方 JSON 僅顯示一般資料欄位。";
  jsonEditorFunctionalControlsEl.appendChild(noteEl);

  const addBarEl = document.createElement("div");
  addBarEl.className = "json-editor-functional-addbar";
  const addTypeSelectEl = document.createElement("select");
  addTypeSelectEl.className = "json-editor-functional-action-select";
  [
    { value: "ref", label: "連接" },
    { value: "check", label: "勾選" },
    { value: "text", label: "文字" },
    { value: "date", label: "日期" },
  ].forEach((item) => {
    const optionEl = document.createElement("option");
    optionEl.value = item.value;
    optionEl.textContent = item.label;
    addTypeSelectEl.appendChild(optionEl);
  });
  const addSideSelectEl = document.createElement("select");
  addSideSelectEl.className = "json-editor-functional-action-select";
  [
    { value: "right", label: "右" },
    { value: "top", label: "上" },
    { value: "left", label: "左" },
    { value: "bottom", label: "下" },
  ].forEach((item) => {
    const optionEl = document.createElement("option");
    optionEl.value = item.value;
    optionEl.textContent = item.label;
    addSideSelectEl.appendChild(optionEl);
  });
  const syncAddSideSelectVisibility = () => {
    addSideSelectEl.hidden = addTypeSelectEl.value !== "ref";
  };
  addTypeSelectEl.addEventListener("change", syncAddSideSelectVisibility);
  syncAddSideSelectVisibility();
  const addBtnEl = document.createElement("button");
  addBtnEl.type = "button";
  addBtnEl.className = "json-editor-functional-action-btn";
  addBtnEl.textContent = "新增參數";
  addBtnEl.addEventListener("click", () => {
    const latest = sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig);
    if (!latest) {
      return;
    }
    if (!isPlainRecord(latest.valueControls)) {
      latest.valueControls = {};
    }
    if (addTypeSelectEl.value === "ref") {
      const side = HANDLE_SIDES.includes(addSideSelectEl.value) ? addSideSelectEl.value : "right";
      const nextIndex = findNextHandleIndexForSide(latest, side);
      latest.handleCounts[side] = Math.max(latest.handleCounts[side], nextIndex);
      const handleKey = buildHandleSideKey(side, nextIndex);
      if (handleKey) {
        latest.refs[handleKey] = "";
      }
    } else if (addTypeSelectEl.value === "check") {
      const nextIndex = findNextCheckIndex(latest);
      latest.checks[`check_${nextIndex}`] = "";
    } else if (addTypeSelectEl.value === "text") {
      const nextIndex = findNextValueControlIndex(latest, "text");
      latest.valueControls[`text_${nextIndex}`] = "";
    } else if (addTypeSelectEl.value === "date") {
      const nextIndex = findNextValueControlIndex(latest, "date");
      latest.valueControls[`date_yyyy_mm_dd_${nextIndex}`] = "";
    }
    state.jsonEditorFunctionalConfig = latest;
    renderJsonEditorFunctionalControls();
  });
  addBarEl.append(addTypeSelectEl, addSideSelectEl, addBtnEl);
  jsonEditorFunctionalControlsEl.appendChild(addBarEl);

  const appendSelect = (groupEl, kind, key, labelText, currentValue) => {
    const itemEl = document.createElement("div");
    itemEl.className = "json-editor-functional-item";

    const labelEl = document.createElement("span");
    labelEl.className = "json-editor-functional-label";
    labelEl.textContent = labelText;
    labelEl.setAttribute("role", "button");
    labelEl.setAttribute("tabindex", "0");

    const selectEl = document.createElement("select");
    selectEl.className = "json-editor-functional-select";
    selectEl.dataset.kind = kind;
    selectEl.dataset.key = key;

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "未綁定";
    selectEl.appendChild(emptyOption);

    const trimmedCurrent = typeof currentValue === "string" ? currentValue.trim() : "";
    if (trimmedCurrent && !optionValues.has(trimmedCurrent)) {
      const optionEl = document.createElement("option");
      optionEl.value = trimmedCurrent;
      optionEl.textContent = `${trimmedCurrent} (目前)`;
      selectEl.appendChild(optionEl);
    }

    options.forEach((option) => {
      const optionEl = document.createElement("option");
      optionEl.value = option.value;
      optionEl.textContent = option.label;
      selectEl.appendChild(optionEl);
    });

    selectEl.value = trimmedCurrent;
    selectEl.addEventListener("change", onJsonEditorFunctionalSelectChange);
    itemEl.addEventListener("mouseenter", () => navigateJsonEditorToBindingTarget(selectEl.value));
    labelEl.addEventListener("click", () => navigateJsonEditorToBindingTarget(selectEl.value, { focus: true }));
    labelEl.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        navigateJsonEditorToBindingTarget(selectEl.value, { focus: true });
      }
    });
    selectEl.addEventListener("focus", () => navigateJsonEditorToBindingTarget(selectEl.value));

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "json-editor-functional-delete-btn";
    deleteBtn.textContent = "刪除";
    deleteBtn.dataset.kind = kind;
    deleteBtn.dataset.key = key;
    deleteBtn.addEventListener("click", onJsonEditorFunctionalDeleteClick);

    itemEl.append(labelEl, selectEl, deleteBtn);
    groupEl.appendChild(itemEl);
  };

  const appendGroupTitle = (groupEl, titleText) => {
    const titleEl = document.createElement("h3");
    titleEl.textContent = titleText;
    groupEl.appendChild(titleEl);
  };

  const handleKeys = buildHandleSideKeysFromCounts(config.handleCounts).sort(compareBindingHandleKeys);
  const refsGroup = document.createElement("section");
  refsGroup.className = "json-editor-functional-group";
  appendGroupTitle(refsGroup, "連接配對");

  if (handleKeys.length > 0) {
    handleKeys.forEach((handleKey) => {
      appendSelect(
        refsGroup,
        "ref",
        handleKey,
        formatBindingDirectiveFromHandleKey(handleKey),
        config.refs?.[handleKey] ?? ""
      );
    });
  } else {
    const emptyEl = document.createElement("p");
    emptyEl.className = "json-editor-functional-note";
    emptyEl.textContent = "尚無連接參數，請先新增。";
    refsGroup.appendChild(emptyEl);
  }
  jsonEditorFunctionalControlsEl.appendChild(refsGroup);

  const checkKeys = Object.keys(config.checks).sort(
    (a, b) => (parseBindingCheckDirectiveKey(a)?.index ?? 1) - (parseBindingCheckDirectiveKey(b)?.index ?? 1)
  );
  const checksGroup = document.createElement("section");
  checksGroup.className = "json-editor-functional-group";
  appendGroupTitle(checksGroup, "勾選配對");

  if (checkKeys.length > 0) {
    checkKeys.forEach((checkKey) => {
      appendSelect(checksGroup, "check", checkKey, checkKey, config.checks?.[checkKey] ?? "");
    });
  } else {
    const emptyEl = document.createElement("p");
    emptyEl.className = "json-editor-functional-note";
    emptyEl.textContent = "尚無勾選參數，請先新增。";
    checksGroup.appendChild(emptyEl);
  }
  jsonEditorFunctionalControlsEl.appendChild(checksGroup);

  const valueControlEntries = Object.entries(config.valueControls ?? {})
    .map(([fieldKey, fieldValue]) => {
      const parsed = parseBindingValueControlField(fieldKey);
      if (!parsed) {
        return null;
      }
      return {
        fieldKey: parsed.fieldName,
        kind: parsed.kind,
        index: parsed.index,
        value: fieldValue,
      };
    })
    .filter(Boolean)
    .sort((a, b) => {
      const kindOrder = { text: 0, date: 1 };
      const kindDiff = (kindOrder[a.kind] ?? 9) - (kindOrder[b.kind] ?? 9);
      if (kindDiff !== 0) {
        return kindDiff;
      }
      return a.index - b.index;
    });

  const valuesGroup = document.createElement("section");
  valuesGroup.className = "json-editor-functional-group";
  appendGroupTitle(valuesGroup, "欄位配對");
  if (valueControlEntries.length > 0) {
    valueControlEntries.forEach((entry) => {
      const itemEl = document.createElement("div");
      itemEl.className = "json-editor-functional-item";

      const labelEl = document.createElement("span");
      labelEl.className = "json-editor-functional-label";
      labelEl.textContent = entry.fieldKey;

      const inputEl = document.createElement("input");
      inputEl.className = "json-editor-functional-value-input";
      inputEl.dataset.kind = "value-control";
      inputEl.dataset.key = entry.fieldKey;
      if (entry.kind === "date") {
        inputEl.type = "date";
        inputEl.value = normalizeBindingDateValue(entry.value);
      } else {
        inputEl.type = "text";
        inputEl.value = normalizeBindingTextValue(entry.value);
        inputEl.placeholder = "預設文字";
      }
      inputEl.addEventListener("change", onJsonEditorFunctionalValueInputChange);

      const deleteBtn = document.createElement("button");
      deleteBtn.type = "button";
      deleteBtn.className = "json-editor-functional-delete-btn";
      deleteBtn.textContent = "刪除";
      deleteBtn.dataset.kind = "value-control";
      deleteBtn.dataset.key = entry.fieldKey;
      deleteBtn.addEventListener("click", onJsonEditorFunctionalDeleteClick);

      itemEl.append(labelEl, inputEl, deleteBtn);
      valuesGroup.appendChild(itemEl);
    });
  } else {
    const emptyEl = document.createElement("p");
    emptyEl.className = "json-editor-functional-note";
    emptyEl.textContent = "尚無文字/日期參數，請先新增。";
    valuesGroup.appendChild(emptyEl);
  }
  jsonEditorFunctionalControlsEl.appendChild(valuesGroup);
}

function onJsonEditorFunctionalSelectChange(event) {
  const selectEl = event.target;
  if (!(selectEl instanceof HTMLSelectElement)) {
    return;
  }

  const config = sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig);
  if (!config) {
    return;
  }
  const kind = selectEl.dataset.kind;
  const key = selectEl.dataset.key ?? "";
  const rawTarget = selectEl.value.trim();
  const normalizedTarget = rawTarget ? normalizeBindingTargetExpression(rawTarget) ?? rawTarget : "";

  if (kind === "ref") {
    const normalizedHandle = normalizeHandleSideKey(key);
    if (!normalizedHandle) {
      return;
    }
    config.refs[normalizedHandle] = normalizedTarget;
  } else if (kind === "check") {
    const parsedCheck = parseBindingCheckDirectiveKey(key);
    if (!parsedCheck) {
      return;
    }
    config.checks[parsedCheck.checkKey] = normalizedTarget;
  } else {
    return;
  }

  state.jsonEditorFunctionalConfig = config;
  navigateJsonEditorToBindingTarget(normalizedTarget, { focus: true });
}

function onJsonEditorFunctionalValueInputChange(event) {
  const inputEl = event.target;
  if (!(inputEl instanceof HTMLInputElement)) {
    return;
  }
  const fieldKey = inputEl.dataset.key ?? "";
  const parsedField = parseBindingValueControlField(fieldKey);
  if (!parsedField) {
    return;
  }

  const config = sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig);
  if (!config) {
    return;
  }
  if (!isPlainRecord(config.valueControls)) {
    config.valueControls = {};
  }

  if (parsedField.kind === "date") {
    config.valueControls[parsedField.fieldName] = normalizeBindingDateValue(inputEl.value);
  } else {
    config.valueControls[parsedField.fieldName] = normalizeBindingTextValue(inputEl.value);
  }
  state.jsonEditorFunctionalConfig = config;
}

function onJsonEditorFunctionalDeleteClick(event) {
  event.preventDefault();
  const buttonEl = event.target;
  if (!(buttonEl instanceof HTMLButtonElement)) {
    return;
  }
  const kind = buttonEl.dataset.kind;
  const key = buttonEl.dataset.key ?? "";
  const config = sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig);
  if (!config) {
    return;
  }

  if (kind === "ref") {
    if (!removeHandleFromFunctionalConfig(config, key)) {
      return;
    }
  } else if (kind === "check") {
    if (!removeCheckFromFunctionalConfig(config, key)) {
      return;
    }
  } else if (kind === "value-control") {
    if (!removeValueControlFromFunctionalConfig(config, key)) {
      return;
    }
  } else {
    return;
  }

  state.jsonEditorFunctionalConfig = config;
  renderJsonEditorFunctionalControls();
}

function onJsonEditorTextareaInput() {
  if (state.jsonEditorFunctionalConfig) {
    renderJsonEditorFunctionalControls();
  }
}

function parseBindingValueControlField(fieldName) {
  if (typeof fieldName !== "string") {
    return null;
  }
  const normalized = fieldName.trim();
  if (!normalized) {
    return null;
  }

  const dateMatch = BINDING_VALUE_DATE_FIELD_PATTERN.exec(normalized);
  if (dateMatch) {
    return {
      fieldName: normalized,
      kind: "date",
      index: Math.max(1, Number.parseInt(dateMatch[1], 10)),
    };
  }

  const textMatch = BINDING_VALUE_TEXT_FIELD_PATTERN.exec(normalized);
  if (textMatch) {
    return {
      fieldName: normalized,
      kind: "text",
      index: Math.max(1, Number.parseInt(textMatch[1], 10)),
    };
  }

  return null;
}

function resolveBindingValueControlEntries(metadata) {
  if (!isBindingMetadata(metadata) || !isPlainRecord(metadata.values)) {
    return [];
  }

  const entries = Object.keys(metadata.values)
    .map((fieldName) => parseBindingValueControlField(fieldName))
    .filter(Boolean);
  entries.sort((a, b) => {
    const kindOrder = {
      date: 0,
      text: 1,
    };
    const orderDiff = (kindOrder[a.kind] ?? 99) - (kindOrder[b.kind] ?? 99);
    if (orderDiff !== 0) {
      return orderDiff;
    }
    return a.index - b.index;
  });
  return entries;
}

function normalizeBindingDateValue(value) {
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }
    const parsed = new Date(trimmed);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toISOString().slice(0, 10);
    }
  }
  return "";
}

function normalizeBindingTextValue(value) {
  if (value == null) {
    return "";
  }
  return String(value);
}

function isSafeBindingPathIdentifier(value) {
  return typeof value === "string" && /^[A-Za-z_][A-Za-z0-9_]*$/.test(value);
}

function parseBindingTargetPathTokens(expression) {
  if (typeof expression !== "string") {
    return null;
  }
  const text = expression.trim();
  if (!text.startsWith("$")) {
    return null;
  }

  const tokens = [];
  let index = 1;
  while (index < text.length) {
    const char = text[index];
    if (char === ".") {
      index += 1;
      const start = index;
      while (index < text.length && /[A-Za-z0-9_]/.test(text[index])) {
        index += 1;
      }
      if (start === index) {
        return null;
      }
      tokens.push({ type: "key", value: text.slice(start, index) });
      continue;
    }

    if (char === "[") {
      index += 1;
      if (index >= text.length) {
        return null;
      }
      const quote = text[index];
      if (quote === "'" || quote === "\"") {
        index += 1;
        let key = "";
        while (index < text.length) {
          const current = text[index];
          if (current === "\\") {
            if (index + 1 >= text.length) {
              return null;
            }
            key += text[index + 1];
            index += 2;
            continue;
          }
          if (current === quote) {
            index += 1;
            break;
          }
          key += current;
          index += 1;
        }
        if (text[index] !== "]") {
          return null;
        }
        index += 1;
        tokens.push({ type: "key", value: key });
        continue;
      }

      const numberStart = index;
      while (index < text.length && /[0-9]/.test(text[index])) {
        index += 1;
      }
      if (numberStart === index || text[index] !== "]") {
        return null;
      }
      const parsedIndex = Number.parseInt(text.slice(numberStart, index), 10);
      index += 1;
      tokens.push({ type: "index", value: Math.max(0, parsedIndex) });
      continue;
    }

    return null;
  }
  return tokens;
}

function buildBindingTargetPathExpression(tokens) {
  if (!Array.isArray(tokens)) {
    return "$";
  }
  let output = "$";
  tokens.forEach((token) => {
    if (!token || typeof token !== "object") {
      return;
    }
    if (token.type === "index") {
      output += `[${Math.max(0, Number(token.value) || 0)}]`;
      return;
    }
    const key = String(token.value ?? "");
    if (isSafeBindingPathIdentifier(key)) {
      output += `.${key}`;
      return;
    }
    const escaped = key.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    output += `['${escaped}']`;
  });
  return output;
}

function normalizeBindingTargetExpression(value) {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  if (!trimmed.startsWith("$")) {
    return trimmed;
  }
  const tokens = parseBindingTargetPathTokens(trimmed);
  if (!tokens) {
    return null;
  }
  return buildBindingTargetPathExpression(tokens);
}

function getBindingValueByTarget(root, targetExpression) {
  const normalizedTarget = normalizeBindingTargetExpression(targetExpression);
  if (!normalizedTarget || !isPlainRecord(root)) {
    return undefined;
  }

  if (!normalizedTarget.startsWith("$")) {
    return root[normalizedTarget];
  }

  const tokens = parseBindingTargetPathTokens(normalizedTarget);
  if (!tokens) {
    return undefined;
  }

  let current = root;
  for (const token of tokens) {
    if (token.type === "index") {
      if (!Array.isArray(current)) {
        return undefined;
      }
      current = current[token.value];
      continue;
    }
    if (!current || typeof current !== "object") {
      return undefined;
    }
    current = current[token.value];
  }
  return current;
}

function setBindingValueByTarget(root, targetExpression, value) {
  const normalizedTarget = normalizeBindingTargetExpression(targetExpression);
  if (!normalizedTarget || !isPlainRecord(root)) {
    return false;
  }

  if (!normalizedTarget.startsWith("$")) {
    root[normalizedTarget] = value;
    return true;
  }

  const tokens = parseBindingTargetPathTokens(normalizedTarget);
  if (!tokens || tokens.length === 0) {
    return false;
  }

  let current = root;
  for (let index = 0; index < tokens.length - 1; index += 1) {
    const token = tokens[index];
    const next = tokens[index + 1];
    if (token.type === "index") {
      if (!Array.isArray(current)) {
        return false;
      }
      let child = current[token.value];
      if (next.type === "index") {
        if (!Array.isArray(child)) {
          child = [];
          current[token.value] = child;
        }
      } else if (!isPlainRecord(child)) {
        child = {};
        current[token.value] = child;
      }
      current = child;
      continue;
    }

    let child = current[token.value];
    if (next.type === "index") {
      if (!Array.isArray(child)) {
        child = [];
        current[token.value] = child;
      }
    } else if (!isPlainRecord(child)) {
      child = {};
      current[token.value] = child;
    }
    current = child;
  }

  const last = tokens[tokens.length - 1];
  if (last.type === "index") {
    if (!Array.isArray(current)) {
      return false;
    }
    current[last.value] = value;
    return true;
  }
  if (!current || typeof current !== "object") {
    return false;
  }
  current[last.value] = value;
  return true;
}

function collectBindingTargetOptions(valuesRoot) {
  if (!isPlainRecord(valuesRoot)) {
    return [];
  }

  const options = [];
  const seen = new Set();
  const pushOption = (expression) => {
    const normalized = normalizeBindingTargetExpression(expression);
    if (!normalized || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    options.push({
      value: normalized,
      label: normalized,
    });
  };

  Object.keys(valuesRoot).forEach((key) => {
    pushOption(key);
  });

  const visit = (value, tokens, depth = 0) => {
    if (depth > 8) {
      return;
    }
    if (isPlainRecord(value)) {
      const keys = Object.keys(value);
      if (keys.length === 0) {
        pushOption(buildBindingTargetPathExpression(tokens));
        return;
      }
      keys.forEach((key) => visit(value[key], tokens.concat({ type: "key", value: key }), depth + 1));
      return;
    }
    if (Array.isArray(value)) {
      if (value.length === 0) {
        pushOption(buildBindingTargetPathExpression(tokens));
        return;
      }
      const limit = Math.min(value.length, 20);
      for (let idx = 0; idx < limit; idx += 1) {
        visit(value[idx], tokens.concat({ type: "index", value: idx }), depth + 1);
      }
      return;
    }
    pushOption(buildBindingTargetPathExpression(tokens));
  };

  visit(valuesRoot, []);
  return options.sort((a, b) => a.label.localeCompare(b.label));
}

function createDefaultHandleCounts() {
  return {
    top: 0,
    right: 0,
    bottom: 0,
    left: 0,
  };
}

function createDefaultBindingDirectiveSeed() {
  return {
    node_Right_1: null,
    node_Top_1: null,
    node_Left_1: null,
    node_Bottom_1: null,
    check_1: null,
    check_2: null,
  };
}

function ensureDefinitionDirectiveDefaultsObject(source) {
  const base = isPlainRecord(source) ? deepCloneJsonValue(source) ?? {} : {};
  return isPlainRecord(base) ? base : {};
}

function extractNodeDefinitionMatchFromTemplate(template) {
  if (!isPlainRecord(template)) {
    return { requiredKeys: [], typeField: null, typeValue: null };
  }

  const requiredKeys = Object.keys(template)
    .filter((key) => !resolveFunctionalParameterSortInfo(key))
    .filter((key) => key !== NODE_DEFINITION_TITLE_KEY && key !== NODE_DEFINITION_DESCRIPTION_KEY)
    .sort((a, b) => a.localeCompare(b));
  const match = {
    requiredKeys,
    typeField: null,
    typeValue: null,
  };

  if (Object.prototype.hasOwnProperty.call(template, "type")) {
    const typeValue = template.type;
    if (typeof typeValue === "string" || typeof typeValue === "number" || typeof typeValue === "boolean") {
      match.typeField = "type";
      match.typeValue = typeValue;
    }
  }

  return match;
}

function normalizeNodeDefinitionMatch(rawMatch, template) {
  const fromTemplate = extractNodeDefinitionMatchFromTemplate(template);
  if (!isPlainRecord(rawMatch)) {
    return fromTemplate;
  }

  const rawRequiredKeys = Array.isArray(rawMatch.requiredKeys) ? rawMatch.requiredKeys : fromTemplate.requiredKeys;
  const requiredKeys = Array.from(
    new Set(
      rawRequiredKeys
        .filter((key) => typeof key === "string")
        .map((key) => key.trim())
        .filter((key) => key !== NODE_DEFINITION_TITLE_KEY && key !== NODE_DEFINITION_DESCRIPTION_KEY)
        .filter((key) => !resolveFunctionalParameterSortInfo(key))
        .filter((key) => key !== "")
    )
  ).sort((a, b) => a.localeCompare(b));

  const typeField = typeof rawMatch.typeField === "string" && rawMatch.typeField.trim() !== "" ? rawMatch.typeField.trim() : null;
  const rawTypeValue = rawMatch.typeValue;
  const typeValue =
    typeof rawTypeValue === "string" || typeof rawTypeValue === "number" || typeof rawTypeValue === "boolean" ? rawTypeValue : null;

  return {
    requiredKeys: requiredKeys.length > 0 ? requiredKeys : fromTemplate.requiredKeys,
    typeField,
    typeValue,
  };
}

function buildNodeDefinitionMatchSignature(match) {
  return JSON.stringify({
    requiredKeys: Array.isArray(match?.requiredKeys) ? match.requiredKeys : [],
    typeField: match?.typeField ?? null,
    typeValue: match?.typeValue ?? null,
  });
}

function normalizeNodeDefinitionRecord(rawDefinition, fallbackIndex = 1) {
  if (!isPlainRecord(rawDefinition)) {
    return null;
  }

  const templateSource =
    isPlainRecord(rawDefinition.template) ? rawDefinition.template : isPlainRecord(rawDefinition.payload) ? rawDefinition.payload : null;
  if (!templateSource) {
    return null;
  }

  const template = ensureDefinitionDirectiveDefaultsObject(templateSource);
  const match = normalizeNodeDefinitionMatch(rawDefinition.match, template);
  const rawId = typeof rawDefinition.id === "string" ? rawDefinition.id.trim() : "";
  const id = rawId || `def_${Date.now().toString(36)}_${fallbackIndex}`;
  const templateTitle =
    typeof template[NODE_DEFINITION_TITLE_KEY] === "string" && template[NODE_DEFINITION_TITLE_KEY].trim() !== ""
      ? normalizeNodeTitle(template[NODE_DEFINITION_TITLE_KEY], fallbackIndex)
      : null;
  const name = typeof rawDefinition.name === "string" && rawDefinition.name.trim() !== "" ? rawDefinition.name.trim() : templateTitle;
  const templateDescription =
    typeof template[NODE_DEFINITION_DESCRIPTION_KEY] === "string" && template[NODE_DEFINITION_DESCRIPTION_KEY].trim() !== ""
      ? normalizeNodeContent(template[NODE_DEFINITION_DESCRIPTION_KEY])
      : null;
  const description =
    typeof rawDefinition.description === "string" && rawDefinition.description.trim() !== ""
      ? normalizeNodeContent(rawDefinition.description)
      : templateDescription;

  return {
    id,
    name,
    description,
    version: NODE_DEFINITION_VERSION,
    match,
    template,
  };
}

function normalizeNodeDefinitionsPayload(rawDefinitions) {
  if (!Array.isArray(rawDefinitions)) {
    return [];
  }

  const normalized = [];
  const seenMatch = new Set();
  rawDefinitions.forEach((item, index) => {
    const definition = normalizeNodeDefinitionRecord(item, index + 1);
    if (!definition) {
      return;
    }
    const signature = buildNodeDefinitionMatchSignature(definition.match);
    if (seenMatch.has(signature)) {
      return;
    }
    seenMatch.add(signature);
    normalized.push(definition);
  });
  return normalized;
}

function upsertNodeDefinition(definition) {
  const normalized = normalizeNodeDefinitionRecord(definition, state.nodeDefinitions.length + 1);
  if (!normalized) {
    return false;
  }

  const signature = buildNodeDefinitionMatchSignature(normalized.match);
  const index = state.nodeDefinitions.findIndex(
    (item) => buildNodeDefinitionMatchSignature(item.match) === signature
  );
  if (index >= 0) {
    const previous = state.nodeDefinitions[index];
    state.nodeDefinitions[index] = {
      ...normalized,
      id: previous.id,
      name: normalized.name ?? previous.name ?? null,
      description: normalized.description ?? previous.description ?? null,
    };
  } else {
    state.nodeDefinitions.push(normalized);
  }
  renderDefinedNodeTypeList();
  return true;
}

function doesNodeDefinitionMatch(definition, rawObject) {
  if (!definition || !isPlainRecord(rawObject)) {
    return false;
  }

  const requiredKeys = Array.isArray(definition.match?.requiredKeys) ? definition.match.requiredKeys : [];
  if (requiredKeys.some((key) => !Object.prototype.hasOwnProperty.call(rawObject, key))) {
    return false;
  }

  const typeField = definition.match?.typeField;
  if (typeof typeField === "string" && typeField !== "") {
    if (!Object.prototype.hasOwnProperty.call(rawObject, typeField)) {
      return false;
    }
    if (rawObject[typeField] !== definition.match.typeValue) {
      return false;
    }
  }

  return true;
}

function selectNodeDefinitionForObject(rawObject) {
  if (!isPlainRecord(rawObject) || state.nodeDefinitions.length === 0) {
    return null;
  }

  const matched = state.nodeDefinitions.filter((definition) => doesNodeDefinitionMatch(definition, rawObject));
  if (matched.length === 0) {
    return null;
  }

  matched.sort((a, b) => {
    const scoreA = (a.match?.requiredKeys?.length ?? 0) + (a.match?.typeField ? 1000 : 0);
    const scoreB = (b.match?.requiredKeys?.length ?? 0) + (b.match?.typeField ? 1000 : 0);
    return scoreB - scoreA;
  });
  return matched[0];
}

function parseBindingSyntaxObject(source) {
  if (!isPlainRecord(source)) {
    return null;
  }

  const values = {};
  const refs = {};
  const checks = {};
  const handleCounts = createDefaultHandleCounts();
  let hasSpecialDirective = false;

  Object.entries(source).forEach(([key, value]) => {
    if (key === NODE_DEFINITION_TITLE_KEY || key === NODE_DEFINITION_DESCRIPTION_KEY) {
      return;
    }
    const directive = parseBindingDirectiveKey(key);
    if (directive) {
      hasSpecialDirective = true;
      handleCounts[directive.side] = Math.max(handleCounts[directive.side], directive.index);
      if (typeof value === "string" && value.trim() !== "") {
        const normalizedTarget = normalizeBindingTargetExpression(value.trim());
        if (normalizedTarget) {
          refs[directive.handleKey] = normalizedTarget;
        }
      }
      return;
    }

    const checkDirective = parseBindingCheckDirectiveKey(key);
    if (checkDirective) {
      hasSpecialDirective = true;
      if (typeof value === "string" && value.trim() !== "") {
        const normalizedTarget = normalizeBindingTargetExpression(value.trim()) ?? value.trim();
        checks[checkDirective.checkKey] = normalizedTarget;
      }
      return;
    }

    values[key] = deepCloneJsonValue(value);
  });

  if (!hasSpecialDirective) {
    return null;
  }

  ensureBindingValueDefaultsFromChecks(values, checks, { coerceExisting: true });
  delete values[NODE_DEFINITION_TITLE_KEY];
  delete values[NODE_DEFINITION_DESCRIPTION_KEY];

  return {
    schema: BINDING_NODE_SCHEMA,
    version: BINDING_NODE_VERSION,
    values,
    refs,
    checks,
    handleCounts,
  };
}

function normalizeBindingMetadata(metadata) {
  if (!isPlainRecord(metadata)) {
    return null;
  }

  const valuesSource = isPlainRecord(metadata.values) ? metadata.values : {};
  const refsSource = isPlainRecord(metadata.refs) ? metadata.refs : {};
  const checksSource = isPlainRecord(metadata.checks) ? metadata.checks : {};
  const countsSource = isPlainRecord(metadata.handleCounts) ? metadata.handleCounts : {};
  const displaySource = isPlainRecord(metadata.display) ? metadata.display : {};
  const values = deepCloneJsonValue(valuesSource) ?? {};
  const refs = {};
  const checks = {};
  const handleCounts = createDefaultHandleCounts();

  Object.entries(countsSource).forEach(([side, rawCount]) => {
    if (!HANDLE_SIDES.includes(side)) {
      return;
    }
    if (!Number.isFinite(rawCount)) {
      return;
    }
    handleCounts[side] = Math.max(0, Math.floor(rawCount));
  });

  Object.entries(refsSource).forEach(([rawHandle, rawField]) => {
    const normalizedHandle = normalizeHandleSideKey(rawHandle) ?? parseBindingDirectiveKey(rawHandle)?.handleKey ?? null;
    if (!normalizedHandle) {
      return;
    }
    const parsedHandle = parseHandleSideReference(normalizedHandle);
    if (!parsedHandle) {
      return;
    }
    handleCounts[parsedHandle.side] = Math.max(handleCounts[parsedHandle.side], parsedHandle.index);
    if (typeof rawField === "string" && rawField.trim() !== "") {
      const normalizedTarget = normalizeBindingTargetExpression(rawField.trim());
      if (normalizedTarget) {
        refs[normalizedHandle] = normalizedTarget;
      }
    }
  });

  Object.entries(checksSource).forEach(([rawCheckKey, rawField]) => {
    const parsedCheck = parseBindingCheckDirectiveKey(rawCheckKey);
    if (!parsedCheck) {
      return;
    }
    if (typeof rawField !== "string" || rawField.trim() === "") {
      return;
    }
    const normalizedTarget = normalizeBindingTargetExpression(rawField.trim()) ?? rawField.trim();
    checks[parsedCheck.checkKey] = normalizedTarget;
  });

  ensureBindingValueDefaultsFromChecks(values, checks, { coerceExisting: true });

  const display = {};
  if (typeof displaySource.description === "string" && displaySource.description.trim() !== "") {
    display.description = normalizeNodeContent(displaySource.description);
  }

  return {
    schema: BINDING_NODE_SCHEMA,
    version: BINDING_NODE_VERSION,
    values: isPlainRecord(values) ? values : {},
    refs,
    checks,
    handleCounts,
    display,
  };
}

function normalizeNodeMetadata(metadata) {
  if (!isPlainRecord(metadata)) {
    return metadata ?? null;
  }

  if (isBindingMetadata(metadata)) {
    return normalizeBindingMetadata(metadata) ?? null;
  }

  const parsedBinding = parseBindingSyntaxObject(metadata);
  if (parsedBinding) {
    return parsedBinding;
  }

  return metadata;
}

function resolveNodeHandleCounts(metadata) {
  const counts = createDefaultHandleCounts();
  if (!isBindingMetadata(metadata)) {
    return counts;
  }

  const normalized = normalizeBindingMetadata(metadata);
  if (!normalized) {
    return counts;
  }

  HANDLE_SIDES.forEach((side) => {
    counts[side] = Math.max(0, Number(normalized.handleCounts?.[side]) || 0);
  });

  Object.keys(normalized.refs).forEach((handleKey) => {
    const parsed = parseHandleSideReference(handleKey);
    if (!parsed) {
      return;
    }
    counts[parsed.side] = Math.max(counts[parsed.side], parsed.index);
  });

  return counts;
}

function getNodeHandleSideKeys(node) {
  const counts = node?.handleCounts ?? resolveNodeHandleCounts(node?.metadata);
  return buildHandleSideKeysFromCounts(counts);
}

function buildHandleSideKeysFromCounts(counts) {
  const keys = [];
  HANDLE_SIDES.forEach((side) => {
    const total = Math.max(0, Number(counts?.[side]) || 0);
    for (let index = 1; index <= total; index += 1) {
      keys.push(buildHandleSideKey(side, index));
    }
  });
  return keys.filter(Boolean);
}

function computeNodeMinHeightFromHandleCounts(counts) {
  const sideCount = Math.max(1, Number(counts?.left) || 1, Number(counts?.right) || 1);
  return Math.max(NODE_DEFAULT_HEIGHT, 56 + sideCount * 26);
}

function computeNodeMinWidthFromHandleCounts(counts) {
  const sideCount = Math.max(1, Number(counts?.top) || 1, Number(counts?.bottom) || 1);
  return Math.max(NODE_DEFAULT_WIDTH, 56 + sideCount * 26);
}

function positionHandleElement(handleEl, side, index, total) {
  const ratio = index / (total + 1);
  if (side === "top") {
    handleEl.style.left = `${ratio * 100}%`;
    handleEl.style.top = "-7px";
    handleEl.style.right = "";
    handleEl.style.bottom = "";
    handleEl.style.transform = "translateX(-50%)";
    return;
  }
  if (side === "right") {
    handleEl.style.right = "-7px";
    handleEl.style.top = `${ratio * 100}%`;
    handleEl.style.left = "";
    handleEl.style.bottom = "";
    handleEl.style.transform = "translateY(-50%)";
    return;
  }
  if (side === "bottom") {
    handleEl.style.left = `${ratio * 100}%`;
    handleEl.style.bottom = "-7px";
    handleEl.style.right = "";
    handleEl.style.top = "";
    handleEl.style.transform = "translateX(-50%)";
    return;
  }

  handleEl.style.left = "-7px";
  handleEl.style.top = `${ratio * 100}%`;
  handleEl.style.right = "";
  handleEl.style.bottom = "";
  handleEl.style.transform = "translateY(-50%)";
}

function renderNodeHandles(nodeModel) {
  if (!nodeModel?.el) {
    return;
  }

  const counts = resolveNodeHandleCounts(nodeModel.metadata);
  nodeModel.handleCounts = counts;
  nodeModel.el.style.minHeight = `${computeNodeMinHeightFromHandleCounts(counts)}px`;
  nodeModel.el.style.minWidth = `${computeNodeMinWidthFromHandleCounts(counts)}px`;

  nodeModel.el.querySelectorAll(".handle").forEach((handle) => handle.remove());

  const sideLabelMap = {
    top: "上",
    right: "右",
    bottom: "下",
    left: "左",
  };

  HANDLE_SIDES.forEach((side) => {
    const total = Math.max(0, Number(counts?.[side]) || 0);
    for (let index = 1; index <= total; index += 1) {
      const sideKey = buildHandleSideKey(side, index);
      if (!sideKey) {
        continue;
      }

      const handleEl = document.createElement("button");
      handleEl.type = "button";
      handleEl.className = `handle handle-${side}`;
      handleEl.dataset.side = sideKey;
      handleEl.setAttribute("aria-label", `${sideLabelMap[side]}方連接點 ${index}`);
      positionHandleElement(handleEl, side, index, total);
      handleEl.addEventListener("pointerdown", (event) => startLink(event, nodeModel.id, sideKey));
      handleEl.addEventListener("pointerup", (event) => completeLink(event, nodeModel.id, sideKey));
      nodeModel.el.appendChild(handleEl);
    }
  });
}

function resolveConnectedNodeIdsAtHandle(nodeId, handleKey) {
  const normalizedHandle = normalizeHandleSideKey(handleKey);
  if (!normalizedHandle) {
    return [];
  }

  const ids = [];
  state.connections.forEach((connection) => {
    const fromSide = normalizeHandleSideKey(connection.fromSide);
    const toSide = normalizeHandleSideKey(connection.toSide);
    if (connection.fromNodeId === nodeId && fromSide === normalizedHandle) {
      ids.push(connection.toNodeId);
    }
    if (connection.toNodeId === nodeId && toSide === normalizedHandle) {
      ids.push(connection.fromNodeId);
    }
  });

  return Array.from(new Set(ids));
}

function resolveBindingFieldValue(nodeId, handleKey) {
  const connected = resolveConnectedNodeIdsAtHandle(nodeId, handleKey);
  if (connected.length === 0) {
    return null;
  }
  if (connected.length === 1) {
    return connected[0];
  }
  return connected;
}

function formatBindingValuePreview(value) {
  if (value == null) {
    return "null";
  }
  if (Array.isArray(value)) {
    return `[${value.join(", ")}]`;
  }
  if (typeof value === "object") {
    return "{...}";
  }
  const text = String(value);
  return text.length > 48 ? `${text.slice(0, 45)}...` : text;
}

function formatBindingDirectiveFromHandleKey(handleKey) {
  const parsed = parseHandleSideReference(handleKey);
  if (!parsed) {
    return handleKey;
  }
  const sideMap = {
    top: "Top",
    right: "Right",
    bottom: "Bottom",
    left: "Left",
  };
  const sideLabel = sideMap[parsed.side] ?? parsed.side;
  return `node_${sideLabel}_${parsed.index}`;
}

function buildBindingNodeContent(metadata) {
  if (!isBindingMetadata(metadata)) {
    return "JSON 綁定節點";
  }
  const customDescription =
    typeof metadata?.display?.description === "string" && metadata.display.description.trim() !== ""
      ? metadata.display.description.trim()
      : "";
  if (customDescription) {
    return customDescription;
  }

  const refs = isPlainRecord(metadata.refs) ? metadata.refs : {};
  const values = isPlainRecord(metadata.values) ? metadata.values : {};
  const refEntries = Object.entries(refs);
  const checkEntries = resolveBindingCheckEntries(metadata);
  const valueControlEntries = resolveBindingValueControlEntries(metadata);
  const checkSummary =
    checkEntries.length > 0
      ? `勾選: ${checkEntries
          .map(({ fieldName }) => `${fieldName}=${formatBindingValuePreview(getBindingValueByTarget(values, fieldName))}`)
          .join(", ")}`
      : "";
  const valueControlSummary =
    valueControlEntries.length > 0
      ? `欄位: ${valueControlEntries
          .map(({ fieldName }) => `${fieldName}=${formatBindingValuePreview(getBindingValueByTarget(values, fieldName))}`)
          .join(", ")}`
      : "";

  if (refEntries.length === 0) {
    const summaryParts = [checkSummary, valueControlSummary].filter(Boolean);
    return summaryParts.length > 0 ? `JSON 綁定 · ${summaryParts.join(" | ")}` : "JSON 綁定節點 · 尚未設定綁定";
  }

  refEntries.sort((a, b) => compareBindingHandleKeys(a[0], b[0]));

  const refSummary = refEntries
    .map(([handleKey, fieldName]) => {
      const preview = formatBindingValuePreview(getBindingValueByTarget(values, fieldName));
      return `${fieldName} <= ${formatBindingDirectiveFromHandleKey(handleKey)} (${preview})`;
    })
    .join(" | ");
  const summaryParts = [refSummary, checkSummary, valueControlSummary].filter(Boolean);
  return `JSON 綁定 · ${summaryParts.join(" | ")}`;
}

function renderBindingToggleControls(nodeModel) {
  if (!nodeModel?.el) {
    return;
  }

  nodeModel.el.querySelectorAll(".binding-ref-list, .binding-toggle-list, .binding-field-list").forEach((el) => el.remove());
  if (!isBindingMetadata(nodeModel.metadata)) {
    return;
  }

  const checkEntries = resolveBindingCheckEntries(nodeModel.metadata);
  const fieldEntries = resolveBindingValueControlEntries(nodeModel.metadata);
  if (checkEntries.length === 0 && fieldEntries.length === 0) {
    return;
  }

  if (!isPlainRecord(nodeModel.metadata.values)) {
    nodeModel.metadata.values = {};
  }
  const values = nodeModel.metadata.values;
  ensureBindingValueDefaultsFromChecks(values, nodeModel.metadata.checks ?? {});

  const contentEl = nodeModel.el.querySelector(".node-content");
  const insertControlBeforeContent = (element) => {
    if (contentEl) {
      nodeModel.el.insertBefore(element, contentEl);
    } else {
      nodeModel.el.appendChild(element);
    }
  };

  if (checkEntries.length > 0) {
    const listEl = document.createElement("div");
    listEl.className = "binding-toggle-list";
    listEl.setAttribute("contenteditable", "false");
    listEl.addEventListener("pointerdown", (event) => event.stopPropagation());
    listEl.addEventListener("dblclick", (event) => event.stopPropagation());

    checkEntries.forEach((entry) => {
      const itemEl = document.createElement("label");
      itemEl.className = "binding-toggle-item";
      itemEl.dataset.field = entry.fieldName;

      const inputEl = document.createElement("input");
      inputEl.type = "checkbox";
      inputEl.className = "binding-toggle-input";
      inputEl.checked = coerceBindingToggleValue(getBindingValueByTarget(values, entry.fieldName));

      const textEl = document.createElement("span");
      textEl.className = "binding-toggle-label";
      textEl.textContent = entry.fieldName;

      if (inputEl.checked) {
        itemEl.classList.add("active");
      }

      inputEl.addEventListener("pointerdown", (event) => event.stopPropagation());
      inputEl.addEventListener("click", (event) => event.stopPropagation());
      inputEl.addEventListener("change", (event) => onBindingToggleInputChange(event, nodeModel.id, entry.fieldName));

      itemEl.append(inputEl, textEl);
      listEl.appendChild(itemEl);
    });

    insertControlBeforeContent(listEl);
  }

  if (fieldEntries.length > 0) {
    const fieldListEl = document.createElement("div");
    fieldListEl.className = "binding-field-list";
    fieldListEl.setAttribute("contenteditable", "false");
    fieldListEl.addEventListener("pointerdown", (event) => event.stopPropagation());
    fieldListEl.addEventListener("dblclick", (event) => event.stopPropagation());

    fieldEntries.forEach((entry) => {
      const itemEl = document.createElement("label");
      itemEl.className = "binding-field-item";
      itemEl.dataset.field = entry.fieldName;

      const labelEl = document.createElement("span");
      labelEl.className = "binding-field-label";
      labelEl.textContent = entry.fieldName;

      const inputEl = document.createElement("input");
      inputEl.className = "binding-field-input";
      inputEl.dataset.field = entry.fieldName;
      inputEl.addEventListener("pointerdown", (event) => event.stopPropagation());
      inputEl.addEventListener("click", (event) => event.stopPropagation());

      if (entry.kind === "date") {
        inputEl.type = "date";
        inputEl.value = normalizeBindingDateValue(getBindingValueByTarget(values, entry.fieldName));
      } else {
        inputEl.type = "text";
        inputEl.placeholder = "輸入文字";
        inputEl.value = normalizeBindingTextValue(getBindingValueByTarget(values, entry.fieldName));
      }
      inputEl.addEventListener("change", (event) =>
        onBindingValueFieldInputChange(event, nodeModel.id, entry.fieldName, entry.kind)
      );

      itemEl.append(labelEl, inputEl);
      fieldListEl.appendChild(itemEl);
    });

    insertControlBeforeContent(fieldListEl);
  }

}

function onBindingReferenceSelectChange(event, nodeId, handleKey) {
  event.stopPropagation();

  const node = state.nodes.get(nodeId);
  if (!node || !isBindingMetadata(node.metadata)) {
    return;
  }

  const selectEl = event.target;
  if (!(selectEl instanceof HTMLSelectElement)) {
    return;
  }

  const normalizedHandle = normalizeHandleSideKey(handleKey);
  if (!normalizedHandle) {
    return;
  }

  if (!isPlainRecord(node.metadata.refs)) {
    node.metadata.refs = {};
  }
  if (!isPlainRecord(node.metadata.values)) {
    node.metadata.values = {};
  }

  const refs = node.metadata.refs;
  const values = node.metadata.values;
  const previousTarget = typeof refs[normalizedHandle] === "string" ? refs[normalizedHandle] : "";
  const rawNextTarget = selectEl.value.trim();
  const nextTarget = rawNextTarget ? normalizeBindingTargetExpression(rawNextTarget) ?? rawNextTarget : "";

  if (nextTarget) {
    refs[normalizedHandle] = nextTarget;
    const linkedValue = resolveBindingFieldValue(node.id, normalizedHandle);
    if (!setBindingValueByTarget(values, nextTarget, linkedValue)) {
      values[nextTarget] = linkedValue;
    }
  } else {
    delete refs[normalizedHandle];
  }

  node.content = buildBindingNodeContent(node.metadata);
  const contentEl = node.el.querySelector(".node-content");
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute("title", "雙擊編輯綁定 JSON");
  }

  refreshBindingNodes();

  if (previousTarget !== nextTarget) {
    commitHistory();
  }
}

function onBindingToggleInputChange(event, nodeId, fieldName) {
  event.stopPropagation();

  const node = state.nodes.get(nodeId);
  if (!node || !isBindingMetadata(node.metadata) || !isPlainRecord(node.metadata.values)) {
    return;
  }

  const inputEl = event.target;
  if (!(inputEl instanceof HTMLInputElement)) {
    return;
  }

  const nextValue = Boolean(inputEl.checked);
  const normalizedField = normalizeBindingTargetExpression(fieldName) ?? fieldName;
  const prevValue = coerceBindingToggleValue(getBindingValueByTarget(node.metadata.values, normalizedField));
  if (!setBindingValueByTarget(node.metadata.values, normalizedField, nextValue)) {
    node.metadata.values[normalizedField] = nextValue;
  }

  const itemEl = inputEl.closest(".binding-toggle-item");
  itemEl?.classList.toggle("active", nextValue);

  node.content = buildBindingNodeContent(node.metadata);
  const contentEl = node.el.querySelector(".node-content");
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute("title", "雙擊編輯綁定 JSON");
  }

  if (nextValue !== prevValue) {
    commitHistory();
  }
}

function onBindingValueFieldInputChange(event, nodeId, fieldName, kind) {
  event.stopPropagation();

  const node = state.nodes.get(nodeId);
  if (!node || !isBindingMetadata(node.metadata) || !isPlainRecord(node.metadata.values)) {
    return;
  }

  const inputEl = event.target;
  if (!(inputEl instanceof HTMLInputElement)) {
    return;
  }

  const normalizedField = normalizeBindingTargetExpression(fieldName) ?? fieldName;
  const prevValue = getBindingValueByTarget(node.metadata.values, normalizedField);
  const nextValue = kind === "date" ? normalizeBindingDateValue(inputEl.value) : normalizeBindingTextValue(inputEl.value);
  if (!setBindingValueByTarget(node.metadata.values, normalizedField, nextValue)) {
    node.metadata.values[normalizedField] = nextValue;
  }

  node.content = buildBindingNodeContent(node.metadata);
  const contentEl = node.el.querySelector(".node-content");
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute("title", "雙擊編輯綁定 JSON");
  }

  if (JSON.stringify(prevValue) !== JSON.stringify(nextValue)) {
    commitHistory();
  }
}

function refreshBindingNodes() {
  let changed = false;

  state.nodes.forEach((node) => {
    if (!isBindingMetadata(node.metadata)) {
      return;
    }

    const normalizedBinding = normalizeBindingMetadata(node.metadata);
    if (!normalizedBinding) {
      return;
    }

    node.metadata = normalizedBinding;
    const refs = normalizedBinding.refs;
    const values = normalizedBinding.values;
    Object.entries(refs).forEach(([handleKey, fieldName]) => {
      const linkedValue = resolveBindingFieldValue(node.id, handleKey);
      if (!setBindingValueByTarget(values, fieldName, linkedValue)) {
        values[fieldName] = linkedValue;
      }
    });

    const content = buildBindingNodeContent(node.metadata);
    if (node.content !== content) {
      node.content = content;
      const contentEl = node.el.querySelector(".node-content");
      if (contentEl) {
        contentEl.textContent = content;
      }
      changed = true;
    }

    renderBindingToggleControls(node);

    const nextCounts = resolveNodeHandleCounts(node.metadata);
    const previousCounts = node.handleCounts ?? createDefaultHandleCounts();
    const countsChanged = HANDLE_SIDES.some((side) => previousCounts[side] !== nextCounts[side]);
    if (countsChanged) {
      renderNodeHandles(node);
      changed = true;
    }
  });

  return changed;
}

function resolveNodeVisualType(metadata) {
  if (!isMediaMetadata(metadata)) {
    return "default";
  }

  const mediaKind = metadata?.file?.kind;
  if (NODE_VISUAL_TYPES.includes(mediaKind)) {
    return mediaKind;
  }
  return "media";
}

function syncVisualTypeClass(element, classPrefix, visualType) {
  if (!element) {
    return;
  }
  NODE_VISUAL_TYPES.forEach((type) => {
    element.classList.remove(`${classPrefix}${type}`);
  });
  element.classList.add(`${classPrefix}${visualType}`);
}

function applyNodeVisualType(nodeModel) {
  if (!nodeModel?.el) {
    return;
  }
  const visualType = resolveNodeVisualType(nodeModel.metadata);
  nodeModel.el.dataset.nodeType = visualType;
  syncVisualTypeClass(nodeModel.el, NODE_VISUAL_CLASS_PREFIX, visualType);
}

function onNodePointerDown(event, nodeId) {
  if (event.button !== 0) {
    return;
  }
  if (
    event.target.closest(".handle") ||
    event.target.closest(".node-delete-btn") ||
    event.target.closest(".binding-ref-list") ||
    event.target.closest(".binding-toggle-list") ||
    event.target.closest(".binding-field-list") ||
    event.target.closest(".media-thumbnail")
  ) {
    return;
  }

  if (event.detail >= 2) {
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
    pointerStartClientX: event.clientX,
    pointerStartClientY: event.clientY,
    started: false,
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
  const normalizedSide = normalizeHandleSideKey(side);
  if (!normalizedSide) {
    return;
  }

  state.linking = {
    fromNodeId: nodeId,
    fromSide: normalizedSide,
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

  const normalizedSide = normalizeHandleSideKey(side);
  if (!normalizedSide) {
    return false;
  }

  const { fromNodeId, fromSide } = state.linking;
  if (fromNodeId === nodeId && fromSide === normalizedSide) {
    return false;
  }

  const created = createConnection({
    fromNodeId,
    fromSide,
    toNodeId: nodeId,
    toSide: normalizedSide,
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
  const excluded = normalizeHandleSideKey(excludeSide);

  state.nodes.forEach((node) => {
    getNodeHandleSideKeys(node).forEach((sideKey) => {
      if (node.id === excludeNodeId && sideKey === excluded) {
        return;
      }

      const handlePoint = getHandlePointWorld(node, sideKey);
      const distance = Math.hypot(worldPoint.x - handlePoint.x, worldPoint.y - handlePoint.y);
      if (distance > snapRadiusWorld) {
        return;
      }

      if (!closest || distance < closest.distance) {
        closest = { nodeId: node.id, side: sideKey, distance };
      }
    });
  });

  return closest;
}

function createConnection(link) {
  const fromSide = normalizeHandleSideKey(link.fromSide);
  const toSide = normalizeHandleSideKey(link.toSide);
  if (!fromSide || !toSide) {
    return false;
  }

  if (link.fromNodeId === link.toNodeId) {
    return false;
  }

  const duplicated = state.connections.some((connection) => {
    const sameForward =
      connection.fromNodeId === link.fromNodeId &&
      normalizeHandleSideKey(connection.fromSide) === fromSide &&
      connection.toNodeId === link.toNodeId &&
      normalizeHandleSideKey(connection.toSide) === toSide;

    const sameReverse =
      connection.fromNodeId === link.toNodeId &&
      normalizeHandleSideKey(connection.fromSide) === toSide &&
      connection.toNodeId === link.fromNodeId &&
      normalizeHandleSideKey(connection.toSide) === fromSide;

    return sameForward || sameReverse;
  });

  if (duplicated) {
    return false;
  }

  state.connections.push({
    id: String(state.nextConnectionId++),
    ...link,
    fromSide,
    toSide,
  });
  refreshBindingNodes();
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
  refreshBindingNodes();
  renderConnections();
  commitHistory();
}

function onPointerDown(event) {
  hideContextMenu();
  hideDefinedNodePanel();

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
    updateEditorGridBackground();

    state.nodes.forEach((node) => applyNodePosition(node));
    renderConnections();
    state.pointerWorld = editorToWorld(state.pointerScreen.x, state.pointerScreen.y);
  }

  if (state.draggingNode) {
    const dragDistance = Math.hypot(
      event.clientX - state.draggingNode.pointerStartClientX,
      event.clientY - state.draggingNode.pointerStartClientY
    );
    if (!state.draggingNode.started) {
      if (dragDistance < 4) {
        updateNodeProximity();
        updateConnectionProximity();
        return;
      }
      state.draggingNode.started = true;
    }

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
  hideDefinedNodePanel();
  const worldPoint = pointerToWorld(event.clientX, event.clientY);
  const targetNodeEl = event.target.closest(".node");
  const targetNodeId = targetNodeEl?.dataset.nodeId ?? null;

  if (targetNodeId && !state.selectedNodeIds.has(targetNodeId)) {
    setSelection([targetNodeId]);
  }

  showContextMenu(event.clientX, event.clientY, worldPoint, targetNodeId);
}

function onWindowPointerDown(event) {
  if (state.contextMenu) {
    if (event.target.closest("#context-menu")) {
      return;
    }
    hideContextMenu();
  }

  if (definedNodePanelEl?.classList.contains("visible")) {
    if (definedNodeMenuEl && event.target instanceof Node && definedNodeMenuEl.contains(event.target)) {
      return;
    }
    hideDefinedNodePanel();
  }
}

function onWindowPaste(event) {
  if (isTypingTarget(event.target)) {
    return;
  }

  const mediaFiles = extractFilesFromClipboardEvent(event);
  if (mediaFiles.length > 0) {
    event.preventDefault();
    const anchor = pointerInsideEditor() ? state.pointerWorld : getViewportCenterWorld();
    importMediaFiles(mediaFiles, "paste", anchor);
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

function onEditorDragEnter(event) {
  if (!event.dataTransfer) {
    return;
  }
  event.preventDefault();
  state.mediaDragDepth += 1;
  editor.classList.add("media-drop-ready");
}

function onEditorDragOver(event) {
  if (!event.dataTransfer) {
    return;
  }
  event.preventDefault();
  event.dataTransfer.dropEffect = "copy";
}

function onEditorDragLeave(event) {
  if (!event.dataTransfer) {
    return;
  }
  event.preventDefault();
  state.mediaDragDepth = Math.max(0, state.mediaDragDepth - 1);
  if (state.mediaDragDepth === 0) {
    editor.classList.remove("media-drop-ready");
  }
}

function onEditorDrop(event) {
  if (!event.dataTransfer) {
    return;
  }
  event.preventDefault();
  state.mediaDragDepth = 0;
  editor.classList.remove("media-drop-ready");

  const files = Array.from(event.dataTransfer.files ?? []).filter((file) => file && file.size >= 0);
  if (files.length === 0) {
    return;
  }

  const anchor = pointerToWorld(event.clientX, event.clientY);
  importMediaFiles(files, "drop", anchor);
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
    case "edit-json":
      ensureTargetSelected();
      openJsonEditorFromSelection();
      break;
    case "define-node":
      ensureTargetSelected();
      openNodeDefinitionEditorFromSelection();
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
  const primaryNodeId =
    selectionSize > 0
      ? Array.from(state.selectedNodeIds).find((id) => state.nodes.has(id)) ?? null
      : state.contextMenu.targetNodeId ?? null;
  const primaryNode = primaryNodeId ? state.nodes.get(primaryNodeId) : null;
  const hasClipboard = Boolean(state.clipboard && state.clipboard.nodes.length > 0);
  const canReadSystemClipboard = Boolean(navigator.clipboard && typeof navigator.clipboard.readText === "function");

  setContextMenuItemDisabled("copy", !hasNodeForAction);
  setContextMenuItemDisabled("cut", !hasNodeForAction);
  setContextMenuItemDisabled("duplicate", !hasNodeForAction);
  setContextMenuItemDisabled("edit-json", selectionSize === 0 && !hasTargetNode);
  setContextMenuItemDisabled("define-node", !primaryNode || isMediaMetadata(primaryNode.metadata));
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

function onDefinedNodeMenuButtonClick(event) {
  event.preventDefault();
  event.stopPropagation();
  if (!definedNodePanelEl || !definedNodeMenuBtn) {
    return;
  }
  if (definedNodePanelEl.classList.contains("visible")) {
    hideDefinedNodePanel();
  } else {
    showDefinedNodePanel();
  }
}

function onDefinedNodePanelClick(event) {
  const button = event.target.closest("button[data-action]");
  if (!button || button.disabled) {
    return;
  }
  const action = button.dataset.action;
  if (action !== "create-sample") {
    return;
  }
  const definitionId = button.dataset.definitionId ?? "";
  createSampleNodeFromDefinitionId(definitionId);
}

function showDefinedNodePanel() {
  if (!definedNodePanelEl || !definedNodeMenuBtn) {
    return;
  }
  hideContextMenu();
  renderDefinedNodeTypeList();
  definedNodePanelEl.classList.add("visible");
  definedNodePanelEl.setAttribute("aria-hidden", "false");
  definedNodeMenuBtn.setAttribute("aria-expanded", "true");
}

function hideDefinedNodePanel() {
  if (!definedNodePanelEl || !definedNodeMenuBtn) {
    return;
  }
  definedNodePanelEl.classList.remove("visible");
  definedNodePanelEl.setAttribute("aria-hidden", "true");
  definedNodeMenuBtn.setAttribute("aria-expanded", "false");
}

function renderDefinedNodeTypeList() {
  if (!definedNodeListEl) {
    return;
  }

  const definitions = Array.isArray(state.nodeDefinitions) ? state.nodeDefinitions : [];
  if (definedNodeCountEl) {
    definedNodeCountEl.textContent = `${definitions.length} 類`;
  }
  definedNodeListEl.textContent = "";

  if (definitions.length === 0) {
    const emptyEl = document.createElement("p");
    emptyEl.className = "defined-node-empty";
    emptyEl.textContent = "尚未定義任何節點類型。";
    definedNodeListEl.appendChild(emptyEl);
    return;
  }

  definitions.forEach((definition) => {
    const cardEl = document.createElement("article");
    cardEl.className = "defined-node-item";

    const titleEl = document.createElement("div");
    titleEl.className = "defined-node-item-title";
    const rawName = typeof definition?.name === "string" && definition.name.trim() !== "" ? definition.name.trim() : DEFINED_NODE_TITLE;
    titleEl.textContent = rawName;

    const metaEl = document.createElement("div");
    metaEl.className = "defined-node-item-meta";
    const requiredKeys = Array.isArray(definition?.match?.requiredKeys) ? definition.match.requiredKeys : [];
    const keySummary =
      requiredKeys.length > 0
        ? requiredKeys.slice(0, 4).join(", ") + (requiredKeys.length > 4 ? "..." : "")
        : "(無)";
    const typeSummary =
      definition?.match?.typeField && definition.match.typeValue != null
        ? `${definition.match.typeField}=${String(definition.match.typeValue)}`
        : "未指定";
    metaEl.textContent = `匹配欄位: ${keySummary} | type: ${typeSummary}`;

    const createBtn = document.createElement("button");
    createBtn.type = "button";
    createBtn.className = "defined-node-item-action";
    createBtn.dataset.action = "create-sample";
    createBtn.dataset.definitionId = definition.id;
    createBtn.textContent = "建立範例";

    cardEl.append(titleEl, metaEl, createBtn);
    definedNodeListEl.appendChild(cardEl);
  });
}

function createSampleNodeFromDefinitionId(definitionId) {
  const id = typeof definitionId === "string" ? definitionId.trim() : "";
  if (!id) {
    return;
  }
  const definition = state.nodeDefinitions.find((item) => item.id === id);
  if (!definition || !isPlainRecord(definition.template)) {
    return;
  }
  const seed = deepCloneJsonValue(definition.template) ?? {};
  if (!isPlainRecord(seed)) {
    return;
  }
  const payload = buildClipboardPayloadFromNodeDefinition(seed);
  if (!payload) {
    return;
  }
  importClipboardPayload(payload, getViewportCenterWorld());
  hideDefinedNodePanel();
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
  hideDefinedNodePanel();
  hideGitSyncPanel();
  hideMediaViewer();
  renderShortcutEditorList();
  shortcutPanelEl.classList.add("visible");
  shortcutPanelEl.setAttribute("aria-hidden", "false");
}

function hideShortcutPanel() {
  state.shortcutEditingActionId = null;
  shortcutPanelEl.classList.remove("visible");
  shortcutPanelEl.setAttribute("aria-hidden", "true");
}

function toggleGitSyncPanel() {
  if (gitSyncPanelEl?.classList.contains("visible")) {
    hideGitSyncPanel();
    return;
  }
  showGitSyncPanel();
}

function showGitSyncPanel() {
  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideJsonEditor();
  hideNodeTextEditor();
  hideMediaViewer();
  renderGitSyncPanel();
  gitSyncPanelEl?.classList.add("visible");
  gitSyncPanelEl?.setAttribute("aria-hidden", "false");
  updateGitSyncStatusIndicator();
}

function hideGitSyncPanel() {
  gitSyncPanelEl?.classList.remove("visible");
  gitSyncPanelEl?.setAttribute("aria-hidden", "true");
}

function onGitSyncPanelPointerDown(event) {
  const target = event.target;
  const card = gitSyncPanelEl?.querySelector(".git-sync-card");
  if (!card || !card.contains(target)) {
    hideGitSyncPanel();
    return;
  }
  event.stopPropagation();
}

function normalizeShortcutToken(rawToken) {
  const token = String(rawToken ?? "").trim();
  if (!token) {
    return null;
  }
  const lower = token.toLowerCase();
  const aliasMap = {
    ctrl: "Mod",
    control: "Mod",
    cmd: "Mod",
    command: "Mod",
    meta: "Mod",
    mod: "Mod",
    option: "Alt",
    alt: "Alt",
    shift: "Shift",
    esc: "Escape",
    escape: "Escape",
    del: "Delete",
    delete: "Delete",
    backspace: "Backspace",
    enter: "Enter",
    return: "Enter",
    tab: "Tab",
    space: "Space",
    spacebar: "Space",
    question: "Slash",
    slash: "Slash",
    "/": "Slash",
    "?": "Slash",
    plus: "Equal",
    equal: "Equal",
    "=": "Equal",
    minus: "Minus",
    hyphen: "Minus",
    "-": "Minus",
    underscore: "Minus",
    "_": "Minus",
    arrowup: "ArrowUp",
    up: "ArrowUp",
    arrowdown: "ArrowDown",
    down: "ArrowDown",
    arrowleft: "ArrowLeft",
    left: "ArrowLeft",
    arrowright: "ArrowRight",
    right: "ArrowRight",
    home: "Home",
    end: "End",
    insert: "Insert",
    pageup: "PageUp",
    pagedown: "PageDown",
    comma: "Comma",
    period: "Period",
    semicolon: "Semicolon",
    quote: "Quote",
    backquote: "Backquote",
    "`": "Backquote",
    bracketleft: "BracketLeft",
    "[": "BracketLeft",
    bracketright: "BracketRight",
    "]": "BracketRight",
    backslash: "Backslash",
    "\\": "Backslash",
    numpadadd: "Equal",
    numpadsubtract: "Minus",
  };
  if (aliasMap[lower]) {
    return aliasMap[lower];
  }

  if (/^key[a-z]$/i.test(token)) {
    return `Key${token.slice(3).toUpperCase()}`;
  }
  if (/^[a-z]$/i.test(token)) {
    return `Key${token.toUpperCase()}`;
  }
  if (/^digit[0-9]$/i.test(token)) {
    return `Digit${token.slice(5)}`;
  }
  if (/^[0-9]$/.test(token)) {
    return `Digit${token}`;
  }
  if (/^numpad[0-9]$/i.test(token)) {
    return `Digit${token.slice(-1)}`;
  }
  if (/^f([1-9]|1[0-9]|2[0-4])$/i.test(token)) {
    return token.toUpperCase();
  }
  if (/^arrow(up|down|left|right)$/i.test(token)) {
    return `Arrow${token.slice(5, 6).toUpperCase()}${token.slice(6).toLowerCase()}`;
  }
  if (/^(escape|delete|backspace|enter|tab|space|home|end|insert|pageup|pagedown)$/i.test(token)) {
    return token.slice(0, 1).toUpperCase() + token.slice(1).toLowerCase();
  }
  if (
    /^(equal|minus|slash|backslash|backquote|bracketleft|bracketright|comma|period|semicolon|quote)$/i.test(token)
  ) {
    return token.slice(0, 1).toUpperCase() + token.slice(1);
  }
  return null;
}

function normalizeShortcutComboString(value) {
  if (typeof value !== "string") {
    return null;
  }
  const parts = value
    .split("+")
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length === 0) {
    return null;
  }

  const modifiers = new Set();
  let keyToken = null;
  for (const part of parts) {
    const token = normalizeShortcutToken(part);
    if (!token) {
      return null;
    }
    if (SHORTCUT_MODIFIER_SET.has(token)) {
      modifiers.add(token);
      continue;
    }
    if (keyToken) {
      return null;
    }
    keyToken = token;
  }

  if (!keyToken) {
    return null;
  }
  const orderedModifiers = SHORTCUT_MODIFIER_ORDER.filter((token) => modifiers.has(token));
  return [...orderedModifiers, keyToken].join("+");
}

function buildDefaultShortcutBindings() {
  const defaults = {};
  SHORTCUT_ACTION_DEFS.forEach((action) => {
    defaults[action.id] = normalizeShortcutComboString(action.defaultCombo);
  });
  return defaults;
}

function initializeShortcutBindings() {
  const defaults = buildDefaultShortcutBindings();
  let storedBindings = null;
  try {
    const raw = window.localStorage.getItem(SHORTCUT_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        storedBindings = parsed.bindings && typeof parsed.bindings === "object" ? parsed.bindings : parsed;
      }
    }
  } catch {
    storedBindings = null;
  }

  const merged = {};
  SHORTCUT_ACTION_DEFS.forEach((action) => {
    const storedValue = storedBindings && typeof storedBindings === "object" ? storedBindings[action.id] : null;
    merged[action.id] = normalizeShortcutComboString(storedValue) ?? defaults[action.id] ?? null;
  });
  state.shortcuts = merged;
  saveShortcutBindings();
  renderShortcutEditorList();
}

function saveShortcutBindings() {
  try {
    window.localStorage.setItem(
      SHORTCUT_STORAGE_KEY,
      JSON.stringify({
        type: "node-editor.shortcuts",
        version: 1,
        updatedAt: new Date().toISOString(),
        bindings: state.shortcuts,
      })
    );
  } catch (error) {
    console.warn("save shortcuts failed", error);
  }
}

function formatShortcutKeyToken(token) {
  if (!token) {
    return "";
  }
  if (/^Key[A-Z]$/.test(token)) {
    return token.slice(3);
  }
  if (/^Digit[0-9]$/.test(token)) {
    return token.slice(5);
  }
  const tokenLabelMap = {
    Equal: "=",
    Minus: "-",
    Slash: "/",
    Backslash: "\\",
    Backquote: "`",
    BracketLeft: "[",
    BracketRight: "]",
    Comma: ",",
    Period: ".",
    Semicolon: ";",
    Quote: "'",
    Space: "Space",
    Escape: "Esc",
    ArrowUp: "Up",
    ArrowDown: "Down",
    ArrowLeft: "Left",
    ArrowRight: "Right",
    PageUp: "PageUp",
    PageDown: "PageDown",
  };
  if (tokenLabelMap[token]) {
    return tokenLabelMap[token];
  }
  return token;
}

function formatShortcutComboDisplay(combo) {
  const normalized = normalizeShortcutComboString(combo);
  if (!normalized) {
    return "未綁定";
  }
  const parts = normalized.split("+");
  return parts
    .map((part) => {
      if (part === "Mod") {
        return IS_MAC_PLATFORM ? "Cmd" : "Ctrl";
      }
      return formatShortcutKeyToken(part);
    })
    .join(" + ");
}

function renderShortcutEditorList() {
  if (!shortcutEditorListEl) {
    return;
  }
  const fragment = document.createDocumentFragment();
  SHORTCUT_ACTION_DEFS.forEach((action) => {
    const isEditing = state.shortcutEditingActionId === action.id;
    const itemEl = document.createElement("article");
    itemEl.className = "shortcut-editor-item";
    if (isEditing) {
      itemEl.classList.add("editing");
    }
    itemEl.dataset.actionId = action.id;

    const metaEl = document.createElement("div");
    metaEl.className = "shortcut-editor-meta";
    const labelEl = document.createElement("h3");
    labelEl.textContent = action.label;
    const descEl = document.createElement("p");
    descEl.textContent = action.description;
    metaEl.append(labelEl, descEl);

    const controlsEl = document.createElement("div");
    controlsEl.className = "shortcut-editor-controls";
    const kbdEl = document.createElement("kbd");
    kbdEl.className = "shortcut-editor-kbd";
    kbdEl.textContent = isEditing ? "按下新組合鍵..." : formatShortcutComboDisplay(state.shortcuts[action.id]);

    const changeBtn = document.createElement("button");
    changeBtn.type = "button";
    changeBtn.className = "shortcut-editor-btn";
    changeBtn.dataset.shortcutAction = "change";
    changeBtn.textContent = isEditing ? "取消" : "變更";

    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "shortcut-editor-btn ghost";
    clearBtn.dataset.shortcutAction = "clear";
    clearBtn.textContent = "清除";
    clearBtn.disabled = !state.shortcuts[action.id];

    controlsEl.append(kbdEl, changeBtn, clearBtn);
    itemEl.append(metaEl, controlsEl);
    fragment.appendChild(itemEl);
  });
  shortcutEditorListEl.replaceChildren(fragment);
}

function onShortcutEditorListClick(event) {
  const button = event.target.closest("button[data-shortcut-action]");
  if (!button) {
    return;
  }
  const itemEl = button.closest("[data-action-id]");
  const actionId = itemEl?.dataset.actionId ?? "";
  if (!SHORTCUT_ACTION_IDS.has(actionId)) {
    return;
  }
  const command = button.dataset.shortcutAction;
  if (command === "change") {
    if (state.shortcutEditingActionId === actionId) {
      state.shortcutEditingActionId = null;
      renderShortcutEditorList();
      return;
    }
    startShortcutCapture(actionId);
    return;
  }
  if (command === "clear") {
    clearShortcutBinding(actionId);
  }
}

function startShortcutCapture(actionId) {
  if (!SHORTCUT_ACTION_IDS.has(actionId)) {
    return;
  }
  state.shortcutEditingActionId = actionId;
  renderShortcutEditorList();
}

function clearShortcutBinding(actionId) {
  if (!SHORTCUT_ACTION_IDS.has(actionId)) {
    return;
  }
  state.shortcuts[actionId] = null;
  if (state.shortcutEditingActionId === actionId) {
    state.shortcutEditingActionId = null;
  }
  saveShortcutBindings();
  renderShortcutEditorList();
}

function setShortcutBinding(actionId, combo) {
  if (!SHORTCUT_ACTION_IDS.has(actionId)) {
    return;
  }
  const normalizedCombo = normalizeShortcutComboString(combo);
  if (!normalizedCombo) {
    return;
  }
  SHORTCUT_ACTION_DEFS.forEach((action) => {
    if (action.id === actionId) {
      return;
    }
    if (state.shortcuts[action.id] === normalizedCombo) {
      state.shortcuts[action.id] = null;
    }
  });
  state.shortcuts[actionId] = normalizedCombo;
  state.shortcutEditingActionId = null;
  saveShortcutBindings();
  renderShortcutEditorList();
}

function onShortcutResetClick() {
  state.shortcuts = buildDefaultShortcutBindings();
  state.shortcutEditingActionId = null;
  saveShortcutBindings();
  renderShortcutEditorList();
}

function shortcutComboFromKeyboardEvent(event) {
  if (!(event instanceof KeyboardEvent)) {
    return null;
  }

  const codeToken = normalizeShortcutToken(event.code);
  const keyToken = normalizeShortcutToken(event.key);
  const resolvedKey = codeToken && !SHORTCUT_MODIFIER_SET.has(codeToken) ? codeToken : keyToken;
  if (!resolvedKey || SHORTCUT_MODIFIER_SET.has(resolvedKey)) {
    return null;
  }

  const parts = [];
  if (event.ctrlKey || event.metaKey) {
    parts.push("Mod");
  }
  if (event.altKey) {
    parts.push("Alt");
  }
  if (event.shiftKey) {
    parts.push("Shift");
  }
  parts.push(resolvedKey);
  return normalizeShortcutComboString(parts.join("+"));
}

function findShortcutActionByEvent(event) {
  const combo = shortcutComboFromKeyboardEvent(event);
  if (!combo) {
    return null;
  }
  for (const action of SHORTCUT_ACTION_DEFS) {
    if (state.shortcuts[action.id] === combo) {
      return action.id;
    }
  }
  return null;
}

function handleShortcutCaptureKeyDown(event) {
  if (!state.shortcutEditingActionId) {
    return false;
  }
  event.preventDefault();
  event.stopPropagation();

  if (event.key === "Escape" && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
    state.shortcutEditingActionId = null;
    renderShortcutEditorList();
    return true;
  }

  const combo = shortcutComboFromKeyboardEvent(event);
  if (!combo) {
    return true;
  }
  setShortcutBinding(state.shortcutEditingActionId, combo);
  return true;
}

function triggerShortcutAction(actionId, event) {
  switch (actionId) {
    case "toggleShortcutPanel":
      toggleShortcutPanel();
      return true;
    case "newNode":
      createNodeAtViewportCenter();
      return true;
    case "selectAll":
      setSelection(Array.from(state.nodes.keys()));
      return true;
    case "copy":
      copySelectionToClipboard(false);
      return true;
    case "cut":
      copySelectionToClipboard(true);
      return true;
    case "paste":
      if (shortcutComboFromKeyboardEvent(event) === "Mod+KeyV") {
        return false;
      }
      void pasteClipboard(getViewportCenterWorld());
      return true;
    case "duplicate":
      duplicateSelection();
      return true;
    case "saveProject":
      void onSaveProjectClick();
      return true;
    case "loadProject":
      void onLoadProjectClick();
      return true;
    case "loadLegacyProject":
      void onLoadLegacyProjectClick();
      return true;
    case "undo":
      undoHistory();
      return true;
    case "redo":
    case "redoAlt":
      redoHistory();
      return true;
    case "zoomIn":
    case "zoomInAlt":
      zoomByFactor(1.12);
      return true;
    case "zoomOut":
      zoomByFactor(1 / 1.12);
      return true;
    case "resetZoom":
      resetZoom();
      return true;
    case "fitView":
      fitViewToNodes();
      return true;
    case "deleteSelection":
    case "deleteSelectionAlt":
      deleteSelectionOrHover();
      return true;
    case "cancel":
      cancelCurrentInteraction();
      return true;
    case "moveUp":
      moveSelectionBy(0, -(event.shiftKey ? 30 : 10));
      return true;
    case "moveDown":
      moveSelectionBy(0, event.shiftKey ? 30 : 10);
      return true;
    case "moveLeft":
      moveSelectionBy(-(event.shiftKey ? 30 : 10), 0);
      return true;
    case "moveRight":
      moveSelectionBy(event.shiftKey ? 30 : 10, 0);
      return true;
    default:
      return false;
  }
}

function onNodeTextDoubleClick(event, nodeId) {
  event.stopPropagation();
  showNodeTextEditor(nodeId);
}

function showNodeTextEditor(nodeId) {
  const node = state.nodes.get(nodeId);
  if (!node) {
    return;
  }
  const mediaNode = isMediaMetadata(node.metadata);
  const bindingNode = isBindingMetadata(node.metadata);

  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideJsonEditor();
  hideMediaViewer();

  state.nodeTextEditorNodeId = nodeId;
  nodeTextEditorTitleInput.value = node.title;
  nodeTextEditorContentTextarea.value = node.content;
  nodeTextEditorContentTextarea.readOnly = mediaNode || bindingNode;
  nodeTextEditorContentTextarea.title = mediaNode
    ? "媒體節點內容由 JSON 資料摘要產生"
    : bindingNode
      ? "綁定節點內容由 JSON 綁定規則與連線自動產生"
      : "";
  nodeTextEditorPanelEl.classList.add("visible");
  nodeTextEditorPanelEl.setAttribute("aria-hidden", "false");
  nodeTextEditorTitleInput.focus();
  nodeTextEditorTitleInput.select();
}

function onNodeTextEditorInput() {
  const nodeId = state.nodeTextEditorNodeId;
  if (!nodeId) {
    return;
  }
  const node = state.nodes.get(nodeId);
  if (!node) {
    return;
  }

  const previousTitle = node.title;
  const previousContent = node.content;
  const nextTitle = normalizeNodeTitle(nodeTextEditorTitleInput.value, node.id);
  const nextContent = isMediaMetadata(node.metadata)
    ? buildMediaNodeContent(node.metadata)
    : isBindingMetadata(node.metadata)
      ? buildBindingNodeContent(node.metadata)
      : normalizeNodeContent(nodeTextEditorContentTextarea.value);

  if (nextTitle === previousTitle && nextContent === previousContent) {
    return;
  }

  node.title = nextTitle;
  node.content = nextContent;
  const titleEl = node.el.querySelector(".node-title");
  const contentEl = node.el.querySelector(".node-content");
  if (titleEl) {
    titleEl.textContent = nextTitle;
  }
  if (contentEl) {
    contentEl.textContent = nextContent;
  }

  commitHistory();
}

function hideNodeTextEditor() {
  state.nodeTextEditorNodeId = null;
  nodeTextEditorPanelEl.classList.remove("visible");
  nodeTextEditorPanelEl.setAttribute("aria-hidden", "true");
  nodeTextEditorTitleInput.value = "";
  nodeTextEditorContentTextarea.value = "";
  nodeTextEditorContentTextarea.readOnly = false;
  nodeTextEditorContentTextarea.title = "";
}

function onNodeTextEditorPanelPointerDown(event) {
  if (event.target === nodeTextEditorPanelEl) {
    hideNodeTextEditor();
  }
}

function onNodeTextEditorKeyDown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    hideNodeTextEditor();
    return;
  }
  const mod = event.ctrlKey || event.metaKey;
  if (mod && event.key.toLowerCase() === "s") {
    event.preventDefault();
    onNodeTextEditorSaveClick();
  }
}

function onNodeTextEditorEditJsonClick() {
  const nodeId = state.nodeTextEditorNodeId;
  if (!nodeId || !state.nodes.has(nodeId)) {
    return;
  }
  showJsonEditor(nodeId);
}

function onNodeTextEditorDefineJsonClick() {
  const nodeId = state.nodeTextEditorNodeId;
  if (!nodeId || !state.nodes.has(nodeId)) {
    return;
  }
  openNodeDefinitionEditorForNode(nodeId);
}

function onNodeTextEditorSaveClick() {
  if (!state.nodeTextEditorNodeId) {
    hideNodeTextEditor();
    return;
  }

  const node = state.nodes.get(state.nodeTextEditorNodeId);
  if (!node) {
    hideNodeTextEditor();
    return;
  }

  const previousTitle = node.title;
  const previousContent = node.content;
  const nextTitle = normalizeNodeTitle(nodeTextEditorTitleInput.value, node.id);
  const nextContent = isMediaMetadata(node.metadata)
    ? buildMediaNodeContent(node.metadata)
    : isBindingMetadata(node.metadata)
      ? buildBindingNodeContent(node.metadata)
    : normalizeNodeContent(nodeTextEditorContentTextarea.value);

  node.title = nextTitle;
  node.content = nextContent;
  const titleEl = node.el.querySelector(".node-title");
  const contentEl = node.el.querySelector(".node-content");
  if (titleEl) {
    titleEl.textContent = nextTitle;
  }
  if (contentEl) {
    contentEl.textContent = nextContent;
  }

  if (nextTitle !== previousTitle || nextContent !== previousContent) {
    commitHistory();
  }
  hideNodeTextEditor();
}

function onMediaContentDoubleClick(event, nodeId) {
  event.stopPropagation();
  showJsonEditor(nodeId);
}

function onNodeContentDoubleClick(event, nodeId) {
  event.stopPropagation();
  const node = state.nodes.get(nodeId);
  if (!node) {
    return;
  }
  if (isMediaMetadata(node.metadata) || isBindingMetadata(node.metadata)) {
    showJsonEditor(nodeId);
    return;
  }
  showNodeTextEditor(nodeId);
}

function onNodeDoubleClick(event, nodeId) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }
  if (
    target.closest(".handle") ||
    target.closest(".node-delete-btn") ||
    target.closest(".binding-ref-list") ||
    target.closest(".binding-toggle-list") ||
    target.closest(".binding-field-list") ||
    target.closest(".media-thumbnail") ||
    target.closest(".node-title") ||
    target.closest(".node-content")
  ) {
    return;
  }
  event.stopPropagation();
  const node = state.nodes.get(nodeId);
  if (!node) {
    return;
  }
  if (isMediaMetadata(node.metadata) || isBindingMetadata(node.metadata)) {
    showJsonEditor(nodeId);
    return;
  }
  showNodeTextEditor(nodeId);
}

function openJsonEditorFromSelection() {
  const selectedIds = Array.from(state.selectedNodeIds).filter((id) => state.nodes.has(id));
  if (selectedIds.length <= 0) {
    return;
  }
  if (selectedIds.length === 1) {
    showJsonEditor(selectedIds[0]);
    return;
  }
  showSelectionJsonEditor(selectedIds);
}

function parseJsonObjectTextSafely(text) {
  if (typeof text !== "string" || text.trim() === "") {
    return null;
  }
  try {
    const parsed = JSON.parse(text);
    return isPlainRecord(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

function resolveFunctionalParameterSortInfo(key) {
  const bindingDirective = parseBindingDirectiveKey(key);
  if (bindingDirective) {
    const sideOrder = {
      right: 0,
      top: 1,
      left: 2,
      bottom: 3,
    };
    return {
      group: 1,
      weightA: sideOrder[bindingDirective.side] ?? 99,
      weightB: bindingDirective.index,
    };
  }

  const bindingCheck = parseBindingCheckDirectiveKey(key);
  if (bindingCheck) {
    return {
      group: 2,
      weightA: bindingCheck.index,
      weightB: 0,
    };
  }

  const bindingValueControl = parseBindingValueControlField(key);
  if (bindingValueControl) {
    const kindOrder = {
      date: 0,
      text: 1,
    };
    return {
      group: 3,
      weightA: kindOrder[bindingValueControl.kind] ?? 99,
      weightB: bindingValueControl.index,
    };
  }

  return null;
}

function resolveEditorKeyGroup(key) {
  if (key === NODE_DEFINITION_TITLE_KEY || key === NODE_DEFINITION_DESCRIPTION_KEY) {
    return 0;
  }
  return resolveFunctionalParameterSortInfo(key) ? 1 : 2;
}

function reorderNodeMetaEditorKeys(source) {
  if (!isPlainRecord(source)) {
    return {};
  }
  const ordered = {};
  if (Object.prototype.hasOwnProperty.call(source, NODE_DEFINITION_TITLE_KEY)) {
    ordered[NODE_DEFINITION_TITLE_KEY] = source[NODE_DEFINITION_TITLE_KEY];
  }
  if (Object.prototype.hasOwnProperty.call(source, NODE_DEFINITION_DESCRIPTION_KEY)) {
    ordered[NODE_DEFINITION_DESCRIPTION_KEY] = source[NODE_DEFINITION_DESCRIPTION_KEY];
  }
  const functionalKeys = [];
  const regularKeys = [];
  Object.keys(source).forEach((key) => {
    if (key === NODE_DEFINITION_TITLE_KEY || key === NODE_DEFINITION_DESCRIPTION_KEY) {
      return;
    }
    if (resolveFunctionalParameterSortInfo(key)) {
      functionalKeys.push(key);
    } else {
      regularKeys.push(key);
    }
  });

  functionalKeys
    .sort((a, b) => {
      const pa = resolveFunctionalParameterSortInfo(a);
      const pb = resolveFunctionalParameterSortInfo(b);
      if (!pa || !pb) {
        return a.localeCompare(b);
      }
      if (pa.group !== pb.group) {
        return pa.group - pb.group;
      }
      if (pa.weightA !== pb.weightA) {
        return pa.weightA - pb.weightA;
      }
      if (pa.weightB !== pb.weightB) {
        return pa.weightB - pb.weightB;
      }
      return a.localeCompare(b);
    })
    .forEach((key) => {
      ordered[key] = source[key];
    });

  regularKeys.forEach((key) => {
    ordered[key] = source[key];
  });
  return ordered;
}

function stringifyEditorJsonObject(source) {
  if (!isPlainRecord(source)) {
    return JSON.stringify(source ?? {}, null, 2);
  }

  const ordered = reorderNodeMetaEditorKeys(source);
  const keys = Object.keys(ordered);
  if (keys.length === 0) {
    return "{}";
  }

  const lines = ["{"];
  keys.forEach((key, index) => {
    const valueText = JSON.stringify(ordered[key], null, 2);
    const valueLines = valueText.split("\n");
    const isLast = index === keys.length - 1;

    if (valueLines.length <= 1) {
      lines.push(`  ${JSON.stringify(key)}: ${valueLines[0]}${isLast ? "" : ","}`);
    } else {
      lines.push(`  ${JSON.stringify(key)}: ${valueLines[0]}`);
      for (let lineIndex = 1; lineIndex < valueLines.length; lineIndex += 1) {
        const isValueLastLine = lineIndex === valueLines.length - 1;
        const trailingComma = isValueLastLine && !isLast ? "," : "";
        lines.push(`  ${valueLines[lineIndex]}${trailingComma}`);
      }
    }

    if (!isLast) {
      const currentGroup = resolveEditorKeyGroup(key);
      const nextGroup = resolveEditorKeyGroup(keys[index + 1]);
      if (currentGroup !== nextGroup && nextGroup === 2) {
        lines.push("");
      }
    }
  });
  lines.push("}");
  return lines.join("\n");
}

function buildBindingMetadataFromEditorObject(sourceObject, options = {}) {
  if (!isPlainRecord(sourceObject)) {
    return null;
  }

  let nextMetadata;
  const functionalConfig = sanitizeJsonEditorFunctionalConfig(options.functionalConfig);
  if (functionalConfig) {
    const mergedValues = {
      ...(deepCloneJsonValue(sourceObject) ?? {}),
      ...(deepCloneJsonValue(functionalConfig.valueControls) ?? {}),
    };
    nextMetadata = normalizeBindingMetadata({
      schema: BINDING_NODE_SCHEMA,
      version: BINDING_NODE_VERSION,
      values: mergedValues,
      refs: functionalConfig.refs,
      checks: functionalConfig.checks,
      handleCounts: functionalConfig.handleCounts,
    });
  } else {
    const normalized = normalizeNodeMetadata(sourceObject);
    if (isBindingMetadata(normalized)) {
      nextMetadata = normalizeBindingMetadata(normalized);
    } else {
      nextMetadata = normalizeBindingMetadata({
        schema: BINDING_NODE_SCHEMA,
        version: BINDING_NODE_VERSION,
        values: sourceObject,
        refs: {},
        checks: {},
        handleCounts: createDefaultHandleCounts(),
      });
    }
  }
  if (!nextMetadata) {
    return null;
  }

  if (options.applyDescriptionOverride === true) {
    const normalizedText =
      typeof options.descriptionValue === "string"
        ? String(options.descriptionValue).replace(/\r?\n/g, " ").trim()
        : "";
    if (!isPlainRecord(nextMetadata.display)) {
      nextMetadata.display = {};
    }
    if (normalizedText) {
      nextMetadata.display.description = normalizeNodeContent(normalizedText);
    } else {
      delete nextMetadata.display.description;
    }
  }

  return nextMetadata;
}

function resolveNodeDefinitionSourceObject(node) {
  if (!node) {
    return {
      [NODE_DEFINITION_TITLE_KEY]: "節點",
      [NODE_DEFINITION_DESCRIPTION_KEY]: "拖曳節點或拉線連接",
    };
  }

  let source;
  if (isBindingMetadata(node.metadata)) {
    const text = buildEditableBindingJson(node.metadata);
    const parsed = parseJsonObjectTextSafely(text);
    source = ensureDefinitionDirectiveDefaultsObject(parsed ?? {});
  } else if (isPlainRecord(node.metadata) && !isMediaMetadata(node.metadata)) {
    source = ensureDefinitionDirectiveDefaultsObject(node.metadata);
  } else {
    const contentObject = parseJsonObjectTextSafely(node.content);
    source = contentObject ? ensureDefinitionDirectiveDefaultsObject(contentObject) : ensureDefinitionDirectiveDefaultsObject({});
  }

  if (typeof source[NODE_DEFINITION_TITLE_KEY] !== "string" || source[NODE_DEFINITION_TITLE_KEY].trim() === "") {
    source[NODE_DEFINITION_TITLE_KEY] = node.title;
  }
  if (
    typeof source[NODE_DEFINITION_DESCRIPTION_KEY] !== "string" ||
    source[NODE_DEFINITION_DESCRIPTION_KEY].trim() === ""
  ) {
    source[NODE_DEFINITION_DESCRIPTION_KEY] = node.content;
  }
  return reorderNodeMetaEditorKeys(source);
}

function openNodeDefinitionEditorForNode(nodeId) {
  const node = state.nodes.get(nodeId);
  if (!node) {
    return;
  }

  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideNodeTextEditor();
  hideMediaViewer();

  const sourceObject = resolveNodeDefinitionSourceObject(node);
  const extracted = extractJsonEditorFunctionalConfigFromObject(sourceObject, {
    seedDefaults: true,
    ensureDefaultChecks: true,
  });
  openJsonEditorWithText("node-definition", stringifyEditorJsonObject(extracted.cleanObject), {
    nodeId,
    title: "定義結點",
    functionalConfig: extracted.functionalConfig,
  });
}

function openNodeDefinitionEditorFromSelection() {
  const selectedIds = Array.from(state.selectedNodeIds).filter((id) => state.nodes.has(id));
  if (selectedIds.length <= 0) {
    return;
  }

  openNodeDefinitionEditorForNode(selectedIds[0]);
}

function onMediaThumbnailClick(event, nodeId) {
  event.stopPropagation();
  showMediaViewer(nodeId);
}

function onMediaThumbnailKeyDown(event, nodeId) {
  if (event.key !== "Enter" && event.key !== " ") {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  showMediaViewer(nodeId);
}

async function showMediaViewer(nodeId) {
  const node = state.nodes.get(nodeId);
  if (!node || !isMediaMetadata(node.metadata)) {
    return;
  }

  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideJsonEditor();
  hideNodeTextEditor();

  state.mediaViewerNodeId = nodeId;
  releaseMediaViewerObjectUrl();
  mediaViewerTitleEl.textContent = node.title || "媒體預覽";
  mediaViewerMetaEl.textContent = buildMediaViewerMeta(node.metadata);
  mediaViewerPanelEl.classList.add("visible");
  mediaViewerPanelEl.setAttribute("aria-hidden", "false");
  renderMediaViewerMessage("載入媒體中...");

  const source = await resolveMediaViewerSource(node.metadata);
  if (state.mediaViewerNodeId !== nodeId) {
    if (source?.objectUrl) {
      URL.revokeObjectURL(source.objectUrl);
    }
    return;
  }

  if (!source) {
    renderMediaViewerMessage("找不到媒體檔案。請確認專案路徑與 media 資料夾設定。");
    return;
  }

  if (source.objectUrl) {
    state.mediaViewerObjectUrl = source.objectUrl;
  }

  await renderMediaViewerContent(nodeId, node.metadata, source);
}

function hideMediaViewer() {
  if (state.mediaViewerNodeId == null && !mediaViewerPanelEl.classList.contains("visible")) {
    return;
  }
  state.mediaViewerNodeId = null;
  mediaViewerPanelEl.classList.remove("visible");
  mediaViewerPanelEl.setAttribute("aria-hidden", "true");
  mediaViewerTitleEl.textContent = "媒體預覽";
  mediaViewerMetaEl.textContent = "";
  renderMediaViewerMessage("請點擊節點中的媒體縮圖以預覽。");
  releaseMediaViewerObjectUrl();
}

function onMediaViewerPanelPointerDown(event) {
  if (event.target === mediaViewerPanelEl) {
    hideMediaViewer();
  }
}

function renderMediaViewerMessage(message) {
  mediaViewerBodyEl.textContent = "";
  const note = document.createElement("p");
  note.className = "media-viewer-empty";
  note.textContent = message;
  mediaViewerBodyEl.appendChild(note);
}

function buildMediaViewerMeta(metadata) {
  if (!metadata || typeof metadata !== "object") {
    return "";
  }
  const kindLabel = {
    image: "圖片",
    video: "影片",
    audio: "音訊",
    text: "文字",
    json: "JSON",
    binary: "檔案",
  }[metadata.file?.kind] ?? "媒體";
  const name = metadata.file?.name ?? "未命名";
  const size = metadata.file?.sizeHuman ?? "";
  return size ? `${kindLabel} · ${name} · ${size}` : `${kindLabel} · ${name}`;
}

async function resolveMediaViewerSource(metadata) {
  const fileName = resolveStorageFileNameFromManifestEntry(metadata?.storage);
  if (fileName) {
    try {
      const directory = await ensureMediaDirectory();
      const fileHandle = await directory.getFileHandle(fileName, { create: false });
      const file = await fileHandle.getFile();
      const objectUrl = URL.createObjectURL(file);
      return {
        file,
        objectUrl,
        src: objectUrl,
        mimeType: file.type || metadata?.file?.mimeType || "application/octet-stream",
      };
    } catch (error) {
      console.error("load media file for preview failed", error);
    }
  }

  const thumbnailDataUrl = metadata?.preview?.thumbnailDataUrl;
  if (typeof thumbnailDataUrl === "string" && thumbnailDataUrl.startsWith("data:image/")) {
    return {
      file: null,
      objectUrl: null,
      src: thumbnailDataUrl,
      mimeType: metadata?.preview?.mimeType || "image/png",
      fromThumbnail: true,
    };
  }

  return null;
}

function parseFrameRateValue(value) {
  if (Number.isFinite(value) && value > 0) {
    return Number(value);
  }
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed === "") {
    return null;
  }

  const ratioMatch = /^(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)$/.exec(trimmed);
  if (ratioMatch) {
    const numerator = Number(ratioMatch[1]);
    const denominator = Number(ratioMatch[2]);
    if (Number.isFinite(numerator) && Number.isFinite(denominator) && denominator > 0) {
      const fps = numerator / denominator;
      return fps > 0 ? fps : null;
    }
    return null;
  }

  const numeric = Number(trimmed);
  if (Number.isFinite(numeric) && numeric > 0) {
    return numeric;
  }

  return null;
}

function resolveTimecodeFps(metadata) {
  const candidates = [
    metadata?.analysis?.frameRate,
    metadata?.analysis?.fps,
    metadata?.file?.frameRate,
    metadata?.file?.fps,
  ];

  for (const candidate of candidates) {
    const parsed = parseFrameRateValue(candidate);
    if (parsed) {
      const rounded = Math.round(parsed);
      if (rounded >= 1 && rounded <= 240) {
        return rounded;
      }
    }
  }

  return DEFAULT_TIMECODE_FPS;
}

function formatTimecodeHhMmSsFf(seconds, fps) {
  const safeFps = Math.max(1, Math.floor(Number(fps) || DEFAULT_TIMECODE_FPS));
  const safeSeconds = Number.isFinite(seconds) && seconds >= 0 ? seconds : 0;
  const totalFrames = Math.max(0, Math.floor(safeSeconds * safeFps));

  const framesPerHour = safeFps * 3600;
  const framesPerMinute = safeFps * 60;
  const hours = Math.floor(totalFrames / framesPerHour);
  const remainAfterHours = totalFrames % framesPerHour;
  const minutes = Math.floor(remainAfterHours / framesPerMinute);
  const remainAfterMinutes = remainAfterHours % framesPerMinute;
  const secs = Math.floor(remainAfterMinutes / safeFps);
  const frames = remainAfterMinutes % safeFps;

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(secs).padStart(2, "0")}:${String(frames).padStart(2, "0")}`;
}

function createMediaTimecodePanel(mediaEl, metadata, options = {}) {
  const { allowFullscreen = false, fullscreenTarget = mediaEl } = options;
  const fps = resolveTimecodeFps(metadata);
  const panel = document.createElement("div");
  panel.className = "media-viewer-timecode";

  const playBtn = document.createElement("button");
  playBtn.type = "button";
  playBtn.className = "media-viewer-control-btn";

  const prevFrameBtn = document.createElement("button");
  prevFrameBtn.type = "button";
  prevFrameBtn.className = "media-viewer-control-btn ghost frame-step";
  prevFrameBtn.textContent = "◀";
  prevFrameBtn.title = "後退一幀";
  prevFrameBtn.setAttribute("aria-label", "後退一幀");

  const nextFrameBtn = document.createElement("button");
  nextFrameBtn.type = "button";
  nextFrameBtn.className = "media-viewer-control-btn ghost frame-step";
  nextFrameBtn.textContent = "▶";
  nextFrameBtn.title = "前進一幀";
  nextFrameBtn.setAttribute("aria-label", "前進一幀");

  const seek = document.createElement("input");
  seek.type = "range";
  seek.className = "media-viewer-seek";
  seek.min = "0";
  seek.max = "0";
  seek.step = "0.01";
  seek.value = "0";

  const muteBtn = document.createElement("button");
  muteBtn.type = "button";
  muteBtn.className = "media-viewer-control-btn ghost";

  const rateSelect = document.createElement("select");
  rateSelect.className = "media-viewer-rate-select";
  rateSelect.setAttribute("aria-label", "播放倍率");
  const rateOptions = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3];
  rateOptions.forEach((rate) => {
    const option = document.createElement("option");
    option.value = String(rate);
    option.textContent = `${rate}x`;
    if (rate === 1) {
      option.selected = true;
    }
    rateSelect.appendChild(option);
  });

  const label = document.createElement("span");
  label.className = "media-viewer-timecode-label";
  label.textContent = `${fps}fps`;

  const value = document.createElement("span");
  value.className = "media-viewer-timecode-value";

  const formatDuration = () => {
    const duration = Number(mediaEl.duration);
    if (!Number.isFinite(duration) || duration < 0) {
      return "--:--:--:--";
    }
    return formatTimecodeHhMmSsFf(duration, fps);
  };

  const updateValue = () => {
    const current = Number(mediaEl.currentTime);
    const currentText = formatTimecodeHhMmSsFf(current, fps);
    value.textContent = `${currentText} / ${formatDuration()}`;
  };

  const updateButtons = () => {
    playBtn.textContent = mediaEl.paused ? "播放" : "暫停";
    muteBtn.textContent = mediaEl.muted || Number(mediaEl.volume) <= 0 ? "靜音中" : "有聲";
  };

  const updateRateSelect = () => {
    const currentRate = Number(mediaEl.playbackRate);
    const matched = rateOptions.find((rate) => Math.abs(rate - currentRate) < 0.001);
    if (matched != null) {
      rateSelect.value = String(matched);
      return;
    }
    const rounded = Number.isFinite(currentRate) && currentRate > 0 ? Number(currentRate.toFixed(2)) : 1;
    const fallback = String(rounded);
    if (!Array.from(rateSelect.options).some((option) => option.value === fallback)) {
      const extra = document.createElement("option");
      extra.value = fallback;
      extra.textContent = `${fallback}x`;
      rateSelect.appendChild(extra);
    }
    rateSelect.value = fallback;
  };

  const updateSeek = () => {
    const duration = Number(mediaEl.duration);
    const current = Number(mediaEl.currentTime);
    if (!Number.isFinite(duration) || duration <= 0) {
      seek.disabled = true;
      seek.max = "0";
      seek.value = "0";
      return;
    }
    seek.disabled = false;
    seek.max = String(duration);
    const safeCurrent = Number.isFinite(current) ? clamp(current, 0, duration) : 0;
    seek.value = String(safeCurrent);
  };

  const stepByFrame = (frameDelta) => {
    const duration = Number(mediaEl.duration);
    const current = Number(mediaEl.currentTime);
    const safeCurrent = Number.isFinite(current) ? current : 0;
    const stepSec = frameDelta / Math.max(1, fps);
    const target = safeCurrent + stepSec;
    const maxTime = Number.isFinite(duration) && duration >= 0 ? duration : Number.POSITIVE_INFINITY;
    mediaEl.currentTime = clamp(target, 0, maxTime);
    updateSeek();
    updateValue();
  };

  let rafId = 0;
  const stopRaf = () => {
    if (rafId !== 0) {
      cancelAnimationFrame(rafId);
      rafId = 0;
    }
  };

  const tick = () => {
    updateSeek();
    updateValue();
    if (!mediaEl.paused && !mediaEl.ended) {
      rafId = requestAnimationFrame(tick);
    } else {
      rafId = 0;
    }
  };

  const startRaf = () => {
    if (rafId !== 0) {
      return;
    }
    rafId = requestAnimationFrame(tick);
  };

  playBtn.addEventListener("click", async () => {
    try {
      if (mediaEl.paused) {
        await mediaEl.play();
      } else {
        mediaEl.pause();
      }
    } catch {
      // ignore playback rejection
    }
    updateButtons();
    updateSeek();
    updateValue();
  });

  prevFrameBtn.addEventListener("click", () => {
    mediaEl.pause();
    stepByFrame(-1);
  });

  nextFrameBtn.addEventListener("click", () => {
    mediaEl.pause();
    stepByFrame(1);
  });

  seek.addEventListener("input", () => {
    const next = Number(seek.value);
    if (Number.isFinite(next)) {
      mediaEl.currentTime = next;
    }
    updateValue();
  });
  seek.addEventListener("change", () => {
    const next = Number(seek.value);
    if (Number.isFinite(next)) {
      mediaEl.currentTime = next;
    }
    updateSeek();
    updateValue();
  });

  muteBtn.addEventListener("click", () => {
    mediaEl.muted = !mediaEl.muted;
    updateButtons();
  });

  rateSelect.addEventListener("change", () => {
    const nextRate = Number(rateSelect.value);
    if (!Number.isFinite(nextRate) || nextRate <= 0) {
      return;
    }
    mediaEl.playbackRate = nextRate;
    updateRateSelect();
  });

  panel.appendChild(playBtn);
  panel.appendChild(prevFrameBtn);
  panel.appendChild(nextFrameBtn);
  panel.appendChild(seek);
  panel.appendChild(value);
  panel.appendChild(muteBtn);
  panel.appendChild(rateSelect);

  if (allowFullscreen) {
    const fullscreenBtn = document.createElement("button");
    fullscreenBtn.type = "button";
    fullscreenBtn.className = "media-viewer-control-btn ghost";
    fullscreenBtn.textContent = "全螢幕";
    fullscreenBtn.addEventListener("click", async () => {
      if (!fullscreenTarget || typeof fullscreenTarget.requestFullscreen !== "function") {
        return;
      }
      try {
        if (document.fullscreenElement === fullscreenTarget) {
          await document.exitFullscreen();
        } else if (!document.fullscreenElement) {
          await fullscreenTarget.requestFullscreen();
        } else {
          await document.exitFullscreen();
          await fullscreenTarget.requestFullscreen();
        }
      } catch {
        // ignore fullscreen errors
      }
    });
    panel.appendChild(fullscreenBtn);
  }

  panel.appendChild(label);

  mediaEl.addEventListener("loadedmetadata", () => {
    updateButtons();
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("durationchange", () => {
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("timeupdate", () => {
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("seeking", () => {
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("seeked", () => {
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("ratechange", updateValue);
  mediaEl.addEventListener("ratechange", updateRateSelect);
  mediaEl.addEventListener("play", () => {
    updateButtons();
    startRaf();
  });
  mediaEl.addEventListener("pause", () => {
    stopRaf();
    updateButtons();
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("ended", () => {
    stopRaf();
    updateButtons();
    updateSeek();
    updateValue();
  });
  mediaEl.addEventListener("emptied", stopRaf);
  mediaEl.addEventListener("volumechange", updateButtons);

  updateButtons();
  updateSeek();
  updateValue();
  updateRateSelect();
  return panel;
}

async function renderMediaViewerContent(nodeId, metadata, source) {
  const kind = metadata?.file?.kind ?? "binary";
  mediaViewerBodyEl.textContent = "";

  if ((kind === "text" || kind === "json") && source.file) {
    let text = "";
    try {
      text = await source.file.text();
    } catch {
      text = "";
    }
    if (state.mediaViewerNodeId !== nodeId) {
      return;
    }
    if (!text && typeof metadata?.analysis?.preview === "string") {
      text = metadata.analysis.preview;
    }
    if (!text) {
      renderMediaViewerMessage("此文字媒體無法讀取內容。");
      return;
    }
    const pre = document.createElement("pre");
    pre.className = "media-viewer-text";
    pre.textContent = text;
    mediaViewerBodyEl.appendChild(pre);
    appendMediaViewerActions(source.src, metadata?.file?.name);
    return;
  }

  if (kind === "image") {
    const image = document.createElement("img");
    image.className = "media-viewer-image";
    image.src = source.src;
    image.alt = metadata?.file?.name || "media";
    image.draggable = false;
    mediaViewerBodyEl.appendChild(image);
    appendMediaViewerActions(source.src, metadata?.file?.name);
    return;
  }

  if (kind === "video") {
    const videoWrap = document.createElement("div");
    videoWrap.className = "media-viewer-player-wrap";
    const video = document.createElement("video");
    video.className = "media-viewer-video";
    video.controls = false;
    video.playsInline = true;
    video.preload = "metadata";
    video.src = source.src;
    video.addEventListener("click", async () => {
      try {
        if (video.paused) {
          await video.play();
        } else {
          video.pause();
        }
      } catch {
        // ignore playback rejection
      }
    });
    videoWrap.appendChild(video);
    videoWrap.appendChild(createMediaTimecodePanel(video, metadata, { allowFullscreen: true, fullscreenTarget: videoWrap }));
    mediaViewerBodyEl.appendChild(videoWrap);
    appendMediaViewerActions(source.src, metadata?.file?.name);
    return;
  }

  if (kind === "audio") {
    const audioWrap = document.createElement("div");
    audioWrap.className = "media-viewer-audio-wrap";
    const audio = document.createElement("audio");
    audio.className = "media-viewer-audio";
    audio.controls = false;
    audio.preload = "metadata";
    audio.src = source.src;
    audioWrap.appendChild(audio);
    audioWrap.appendChild(createMediaTimecodePanel(audio, metadata));
    mediaViewerBodyEl.appendChild(audioWrap);
    appendMediaViewerActions(source.src, metadata?.file?.name);
    return;
  }

  if (String(source.mimeType || "").startsWith("application/pdf")) {
    const frame = document.createElement("iframe");
    frame.className = "media-viewer-frame";
    frame.src = source.src;
    frame.title = metadata?.file?.name || "PDF";
    mediaViewerBodyEl.appendChild(frame);
    appendMediaViewerActions(source.src, metadata?.file?.name);
    return;
  }

  if (source.fromThumbnail) {
    const fallback = document.createElement("img");
    fallback.className = "media-viewer-image";
    fallback.src = source.src;
    fallback.alt = metadata?.file?.name || "media thumbnail";
    fallback.draggable = false;
    mediaViewerBodyEl.appendChild(fallback);
    return;
  }

  renderMediaViewerMessage("此類型目前無法直接內嵌預覽，可用下方按鈕開啟。");
  appendMediaViewerActions(source.src, metadata?.file?.name);
}

function appendMediaViewerActions(src, fileName) {
  if (typeof src !== "string" || src.trim() === "") {
    return;
  }
  const actions = document.createElement("div");
  actions.className = "media-viewer-actions";

  const link = document.createElement("a");
  link.className = "media-viewer-open-link";
  link.href = src;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  if (fileName) {
    link.download = fileName;
  }
  link.textContent = "在新分頁開啟";

  actions.appendChild(link);
  mediaViewerBodyEl.appendChild(actions);
}

function releaseMediaViewerObjectUrl() {
  if (state.mediaViewerObjectUrl) {
    URL.revokeObjectURL(state.mediaViewerObjectUrl);
    state.mediaViewerObjectUrl = null;
  }
}

function showJsonEditor(nodeId) {
  const node = state.nodes.get(nodeId);
  if (!node) {
    return;
  }

  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideNodeTextEditor();
  hideMediaViewer();

  if (isMediaMetadata(node.metadata)) {
    const { text, thumbnailDataUrl } = buildEditableMediaJson(node.metadata);
    openJsonEditorWithText("media", text, {
      nodeId,
      originalThumbnailDataUrl: thumbnailDataUrl,
      title: "筆記資料編輯",
    });
    return;
  }

  if (isBindingMetadata(node.metadata)) {
    const normalizedBinding = normalizeBindingMetadata(node.metadata);
    if (!normalizedBinding) {
      openJsonEditorWithText("binding", "{}", {
        nodeId,
        title: "筆記資料編輯",
      });
      return;
    }
    const values = deepCloneJsonValue(normalizedBinding.values);
    const editable = isPlainRecord(values) ? values : {};
    editable[NODE_DEFINITION_TITLE_KEY] = node.title;
    editable[NODE_DEFINITION_DESCRIPTION_KEY] =
      typeof normalizedBinding?.display?.description === "string" ? normalizedBinding.display.description : "";
    const functionalConfig = createJsonEditorFunctionalConfigFromBindingMetadata(normalizedBinding);
    const extracted = extractJsonEditorFunctionalConfigFromObject(editable);
    openJsonEditorWithText(
      "binding",
      stringifyEditorJsonObject(extracted.cleanObject),
      {
        nodeId,
        title: "筆記資料編輯",
        functionalConfig: sanitizeJsonEditorFunctionalConfig(functionalConfig) ?? extracted.functionalConfig,
      }
    );
    return;
  }

  const genericSource = isPlainRecord(node.metadata)
    ? deepCloneJsonValue(node.metadata) ?? {}
    : parseJsonObjectTextSafely(node.content) ?? {};
  openJsonEditorWithText("node-json", stringifyEditorJsonObject(genericSource), {
    nodeId,
    title: "筆記資料編輯",
  });
}

function showSelectionJsonEditor(nodeIds) {
  const orderedIds = Array.from(new Set(nodeIds))
    .map((id) => String(id))
    .filter((id) => state.nodes.has(id))
    .sort((a, b) => Number(a) - Number(b));

  if (orderedIds.length <= 1) {
    return;
  }

  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideNodeTextEditor();
  hideMediaViewer();

  const payload = buildSelectionJsonEditorPayload(orderedIds);
  openJsonEditorWithText("selection", JSON.stringify(payload, null, 2), {
    selectionNodeIds: orderedIds,
    title: "多筆記 JSON 編輯",
  });
}

function openJsonEditorWithText(mode, text, options = {}) {
  state.jsonEditorMode = mode;
  state.jsonEditorNodeId = options.nodeId ?? null;
  state.jsonEditorSelectionNodeIds = Array.isArray(options.selectionNodeIds)
    ? options.selectionNodeIds.map((id) => String(id))
    : [];
  state.jsonEditorOriginalThumbnailDataUrl = options.originalThumbnailDataUrl ?? null;
  state.jsonEditorFunctionalConfig = sanitizeJsonEditorFunctionalConfig(options.functionalConfig);
  if (jsonEditorTitleEl) {
    jsonEditorTitleEl.textContent = options.title ?? "筆記資料編輯";
  }
  jsonEditorTextarea.value = text;
  renderJsonEditorFunctionalControls();
  jsonEditorPanelEl.classList.add("visible");
  jsonEditorPanelEl.setAttribute("aria-hidden", "false");
  jsonEditorTextarea.focus();
  jsonEditorTextarea.setSelectionRange(0, 0);
}

function hideJsonEditor() {
  state.jsonEditorMode = null;
  state.jsonEditorNodeId = null;
  state.jsonEditorSelectionNodeIds = [];
  state.jsonEditorOriginalThumbnailDataUrl = null;
  state.jsonEditorFunctionalConfig = null;
  jsonEditorPanelEl.classList.remove("visible");
  jsonEditorPanelEl.setAttribute("aria-hidden", "true");
  jsonEditorTextarea.value = "";
  renderJsonEditorFunctionalControls();
  if (jsonEditorTitleEl) {
    jsonEditorTitleEl.textContent = "筆記資料編輯";
  }
}

function onJsonEditorPanelPointerDown(event) {
  if (event.target === jsonEditorPanelEl) {
    hideJsonEditor();
  }
}

function onJsonEditorTextareaKeyDown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    hideJsonEditor();
    return;
  }
  const mod = event.ctrlKey || event.metaKey;
  if (mod && event.key.toLowerCase() === "s") {
    event.preventDefault();
    onJsonEditorSaveClick();
  }
}

function onJsonEditorSaveClick() {
  if (state.jsonEditorMode === "selection") {
    applySelectionJsonEditorSave();
    return;
  }

  if (state.jsonEditorMode === "node-json") {
    applyNodeJsonEditorSave();
    return;
  }

  if (state.jsonEditorMode === "media") {
    applyMediaJsonEditorSave();
    return;
  }

  if (state.jsonEditorMode === "binding") {
    applyBindingJsonEditorSave();
    return;
  }

  if (state.jsonEditorMode === "node-definition") {
    applyNodeDefinitionJsonEditorSave();
    return;
  }

  hideJsonEditor();
}

function applyNodeJsonEditorSave() {
  if (!state.jsonEditorNodeId) {
    hideJsonEditor();
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonEditorTextarea.value);
  } catch {
    alert("JSON 格式錯誤，請修正後再儲存。");
    jsonEditorTextarea.focus();
    return;
  }

  if (!isPlainRecord(parsed)) {
    alert("JSON 內容必須是物件。");
    jsonEditorTextarea.focus();
    return;
  }

  const node = state.nodes.get(state.jsonEditorNodeId);
  if (!node) {
    hideJsonEditor();
    return;
  }

  const shouldSyncGenericContent =
    !isPlainRecord(node.metadata) || parseJsonObjectTextSafely(node.content) != null || node.title === "未定義結點";
  const nextMetadata = normalizeNodeMetadata(parsed);
  node.metadata = nextMetadata;

  if (isMediaMetadata(nextMetadata)) {
    node.content = buildMediaNodeContent(nextMetadata);
  } else if (isBindingMetadata(nextMetadata)) {
    node.content = buildBindingNodeContent(nextMetadata);
  } else if (shouldSyncGenericContent) {
    node.content = normalizeNodeContent(buildJsonNodeSummaryText(parsed));
  } else {
    node.content = normalizeNodeContent(node.content);
  }

  const contentEl = node.el.querySelector(".node-content");
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute(
      "title",
      isMediaMetadata(nextMetadata) ? "雙擊編輯 JSON" : isBindingMetadata(nextMetadata) ? "雙擊編輯綁定 JSON" : "雙擊編輯文字"
    );
  }
  renderNodeHandles(node);
  renderNodeMediaPreview(node);
  renderBindingToggleControls(node);
  refreshBindingNodes();
  renderConnections();
  commitHistory();
  hideJsonEditor();
}

function applyMediaJsonEditorSave() {
  if (!state.jsonEditorNodeId) {
    hideJsonEditor();
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonEditorTextarea.value);
  } catch {
    alert("JSON 格式錯誤，請修正後再儲存。");
    jsonEditorTextarea.focus();
    return;
  }

  if (!parsed || typeof parsed !== "object") {
    alert("JSON 內容必須是物件。");
    jsonEditorTextarea.focus();
    return;
  }

  if (
    state.jsonEditorOriginalThumbnailDataUrl &&
    parsed.preview &&
    parsed.preview.thumbnailDataUrl === MEDIA_THUMBNAIL_TOKEN
  ) {
    parsed.preview.thumbnailDataUrl = state.jsonEditorOriginalThumbnailDataUrl;
  }

  const node = state.nodes.get(state.jsonEditorNodeId);
  if (!node) {
    hideJsonEditor();
    return;
  }

  const nextMetadata = normalizeNodeMetadata(parsed);
  node.metadata = nextMetadata;
  if (isMediaMetadata(nextMetadata)) {
    node.content = buildMediaNodeContent(nextMetadata);
  } else if (isBindingMetadata(nextMetadata)) {
    node.content = buildBindingNodeContent(nextMetadata);
  } else {
    node.content = normalizeNodeContent(node.content);
  }
  const contentEl = node.el.querySelector(".node-content");
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute(
      "title",
      isMediaMetadata(nextMetadata) ? "雙擊編輯 JSON" : isBindingMetadata(nextMetadata) ? "雙擊編輯綁定 JSON" : "雙擊編輯文字"
    );
  }
  renderNodeHandles(node);
  renderNodeMediaPreview(node);
  renderBindingToggleControls(node);
  refreshBindingNodes();
  commitHistory();
  hideJsonEditor();
}

function applyBindingJsonEditorSave() {
  if (!state.jsonEditorNodeId) {
    hideJsonEditor();
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonEditorTextarea.value);
  } catch {
    alert("JSON 格式錯誤，請修正後再儲存。");
    jsonEditorTextarea.focus();
    return;
  }

  if (!isPlainRecord(parsed)) {
    alert("JSON 內容必須是物件。");
    jsonEditorTextarea.focus();
    return;
  }
  const extracted = extractJsonEditorFunctionalConfigFromObject(parsed);
  const parsedWithoutFunctional = ensureDefinitionDirectiveDefaultsObject(extracted.cleanObject);

  const node = state.nodes.get(state.jsonEditorNodeId);
  if (!node) {
    hideJsonEditor();
    return;
  }

  const editedTitle =
    typeof parsedWithoutFunctional[NODE_DEFINITION_TITLE_KEY] === "string"
      ? normalizeNodeTitle(parsedWithoutFunctional[NODE_DEFINITION_TITLE_KEY], node.id)
      : node.title;
  const hasDescriptionOverride = Object.prototype.hasOwnProperty.call(
    parsedWithoutFunctional,
    NODE_DEFINITION_DESCRIPTION_KEY
  );
  const descriptionValue = hasDescriptionOverride ? parsedWithoutFunctional[NODE_DEFINITION_DESCRIPTION_KEY] : null;

  const metadataSource = deepCloneJsonValue(parsedWithoutFunctional) ?? {};
  if (!isPlainRecord(metadataSource)) {
    alert("JSON 內容必須是物件。");
    jsonEditorTextarea.focus();
    return;
  }
  delete metadataSource[NODE_DEFINITION_TITLE_KEY];
  delete metadataSource[NODE_DEFINITION_DESCRIPTION_KEY];
  const functionalConfig =
    sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig) ??
    extracted.functionalConfig;

  const nextMetadata = buildBindingMetadataFromEditorObject(metadataSource, {
    applyDescriptionOverride: hasDescriptionOverride,
    descriptionValue,
    functionalConfig,
  });
  if (!nextMetadata) {
    alert("綁定節點 JSON 內容無效。");
    jsonEditorTextarea.focus();
    return;
  }

  node.metadata = nextMetadata;
  node.title = editedTitle;
  node.content = buildBindingNodeContent(nextMetadata);
  const titleEl = node.el.querySelector(".node-title");
  const contentEl = node.el.querySelector(".node-content");
  if (titleEl) {
    titleEl.textContent = node.title;
  }
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute("title", "雙擊編輯綁定 JSON");
  }
  renderNodeHandles(node);
  renderNodeMediaPreview(node);
  renderBindingToggleControls(node);
  refreshBindingNodes();
  renderConnections();
  commitHistory();
  hideJsonEditor();
}

function applyNodeDefinitionJsonEditorSave() {
  if (!state.jsonEditorNodeId) {
    hideJsonEditor();
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonEditorTextarea.value);
  } catch {
    alert("JSON 格式錯誤，請修正後再儲存。");
    jsonEditorTextarea.focus();
    return;
  }

  if (!isPlainRecord(parsed)) {
    alert("JSON 內容必須是物件。");
    jsonEditorTextarea.focus();
    return;
  }
  const extracted = extractJsonEditorFunctionalConfigFromObject(parsed);
  const sourceWithoutFunctional = ensureDefinitionDirectiveDefaultsObject(extracted.cleanObject);
  const functionalConfig =
    sanitizeJsonEditorFunctionalConfig(state.jsonEditorFunctionalConfig) ??
    extracted.functionalConfig;
  const definitionTemplate = mergeJsonEditorFunctionalConfigIntoObject(sourceWithoutFunctional, functionalConfig, {
    includeEmptyRefs: true,
    includeEmptyChecks: true,
  });
  definitionTemplate[NODE_DEFINITION_TITLE_KEY] = DEFINED_NODE_TITLE;
  const definitionMatch = extractNodeDefinitionMatchFromTemplate(definitionTemplate);
  if (definitionMatch.requiredKeys.length === 0 && !definitionMatch.typeField) {
    alert("請至少保留一個一般資料欄位（非 node_Right_1/node_Left_1/node_Top_1/node_Bottom_1/check）作為格式匹配依據。");
    jsonEditorTextarea.focus();
    return;
  }

  const node = state.nodes.get(state.jsonEditorNodeId);
  if (!node) {
    hideJsonEditor();
    return;
  }

  const editedTitle = DEFINED_NODE_TITLE;
  const editedDescription =
    typeof definitionTemplate[NODE_DEFINITION_DESCRIPTION_KEY] === "string"
      ? normalizeNodeContent(definitionTemplate[NODE_DEFINITION_DESCRIPTION_KEY])
      : node.content;
  const metadataSource = deepCloneJsonValue(sourceWithoutFunctional) ?? {};
  if (!isPlainRecord(metadataSource)) {
    alert("JSON 內容必須是物件。");
    jsonEditorTextarea.focus();
    return;
  }
  delete metadataSource[NODE_DEFINITION_TITLE_KEY];
  delete metadataSource[NODE_DEFINITION_DESCRIPTION_KEY];

  const normalizedFunctional = sanitizeJsonEditorFunctionalConfig(functionalConfig);
  const normalizedNodeMetadata = normalizeNodeMetadata(metadataSource);
  const nextMetadata = normalizedFunctional
    ? buildBindingMetadataFromEditorObject(metadataSource, {
        applyDescriptionOverride: true,
        descriptionValue: editedDescription,
        functionalConfig: normalizedFunctional,
      })
    : normalizedNodeMetadata;

  const definition = {
    name: editedTitle || DEFINED_NODE_TITLE,
    description: editedDescription,
    match: definitionMatch,
    template: definitionTemplate,
  };
  upsertNodeDefinition(definition);

  node.metadata = nextMetadata ?? null;
  node.title = editedTitle;
  node.content = isBindingMetadata(nextMetadata) ? buildBindingNodeContent(nextMetadata) : editedDescription;
  const titleEl = node.el.querySelector(".node-title");
  const contentEl = node.el.querySelector(".node-content");
  if (titleEl) {
    titleEl.textContent = node.title;
  }
  if (contentEl) {
    contentEl.textContent = node.content;
    contentEl.setAttribute(
      "title",
      isBindingMetadata(nextMetadata) ? "雙擊編輯綁定 JSON" : "雙擊編輯文字"
    );
  }
  renderNodeHandles(node);
  renderNodeMediaPreview(node);
  renderBindingToggleControls(node);
  refreshBindingNodes();
  renderConnections();
  commitHistory();
  hideJsonEditor();
}

function buildSelectionJsonEditorPayload(nodeIds) {
  const orderedIds = Array.from(new Set(nodeIds))
    .map((id) => String(id))
    .filter((id) => state.nodes.has(id))
    .sort((a, b) => Number(a) - Number(b));
  const selectedSet = new Set(orderedIds);

  const nodes = orderedIds.map((id) => {
    const node = state.nodes.get(id);
    return {
      id: node.id,
      x: node.x,
      y: node.y,
      title: node.title,
      content: node.content,
      metadata: deepCloneJsonValue(node.metadata) ?? null,
    };
  });

  const connections = state.connections
    .filter((connection) => selectedSet.has(connection.fromNodeId) && selectedSet.has(connection.toNodeId))
    .map((connection) => ({
      id: connection.id,
      fromNodeId: connection.fromNodeId,
      fromSide: connection.fromSide,
      toNodeId: connection.toNodeId,
      toSide: connection.toSide,
    }));

  return {
    type: SELECTION_JSON_EDIT_SCHEMA,
    version: SELECTION_JSON_EDIT_VERSION,
    nodes,
    connections,
  };
}

function parseSelectionJsonEditorPayload(payload, expectedNodeIds) {
  if (!payload || typeof payload !== "object") {
    return { error: "JSON 內容必須是物件。" };
  }

  const rawNodes = Array.isArray(payload.nodes) ? payload.nodes : null;
  if (!rawNodes) {
    return { error: "必須包含 nodes 陣列。" };
  }

  const expectedSet = new Set(expectedNodeIds.map((id) => String(id)));
  const updates = new Map();

  for (const rawNode of rawNodes) {
    if (!rawNode || typeof rawNode !== "object") {
      return { error: "nodes 內的每個項目都必須是物件。" };
    }
    const nodeId = String(rawNode.id ?? "");
    if (!expectedSet.has(nodeId)) {
      return { error: `nodes 內含未選取的節點 id：${nodeId}` };
    }
    if (updates.has(nodeId)) {
      return { error: `nodes 有重複的節點 id：${nodeId}` };
    }

    const hasMetadata = Object.prototype.hasOwnProperty.call(rawNode, "metadata");
    if (hasMetadata && rawNode.metadata !== null && !isPlainRecord(rawNode.metadata)) {
      return { error: `節點 ${nodeId} 的 metadata 必須是物件或 null。` };
    }

    updates.set(nodeId, {
      title: typeof rawNode.title === "string" ? rawNode.title : null,
      content: typeof rawNode.content === "string" ? rawNode.content : null,
      x: Number.isFinite(rawNode.x) ? Number(rawNode.x) : null,
      y: Number.isFinite(rawNode.y) ? Number(rawNode.y) : null,
      hasMetadata,
      metadata: hasMetadata ? normalizeNodeMetadata(rawNode.metadata) : undefined,
    });
  }

  if (updates.size !== expectedSet.size) {
    return { error: "nodes 必須完整包含目前選取的所有節點。" };
  }

  const rawConnections = Array.isArray(payload.connections) ? payload.connections : null;
  if (!rawConnections) {
    return { error: "必須包含 connections 陣列。" };
  }

  const connections = [];
  for (const rawConnection of rawConnections) {
    if (!rawConnection || typeof rawConnection !== "object") {
      return { error: "connections 內的每個項目都必須是物件。" };
    }

    const fromNodeId = String(rawConnection.fromNodeId ?? rawConnection.fromId ?? "");
    const toNodeId = String(rawConnection.toNodeId ?? rawConnection.toId ?? "");
    const fromSide = normalizeHandleSideKey(rawConnection.fromSide);
    const toSide = normalizeHandleSideKey(rawConnection.toSide);

    if (!expectedSet.has(fromNodeId) || !expectedSet.has(toNodeId)) {
      return { error: "connections 僅能使用目前選取的節點。" };
    }
    if (!fromSide || !toSide) {
      return { error: "connections 的 fromSide / toSide 格式無效。" };
    }
    if (fromNodeId === toNodeId) {
      return { error: "connections 不允許節點連到自己。" };
    }

    const duplicated = connections.some((connection) => {
      const sameForward =
        connection.fromNodeId === fromNodeId &&
        connection.fromSide === fromSide &&
        connection.toNodeId === toNodeId &&
        connection.toSide === toSide;
      const sameReverse =
        connection.fromNodeId === toNodeId &&
        connection.fromSide === toSide &&
        connection.toNodeId === fromNodeId &&
        connection.toSide === fromSide;
      return sameForward || sameReverse;
    });
    if (duplicated) {
      return { error: "connections 有重複連線。" };
    }

    connections.push({
      fromNodeId,
      fromSide,
      toNodeId,
      toSide,
    });
  }

  return { updates, connections };
}

function validateSelectionConnectionsAgainstProjectedHandles(expectedNodeIds, updates, connections) {
  const handleSetByNodeId = new Map();

  expectedNodeIds.forEach((nodeId) => {
    const node = state.nodes.get(nodeId);
    if (!node) {
      return;
    }
    const update = updates.get(nodeId);
    const metadata = update?.hasMetadata ? update.metadata : node.metadata;
    const counts = resolveNodeHandleCounts(metadata);
    handleSetByNodeId.set(nodeId, new Set(buildHandleSideKeysFromCounts(counts)));
  });

  for (const connection of connections) {
    const fromSet = handleSetByNodeId.get(connection.fromNodeId);
    const toSet = handleSetByNodeId.get(connection.toNodeId);
    if (!fromSet?.has(connection.fromSide)) {
      return `連線端點不存在：${connection.fromNodeId}.${connection.fromSide}`;
    }
    if (!toSet?.has(connection.toSide)) {
      return `連線端點不存在：${connection.toNodeId}.${connection.toSide}`;
    }
  }

  return null;
}

function applySelectionJsonEditorSave() {
  const expectedNodeIds = state.jsonEditorSelectionNodeIds.filter((id) => state.nodes.has(id));
  if (expectedNodeIds.length <= 1) {
    hideJsonEditor();
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonEditorTextarea.value);
  } catch {
    alert("JSON 格式錯誤，請修正後再儲存。");
    jsonEditorTextarea.focus();
    return;
  }

  const normalized = parseSelectionJsonEditorPayload(parsed, expectedNodeIds);
  if (normalized.error) {
    alert(normalized.error);
    jsonEditorTextarea.focus();
    return;
  }

  const handleError = validateSelectionConnectionsAgainstProjectedHandles(
    expectedNodeIds,
    normalized.updates,
    normalized.connections
  );
  if (handleError) {
    alert(handleError);
    jsonEditorTextarea.focus();
    return;
  }

  expectedNodeIds.forEach((nodeId) => {
    const node = state.nodes.get(nodeId);
    const update = normalized.updates.get(nodeId);
    if (!node || !update) {
      return;
    }

    const nextTitle = update.title != null ? normalizeNodeTitle(update.title, node.id) : node.title;
    const nextMetadata = update.hasMetadata ? update.metadata : node.metadata;
    let nextContent = node.content;
    if (isMediaMetadata(nextMetadata)) {
      nextContent = buildMediaNodeContent(nextMetadata);
    } else if (isBindingMetadata(nextMetadata)) {
      nextContent = buildBindingNodeContent(nextMetadata);
    } else {
      nextContent = normalizeNodeContent(update.content != null ? update.content : node.content);
    }

    node.title = nextTitle;
    node.metadata = nextMetadata;
    node.content = nextContent;
    if (update.x != null) {
      node.x = update.x;
    }
    if (update.y != null) {
      node.y = update.y;
    }

    const titleEl = node.el.querySelector(".node-title");
    const contentEl = node.el.querySelector(".node-content");
    if (titleEl) {
      titleEl.textContent = node.title;
    }
    if (contentEl) {
      contentEl.textContent = node.content;
      contentEl.setAttribute(
        "title",
        isMediaMetadata(nextMetadata)
          ? "雙擊編輯 JSON"
          : isBindingMetadata(nextMetadata)
            ? "雙擊編輯綁定 JSON"
            : "雙擊編輯文字"
      );
    }

    renderNodeHandles(node);
    renderNodeMediaPreview(node);
    renderBindingToggleControls(node);
    applyNodePosition(node);
  });

  const selectedSet = new Set(expectedNodeIds);
  const preservedConnections = state.connections.filter(
    (connection) => !(selectedSet.has(connection.fromNodeId) && selectedSet.has(connection.toNodeId))
  );
  const editedConnections = normalized.connections.map((connection) => ({
    id: String(state.nextConnectionId++),
    ...connection,
  }));

  state.connections = preservedConnections.concat(editedConnections);
  refreshBindingNodes();
  renderConnections();
  commitHistory();
  hideJsonEditor();
}

function buildEditableMediaJson(metadata) {
  const cloned = JSON.parse(JSON.stringify(metadata ?? {}));
  const thumbnailDataUrl = cloned?.preview?.thumbnailDataUrl;
  if (typeof thumbnailDataUrl === "string" && thumbnailDataUrl.startsWith("data:image/")) {
    cloned.preview.thumbnailDataUrl = MEDIA_THUMBNAIL_TOKEN;
  }
  return {
    text: JSON.stringify(cloned, null, 2),
    thumbnailDataUrl: typeof thumbnailDataUrl === "string" ? thumbnailDataUrl : null,
  };
}

function buildEditableBindingJson(metadata, options = {}) {
  const normalized = normalizeBindingMetadata(metadata);
  if (!normalized) {
    return "{}";
  }

  const values = deepCloneJsonValue(normalized.values);
  const editable = isPlainRecord(values) ? values : {};
  const refs = isPlainRecord(normalized.refs) ? normalized.refs : {};
  const checks = isPlainRecord(normalized.checks) ? normalized.checks : {};
  const counts = resolveNodeHandleCounts(normalized);
  const directiveSideOrder = ["right", "top", "left", "bottom"];

  directiveSideOrder.forEach((side) => {
    const total = Math.max(0, Number(counts?.[side]) || 0);
    for (let index = 1; index <= total; index += 1) {
      const handleKey = buildHandleSideKey(side, index);
      if (!handleKey) {
        continue;
      }
      const directiveKey = formatBindingDirectiveFromHandleKey(handleKey);
      const fieldName = refs[handleKey];
      editable[directiveKey] = typeof fieldName === "string" && fieldName.trim() !== "" ? fieldName.trim() : null;
    }
  });

  Object.entries(checks)
    .sort((a, b) => {
      const ca = parseBindingCheckDirectiveKey(a[0]);
      const cb = parseBindingCheckDirectiveKey(b[0]);
      return (ca?.index ?? 1) - (cb?.index ?? 1);
    })
    .forEach(([checkKey, fieldName]) => {
      if (typeof fieldName !== "string" || fieldName.trim() === "") {
        return;
      }
      editable[checkKey] = fieldName.trim();
    });

  if (!Object.prototype.hasOwnProperty.call(editable, "check_1")) {
    editable.check_1 = null;
  }
  if (!Object.prototype.hasOwnProperty.call(editable, "check_2")) {
    editable.check_2 = null;
  }

  if (typeof options.nodeTitle === "string" && options.nodeTitle.trim() !== "") {
    editable[NODE_DEFINITION_TITLE_KEY] = options.nodeTitle.trim();
  }
  if (options.includeNodeDescriptionField === true) {
    editable[NODE_DEFINITION_DESCRIPTION_KEY] = typeof options.nodeDescription === "string" ? options.nodeDescription : "";
  }

  return stringifyEditorJsonObject(editable);
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
  updateEditorGridBackground();
  state.nodes.forEach((node) => applyNodePosition(node));
  renderConnections();

  if (state.linking) {
    updatePreviewPath();
  }
}

function onKeyDown(event) {
  if (handleShortcutCaptureKeyDown(event)) {
    return;
  }

  if (isTypingTarget(event.target)) {
    return;
  }

  const actionId = findShortcutActionByEvent(event);

  if (actionId === "cancel" && shortcutPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideShortcutPanel();
    return;
  }

  if (event.key === "Escape" && shortcutPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideShortcutPanel();
    return;
  }

  if (actionId === "cancel" && mediaViewerPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideMediaViewer();
    return;
  }

  if (event.key === "Escape" && mediaViewerPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideMediaViewer();
    return;
  }

  if (actionId === "cancel" && jsonEditorPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideJsonEditor();
    return;
  }

  if (actionId === "cancel" && nodeTextEditorPanelEl.classList.contains("visible")) {
    event.preventDefault();
    hideNodeTextEditor();
    return;
  }

  if (actionId === "toggleShortcutPanel") {
    event.preventDefault();
    toggleShortcutPanel();
    return;
  }

  if (shortcutPanelEl.classList.contains("visible")) {
    return;
  }

  if (mediaViewerPanelEl.classList.contains("visible")) {
    return;
  }

  if (jsonEditorPanelEl.classList.contains("visible")) {
    return;
  }

  if (nodeTextEditorPanelEl.classList.contains("visible")) {
    return;
  }

  if (event.key === "Escape" && state.contextMenu) {
    event.preventDefault();
    hideContextMenu();
    return;
  }

  if (actionId === "cancel" && state.contextMenu) {
    event.preventDefault();
    hideContextMenu();
    return;
  }
  if (!actionId) {
    return;
  }
  if (triggerShortcutAction(actionId, event)) {
    event.preventDefault();
  }
}

function onGamepadConnected() {
  startGamepadPolling();
}

function onGamepadDisconnected() {
  if (!getPrimaryConnectedGamepad()) {
    resetGamepadInputState();
  }
}

function startGamepadPolling() {
  if (state.gamepadRafId != null) {
    return;
  }

  const tick = (timestamp) => {
    pollGamepadShortcuts(timestamp);
    state.gamepadRafId = window.requestAnimationFrame(tick);
  };

  state.gamepadRafId = window.requestAnimationFrame(tick);
}

function getPrimaryConnectedGamepad() {
  if (typeof navigator.getGamepads !== "function") {
    return null;
  }

  const gamepads = Array.from(navigator.getGamepads() ?? []).filter(Boolean);
  if (gamepads.length === 0) {
    return null;
  }

  const preferred = gamepads.find((pad) => {
    if (!pad?.connected) {
      return false;
    }
    if (pad.mapping === "standard") {
      return true;
    }
    return /xinput|xbox|controller/i.test(pad.id ?? "");
  });

  return preferred ?? gamepads.find((pad) => pad?.connected) ?? null;
}

function isGamepadButtonPressed(button, threshold = 0.55) {
  if (!button) {
    return false;
  }
  if (button.pressed) {
    return true;
  }
  return Number(button.value) >= threshold;
}

function readGamepadButtonInput(gamepad, buttonIndex) {
  const key = `${gamepad.index}:${buttonIndex}`;
  const pressed = isGamepadButtonPressed(gamepad.buttons?.[buttonIndex]);
  const previous = state.gamepadButtonStates.get(key) === true;
  state.gamepadButtonStates.set(key, pressed);
  return {
    pressed,
    justPressed: pressed && !previous,
  };
}

function resetGamepadInputState() {
  state.gamepadButtonStates.clear();
  state.gamepadRepeatDueAt = {};
}

function runGamepadRepeatAction(timestamp, key, active, action, initialDelay = GAMEPAD_REPEAT_INITIAL_MS, repeatDelay = GAMEPAD_REPEAT_INTERVAL_MS) {
  if (!active) {
    delete state.gamepadRepeatDueAt[key];
    return;
  }

  const nextDueAt = Number(state.gamepadRepeatDueAt[key] ?? 0);
  if (timestamp < nextDueAt) {
    return;
  }

  action();
  state.gamepadRepeatDueAt[key] = timestamp + (nextDueAt > 0 ? repeatDelay : initialDelay);
}

function handleGamepadCancelOrDelete() {
  if (shortcutPanelEl.classList.contains("visible")) {
    hideShortcutPanel();
    return;
  }
  if (mediaViewerPanelEl.classList.contains("visible")) {
    hideMediaViewer();
    return;
  }
  if (jsonEditorPanelEl.classList.contains("visible")) {
    hideJsonEditor();
    return;
  }
  if (nodeTextEditorPanelEl.classList.contains("visible")) {
    hideNodeTextEditor();
    return;
  }
  if (state.contextMenu) {
    hideContextMenu();
    return;
  }
  if (!deleteSelectionOrHover()) {
    cancelCurrentInteraction();
  }
}

function pollGamepadShortcuts(timestampRaw) {
  const timestamp = Number.isFinite(timestampRaw) ? timestampRaw : performance.now();
  const gamepad = getPrimaryConnectedGamepad();
  if (!gamepad) {
    resetGamepadInputState();
    return;
  }

  const btnA = readGamepadButtonInput(gamepad, 0);
  const btnB = readGamepadButtonInput(gamepad, 1);
  const btnX = readGamepadButtonInput(gamepad, 2);
  const btnY = readGamepadButtonInput(gamepad, 3);
  const btnLB = readGamepadButtonInput(gamepad, 4);
  const btnRB = readGamepadButtonInput(gamepad, 5);
  const btnLT = readGamepadButtonInput(gamepad, 6);
  const btnRT = readGamepadButtonInput(gamepad, 7);
  const btnView = readGamepadButtonInput(gamepad, 8);
  const btnStart = readGamepadButtonInput(gamepad, 9);
  const btnLStick = readGamepadButtonInput(gamepad, 10);
  const btnRStick = readGamepadButtonInput(gamepad, 11);
  const dpadUp = readGamepadButtonInput(gamepad, 12);
  const dpadDown = readGamepadButtonInput(gamepad, 13);
  const dpadLeft = readGamepadButtonInput(gamepad, 14);
  const dpadRight = readGamepadButtonInput(gamepad, 15);

  const typing = isTypingTarget(document.activeElement);
  const panelVisibleForCancel =
    shortcutPanelEl.classList.contains("visible") ||
    mediaViewerPanelEl.classList.contains("visible") ||
    jsonEditorPanelEl.classList.contains("visible") ||
    nodeTextEditorPanelEl.classList.contains("visible") ||
    Boolean(state.contextMenu);

  if (btnB.justPressed) {
    if (!typing || panelVisibleForCancel) {
      handleGamepadCancelOrDelete();
    }
    return;
  }

  if (typing) {
    return;
  }

  if (btnStart.justPressed) {
    toggleShortcutPanel();
    return;
  }

  const panelBlocking =
    shortcutPanelEl.classList.contains("visible") ||
    mediaViewerPanelEl.classList.contains("visible") ||
    jsonEditorPanelEl.classList.contains("visible") ||
    nodeTextEditorPanelEl.classList.contains("visible");

  if (panelBlocking) {
    return;
  }

  if (btnY.justPressed) {
    createNodeAtViewportCenter();
  }
  if (btnX.justPressed) {
    copySelectionToClipboard(false);
  }
  if (btnA.justPressed) {
    void pasteClipboard(getViewportCenterWorld());
  }
  if (btnLStick.justPressed) {
    duplicateSelection();
  }
  if (btnRStick.justPressed) {
    deleteSelectionOrHover();
  }
  if (btnLB.justPressed) {
    undoHistory();
  }
  if (btnRB.justPressed) {
    redoHistory();
  }
  if (btnView.justPressed) {
    fitViewToNodes();
  }

  const moveStep = 10;
  const dx = (dpadRight.pressed ? moveStep : 0) - (dpadLeft.pressed ? moveStep : 0);
  const dy = (dpadDown.pressed ? moveStep : 0) - (dpadUp.pressed ? moveStep : 0);
  runGamepadRepeatAction(timestamp, "move", dx !== 0 || dy !== 0, () => {
    moveSelectionBy(dx, dy);
  });

  const bothTriggers = btnLT.pressed && btnRT.pressed;
  runGamepadRepeatAction(timestamp, "zoom-in", !bothTriggers && btnRT.pressed, () => zoomByFactor(1.04));
  runGamepadRepeatAction(timestamp, "zoom-out", !bothTriggers && btnLT.pressed, () => zoomByFactor(1 / 1.04));
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
    path.setAttribute("d", buildPath(fromPoint, toPoint, connection.fromSide, connection.toSide));
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

  preview.setAttribute("d", buildPath(fromPoint, toPoint, state.linking.fromSide, null));
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
    const { c1, c2 } = buildBezierControlPoints(from, to, connection.fromSide, connection.toSide);
    const distance = pointToCubicBezierDistance(state.pointerWorld, from, c1, c2, to);
    const midpoint = evaluateCubicBezierPoint(from, c1, c2, to, 0.5);

    if (!nearest || distance < nearest.distance) {
      nearest = {
        connectionId: connection.id,
        distance,
        midpoint,
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
  const parsedSide = parseHandleSideReference(side) ?? { side: "right", index: 1 };
  const counts = node?.handleCounts ?? resolveNodeHandleCounts(node?.metadata);
  const total = Math.max(1, Number(counts?.[parsedSide.side]) || 1);
  const index = clamp(parsedSide.index, 1, total);
  const ratio = index / (total + 1);
  const width = node.el.offsetWidth;
  const height = node.el.offsetHeight;

  switch (parsedSide.side) {
    case "top":
      return { x: node.x + width * ratio, y: node.y };
    case "right":
      return { x: node.x + width, y: node.y + height * ratio };
    case "bottom":
      return { x: node.x + width * ratio, y: node.y + height };
    case "left":
      return { x: node.x, y: node.y + height * ratio };
    default:
      return { x: node.x + width / 2, y: node.y + height / 2 };
  }
}

function resolveHandleDirectionVector(sideReference) {
  const parsed = parseHandleSideReference(sideReference);
  const side = parsed?.side ?? (typeof sideReference === "string" && HANDLE_SIDES.includes(sideReference) ? sideReference : null);
  switch (side) {
    case "top":
      return { x: 0, y: -1 };
    case "right":
      return { x: 1, y: 0 };
    case "bottom":
      return { x: 0, y: 1 };
    case "left":
      return { x: -1, y: 0 };
    default:
      return null;
  }
}

function normalizeVector(vx, vy) {
  const length = Math.hypot(vx, vy);
  if (!Number.isFinite(length) || length < 1e-6) {
    return null;
  }
  return {
    x: vx / length,
    y: vy / length,
  };
}

function computeBezierControlOffset(start, end, dir) {
  if (!dir) {
    return 64;
  }
  const dx = Math.abs(end.x - start.x);
  const dy = Math.abs(end.y - start.y);
  const primaryDelta = dir.x !== 0 ? dx : dy;
  const secondaryDelta = dir.x !== 0 ? dy : dx;
  const geometricBase = Math.hypot(dx, dy) * 0.28;
  return clamp(Math.max(36, primaryDelta * 0.56 + secondaryDelta * 0.08, geometricBase), 36, 260);
}

function buildBezierControlPoints(start, end, startSide = null, endSide = null) {
  const startToEnd = normalizeVector(end.x - start.x, end.y - start.y) ?? { x: 1, y: 0 };
  const endToStart = normalizeVector(start.x - end.x, start.y - end.y) ?? { x: -1, y: 0 };
  const startDir = resolveHandleDirectionVector(startSide) ?? startToEnd;
  const endDir = resolveHandleDirectionVector(endSide) ?? endToStart;

  const startOffset = computeBezierControlOffset(start, end, startDir);
  const endOffset = computeBezierControlOffset(end, start, endDir);

  const c1 = {
    x: start.x + startDir.x * startOffset,
    y: start.y + startDir.y * startOffset,
  };
  const c2 = {
    x: end.x + endDir.x * endOffset,
    y: end.y + endDir.y * endOffset,
  };
  return { c1, c2 };
}

function buildPath(start, end, startSide = null, endSide = null) {
  const { c1, c2 } = buildBezierControlPoints(start, end, startSide, endSide);
  return `M ${start.x} ${start.y} C ${c1.x} ${c1.y}, ${c2.x} ${c2.y}, ${end.x} ${end.y}`;
}

function evaluateCubicBezierPoint(p0, p1, p2, p3, t) {
  const u = 1 - t;
  const tt = t * t;
  const uu = u * u;
  const uuu = uu * u;
  const ttt = tt * t;
  return {
    x: uuu * p0.x + 3 * uu * t * p1.x + 3 * u * tt * p2.x + ttt * p3.x,
    y: uuu * p0.y + 3 * uu * t * p1.y + 3 * u * tt * p2.y + ttt * p3.y,
  };
}

function pointToCubicBezierDistance(point, p0, p1, p2, p3, segments = 28) {
  let minDistance = Number.POSITIVE_INFINITY;
  let prev = p0;
  for (let step = 1; step <= segments; step += 1) {
    const t = step / segments;
    const curr = evaluateCubicBezierPoint(p0, p1, p2, p3, t);
    const distance = pointToSegmentDistance(point, prev, curr);
    if (distance < minDistance) {
      minDistance = distance;
    }
    prev = curr;
  }
  return minDistance;
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
  updateEditorGridBackground();
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
    const visualType = resolveNodeVisualType(node.metadata);
    syncVisualTypeClass(rect, NAVIGATOR_NODE_VISUAL_CLASS_PREFIX, visualType);
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

function resolveNextAutoIndexedTitle(baseTitle) {
  const base = normalizeNodeTitle(baseTitle, "");
  if (!base) {
    return "節點 01";
  }

  const escapedBase = base.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const suffixPattern = new RegExp(`^${escapedBase}(\\d+)$`);
  let maxIndex = 0;

  state.nodes.forEach((node) => {
    if (!node || typeof node.title !== "string") {
      return;
    }
    const match = suffixPattern.exec(node.title.trim());
    if (!match) {
      return;
    }
    const index = Number.parseInt(match[1], 10);
    if (Number.isFinite(index)) {
      maxIndex = Math.max(maxIndex, index);
    }
  });

  return `${base}${String(maxIndex + 1).padStart(2, "0")}`;
}

function createNodeAtViewportCenter() {
  const center = getViewportCenterWorld();
  createNodeAtWorld(center);
}

function createNodeAtWorld(worldPoint) {
  const node = createNode(worldPoint.x - NODE_DEFAULT_WIDTH / 2, worldPoint.y - NODE_DEFAULT_HEIGHT / 2, {
    metadata: createDefaultBindingDirectiveSeed(),
  });
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

function buildNodeMediaDeletionCandidate(node) {
  if (!node || !isMediaMetadata(node.metadata)) {
    return null;
  }
  const storage = node.metadata?.storage;
  if (!storage || typeof storage !== "object") {
    return null;
  }

  const key = buildMediaStorageKeyFromEntry(storage);
  const fileName = resolveStorageFileNameFromManifestEntry(storage);
  if (!key || !fileName || !key.startsWith("media/")) {
    return null;
  }

  return { key, fileName };
}

function queueOrphanedMediaCleanup(removedCandidates) {
  if (!Array.isArray(removedCandidates) || removedCandidates.length === 0) {
    return;
  }

  const remainingKeys = new Set();
  state.nodes.forEach((node) => {
    const candidate = buildNodeMediaDeletionCandidate(node);
    if (candidate) {
      remainingKeys.add(candidate.key);
    }
  });

  const orphanedByKey = new Map();
  removedCandidates.forEach((candidate) => {
    if (!candidate || !candidate.key || remainingKeys.has(candidate.key)) {
      return;
    }
    if (!orphanedByKey.has(candidate.key)) {
      orphanedByKey.set(candidate.key, candidate);
    }
  });

  if (orphanedByKey.size === 0) {
    return;
  }

  void deleteOrphanedMediaFiles(Array.from(orphanedByKey.values()));
}

async function resolveMediaDirectoryForCleanup() {
  if (state.mediaStoreHandle) {
    const allowed = await verifyHandlePermission(state.mediaStoreHandle, false);
    if (allowed) {
      return state.mediaStoreHandle;
    }
  }

  if (state.projectRootHandle) {
    const rootAllowed = await verifyHandlePermission(state.projectRootHandle, false);
    if (rootAllowed) {
      try {
        const mediaDir = await state.projectRootHandle.getDirectoryHandle("media", { create: false });
        state.mediaStoreHandle = mediaDir;
        return mediaDir;
      } catch {
        return null;
      }
    }
  }

  return null;
}

async function deleteOrphanedMediaFiles(candidates) {
  const mediaDirectory = await resolveMediaDirectoryForCleanup();
  if (!mediaDirectory) {
    return;
  }

  const deletedKeys = new Set();
  for (const candidate of candidates) {
    if (!candidate?.fileName || !candidate?.key) {
      continue;
    }
    try {
      await mediaDirectory.removeEntry(candidate.fileName);
      deletedKeys.add(candidate.key);
    } catch (error) {
      if (error?.name === "NotFoundError") {
        deletedKeys.add(candidate.key);
      } else {
        console.error("delete orphaned media failed", error);
      }
    }
  }

  if (deletedKeys.size === 0) {
    return;
  }

  state.mediaManifest = state.mediaManifest.filter((entry) => !deletedKeys.has(buildMediaStorageKeyFromEntry(entry)));
  persistMediaManifest();
}

function deleteNodes(nodeIds, options = {}) {
  const shouldRecord = options.recordHistory !== false;
  const uniqueIds = Array.from(new Set(nodeIds)).filter((id) => state.nodes.has(id));
  if (uniqueIds.length === 0) {
    return false;
  }

  const removedMediaCandidates = [];
  const removedSet = new Set(uniqueIds);
  uniqueIds.forEach((id) => {
    const node = state.nodes.get(id);
    const mediaCandidate = buildNodeMediaDeletionCandidate(node);
    if (mediaCandidate) {
      removedMediaCandidates.push(mediaCandidate);
    }
    node?.el.remove();
    state.nodes.delete(id);
    state.selectedNodeIds.delete(id);
  });

  state.connections = state.connections.filter((connection) => {
    return !removedSet.has(connection.fromNodeId) && !removedSet.has(connection.toNodeId);
  });
  refreshBindingNodes();

  state.hoverConnectionId = null;
  lineDeleteBtn.classList.remove("visible");
  syncNodeSelectionClasses();
  renderConnections();
  if (shouldRecord) {
    commitHistory();
  }
  queueOrphanedMediaCleanup(removedMediaCandidates);
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
      metadata: nodeData.metadata ?? null,
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

  refreshBindingNodes();
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
      metadata: node.metadata ?? null,
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

function buildClipboardPayloadFromBindingSyntax(rawObject) {
  const metadata = normalizeNodeMetadata(rawObject);
  if (!isBindingMetadata(metadata)) {
    return null;
  }

  const counts = resolveNodeHandleCounts(metadata);
  const width = computeNodeMinWidthFromHandleCounts(counts);
  const height = computeNodeMinHeightFromHandleCounts(counts);
  const titleSource = isPlainRecord(rawObject) ? rawObject.title : null;
  const title = typeof titleSource === "string" && titleSource.trim() !== "" ? titleSource.trim() : "JSON 綁定節點";

  return {
    type: CLIPBOARD_SCHEMA,
    version: CLIPBOARD_VERSION,
    nodes: [
      {
        id: "1",
        offsetX: 0,
        offsetY: 0,
        title,
        content: buildBindingNodeContent(metadata),
        metadata,
      },
    ],
    connections: [],
    width,
    height,
  };
}

function buildClipboardPayloadFromNodeDefinition(rawObject) {
  if (!isPlainRecord(rawObject)) {
    return null;
  }

  const definition = selectNodeDefinitionForObject(rawObject);
  if (!definition) {
    return null;
  }

  const merged = deepCloneJsonValue(rawObject) ?? {};
  if (!isPlainRecord(merged)) {
    return null;
  }

  Object.entries(definition.template ?? {}).forEach(([key, value]) => {
    if (key === NODE_DEFINITION_TITLE_KEY || key === NODE_DEFINITION_DESCRIPTION_KEY) {
      merged[key] = deepCloneJsonValue(value);
      return;
    }
    if (parseBindingDirectiveKey(key) || parseBindingCheckDirectiveKey(key)) {
      merged[key] = deepCloneJsonValue(value);
      return;
    }
    if (!Object.prototype.hasOwnProperty.call(merged, key)) {
      merged[key] = deepCloneJsonValue(value);
    }
  });

  const preferredTitle =
    typeof merged[NODE_DEFINITION_TITLE_KEY] === "string" && merged[NODE_DEFINITION_TITLE_KEY].trim() !== ""
      ? merged[NODE_DEFINITION_TITLE_KEY].trim()
      : typeof merged.title === "string" && merged.title.trim() !== ""
      ? merged.title.trim()
      : typeof definition.name === "string" && definition.name.trim() !== ""
        ? definition.name.trim()
        : null;
  const autoIndexedTitle = preferredTitle ? resolveNextAutoIndexedTitle(preferredTitle) : null;
  const preferredDescription =
    typeof merged[NODE_DEFINITION_DESCRIPTION_KEY] === "string" && merged[NODE_DEFINITION_DESCRIPTION_KEY].trim() !== ""
      ? normalizeNodeContent(merged[NODE_DEFINITION_DESCRIPTION_KEY])
      : typeof definition.description === "string" && definition.description.trim() !== ""
        ? normalizeNodeContent(definition.description)
        : "";

  const metadataSeed = deepCloneJsonValue(merged) ?? {};
  if (!isPlainRecord(metadataSeed)) {
    return null;
  }
  delete metadataSeed[NODE_DEFINITION_TITLE_KEY];
  delete metadataSeed[NODE_DEFINITION_DESCRIPTION_KEY];

  const payload = buildClipboardPayloadFromBindingSyntax(metadataSeed);
  if (payload?.nodes?.[0]) {
    if (autoIndexedTitle) {
      payload.nodes[0].title = autoIndexedTitle;
    }
    if (preferredDescription) {
      if (isBindingMetadata(payload.nodes[0].metadata)) {
        const normalizedBinding = normalizeBindingMetadata(payload.nodes[0].metadata);
        if (normalizedBinding) {
          if (!isPlainRecord(normalizedBinding.display)) {
            normalizedBinding.display = {};
          }
          normalizedBinding.display.description = preferredDescription;
          payload.nodes[0].metadata = normalizedBinding;
        }
      }
      payload.nodes[0].content = preferredDescription;
    }
    return payload;
  }

  const content = preferredDescription || normalizeNodeContent(buildJsonNodeSummaryText(rawObject));
  return {
    type: CLIPBOARD_SCHEMA,
    version: CLIPBOARD_VERSION,
    nodes: [
      {
        id: "1",
        offsetX: 0,
        offsetY: 0,
        title: autoIndexedTitle || preferredTitle || "未定義結點",
        content,
        metadata: metadataSeed,
      },
    ],
    connections: [],
    width: NODE_DEFAULT_WIDTH,
    height: NODE_DEFAULT_HEIGHT,
  };
}

function buildClipboardPayloadFromUnknownJson(rawValue, originalText = "") {
  const content = normalizeNodeContent(buildJsonNodeSummaryText(rawValue));
  const metadata = isPlainRecord(rawValue) ? deepCloneJsonValue(rawValue) : null;

  return {
    type: CLIPBOARD_SCHEMA,
    version: CLIPBOARD_VERSION,
    nodes: [
      {
        id: "1",
        offsetX: 0,
        offsetY: 0,
        title: "未定義結點",
        content,
        metadata,
      },
    ],
    connections: [],
    width: NODE_DEFAULT_WIDTH,
    height: NODE_DEFAULT_HEIGHT,
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

  if (!Array.isArray(parsed?.nodes)) {
    const bindingPayload = buildClipboardPayloadFromBindingSyntax(parsed);
    if (bindingPayload) {
      return bindingPayload;
    }
    const definedPayload = buildClipboardPayloadFromNodeDefinition(parsed);
    if (definedPayload) {
      return definedPayload;
    }
  }

  const normalized = normalizeClipboardPayload(parsed);
  if (!normalized) {
    return buildClipboardPayloadFromUnknownJson(parsed, text);
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
      metadata: node.metadata ?? null,
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
      metadata: rawNode.metadata && typeof rawNode.metadata === "object" ? rawNode.metadata : null,
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
    metadata: node.metadata ?? null,
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
      const normalizedFromSide = normalizeHandleSideKey(fromSide);
      const normalizedToSide = normalizeHandleSideKey(toSide);
      if (
        !nodeIdSet.has(fromSourceId) ||
        !nodeIdSet.has(toSourceId) ||
        !normalizedFromSide ||
        !normalizedToSide
      ) {
        return null;
      }
      return { fromSourceId, fromSide: normalizedFromSide, toSourceId, toSide: normalizedToSide };
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

async function onSetMediaPathClick() {
  try {
    await pickAndPersistProjectDirectory();
    if (!state.projectRootHandle) {
      updateSetPathButtonState();
      return;
    }

    const loadResult = await loadProjectStateFromProjectRoot(state.projectRootHandle, {
      suppressErrors: true,
      syncRootFromFile: false,
    });
    if (!loadResult.found) {
      await clearProjectStateFileBinding();
    } else if (loadResult.error) {
      console.error("load project from selected path failed", loadResult.error);
      alert("此路徑中的 project-state.json 無法讀取，已綁定路徑但未載入內容。");
    }
    updateSetPathButtonState();
  } catch (error) {
    console.error("set media path failed", error);
  }
}

async function onSaveProjectClick() {
  try {
    await saveProjectStateNow();
  } catch (error) {
    console.error("manual project save failed", error);
  }
}

async function onLoadProjectClick() {
  try {
    const source = await resolveProjectFileSourceForLoad();
    if (!source?.file) {
      return;
    }
    await loadProjectFromFile(source.file);
    if (source.fileHandle) {
      await setProjectStateFileHandle(source.fileHandle);
      await syncProjectRootFromProjectFile(source.fileHandle);
    }
  } catch (error) {
    if (error?.name === "AbortError") {
      return;
    }
    console.error("load project failed", error);
  }
}

async function onLoadLegacyProjectClick() {
  try {
    const source = await pickProjectFileForLoad();
    if (!source?.file) {
      return;
    }
    await loadLegacyProjectFromFile(source.file);
  } catch (error) {
    if (error?.name === "AbortError") {
      return;
    }
    console.error("load legacy project failed", error);
  }
}

async function initializeProjectFileFromStorage() {
  const storedFileHandle = await loadHandleFromStorage(PROJECT_FILE_HANDLE_KEY);
  if (!storedFileHandle) {
    return null;
  }
  const allowed = await verifyHandlePermission(storedFileHandle, false, "read");
  if (allowed) {
    state.projectStateFileHandle = storedFileHandle;
    await syncProjectRootFromProjectFile(storedFileHandle);
    return storedFileHandle;
  }
  return null;
}

async function resolveProjectFileSourceForLoad(options = {}) {
  const { allowPicker = true } = options;
  const boundHandle = await resolveProjectStateFileHandle({
    forWrite: false,
    canRequestPermission: false,
    allowPrompt: false,
  });
  if (boundHandle) {
    try {
      const file = await boundHandle.getFile();
      return { file, fileHandle: boundHandle };
    } catch {
      state.projectStateFileHandle = null;
    }
  }
  if (!allowPicker) {
    return null;
  }
  return pickProjectFileForLoad();
}

async function autoLoadLastProjectOnStartup() {
  try {
    await initializeProjectFileFromStorage();
    const draftLoaded = loadProjectStateFromDraftStorage();
    if (draftLoaded) {
      return true;
    }
    const gitLoaded = await loadProjectStateFromGitOnStartup();
    if (gitLoaded) {
      return true;
    }
    const source = await resolveProjectFileSourceForLoad({ allowPicker: false });
    if (!source?.file) {
      return false;
    }
    await loadProjectFromFile(source.file);
    if (source.fileHandle) {
      await setProjectStateFileHandle(source.fileHandle);
      await syncProjectRootFromProjectFile(source.fileHandle);
    }
    return true;
  } catch (error) {
    console.error("auto load last project failed", error);
    return false;
  }
}

async function loadProjectStateFromProjectRoot(rootHandle, options = {}) {
  const { suppressErrors = false, syncRootFromFile = true } = options;
  if (!rootHandle) {
    return { found: false, loaded: false, error: null };
  }

  let fileHandle;
  try {
    fileHandle = await rootHandle.getFileHandle(PROJECT_STATE_FILE, { create: false });
  } catch (error) {
    if (error?.name === "NotFoundError") {
      return { found: false, loaded: false, error: null };
    }
    if (!suppressErrors) {
      throw error;
    }
    return { found: false, loaded: false, error };
  }

  try {
    await setProjectStateFileHandle(fileHandle);
    const file = await fileHandle.getFile();
    if (Number(file.size) > 0) {
      await loadProjectFromFile(file);
      if (syncRootFromFile) {
        await syncProjectRootFromProjectFile(fileHandle);
      }
      return { found: true, loaded: true, error: null };
    }
    if (syncRootFromFile) {
      await syncProjectRootFromProjectFile(fileHandle);
    }
    return { found: true, loaded: false, error: null };
  } catch (error) {
    if (!suppressErrors) {
      throw error;
    }
    return { found: true, loaded: false, error };
  }
}

function loadProjectStateFromDraftStorage() {
  try {
    const raw = localStorage.getItem(PROJECT_DRAFT_STORAGE_KEY);
    if (!raw) {
      return false;
    }
    const parsed = JSON.parse(raw);
    const snapshot = isPlainRecord(parsed?.snapshot) ? parsed.snapshot : null;
    if (!snapshot || !Array.isArray(snapshot.nodes) || !Array.isArray(snapshot.connections)) {
      return false;
    }
    restoreSnapshot(snapshot);
    state.history.undo = [snapshot];
    state.history.redo = [];
    state.history.lastHash = hashSnapshot(snapshot);
    state.projectLastSavedSnapshot = cloneSnapshot(snapshot);
    state.projectLastSavedHash = state.history.lastHash;
    updateHistoryToolbarButtonsState();
    updateGitSyncStatusIndicator();
    return true;
  } catch (error) {
    console.warn("load project draft failed", error);
    return false;
  }
}

function persistProjectStateDraft(snapshot) {
  if (!snapshot) {
    return;
  }
  try {
    localStorage.setItem(
      PROJECT_DRAFT_STORAGE_KEY,
      JSON.stringify({
        savedAt: new Date().toISOString(),
        snapshot,
      })
    );
  } catch (error) {
    console.warn("persist project draft failed", error);
  }
}

async function syncProjectRootFromProjectFile(fileHandle) {
  if (!fileHandle || typeof fileHandle.isSameEntry !== "function") {
    return false;
  }

  const candidates = [];
  if (state.projectRootHandle) {
    candidates.push(state.projectRootHandle);
  }
  const storedRoot = await loadHandleFromStorage(PROJECT_ROOT_HANDLE_KEY);
  if (storedRoot && !candidates.includes(storedRoot)) {
    candidates.push(storedRoot);
  }

  for (const rootHandle of candidates) {
    const allowed = await verifyHandlePermission(rootHandle, false, "read");
    if (!allowed) {
      continue;
    }
    try {
      const rootProjectFile = await rootHandle.getFileHandle(PROJECT_STATE_FILE, { create: false });
      const same = await rootProjectFile.isSameEntry(fileHandle);
      if (!same) {
        continue;
      }
      state.projectRootHandle = rootHandle;
      state.mediaStoreMode = "directory";
      try {
        state.mediaStoreHandle = await rootHandle.getDirectoryHandle("media", { create: true });
      } catch {
        state.mediaStoreHandle = null;
      }
      updateSetPathButtonState();
      return true;
    } catch {
      // continue checking next candidate
    }
  }

  return false;
}

async function pickProjectFileForLoad() {
  if (typeof window.showOpenFilePicker === "function") {
    const handles = await window.showOpenFilePicker({
      id: "node-editor-project-load",
      multiple: false,
      types: [
        {
          description: "Note Editor Project",
          accept: {
            "application/json": [".json"],
          },
        },
      ],
    });
    if (!handles || handles.length === 0) {
      return null;
    }
    const fileHandle = handles[0];
    const file = await fileHandle.getFile();
    return { file, fileHandle };
  }

  return pickProjectFileWithInputFallback();
}

function pickProjectFileWithInputFallback() {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json,application/json";
    input.addEventListener(
      "change",
      () => {
        const file = input.files?.[0] ?? null;
        resolve(file ? { file, fileHandle: null } : null);
      },
      { once: true }
    );
    input.click();
  });
}

async function loadProjectFromFile(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const normalized = normalizeProjectPayload(parsed);
  if (!normalized) {
    throw new Error("invalid project file format");
  }
  applyLoadedProjectState(normalized);
}

async function loadLegacyProjectFromFile(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const normalized = normalizeLegacyProjectPayload(parsed);
  if (!normalized) {
    throw new Error("invalid legacy project file format");
  }
  applyLoadedProjectState(normalized);
}

function normalizeProjectPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return null;
  }
  if (payload.type && payload.type !== PROJECT_STATE_SCHEMA) {
    return normalizeLegacyProjectPayload(payload);
  }

  const rawNodes = Array.isArray(payload.nodes) ? payload.nodes : [];
  const nodes = [];
  const seenNodeIds = new Set();
  let fallbackNodeId = 1;

  rawNodes.forEach((rawNode) => {
    if (!rawNode || typeof rawNode !== "object") {
      return;
    }

    let nodeId = String(rawNode.id ?? fallbackNodeId);
    while (seenNodeIds.has(nodeId)) {
      nodeId = `${nodeId}_${fallbackNodeId}`;
      fallbackNodeId += 1;
    }
    seenNodeIds.add(nodeId);
    fallbackNodeId += 1;

    const x = Number.isFinite(rawNode.x) ? Number(rawNode.x) : 0;
    const y = Number.isFinite(rawNode.y) ? Number(rawNode.y) : 0;
    nodes.push({
      id: nodeId,
      x,
      y,
      title: normalizeNodeTitle(rawNode.title, nodeId),
      content: normalizeNodeContent(rawNode.content),
      metadata: rawNode.metadata && typeof rawNode.metadata === "object" ? rawNode.metadata : null,
    });
  });

  const nodeIdSet = new Set(nodes.map((node) => node.id));
  const rawConnections = Array.isArray(payload.connections) ? payload.connections : [];
  const connections = [];
  const seenConnectionIds = new Set();
  let fallbackConnectionId = 1;

  rawConnections.forEach((rawConnection) => {
    if (!rawConnection || typeof rawConnection !== "object") {
      return;
    }
    const fromNodeId = String(rawConnection.fromNodeId ?? rawConnection.fromId ?? "");
    const toNodeId = String(rawConnection.toNodeId ?? rawConnection.toId ?? "");
    const fromSide = normalizeHandleSideKey(rawConnection.fromSide);
    const toSide = normalizeHandleSideKey(rawConnection.toSide);
    if (
      !nodeIdSet.has(fromNodeId) ||
      !nodeIdSet.has(toNodeId) ||
      !fromSide ||
      !toSide
    ) {
      return;
    }

    let connectionId = String(rawConnection.id ?? fallbackConnectionId);
    while (seenConnectionIds.has(connectionId)) {
      connectionId = `${connectionId}_${fallbackConnectionId}`;
      fallbackConnectionId += 1;
    }
    seenConnectionIds.add(connectionId);
    fallbackConnectionId += 1;

    connections.push({
      id: connectionId,
      fromNodeId,
      fromSide,
      toNodeId,
      toSide,
    });
  });

  const selectedNodeIds = Array.isArray(payload.selectedNodeIds)
    ? payload.selectedNodeIds.map((id) => String(id)).filter((id) => nodeIdSet.has(id))
    : [];

  const maxNodeId = nodes.reduce((max, node) => {
    const numeric = Number(node.id);
    return Number.isFinite(numeric) ? Math.max(max, numeric) : max;
  }, 0);
  const maxConnectionId = connections.reduce((max, connection) => {
    const numeric = Number(connection.id);
    return Number.isFinite(numeric) ? Math.max(max, numeric) : max;
  }, 0);

  const nextNodeId = Number.isFinite(payload.nextNodeId)
    ? Math.max(1, Math.floor(payload.nextNodeId))
    : maxNodeId + 1;
  const nextConnectionId = Number.isFinite(payload.nextConnectionId)
    ? Math.max(1, Math.floor(payload.nextConnectionId))
    : maxConnectionId + 1;
  const nodeDefinitions = normalizeNodeDefinitionsPayload(
    payload.nodeDefinitions ?? payload.definitions ?? payload.nodeDefinitionRules ?? null
  );

  const viewport = normalizeViewportPayload(payload.viewport);
  const mediaManifest = normalizeMediaManifestPayload(payload.media?.manifestItems);

  return {
    snapshot: {
      nextNodeId,
      nextConnectionId,
      nodeDefinitions,
      nodes,
      connections,
      selectedNodeIds,
    },
    viewport,
    mediaManifest,
  };
}

function normalizeLegacyProjectPayload(rawPayload) {
  if (!rawPayload || typeof rawPayload !== "object") {
    return null;
  }

  let payload = rawPayload;
  if (Array.isArray(payload)) {
    payload = { nodes: payload };
  }
  if (payload?.snapshot && typeof payload.snapshot === "object") {
    payload = {
      ...payload.snapshot,
      viewport: payload.viewport ?? payload.snapshot.viewport ?? null,
      media: payload.media ?? payload.snapshot.media ?? null,
      mediaManifest: payload.mediaManifest ?? payload.snapshot.mediaManifest ?? null,
    };
  }
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const rawNodes = Array.isArray(payload.nodes) ? payload.nodes : null;
  if (!rawNodes || rawNodes.length === 0) {
    return null;
  }

  const nodes = [];
  const seenNodeIds = new Set();
  let fallbackNodeId = 1;
  let offsetMinX = Number.POSITIVE_INFINITY;
  let offsetMinY = Number.POSITIVE_INFINITY;
  const rawNodeDrafts = [];

  rawNodes.forEach((rawNode) => {
    if (!rawNode || typeof rawNode !== "object") {
      return;
    }

    let nodeId = String(rawNode.id ?? rawNode.localId ?? fallbackNodeId);
    while (seenNodeIds.has(nodeId)) {
      nodeId = `${nodeId}_${fallbackNodeId}`;
      fallbackNodeId += 1;
    }
    seenNodeIds.add(nodeId);
    fallbackNodeId += 1;

    const hasX = Number.isFinite(rawNode.x);
    const hasY = Number.isFinite(rawNode.y);
    const hasOffsetX = Number.isFinite(rawNode.offsetX);
    const hasOffsetY = Number.isFinite(rawNode.offsetY);

    if (!hasX && hasOffsetX) {
      offsetMinX = Math.min(offsetMinX, Number(rawNode.offsetX));
    }
    if (!hasY && hasOffsetY) {
      offsetMinY = Math.min(offsetMinY, Number(rawNode.offsetY));
    }

    rawNodeDrafts.push({
      id: nodeId,
      rawNode,
      hasX,
      hasY,
      hasOffsetX,
      hasOffsetY,
    });
  });

  if (rawNodeDrafts.length === 0) {
    return null;
  }

  if (!Number.isFinite(offsetMinX)) {
    offsetMinX = 0;
  }
  if (!Number.isFinite(offsetMinY)) {
    offsetMinY = 0;
  }

  rawNodeDrafts.forEach((draft) => {
    const { id, rawNode, hasX, hasY, hasOffsetX, hasOffsetY } = draft;
    const x = hasX ? Number(rawNode.x) : hasOffsetX ? Number(rawNode.offsetX) - offsetMinX : 0;
    const y = hasY ? Number(rawNode.y) : hasOffsetY ? Number(rawNode.offsetY) - offsetMinY : 0;
    const metadataValue =
      rawNode.metadata && typeof rawNode.metadata === "object"
        ? rawNode.metadata
        : rawNode.data && typeof rawNode.data === "object"
          ? rawNode.data
          : null;
    const titleSource =
      typeof rawNode.title === "string"
        ? rawNode.title
        : typeof rawNode.name === "string"
          ? rawNode.name
          : `節點 ${id}`;
    const contentSource =
      typeof rawNode.content === "string"
        ? rawNode.content
        : typeof rawNode.text === "string"
          ? rawNode.text
          : "拖曳節點或拉線連接";

    nodes.push({
      id,
      x,
      y,
      title: normalizeNodeTitle(titleSource, id),
      content: normalizeNodeContent(contentSource),
      metadata: metadataValue,
    });
  });

  const nodeIdSet = new Set(nodes.map((node) => node.id));
  const rawConnections = Array.isArray(payload.connections)
    ? payload.connections
    : Array.isArray(payload.links)
      ? payload.links
      : Array.isArray(payload.edges)
        ? payload.edges
        : [];
  const connections = [];
  const seenConnectionIds = new Set();
  let fallbackConnectionId = 1;

  rawConnections.forEach((rawConnection) => {
    if (!rawConnection || typeof rawConnection !== "object") {
      return;
    }

    const fromNodeId = String(
      rawConnection.fromNodeId ??
        rawConnection.fromId ??
        rawConnection.fromLocalId ??
        rawConnection.from?.id ??
        rawConnection.source ??
        ""
    );
    const toNodeId = String(
      rawConnection.toNodeId ??
        rawConnection.toId ??
        rawConnection.toLocalId ??
        rawConnection.to?.id ??
        rawConnection.target ??
        ""
    );

    const fromSide = normalizeHandleSideKey(
      rawConnection.fromSide ?? rawConnection.from?.side ?? rawConnection.sourceSide ?? "right"
    );
    const toSide = normalizeHandleSideKey(rawConnection.toSide ?? rawConnection.to?.side ?? rawConnection.targetSide ?? "left");

    if (!nodeIdSet.has(fromNodeId) || !nodeIdSet.has(toNodeId) || !fromSide || !toSide) {
      return;
    }

    let connectionId = String(rawConnection.id ?? fallbackConnectionId);
    while (seenConnectionIds.has(connectionId)) {
      connectionId = `${connectionId}_${fallbackConnectionId}`;
      fallbackConnectionId += 1;
    }
    seenConnectionIds.add(connectionId);
    fallbackConnectionId += 1;

    connections.push({
      id: connectionId,
      fromNodeId,
      fromSide,
      toNodeId,
      toSide,
    });
  });

  const selectedNodeIds = Array.isArray(payload.selectedNodeIds)
    ? payload.selectedNodeIds.map((id) => String(id)).filter((id) => nodeIdSet.has(id))
    : Array.isArray(payload.selected)
      ? payload.selected.map((id) => String(id)).filter((id) => nodeIdSet.has(id))
      : [];

  const maxNodeId = nodes.reduce((max, node) => {
    const numeric = Number(node.id);
    return Number.isFinite(numeric) ? Math.max(max, numeric) : max;
  }, 0);
  const maxConnectionId = connections.reduce((max, connection) => {
    const numeric = Number(connection.id);
    return Number.isFinite(numeric) ? Math.max(max, numeric) : max;
  }, 0);

  const nextNodeId = Number.isFinite(payload.nextNodeId)
    ? Math.max(1, Math.floor(payload.nextNodeId))
    : maxNodeId + 1;
  const nextConnectionId = Number.isFinite(payload.nextConnectionId)
    ? Math.max(1, Math.floor(payload.nextConnectionId))
    : maxConnectionId + 1;
  const nodeDefinitions = normalizeNodeDefinitionsPayload(
    payload.nodeDefinitions ?? payload.definitions ?? payload.nodeDefinitionRules ?? null
  );

  const viewport = normalizeViewportPayload(payload.viewport ?? payload.view ?? null);
  const mediaManifest = normalizeMediaManifestPayload(payload.media?.manifestItems ?? payload.mediaManifest ?? null);

  return {
    snapshot: {
      nextNodeId,
      nextConnectionId,
      nodeDefinitions,
      nodes,
      connections,
      selectedNodeIds,
    },
    viewport,
    mediaManifest,
  };
}

function normalizeViewportPayload(rawViewport) {
  if (!rawViewport || typeof rawViewport !== "object") {
    return null;
  }
  const viewport = {
    x: Number.isFinite(rawViewport.x) ? Number(rawViewport.x) : 0,
    y: Number.isFinite(rawViewport.y) ? Number(rawViewport.y) : 0,
    zoom: Number.isFinite(rawViewport.zoom) ? clamp(Number(rawViewport.zoom), ZOOM_MIN, ZOOM_MAX) : 1,
  };
  return viewport;
}

function normalizeMediaManifestPayload(rawItems) {
  if (!Array.isArray(rawItems)) {
    return [];
  }
  return rawItems.filter((item) => item && typeof item === "object").map((item) => ({ ...item }));
}

function applyLoadedProjectState(payload) {
  cancelCurrentInteraction();
  hideContextMenu();
  hideDefinedNodePanel();
  hideShortcutPanel();
  hideJsonEditor();
  hideNodeTextEditor();
  hideMediaViewer();
  state.projectSaveQueued = null;
  state.gitSync.queuedSnapshot = null;
  state.gitSync.queuedHash = "";
  if (state.gitSync.timerId != null) {
    clearTimeout(state.gitSync.timerId);
    state.gitSync.timerId = null;
  }

  restoreSnapshot(payload.snapshot);
  persistProjectStateDraft(payload.snapshot);
  if (payload.viewport) {
    applyViewportState(payload.viewport);
  }
  state.mediaManifest = payload.mediaManifest;
  initializeHistory();
}

function applyViewportState(viewport) {
  if (!viewport || typeof viewport !== "object") {
    return;
  }
  if (Number.isFinite(viewport.zoom)) {
    state.zoom = clamp(Number(viewport.zoom), ZOOM_MIN, ZOOM_MAX);
  }
  if (Number.isFinite(viewport.x)) {
    state.viewport.x = Number(viewport.x);
  }
  if (Number.isFinite(viewport.y)) {
    state.viewport.y = Number(viewport.y);
  }
  updateEditorGridBackground();

  state.nodes.forEach((node) => applyNodePosition(node));
  renderConnections();
  updateNodeProximity();
  updateConnectionProximity();
}

async function initializeMediaPathFromStorage() {
  const projectRootHandle = await loadHandleFromStorage(PROJECT_ROOT_HANDLE_KEY);
  if (projectRootHandle) {
    const hasPermission = await verifyHandlePermission(projectRootHandle, false);
    if (hasPermission) {
      state.projectRootHandle = projectRootHandle;
      state.mediaStoreMode = "directory";
      try {
        state.mediaStoreHandle = await projectRootHandle.getDirectoryHandle("media", { create: true });
      } catch {
        state.mediaStoreHandle = null;
      }
      updateSetPathButtonState();
      return;
    }
  }

  const legacyMediaHandle = await loadHandleFromStorage(LEGACY_MEDIA_HANDLE_KEY);
  if (legacyMediaHandle) {
    const hasPermission = await verifyHandlePermission(legacyMediaHandle, false);
    if (hasPermission) {
      state.mediaStoreHandle = legacyMediaHandle;
      state.mediaStoreMode = "legacy";
    }
  }
  updateSetPathButtonState();
}

function updateSetPathButtonState() {
  if (!setMediaPathBtn) {
    return;
  }
  if (state.mediaStoreMode === "directory") {
    setMediaPathBtn.textContent = "設置路徑（已設定）";
    setMediaPathBtn.title = "已設定專案資料夾（節點座標需按存檔才會寫入）";
  } else if (state.mediaStoreMode === "legacy") {
    setMediaPathBtn.textContent = "設置路徑（舊版設定）";
    setMediaPathBtn.title = "目前僅有媒體資料夾設定，請重新設置專案資料夾";
  } else if (state.mediaStoreMode === "opfs") {
    setMediaPathBtn.textContent = "設置路徑（目前內建儲存）";
    setMediaPathBtn.title = "目前使用瀏覽器內建儲存（節點座標需按存檔才會寫入）";
  } else {
    setMediaPathBtn.textContent = "設置路徑";
    setMediaPathBtn.title = "指定專案資料夾（含 media 與存檔）";
  }
}

async function pickAndPersistProjectDirectory() {
  if (typeof window.showDirectoryPicker !== "function") {
    throw new Error("Directory picker is not supported");
  }

  const rootHandle = await window.showDirectoryPicker({ mode: "readwrite", id: "node-editor-media-root" });
  const mediaDir = await rootHandle.getDirectoryHandle("media", { create: true });
  state.projectRootHandle = rootHandle;
  state.mediaStoreHandle = mediaDir;
  state.mediaStoreMode = "directory";
  await saveHandleToStorage(PROJECT_ROOT_HANDLE_KEY, rootHandle);
  updateSetPathButtonState();
  return rootHandle;
}

function openMediaSettingsDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(MEDIA_HANDLE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(MEDIA_HANDLE_STORE)) {
        db.createObjectStore(MEDIA_HANDLE_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveHandleToStorage(key, handle) {
  try {
    const db = await openMediaSettingsDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(MEDIA_HANDLE_STORE, "readwrite");
      tx.objectStore(MEDIA_HANDLE_STORE).put(handle, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
    db.close();
  } catch (error) {
    console.error("save directory handle failed", error);
  }
}

async function clearHandleFromStorage(key) {
  try {
    const db = await openMediaSettingsDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(MEDIA_HANDLE_STORE, "readwrite");
      tx.objectStore(MEDIA_HANDLE_STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
    db.close();
  } catch (error) {
    console.error("clear directory handle failed", error);
  }
}

async function loadHandleFromStorage(key) {
  try {
    const db = await openMediaSettingsDb();
    const handle = await new Promise((resolve, reject) => {
      const tx = db.transaction(MEDIA_HANDLE_STORE, "readonly");
      const request = tx.objectStore(MEDIA_HANDLE_STORE).get(key);
      request.onsuccess = () => resolve(request.result ?? null);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return handle;
  } catch {
    return null;
  }
}

async function verifyHandlePermission(handle, canRequest, mode = "readwrite") {
  if (!handle || typeof handle.queryPermission !== "function") {
    return false;
  }

  try {
    const status = await handle.queryPermission({ mode });
    if (status === "granted") {
      return true;
    }
    if (!canRequest || typeof handle.requestPermission !== "function") {
      return false;
    }
    const requested = await handle.requestPermission({ mode });
    return requested === "granted";
  } catch {
    return false;
  }
}

function extractFilesFromClipboardEvent(event) {
  const files = [];
  const items = Array.from(event.clipboardData?.items ?? []);
  items.forEach((item) => {
    if (item.kind !== "file") {
      return;
    }
    const file = item.getAsFile();
    if (file) {
      files.push(file);
    }
  });
  return files;
}

function pointerInsideEditor() {
  return (
    state.pointerScreen.x >= 0 &&
    state.pointerScreen.y >= 0 &&
    state.pointerScreen.x <= editor.clientWidth &&
    state.pointerScreen.y <= editor.clientHeight
  );
}

async function importMediaFiles(files, source, anchorWorld) {
  const importAnchor = anchorWorld ?? getViewportCenterWorld();
  const prepared = [];

  for (const file of files) {
    try {
      const fileHash = await computeSha256(file);
      const duplicate = findDuplicateMediaManifestItem(file, fileHash);
      const importMode = duplicate ? promptDuplicateMediaImportMode(file.name) : "rename";
      const storage = await storeMediaFileInLibrary(file, {
        mode: importMode,
        duplicateIndex: duplicate?.index ?? -1,
        duplicateEntry: duplicate?.item ?? null,
        fileHash,
      });
      const metadata = await buildMediaMetadata(file, source, storage, fileHash);
      prepared.push(metadata);
    } catch (error) {
      console.error("media import failed", error);
    }
  }

  if (prepared.length === 0) {
    return false;
  }

  const cols = Math.max(1, Math.min(3, Math.ceil(Math.sqrt(prepared.length))));
  const gapX = 280;
  const gapY = 150;
  const startX = importAnchor.x - ((Math.min(prepared.length, cols) - 1) * gapX) / 2;
  const startY = importAnchor.y - (Math.ceil(prepared.length / cols) - 1) * gapY * 0.5;
  const createdNodeIds = [];

  prepared.forEach((record, index) => {
    const col = index % cols;
    const row = Math.floor(index / cols);
    const node = createNode(startX + col * gapX, startY + row * gapY, {
      title: record.file.name,
      content: buildMediaNodeContent(record),
      metadata: record,
    });
    createdNodeIds.push(node.id);
  });

  setSelection(createdNodeIds);
  commitHistory();
  return true;
}

function buildMediaNodeContent(record) {
  if (!isMediaMetadata(record)) {
    return "雙擊編輯 JSON";
  }

  const parts = [];
  const kind = record.file?.kind;
  const sizeText = record.file?.sizeHuman;
  const kindLabel = {
    image: "圖片",
    video: "影片",
    audio: "音訊",
    text: "文字",
    json: "JSON",
    binary: "檔案",
  }[kind] ?? "媒體";
  parts.push(kindLabel);
  if (typeof sizeText === "string" && sizeText.trim() !== "") {
    parts.push(sizeText);
  }

  const width = record.analysis?.width;
  const height = record.analysis?.height;
  if (Number.isFinite(width) && Number.isFinite(height)) {
    parts.push(`${width}x${height}`);
  }

  parts.push("雙擊編輯 JSON");
  return parts.join(" · ");
}

function findDuplicateMediaManifestItem(file, fileHash) {
  if (!file || typeof file.name !== "string") {
    return null;
  }

  const normalizedName = file.name.trim();
  if (!normalizedName) {
    return null;
  }

  const fallbackMimeType = file.type || "application/octet-stream";
  const fallbackSize = Number(file.size);

  for (let index = state.mediaManifest.length - 1; index >= 0; index -= 1) {
    const item = state.mediaManifest[index];
    if (!item || typeof item !== "object") {
      continue;
    }
    if (String(item.originalName ?? "").trim() !== normalizedName) {
      continue;
    }

    const hashMatches =
      typeof fileHash === "string" &&
      fileHash !== "" &&
      typeof item.sha256 === "string" &&
      item.sha256 !== "" &&
      item.sha256 === fileHash;
    if (hashMatches) {
      return { index, item };
    }

    const hashMissing = !fileHash || !item.sha256;
    if (hashMissing) {
      const sizeMatches = Number(item.sizeBytes) === fallbackSize;
      const mimeMatches = String(item.mimeType || "application/octet-stream") === fallbackMimeType;
      if (sizeMatches && mimeMatches) {
        return { index, item };
      }
    }
  }

  return null;
}

function promptDuplicateMediaImportMode(fileName) {
  const normalizedName = String(fileName ?? "").trim() || "未命名媒體";
  const overwrite = window.confirm(
    `偵測到相同名稱且相同內容的媒體：${normalizedName}\n\n要覆蓋既有檔案嗎？\n按「確定」= 覆蓋\n按「取消」= 改名後保留`
  );
  return overwrite ? "overwrite" : "rename";
}

async function buildMediaMetadata(file, source, storage, precomputedHash = null) {
  const importedAt = new Date().toISOString();
  const extension = getFileExtension(file.name);
  const mediaKind = resolveMediaKind(file);
  const metadata = {
    schema: MEDIA_NODE_SCHEMA,
    version: 1,
    source,
    importedAt,
    file: {
      name: file.name,
      extension,
      mimeType: file.type || "application/octet-stream",
      sizeBytes: file.size,
      sizeHuman: formatBytes(file.size),
      lastModified: file.lastModified,
      lastModifiedISO: file.lastModified ? new Date(file.lastModified).toISOString() : null,
      kind: mediaKind,
    },
    storage,
  };

  const preview = await buildMediaPreview(file, mediaKind);
  if (preview) {
    metadata.preview = preview;
  }

  const detail = await extractRichMediaInfo(file, mediaKind);
  if (detail) {
    metadata.analysis = detail;
  }

  const checksum = precomputedHash || (await computeSha256(file));
  if (checksum) {
    metadata.hash = { sha256: checksum };
  }

  return metadata;
}

async function buildMediaPreview(file, mediaKind) {
  if (mediaKind === "image") {
    return buildImageThumbnail(file);
  }
  if (mediaKind === "video") {
    return buildVideoThumbnail(file);
  }
  return null;
}

async function buildImageThumbnail(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    const source = await new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error("image preview decode failed"));
      image.src = objectUrl;
    });
    return renderThumbnailFromSource(source, source.naturalWidth, source.naturalHeight);
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function buildVideoThumbnail(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    const video = await new Promise((resolve, reject) => {
      const el = document.createElement("video");
      el.preload = "metadata";
      el.muted = true;
      el.playsInline = true;
      el.onloadeddata = () => resolve(el);
      el.onerror = () => reject(new Error("video preview decode failed"));
      el.src = objectUrl;
    });
    return renderThumbnailFromSource(video, video.videoWidth, video.videoHeight);
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function renderThumbnailFromSource(source, width, height) {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return null;
  }
  const scale = Math.min(THUMBNAIL_MAX_WIDTH / width, THUMBNAIL_MAX_HEIGHT / height, 1);
  const targetWidth = Math.max(1, Math.round(width * scale));
  const targetHeight = Math.max(1, Math.round(height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = targetWidth;
  canvas.height = targetHeight;
  const context = canvas.getContext("2d");
  if (!context) {
    return null;
  }
  context.drawImage(source, 0, 0, targetWidth, targetHeight);
  const encoded = encodeThumbnailCanvas(canvas);
  return {
    kind: "image",
    width: targetWidth,
    height: targetHeight,
    mimeType: encoded.mimeType,
    thumbnailDataUrl: encoded.dataUrl,
  };
}

function encodeThumbnailCanvas(canvas) {
  try {
    const webp = canvas.toDataURL("image/webp", 0.95);
    if (typeof webp === "string" && webp.startsWith("data:image/webp")) {
      return { mimeType: "image/webp", dataUrl: webp };
    }
  } catch {
    // fallback below
  }

  return {
    mimeType: "image/png",
    dataUrl: canvas.toDataURL("image/png"),
  };
}

function resolveMediaKind(file) {
  const type = file.type || "";
  if (type.startsWith("image/")) {
    return "image";
  }
  if (type.startsWith("video/")) {
    return "video";
  }
  if (type.startsWith("audio/")) {
    return "audio";
  }
  if (type.startsWith("text/")) {
    return "text";
  }
  if (type.includes("json")) {
    return "json";
  }
  return "binary";
}

async function extractRichMediaInfo(file, mediaKind) {
  if (mediaKind === "image") {
    return readImageInfo(file);
  }
  if (mediaKind === "video") {
    return readVideoInfo(file);
  }
  if (mediaKind === "audio") {
    return readAudioInfo(file);
  }
  if (mediaKind === "text" || mediaKind === "json") {
    return readTextInfo(file);
  }
  return null;
}

async function readImageInfo(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    return await new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
      image.onerror = () => reject(new Error("image decode failed"));
      image.src = objectUrl;
    });
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function readVideoInfo(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    return await new Promise((resolve, reject) => {
      const video = document.createElement("video");
      video.preload = "metadata";
      video.onloadedmetadata = () =>
        resolve({
          width: video.videoWidth,
          height: video.videoHeight,
          durationSec: Number.isFinite(video.duration) ? Number(video.duration.toFixed(3)) : null,
        });
      video.onerror = () => reject(new Error("video decode failed"));
      video.src = objectUrl;
    });
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function readAudioInfo(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    return await new Promise((resolve, reject) => {
      const audio = document.createElement("audio");
      audio.preload = "metadata";
      audio.onloadedmetadata = () =>
        resolve({
          durationSec: Number.isFinite(audio.duration) ? Number(audio.duration.toFixed(3)) : null,
        });
      audio.onerror = () => reject(new Error("audio decode failed"));
      audio.src = objectUrl;
    });
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function readTextInfo(file) {
  try {
    const text = await file.text();
    const normalized = text.replace(/\r\n/g, "\n");
    return {
      chars: normalized.length,
      lines: normalized === "" ? 0 : normalized.split("\n").length,
      preview: normalized.slice(0, 240),
    };
  } catch {
    return null;
  }
}

async function computeSha256(file) {
  if (!crypto?.subtle) {
    return null;
  }
  try {
    const buffer = await file.arrayBuffer();
    const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
    return [...new Uint8Array(hashBuffer)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
  } catch {
    return null;
  }
}

async function ensureProjectRootDirectory(options = {}) {
  const {
    canRequestPermission = true,
    allowPrompt = true,
    allowOpfs = true,
    useStoredHandle = true,
  } = options;

  if (state.projectRootHandle) {
    const allowed = await verifyHandlePermission(state.projectRootHandle, canRequestPermission);
    if (allowed) {
      return state.projectRootHandle;
    }
    state.projectRootHandle = null;
    if (state.mediaStoreMode === "directory") {
      state.mediaStoreHandle = null;
      state.mediaStoreMode = null;
      updateSetPathButtonState();
    }
  }

  if (useStoredHandle) {
    const storedProjectHandle = await loadHandleFromStorage(PROJECT_ROOT_HANDLE_KEY);
    if (storedProjectHandle) {
      const allowed = await verifyHandlePermission(storedProjectHandle, canRequestPermission);
      if (allowed) {
        state.projectRootHandle = storedProjectHandle;
        state.mediaStoreMode = "directory";
        updateSetPathButtonState();
        return storedProjectHandle;
      }
    }
  }

  if (allowPrompt && typeof window.showDirectoryPicker === "function") {
    try {
      return await pickAndPersistProjectDirectory();
    } catch {
      // fallback below
    }
  }

  if (allowOpfs && navigator.storage?.getDirectory) {
    const opfsRoot = await navigator.storage.getDirectory();
    state.projectRootHandle = opfsRoot;
    state.mediaStoreMode = "opfs";
    updateSetPathButtonState();
    return opfsRoot;
  }

  return null;
}

async function ensureMediaDirectory() {
  const projectRoot = await ensureProjectRootDirectory({
    canRequestPermission: true,
    allowPrompt: true,
    allowOpfs: true,
    useStoredHandle: true,
  });
  if (projectRoot) {
    try {
      const mediaDir = await projectRoot.getDirectoryHandle("media", { create: true });
      state.mediaStoreHandle = mediaDir;
      return mediaDir;
    } catch (error) {
      console.error("resolve media directory failed", error);
    }
  }

  if (state.mediaStoreHandle) {
    const allowed = await verifyHandlePermission(state.mediaStoreHandle, true);
    if (allowed) {
      if (state.mediaStoreMode == null) {
        state.mediaStoreMode = "legacy";
      }
      updateSetPathButtonState();
      return state.mediaStoreHandle;
    }
    state.mediaStoreHandle = null;
  }

  const legacyHandle = await loadHandleFromStorage(LEGACY_MEDIA_HANDLE_KEY);
  if (legacyHandle) {
    const allowed = await verifyHandlePermission(legacyHandle, true);
    if (allowed) {
      state.mediaStoreHandle = legacyHandle;
      state.mediaStoreMode = "legacy";
      updateSetPathButtonState();
      return legacyHandle;
    }
  }

  throw new Error("No writable directory for media");
}

function resolveStorageFileNameFromManifestEntry(entry) {
  if (!entry || typeof entry !== "object") {
    return "";
  }
  const directName = typeof entry.fileName === "string" ? entry.fileName.trim() : "";
  if (directName) {
    return directName;
  }

  const pathValue =
    typeof entry.storagePath === "string" && entry.storagePath.trim() !== ""
      ? entry.storagePath.trim()
      : typeof entry.path === "string" && entry.path.trim() !== ""
        ? entry.path.trim()
        : "";
  if (pathValue) {
    const normalized = pathValue.replace(/\\/g, "/");
    const segments = normalized.split("/").filter((part) => part && part !== ".");
    return segments.length > 0 ? segments[segments.length - 1] : "";
  }
  return "";
}

function normalizeMediaStoragePath(pathValue) {
  if (typeof pathValue !== "string") {
    return "";
  }
  return pathValue.replace(/\\/g, "/").replace(/^\.?\//, "").replace(/\/+/g, "/").trim();
}

function buildMediaStorageKeyFromEntry(entry) {
  if (!entry || typeof entry !== "object") {
    return "";
  }

  const rawPath =
    typeof entry.path === "string" && entry.path.trim() !== ""
      ? entry.path
      : typeof entry.storagePath === "string" && entry.storagePath.trim() !== ""
        ? entry.storagePath
        : "";
  const normalizedPath = normalizeMediaStoragePath(rawPath);
  if (normalizedPath) {
    return normalizedPath.toLowerCase();
  }

  const fileName = resolveStorageFileNameFromManifestEntry(entry);
  if (!fileName) {
    return "";
  }
  const folder = typeof entry.folder === "string" && entry.folder.trim() !== "" ? entry.folder.trim() : "media";
  return normalizeMediaStoragePath(`${folder}/${fileName}`).toLowerCase();
}

async function storeMediaFileInLibrary(file, options = {}) {
  const {
    mode = "rename",
    duplicateIndex = -1,
    duplicateEntry = null,
    fileHash = null,
  } = options;
  const directory = await ensureMediaDirectory();
  let fileName = buildStoredFileName(file.name);

  if (mode === "overwrite" && duplicateEntry) {
    const existingFileName = resolveStorageFileNameFromManifestEntry(duplicateEntry);
    if (existingFileName) {
      fileName = existingFileName;
    }
  }

  const fileHandle = await directory.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(file);
  await writable.close();

  const storage = {
    folder: "media",
    fileName,
    path: `media/${fileName}`,
    mode: state.mediaStoreMode ?? "opfs",
  };

  const manifestItem = {
    importedAt: new Date().toISOString(),
    originalName: file.name,
    fileName,
    storagePath: storage.path,
    sizeBytes: file.size,
    mimeType: file.type || "application/octet-stream",
    sha256: typeof fileHash === "string" && fileHash !== "" ? fileHash : null,
  };

  if (mode === "overwrite" && duplicateIndex >= 0 && duplicateIndex < state.mediaManifest.length) {
    state.mediaManifest[duplicateIndex] = manifestItem;
  } else {
    state.mediaManifest.push(manifestItem);
  }
  persistMediaManifest();
  return storage;
}

async function persistMediaManifest() {
  if (!state.mediaStoreHandle) {
    return;
  }
  try {
    const fileHandle = await state.mediaStoreHandle.getFileHandle(MEDIA_MANIFEST_FILE, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(
      JSON.stringify(
        {
          schema: `${MEDIA_NODE_SCHEMA}.manifest`,
          version: 1,
          updatedAt: new Date().toISOString(),
          items: state.mediaManifest,
        },
        null,
        2
      )
    );
    await writable.close();
  } catch (error) {
    console.error("manifest write failed", error);
  }
}

function buildStoredFileName(originalName) {
  const base = sanitizeFileName(removeFileExtension(originalName)).slice(0, 48) || "media";
  const ext = getFileExtension(originalName);
  const timestamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
  const nonce = Math.random().toString(36).slice(2, 8);
  return ext ? `${timestamp}_${nonce}_${base}.${ext}` : `${timestamp}_${nonce}_${base}`;
}

function removeFileExtension(fileName) {
  const dot = fileName.lastIndexOf(".");
  if (dot <= 0) {
    return fileName;
  }
  return fileName.slice(0, dot);
}

function getFileExtension(fileName) {
  const dot = fileName.lastIndexOf(".");
  if (dot <= 0 || dot === fileName.length - 1) {
    return "";
  }
  return fileName.slice(dot + 1).toLowerCase();
}

function sanitizeFileName(value) {
  return String(value ?? "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\s+/g, "_")
    .trim();
}

function formatBytes(size) {
  const units = ["B", "KB", "MB", "GB", "TB"];
  let value = Number(size);
  let index = 0;
  while (value >= 1024 && index < units.length - 1) {
    value /= 1024;
    index += 1;
  }
  const digits = value >= 10 || index === 0 ? 0 : 1;
  return `${value.toFixed(digits)} ${units[index]}`;
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
  updateEditorGridBackground();

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
  state.projectLastSavedHash = state.history.lastHash;
  state.projectLastSavedSnapshot = cloneSnapshot(snapshot);
  updateHistoryToolbarButtonsState();
}

function updateHistoryToolbarButtonsState() {
  if (undoBtn) {
    undoBtn.disabled = state.history.undo.length <= 1;
    undoBtn.setAttribute("aria-disabled", String(undoBtn.disabled));
  }
  if (redoBtn) {
    redoBtn.disabled = state.history.redo.length === 0;
    redoBtn.setAttribute("aria-disabled", String(redoBtn.disabled));
  }
}

function captureSnapshot() {
  const nodes = Array.from(state.nodes.values())
    .map((node) => ({
      id: node.id,
      x: node.x,
      y: node.y,
      title: node.title,
      content: node.content,
      metadata: node.metadata ?? null,
    }))
    .sort((a, b) => Number(a.id) - Number(b.id));

  const connections = state.connections
    .map((connection) => ({ ...connection }))
    .sort((a, b) => Number(a.id) - Number(b.id));

  const selectedNodeIds = Array.from(state.selectedNodeIds).sort((a, b) => Number(a) - Number(b));

  return {
    nextNodeId: state.nextNodeId,
    nextConnectionId: state.nextConnectionId,
    nodeDefinitions: normalizeNodeDefinitionsPayload(state.nodeDefinitions),
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
  state.nodeDefinitions = normalizeNodeDefinitionsPayload(snapshot.nodeDefinitions);
  renderDefinedNodeTypeList();

  snapshot.nodes.forEach((nodeData) => {
    createNode(nodeData.x, nodeData.y, {
      id: nodeData.id,
      title: nodeData.title,
      content: nodeData.content,
      metadata: nodeData.metadata ?? null,
    });
  });

  state.connections = snapshot.connections
    .map((connection) => {
      const fromSide = normalizeHandleSideKey(connection.fromSide);
      const toSide = normalizeHandleSideKey(connection.toSide);
      if (!fromSide || !toSide) {
        return null;
      }
      return {
        ...connection,
        fromSide,
        toSide,
      };
    })
    .filter(Boolean);
  state.nextNodeId = snapshot.nextNodeId;
  state.nextConnectionId = snapshot.nextConnectionId;
  refreshBindingNodes();
  setSelection(snapshot.selectedNodeIds);

  state.hoverConnectionId = null;
  lineDeleteBtn.classList.remove("visible");
  editor.classList.remove("panning");
  renderConnections();
  updateNodeProximity();
}

function commitHistory() {
  const snapshot = captureSnapshot();
  const previousSnapshot = state.history.undo[state.history.undo.length - 1] ?? null;
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
  persistProjectStateDraft(snapshot);
  updateHistoryToolbarButtonsState();
  if (shouldAutoPersistSnapshotChange(previousSnapshot, snapshot)) {
    queueProjectAutosave(snapshot, hash);
  }
  queueGitSync(snapshot, hash);
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
  persistProjectStateDraft(previous);
  updateHistoryToolbarButtonsState();
  if (shouldAutoPersistSnapshotChange(current, previous)) {
    queueProjectAutosave(previous, state.history.lastHash);
  }
  queueGitSync(previous, state.history.lastHash);
}

function redoHistory() {
  if (state.history.redo.length === 0) {
    return;
  }
  const snapshot = state.history.redo.pop();
  const previousSnapshot = state.history.undo[state.history.undo.length - 1] ?? null;
  state.history.undo.push(snapshot);
  restoreSnapshot(snapshot);
  state.history.lastHash = hashSnapshot(snapshot);
  persistProjectStateDraft(snapshot);
  updateHistoryToolbarButtonsState();
  if (shouldAutoPersistSnapshotChange(previousSnapshot, snapshot)) {
    queueProjectAutosave(snapshot, state.history.lastHash);
  }
  queueGitSync(snapshot, state.history.lastHash);
}

function shouldAutoPersistSnapshotChange(previousSnapshot, nextSnapshot) {
  if (!previousSnapshot || !nextSnapshot) {
    return true;
  }
  return buildNonPositionSnapshotHash(previousSnapshot) !== buildNonPositionSnapshotHash(nextSnapshot);
}

function buildNonPositionSnapshotHash(snapshot) {
  return JSON.stringify({
    nextNodeId: snapshot.nextNodeId,
    nextConnectionId: snapshot.nextConnectionId,
    nodeDefinitions: Array.isArray(snapshot.nodeDefinitions) ? snapshot.nodeDefinitions : [],
    nodes: snapshot.nodes.map((node) => ({
      id: node.id,
      title: node.title,
      content: node.content,
      metadata: node.metadata ?? null,
    })),
    connections: snapshot.connections.map((connection) => ({ ...connection })),
  });
}

function queueProjectAutosave(snapshot, hash) {
  if (!snapshot || !hash || hash === state.projectLastSavedHash) {
    return;
  }

  state.projectSaveQueued = { snapshot, hash };
  if (state.projectSaveInFlight) {
    return;
  }
  void flushProjectAutosaveQueue();
}

async function flushProjectAutosaveQueue() {
  if (state.projectSaveInFlight) {
    return;
  }
  state.projectSaveInFlight = true;

  while (state.projectSaveQueued) {
    const job = state.projectSaveQueued;
    state.projectSaveQueued = null;
    try {
      const result = await persistProjectState(job.snapshot, job.hash);
      if (result.saved) {
        state.projectLastSavedHash = job.hash;
        state.projectLastSavedSnapshot = cloneSnapshot(result.persistedSnapshot);
      }
    } catch (error) {
      console.error("project autosave failed", error);
    }
  }

  state.projectSaveInFlight = false;
}

async function saveProjectStateNow() {
  const snapshot = captureSnapshot();
  const hash = hashSnapshot(snapshot);
  const gitResult = await syncProjectStateToGit(snapshot, hash, { immediate: true, source: "manual-save" });
  if (gitResult.reason === "already synced") {
    const localResult = await persistProjectState(snapshot, hash, {
      canRequestPermission: false,
      allowPrompt: false,
      lockNodePositions: false,
    });
    if (localResult.saved) {
      state.projectLastSavedHash = hash;
      state.projectLastSavedSnapshot = cloneSnapshot(localResult.persistedSnapshot);
    }
    return true;
  }
  if (gitResult.saved) {
    const localResult = await persistProjectState(snapshot, hash, {
      canRequestPermission: false,
      allowPrompt: false,
      lockNodePositions: false,
    });
    if (localResult.saved) {
      state.projectLastSavedHash = hash;
      state.projectLastSavedSnapshot = cloneSnapshot(localResult.persistedSnapshot);
    }
    return true;
  }

  const localResult = await persistProjectState(snapshot, hash, {
    canRequestPermission: true,
    allowPrompt: true,
    lockNodePositions: false,
  });
  if (localResult.saved) {
    state.projectLastSavedHash = hash;
    state.projectLastSavedSnapshot = cloneSnapshot(localResult.persistedSnapshot);
  }
  return localResult.saved || gitResult.saved;
}

async function setProjectStateFileHandle(fileHandle) {
  if (!fileHandle) {
    return;
  }
  state.projectStateFileHandle = fileHandle;
  await saveHandleToStorage(PROJECT_FILE_HANDLE_KEY, fileHandle);
}

async function clearProjectStateFileBinding() {
  state.projectStateFileHandle = null;
  await clearHandleFromStorage(PROJECT_FILE_HANDLE_KEY);
}

async function resolveProjectStateFileHandle(options = {}) {
  const { forWrite = false, canRequestPermission = false, allowPrompt = false } = options;
  const mode = forWrite ? "readwrite" : "read";

  if (state.projectStateFileHandle) {
    const allowed = await verifyHandlePermission(state.projectStateFileHandle, canRequestPermission, mode);
    if (allowed) {
      return state.projectStateFileHandle;
    }
    state.projectStateFileHandle = null;
  }

  const storedFileHandle = await loadHandleFromStorage(PROJECT_FILE_HANDLE_KEY);
  if (storedFileHandle) {
    const allowed = await verifyHandlePermission(storedFileHandle, canRequestPermission, mode);
    if (allowed) {
      state.projectStateFileHandle = storedFileHandle;
      return storedFileHandle;
    }
  }

  const projectRoot = await ensureProjectRootDirectory({
    canRequestPermission,
    allowPrompt: false,
    allowOpfs: false,
    useStoredHandle: true,
  });
  if (projectRoot) {
    try {
      const fileHandle = await projectRoot.getFileHandle(PROJECT_STATE_FILE, { create: forWrite });
      const allowed = await verifyHandlePermission(fileHandle, canRequestPermission, mode);
      if (allowed) {
        await setProjectStateFileHandle(fileHandle);
        return fileHandle;
      }
    } catch {
      // keep fallback path below
    }
  }

  if (forWrite && allowPrompt && typeof window.showSaveFilePicker === "function") {
    const fileHandle = await window.showSaveFilePicker({
      id: "node-editor-project-save",
      suggestedName: PROJECT_STATE_FILE,
      types: [
        {
          description: "Note Editor Project",
          accept: {
            "application/json": [".json"],
          },
        },
      ],
    });
    await setProjectStateFileHandle(fileHandle);
    return fileHandle;
  }

  return null;
}

async function persistProjectState(snapshot, hash, directoryOptions = {}) {
  const snapshotToPersist = buildSnapshotForPersistence(snapshot, directoryOptions.lockNodePositions !== false);
  const fileHandle = await resolveProjectStateFileHandle({
    forWrite: true,
    canRequestPermission: Boolean(directoryOptions.canRequestPermission),
    allowPrompt: Boolean(directoryOptions.allowPrompt),
  });
  if (!fileHandle) {
    return { saved: false, persistedSnapshot: snapshotToPersist };
  }

  try {
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(buildProjectStatePayload(snapshotToPersist, hash), null, 2));
    await writable.close();
    return { saved: true, persistedSnapshot: snapshotToPersist };
  } catch (error) {
    console.error("write project state failed", error);
    return { saved: false, persistedSnapshot: snapshotToPersist };
  }
}

function buildSnapshotForPersistence(snapshot, lockNodePositions) {
  const cloned = cloneSnapshot(snapshot);
  if (!lockNodePositions || !state.projectLastSavedSnapshot) {
    return cloned;
  }

  const previousNodeMap = new Map(
    state.projectLastSavedSnapshot.nodes.map((node) => [String(node.id), { x: node.x, y: node.y }])
  );
  cloned.nodes = cloned.nodes.map((node) => {
    const previous = previousNodeMap.get(String(node.id));
    if (!previous) {
      return node;
    }
    return {
      ...node,
      x: previous.x,
      y: previous.y,
    };
  });
  return cloned;
}

function cloneSnapshot(snapshot) {
  return {
    nextNodeId: snapshot.nextNodeId,
    nextConnectionId: snapshot.nextConnectionId,
    nodeDefinitions: normalizeNodeDefinitionsPayload(snapshot.nodeDefinitions),
    nodes: snapshot.nodes.map((node) => ({
      id: node.id,
      x: node.x,
      y: node.y,
      title: node.title,
      content: node.content,
      metadata: node.metadata ?? null,
    })),
    connections: snapshot.connections.map((connection) => ({ ...connection })),
    selectedNodeIds: [...snapshot.selectedNodeIds],
  };
}

function buildProjectStatePayload(snapshot, hash) {
  return {
    type: PROJECT_STATE_SCHEMA,
    version: PROJECT_STATE_VERSION,
    savedAt: new Date().toISOString(),
    hash,
    viewport: {
      x: state.viewport.x,
      y: state.viewport.y,
      zoom: state.zoom,
    },
    nextNodeId: snapshot.nextNodeId,
    nextConnectionId: snapshot.nextConnectionId,
    nodeDefinitions: normalizeNodeDefinitionsPayload(snapshot.nodeDefinitions),
    nodes: snapshot.nodes.map((node) => ({
      id: node.id,
      x: node.x,
      y: node.y,
      title: node.title,
      content: node.content,
      metadata: node.metadata ?? null,
    })),
    connections: snapshot.connections.map((connection) => ({ ...connection })),
    selectedNodeIds: [...snapshot.selectedNodeIds],
    media: {
      folder: "media",
      manifestFile: MEDIA_MANIFEST_FILE,
      manifestItems: state.mediaManifest.map((item) => ({ ...item })),
    },
  };
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

function positiveModulo(value, divisor) {
  if (!Number.isFinite(value) || !Number.isFinite(divisor) || divisor <= 0) {
    return 0;
  }
  return ((value % divisor) + divisor) % divisor;
}

function updateEditorGridBackground() {
  const spacing = Math.max(4, GRID_WORLD_SIZE * state.zoom);
  const offsetX = positiveModulo(state.viewport.x, spacing);
  const offsetY = positiveModulo(state.viewport.y, spacing);
  editor.style.setProperty("--grid-size", `${spacing}px`);
  editor.style.setProperty("--grid-offset-x", `${offsetX}px`);
  editor.style.setProperty("--grid-offset-y", `${offsetY}px`);
}

function inferGitSyncDefaultsFromLocation() {
  const defaults = {
    owner: "",
    repo: "",
    branch: "main",
    path: PROJECT_STATE_FILE,
    token: "",
    autosync: false,
    loadOnStartup: false,
    commitMessage: GIT_SYNC_DEFAULT_COMMIT_MESSAGE,
  };

  const hostname = String(window.location.hostname ?? "").toLowerCase();
  if (!hostname.endsWith(".github.io")) {
    return defaults;
  }

  const owner = hostname.split(".")[0]?.trim();
  if (owner) {
    defaults.owner = owner;
  }

  const pathSegments = String(window.location.pathname ?? "")
    .split("/")
    .map((segment) => segment.trim())
    .filter(Boolean);
  if (pathSegments.length > 0) {
    try {
      defaults.repo = decodeURIComponent(pathSegments[0]);
    } catch {
      defaults.repo = pathSegments[0];
    }
    if (defaults.repo) {
      defaults.branch = "master";
    }
  }

  return defaults;
}

function normalizeGitSyncPath(rawPath) {
  const segments = String(rawPath ?? "")
    .split("/")
    .map((segment) => segment.trim())
    .filter(Boolean);
  return segments.length > 0 ? segments.join("/") : PROJECT_STATE_FILE;
}

function normalizeGitSyncSettings(rawSettings) {
  const defaults = inferGitSyncDefaultsFromLocation();
  const source = isPlainRecord(rawSettings) ? rawSettings : {};
  const owner = String(source.owner ?? "").trim() || defaults.owner;
  const repo = String(source.repo ?? "").trim() || defaults.repo;
  const branch = String(source.branch ?? "").trim() || defaults.branch;
  const path = normalizeGitSyncPath(source.path ?? defaults.path);

  return {
    owner,
    repo,
    branch,
    path,
    token: String(source.token ?? "").trim(),
    autosync: Boolean(source.autosync),
    loadOnStartup: Boolean(source.loadOnStartup),
    commitMessage: String(source.commitMessage ?? "").trim() || defaults.commitMessage,
  };
}

function loadGitSyncSettingsFromStorage() {
  const defaults = normalizeGitSyncSettings({});
  try {
    const raw = localStorage.getItem(GIT_SYNC_STORAGE_KEY);
    if (!raw) {
      return defaults;
    }
    const parsed = JSON.parse(raw);
    const source = isPlainRecord(parsed?.settings) ? parsed.settings : parsed;
    return normalizeGitSyncSettings(source);
  } catch (error) {
    console.warn("load git sync settings failed", error);
    return defaults;
  }
}

function saveGitSyncSettingsToStorage(settings) {
  try {
    localStorage.setItem(
      GIT_SYNC_STORAGE_KEY,
      JSON.stringify({
        version: 1,
        settings: normalizeGitSyncSettings(settings),
      })
    );
    return true;
  } catch (error) {
    console.warn("save git sync settings failed", error);
    return false;
  }
}

function getGitSyncSettings() {
  if (!state.gitSync.settings) {
    state.gitSync.settings = loadGitSyncSettingsFromStorage();
  }
  return state.gitSync.settings;
}

function isGitSyncTargetConfigured(settings = getGitSyncSettings()) {
  return Boolean(settings?.owner && settings?.repo && settings?.path);
}

function isGitSyncWriteReady(settings = getGitSyncSettings()) {
  return isGitSyncTargetConfigured(settings) && Boolean(settings?.token);
}

function buildGitSyncTargetLabel(settings = getGitSyncSettings()) {
  if (!isGitSyncTargetConfigured(settings)) {
    return "未設定";
  }
  return `${settings.owner}/${settings.repo}@${settings.branch}:${settings.path}`;
}

function renderGitSyncMessage(message, kind = "") {
  if (!gitSyncMessageEl) {
    return;
  }
  gitSyncMessageEl.textContent = String(message ?? "");
  gitSyncMessageEl.className = "git-sync-message";
  if (kind) {
    gitSyncMessageEl.classList.add(kind);
  }
}

function updateGitSyncStatusIndicator() {
  if (!gitSyncStatusEl) {
    return;
  }

  const settings = getGitSyncSettings();
  let label = "備份: 關閉";
  let kind = "";

  if (state.gitSync.inFlight) {
    label = "備份: 同步中";
    kind = "syncing";
  } else if (state.gitSync.lastError) {
    label = "備份: 錯誤";
    kind = "error";
  } else if (isGitSyncWriteReady(settings)) {
    label = settings.autosync ? "備份: 自動" : "備份: 就緒";
    kind = "ready";
  } else if (isGitSyncTargetConfigured(settings)) {
    label = "備份: 唯讀";
    kind = "ready";
  }

  gitSyncStatusEl.textContent = label;
  gitSyncStatusEl.className = "git-sync-status";
  if (kind) {
    gitSyncStatusEl.classList.add(kind);
  }
  const statusSuffix = state.gitSync.lastError
    ? state.gitSync.lastError
    : state.gitSync.lastMessage || (!isGitSyncTargetConfigured(settings) ? "同步備份未啟用" : "");
  gitSyncStatusEl.title = [buildGitSyncTargetLabel(settings), statusSuffix].filter(Boolean).join(" | ");
  gitSyncStatusEl.dataset.status = statusSuffix;
}

function updateGitSyncPanelState() {
  const settings = getGitSyncSettings();
  const hasTarget = isGitSyncTargetConfigured(settings);
  const canWrite = isGitSyncWriteReady(settings);
  const hasQueuedChange =
    Boolean(state.gitSync.queuedSnapshot) && Boolean(state.gitSync.queuedHash) && state.gitSync.queuedHash !== state.gitSync.lastSyncedHash;

  if (gitSyncLoadBtn) {
    gitSyncLoadBtn.disabled = state.gitSync.inFlight || !hasTarget;
  }
  if (gitSyncNowBtn) {
    gitSyncNowBtn.disabled = state.gitSync.inFlight || !canWrite;
  }
  if (gitSyncClearBtn) {
    gitSyncClearBtn.disabled = state.gitSync.inFlight || !settings.token;
  }
  if (gitSyncAutosyncInput) {
    gitSyncAutosyncInput.disabled = state.gitSync.inFlight || !hasTarget;
  }
  if (gitSyncStartupInput) {
    gitSyncStartupInput.disabled = state.gitSync.inFlight || !hasTarget;
  }

  if (state.gitSync.lastError) {
    renderGitSyncMessage(state.gitSync.lastError, "error");
    updateGitSyncStatusIndicator();
    return;
  }

  if (hasQueuedChange) {
    renderGitSyncMessage(
    settings.autosync && canWrite ? "變更已排程備份，稍候會自動寫入雲端。" : "目前有未同步的變更，請按「立即備份」。",
      "warning"
    );
    updateGitSyncStatusIndicator();
    return;
  }

  if (state.gitSync.lastMessage) {
    renderGitSyncMessage(state.gitSync.lastMessage, state.gitSync.inFlight ? "warning" : "success");
    updateGitSyncStatusIndicator();
    return;
  }

  if (!hasTarget) {
    renderGitSyncMessage("請先設定 Owner / Repository / File path。", "warning");
  } else if (!settings.token) {
    renderGitSyncMessage("可讀取遠端檔案，寫入 Git 需要 Token。", "warning");
  } else if (settings.autosync) {
    renderGitSyncMessage("已啟用自動備份，變更會延遲寫入雲端。", "success");
  } else {
    renderGitSyncMessage("設定已就緒，按「立即備份」可寫入雲端。", "warning");
  }

  updateGitSyncStatusIndicator();
}

function renderGitSyncPanel() {
  const settings = getGitSyncSettings();
  if (gitSyncOwnerInput) {
    gitSyncOwnerInput.value = settings.owner;
  }
  if (gitSyncRepoInput) {
    gitSyncRepoInput.value = settings.repo;
  }
  if (gitSyncBranchInput) {
    gitSyncBranchInput.value = settings.branch;
  }
  if (gitSyncPathInput) {
    gitSyncPathInput.value = settings.path;
  }
  if (gitSyncTokenInput) {
    gitSyncTokenInput.value = settings.token;
  }
  if (gitSyncAutosyncInput) {
    gitSyncAutosyncInput.checked = Boolean(settings.autosync);
  }
  if (gitSyncStartupInput) {
    gitSyncStartupInput.checked = Boolean(settings.loadOnStartup);
  }

  updateGitSyncPanelState();
}

function readGitSyncSettingsFromInputs() {
  return normalizeGitSyncSettings({
    owner: gitSyncOwnerInput?.value ?? "",
    repo: gitSyncRepoInput?.value ?? "",
    branch: gitSyncBranchInput?.value ?? "",
    path: gitSyncPathInput?.value ?? "",
    token: gitSyncTokenInput?.value ?? "",
    autosync: Boolean(gitSyncAutosyncInput?.checked),
    loadOnStartup: Boolean(gitSyncStartupInput?.checked),
    commitMessage: getGitSyncSettings().commitMessage,
  });
}

function syncGitSyncSettingsFromInputs(options = {}) {
  const nextSettings = readGitSyncSettingsFromInputs();
  state.gitSync.settings = nextSettings;
  if (options.persist !== false) {
    saveGitSyncSettingsToStorage(nextSettings);
  }
  if (options.clearMessages !== false) {
    state.gitSync.lastMessage = "";
    state.gitSync.lastError = "";
  }
  updateGitSyncPanelState();

  if (nextSettings.autosync && state.gitSync.queuedSnapshot && isGitSyncWriteReady(nextSettings) && !state.gitSync.inFlight) {
    scheduleGitSyncFlush();
  }

  return nextSettings;
}

function onGitSyncSettingsInputChange() {
  syncGitSyncSettingsFromInputs({ persist: true, clearMessages: true });
}

async function initializeGitSyncFromStorage() {
  state.gitSync.settings = loadGitSyncSettingsFromStorage();
  renderGitSyncPanel();
  updateGitSyncStatusIndicator();
  return state.gitSync.settings;
}

async function onGitSyncLoadClick() {
  try {
    syncGitSyncSettingsFromInputs({ persist: true, clearMessages: true });
    if (!isGitSyncTargetConfigured()) {
      state.gitSync.lastError = "請先設定 Owner / Repository / File path。";
      state.gitSync.lastMessage = "";
      updateGitSyncPanelState();
      return;
    }

    const loaded = await loadProjectFromGit({ source: "manual-load" });
    if (!loaded && !state.gitSync.lastError && !state.gitSync.lastMessage) {
      state.gitSync.lastError = "從備份載入失敗。";
      state.gitSync.lastMessage = "";
      updateGitSyncPanelState();
    }
  } catch (error) {
    console.error("manual git load failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "從備份載入失敗";
    state.gitSync.lastMessage = "";
    updateGitSyncPanelState();
  }
}

async function onGitSyncNowClick() {
  try {
    syncGitSyncSettingsFromInputs({ persist: true, clearMessages: true });
    if (!isGitSyncWriteReady()) {
      state.gitSync.lastError = "寫入 Git 需要 Owner / Repository / File path 與 Token。";
      state.gitSync.lastMessage = "";
      updateGitSyncPanelState();
      return;
    }

    const snapshot = captureSnapshot();
    const hash = hashSnapshot(snapshot);
    const result = await syncProjectStateToGit(snapshot, hash, { immediate: true, source: "manual-now" });
    if (result.reason === "already synced") {
      state.gitSync.lastMessage = "備份已經是最新狀態。";
      state.gitSync.lastError = "";
      updateGitSyncPanelState();
    }
  } catch (error) {
    console.error("manual git sync failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "備份同步失敗";
    state.gitSync.lastMessage = "";
    updateGitSyncPanelState();
  }
}

async function onGitSyncSaveClick() {
  try {
    const settings = syncGitSyncSettingsFromInputs({ persist: true, clearMessages: false });
    if (!isGitSyncTargetConfigured(settings)) {
      state.gitSync.lastError = "請先設定 Owner / Repository / File path。";
      state.gitSync.lastMessage = "";
      updateGitSyncPanelState();
      return;
    }

    state.gitSync.lastMessage = "備份設定已儲存。";
    state.gitSync.lastError = "";
    updateGitSyncPanelState();
  } catch (error) {
    console.error("save git sync settings failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "儲存同步備份設定失敗";
    state.gitSync.lastMessage = "";
    updateGitSyncPanelState();
  }
}

async function onGitSyncClearTokenClick() {
  try {
    if (gitSyncTokenInput) {
      gitSyncTokenInput.value = "";
    }
    syncGitSyncSettingsFromInputs({ persist: true, clearMessages: true });
    state.gitSync.lastMessage = "GitHub Token 已清除。";
    state.gitSync.lastError = "";
    updateGitSyncPanelState();
  } catch (error) {
    console.error("clear git sync token failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "清除 Token 失敗";
    state.gitSync.lastMessage = "";
    updateGitSyncPanelState();
  }
}

function queueGitSync(snapshot, hash) {
  if (!snapshot || !hash) {
    return;
  }

  const settings = getGitSyncSettings();
  if (hash === state.gitSync.lastSyncedHash && !state.gitSync.inFlight) {
    state.gitSync.queuedSnapshot = null;
    state.gitSync.queuedHash = "";
    if (state.gitSync.timerId != null) {
      clearTimeout(state.gitSync.timerId);
      state.gitSync.timerId = null;
    }
    updateGitSyncStatusIndicator();
    return;
  }

  state.gitSync.queuedSnapshot = cloneSnapshot(snapshot);
  state.gitSync.queuedHash = hash;
  if (hash !== state.gitSync.lastSyncedHash) {
    state.gitSync.lastMessage = "";
  }
  if (state.gitSync.timerId != null) {
    clearTimeout(state.gitSync.timerId);
    state.gitSync.timerId = null;
  }
  if (!isGitSyncWriteReady(settings)) {
    updateGitSyncStatusIndicator();
    return;
  }
  scheduleGitSyncFlush();
}

function scheduleGitSyncFlush() {
  if (state.gitSync.timerId != null) {
    clearTimeout(state.gitSync.timerId);
  }
  state.gitSync.timerId = window.setTimeout(() => {
    state.gitSync.timerId = null;
    void flushGitSyncQueue();
  }, GIT_SYNC_DEBOUNCE_MS);
  updateGitSyncStatusIndicator();
}

async function flushGitSyncQueue() {
  if (state.gitSync.inFlight) {
    return false;
  }

  const settings = getGitSyncSettings();
  if (!isGitSyncWriteReady(settings) || !state.gitSync.queuedSnapshot || !state.gitSync.queuedHash) {
    updateGitSyncStatusIndicator();
    return false;
  }

  const snapshot = state.gitSync.queuedSnapshot;
  const hash = state.gitSync.queuedHash;
  state.gitSync.inFlight = true;
    state.gitSync.lastMessage = "正在同步備份...";
  state.gitSync.lastError = "";
  updateGitSyncPanelState();

  try {
    const result = await performGitSync(snapshot, hash, settings, { source: "autosync" });
    if (result.saved && state.gitSync.queuedHash === hash) {
      state.gitSync.queuedSnapshot = null;
      state.gitSync.queuedHash = "";
    }
    return result.saved;
  } catch (error) {
    console.error("git autosync failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "自動備份失敗";
    state.gitSync.lastMessage = "";
    return false;
  } finally {
    state.gitSync.inFlight = false;
    updateGitSyncPanelState();
    if (
      state.gitSync.queuedSnapshot &&
      state.gitSync.queuedHash &&
      state.gitSync.queuedHash !== hash &&
      getGitSyncSettings().autosync &&
      isGitSyncWriteReady(getGitSyncSettings())
    ) {
      scheduleGitSyncFlush();
    }
  }
}

async function syncProjectStateToGit(snapshot, hash, options = {}) {
  const settings = getGitSyncSettings();
  if (!snapshot || !hash || !isGitSyncWriteReady(settings)) {
    return { saved: false, reason: "git sync not ready" };
  }

  if (!options.immediate) {
    queueGitSync(snapshot, hash);
    return { saved: false, queued: true };
  }

  if (state.gitSync.inFlight) {
    await waitForGitSyncIdle();
  }

  return performGitSync(snapshot, hash, settings, options);
}

async function performGitSync(snapshot, hash, settings, options = {}) {
  if (!snapshot || !hash) {
    return { saved: false, reason: "missing snapshot" };
  }

  const resolvedSettings = normalizeGitSyncSettings(settings);
  if (!isGitSyncWriteReady(resolvedSettings)) {
    return { saved: false, reason: "git sync not ready" };
  }
  if (hash === state.gitSync.lastSyncedHash && state.gitSync.lastSyncedHash) {
      state.gitSync.lastMessage = "備份已經是最新狀態。";
    state.gitSync.lastError = "";
    updateGitSyncPanelState();
    return { saved: false, reason: "already synced" };
  }

  state.gitSync.inFlight = true;
  state.gitSync.lastError = "";
    state.gitSync.lastMessage = options.source === "manual-save" ? "正在儲存到備份..." : "正在同步備份...";
  updateGitSyncPanelState();

  try {
    const snapshotToPersist = buildSnapshotForPersistence(snapshot, false);
    const payload = buildProjectStatePayload(snapshotToPersist, hash);
    const payloadText = JSON.stringify(payload, null, 2);
    const fileRecord = await getGitHubFileRecord(resolvedSettings);
    const nextRemoteSha = fileRecord?.sha ? String(fileRecord.sha) : "";
    const writeResult = await putGitHubFileRecord(resolvedSettings, payloadText, nextRemoteSha);

    if (!writeResult.saved) {
      throw new Error(writeResult.errorMessage || "Git 寫入失敗");
    }

    state.gitSync.lastRemoteSha = writeResult.remoteSha || nextRemoteSha;
    state.gitSync.lastSyncedHash = hash;
    state.gitSync.lastSyncedAt = new Date().toISOString();
    state.gitSync.lastMessage = `已同步備份：${resolvedSettings.owner}/${resolvedSettings.repo}/${resolvedSettings.path}`;
    state.gitSync.lastError = "";
    if (state.gitSync.queuedHash === hash) {
      state.gitSync.queuedSnapshot = null;
      state.gitSync.queuedHash = "";
      if (state.gitSync.timerId != null) {
        clearTimeout(state.gitSync.timerId);
        state.gitSync.timerId = null;
      }
    }
    updateGitSyncPanelState();
    return {
      saved: true,
      remoteSha: state.gitSync.lastRemoteSha,
      commitSha: writeResult.commitSha,
    };
  } catch (error) {
    console.error("git sync failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "備份同步失敗";
    state.gitSync.lastMessage = "";
    updateGitSyncPanelState();
    return { saved: false, error };
  } finally {
    state.gitSync.inFlight = false;
    updateGitSyncPanelState();
  }
}

async function waitForGitSyncIdle() {
  while (state.gitSync.inFlight) {
    await new Promise((resolve) => window.setTimeout(resolve, 60));
  }
}

async function getGitHubFileRecord(settings) {
  const response = await fetchGitHubContents(settings, "GET");
  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    throw new Error(await readGitHubErrorMessage(response));
  }
  return response.json();
}

async function putGitHubFileRecord(settings, payloadText, remoteSha = "") {
  const body = {
    message: settings.commitMessage || GIT_SYNC_DEFAULT_COMMIT_MESSAGE,
    content: encodeUtf8Base64(payloadText),
    branch: settings.branch,
  };
  if (remoteSha) {
    body.sha = remoteSha;
  }

  let response = await fetchGitHubContents(settings, "PUT", body);
  if (response.status === 409) {
    const latest = await getGitHubFileRecord(settings);
    const latestSha = latest?.sha ? String(latest.sha) : "";
    if (!latestSha) {
      response = await fetchGitHubContents(settings, "PUT", body);
    } else {
      body.sha = latestSha;
      response = await fetchGitHubContents(settings, "PUT", body);
    }
  }

  if (!response.ok) {
    return {
      saved: false,
      errorMessage: await readGitHubErrorMessage(response),
    };
  }

  const json = await response.json();
  return {
    saved: true,
    remoteSha: String(json?.content?.sha ?? body.sha ?? remoteSha ?? ""),
    commitSha: String(json?.commit?.sha ?? ""),
  };
}

async function fetchGitHubContents(settings, method, body = null) {
  const url = buildGitHubContentsUrl(settings, method === "GET" ? settings.branch : "");
  const headers = {
    Accept: GITHUB_CONTENTS_ACCEPT,
    "X-GitHub-Api-Version": GITHUB_API_VERSION,
  };
  if (settings.token) {
    headers.Authorization = `Bearer ${settings.token}`;
  }
  if (method !== "GET") {
    headers["Content-Type"] = "application/json";
  }

  return fetch(url, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
    cache: "no-store",
  });
}

function buildGitHubContentsUrl(settings, ref = "") {
  const path = normalizeGitSyncPath(settings.path)
    .split("/")
    .map((segment) => encodeURIComponent(segment))
    .join("/");
  const baseUrl = `${GITHUB_API_BASE}/repos/${encodeURIComponent(settings.owner)}/${encodeURIComponent(settings.repo)}/contents/${path}`;
  if (!ref) {
    return baseUrl;
  }
  return `${baseUrl}?ref=${encodeURIComponent(ref)}`;
}

function encodeUtf8Base64(text) {
  const bytes = new TextEncoder().encode(String(text ?? ""));
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

async function readGitHubErrorMessage(response) {
  const rawText = await response.text();
  if (!rawText) {
    return response.statusText || `HTTP ${response.status}`;
  }

  try {
    const parsed = JSON.parse(rawText);
    const message = typeof parsed?.message === "string" ? parsed.message : "";
    const details = Array.isArray(parsed?.errors)
      ? parsed.errors
          .map((item) => {
            if (typeof item === "string") {
              return item;
            }
            if (item && typeof item === "object") {
              return String(item.message ?? item.code ?? item.resource ?? "");
            }
            return "";
          })
          .filter(Boolean)
          .join("; ")
      : "";
    return [message, details].filter(Boolean).join(" - ") || response.statusText || `HTTP ${response.status}`;
  } catch {
    return rawText.trim() || response.statusText || `HTTP ${response.status}`;
  }
}

async function loadProjectFromGit(options = {}) {
  const settings = getGitSyncSettings();
  if (!isGitSyncTargetConfigured(settings)) {
    return false;
  }

  if (state.gitSync.inFlight) {
    await waitForGitSyncIdle();
  }

  state.gitSync.inFlight = true;
  state.gitSync.lastError = "";
    state.gitSync.lastMessage = options.source === "startup" ? "正在從備份載入..." : "正在載入備份版本...";
  updateGitSyncPanelState();

  try {
    const response = await fetchGitHubContents(settings, "GET");
    if (response.status === 404) {
    const missingMessage = "備份空間尚未建立 project-state.json，請先按「立即備份」建立遠端檔案。";
      if (options.source === "startup") {
        state.gitSync.lastMessage = "";
        state.gitSync.lastError = "";
        return false;
      }
      state.gitSync.lastMessage = missingMessage;
      state.gitSync.lastError = "";
      updateGitSyncPanelState();
      return false;
    }
    if (!response.ok) {
      throw new Error(await readGitHubErrorMessage(response));
    }

    const json = await response.json();
    const content = typeof json?.content === "string" ? json.content.trim() : "";
    if (!content) {
      throw new Error("遠端檔案沒有可讀取的內容");
    }

    const decoded = decodeBase64Utf8(content);
    const parsed = JSON.parse(decoded);
    const normalized = normalizeProjectPayload(parsed) || normalizeLegacyProjectPayload(parsed);
    if (!normalized) {
      throw new Error("遠端 JSON 格式無法載入");
    }

    applyLoadedProjectState(normalized);
    const loadedHash = hashSnapshot(normalized.snapshot);
    state.gitSync.lastRemoteSha = String(json?.sha ?? "");
    state.gitSync.lastSyncedHash = loadedHash;
    state.gitSync.lastSyncedAt = new Date().toISOString();
    state.gitSync.lastMessage = `已從備份載入：${settings.owner}/${settings.repo}/${settings.path}`;
    state.gitSync.lastError = "";
    updateGitSyncPanelState();
    return true;
  } catch (error) {
    console.error("load project from git failed", error);
    state.gitSync.lastError = error?.message ? String(error.message) : "從備份載入失敗";
    state.gitSync.lastMessage = "";
    updateGitSyncPanelState();
    return false;
  } finally {
    state.gitSync.inFlight = false;
    updateGitSyncPanelState();
  }
}

async function loadProjectStateFromGitOnStartup() {
  const settings = getGitSyncSettings();
  if (!settings.loadOnStartup || !isGitSyncTargetConfigured(settings)) {
    return false;
  }
  return loadProjectFromGit({ source: "startup" });
}

function decodeBase64Utf8(base64Text) {
  const cleaned = String(base64Text ?? "").replace(/\s+/g, "");
  const binary = atob(cleaned);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new TextDecoder().decode(bytes);
}
