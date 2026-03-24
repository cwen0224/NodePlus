export type GraphSchemaName = "vn-node-graph";
export type GraphFragmentSchemaName = "vn-node-graph.fragment";
export type GraphLegacyClipboardSchemaName = "node-editor.clipboard";
export type GraphVersion = string;

export type GraphId = string;
export type GraphRecord = Record<string, unknown>;

export type GraphPortSide = "top" | "right" | "bottom" | "left";
export type GraphPortRole = "input" | "output" | "bidirectional" | "event" | "anchor";
export type GraphNodeType =
  | "vn.dialogue"
  | "vn.choice"
  | "vn.branch"
  | "vn.command"
  | "vn.media"
  | "vn.end"
  | (string & {});
export type GraphEdgeKind =
  | "flow"
  | "branch"
  | "reference"
  | "condition"
  | "control"
  | "data"
  | (string & {});
export type GraphMediaKind = "image" | "video" | "audio" | "live2d" | "text" | "binary" | (string & {});

export interface GraphPoint {
  x: number;
  y: number;
}

export interface GraphSize {
  width: number;
  height: number;
}

export interface GraphRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface GraphDocumentMeta {
  id?: GraphId;
  title?: string;
  language?: string;
  description?: string;
  author?: string;
  createdAt?: string;
  updatedAt?: string;
  tags?: string[];
}

export interface GraphViewport {
  x: number;
  y: number;
  zoom: number;
}

export interface GraphEditorState {
  viewport?: GraphViewport;
  snapToGrid?: boolean;
  gridSize?: number;
  theme?: string;
  ui?: GraphRecord;
}

export interface GraphAssetRef {
  id: GraphId;
  kind: GraphMediaKind;
  name?: string;
  src?: string;
  path?: string;
  url?: string;
  mimeType?: string;
  hash?: string;
  durationMs?: number;
  sizeBytes?: number;
  meta?: GraphRecord;
}

export interface GraphMediaRef {
  assetId?: GraphId;
  kind?: GraphMediaKind;
  name?: string;
  src?: string;
  path?: string;
  url?: string;
  mimeType?: string;
  hash?: string;
  durationMs?: number;
  sizeBytes?: number;
  meta?: GraphRecord;
}

export interface GraphConditionObject {
  op: string;
  left?: GraphCondition;
  right?: GraphCondition;
  path?: string;
  value?: unknown;
  values?: unknown[];
  metadata?: GraphRecord;
}

export type GraphCondition = string | GraphConditionObject;

export interface GraphEffect {
  op: string;
  path?: string;
  value?: unknown;
  args?: unknown[];
  when?: GraphCondition;
  metadata?: GraphRecord;
}

export interface GraphChoice {
  label: string;
  edgeKey: string;
  condition?: GraphCondition;
  metadata?: GraphRecord;
}

export interface GraphNodeMediaBundle {
  background?: GraphMediaRef | null;
  portrait?: GraphMediaRef | null;
  voice?: GraphMediaRef | null;
  sfx?: GraphMediaRef[];
  image?: GraphMediaRef | null;
  video?: GraphMediaRef | null;
  live2d?: GraphMediaRef | null;
}

export interface GraphNodeData {
  speaker?: string;
  text?: string;
  textFormat?: "plain" | "markdown" | "rich" | string;
  media?: GraphNodeMediaBundle;
  choices?: GraphChoice[];
  effects?: GraphEffect[];
  variables?: GraphRecord;
  runtime?: GraphRecord;
  [key: string]: unknown;
}

export interface GraphPort {
  id: string;
  side: GraphPortSide;
  role: GraphPortRole;
  label?: string;
  maxConnections?: number;
  required?: boolean;
  layoutIndex?: number;
  metadata?: GraphRecord;
}

export interface GraphNodeUiState {
  color?: string;
  collapsed?: boolean;
  locked?: boolean;
  hidden?: boolean;
  selected?: boolean;
  notes?: string;
  [key: string]: unknown;
}

export interface GraphNode {
  id: GraphId;
  type: GraphNodeType;
  title: string;
  position: GraphPoint;
  size?: GraphSize;
  ports?: GraphPort[];
  data: GraphNodeData;
  ui?: GraphNodeUiState;
  metadata?: GraphRecord;
}

export interface GraphEdge {
  id: GraphId;
  kind: GraphEdgeKind;
  key: string;
  label?: string;
  fromNodeId: GraphId;
  fromPortId: string;
  toNodeId: GraphId;
  toPortId: string;
  points?: GraphPoint[];
  condition?: GraphCondition;
  effects?: GraphEffect[];
  metadata?: GraphRecord;
}

export interface GraphDocument {
  schema: GraphSchemaName;
  version: GraphVersion;
  document: GraphDocumentMeta;
  entryNodeId: GraphId;
  nodes: GraphNode[];
  edges: GraphEdge[];
  assets?: GraphAssetRef[];
  variables?: GraphRecord;
  editor?: GraphEditorState;
  metadata?: GraphRecord;
}

export interface GraphFragment {
  schema: GraphFragmentSchemaName;
  version: GraphVersion;
  bounds: GraphRect;
  nodes: GraphNode[];
  edges: GraphEdge[];
  metadata?: GraphRecord;
}

export interface GraphClipboardEnvelope {
  type?: string;
  schema?: string;
  version?: GraphVersion | number;
  nodes?: unknown[];
  connections?: unknown[];
  bounds?: GraphRect;
  width?: number;
  height?: number;
  [key: string]: unknown;
}

export interface GraphInsertOptions {
  anchor: GraphPoint;
  snapToGrid?: boolean;
  preserveRelativeSpacing?: boolean;
}

export interface GraphInsertPatch {
  nodes: GraphNode[];
  edges: GraphEdge[];
  bounds: GraphRect;
  idMap: Record<string, GraphId>;
}

export interface FlowNodeAdapter {
  id: GraphId;
  type: string;
  position: GraphPoint;
  width?: number;
  height?: number;
  data: GraphNode;
}

export interface FlowEdgeAdapter {
  id: GraphId;
  source: GraphId;
  target: GraphId;
  sourceHandle: string;
  targetHandle: string;
  type?: string;
  label?: string;
  data: GraphEdge;
}

export interface GraphSelectionExportOptions {
  includeDanglingEdges?: boolean;
  includeAssets?: boolean;
}

export declare function parseGraphDocument(input: unknown): GraphDocument | null;
export declare function parseGraphFragment(input: unknown): GraphFragment | null;
export declare function parseGraphClipboardEnvelope(input: unknown): GraphClipboardEnvelope | null;
export declare function serializeGraphDocument(document: GraphDocument): string;
export declare function serializeGraphFragment(fragment: GraphFragment): string;
export declare function documentToFragment(
  document: GraphDocument,
  nodeIds: GraphId[],
  options?: GraphSelectionExportOptions
): GraphFragment;
export declare function fragmentToInsertPatch(
  fragment: GraphFragment,
  options: GraphInsertOptions
): GraphInsertPatch;
export declare function documentToFlowNodes(document: GraphDocument): FlowNodeAdapter[];
export declare function documentToFlowEdges(document: GraphDocument): FlowEdgeAdapter[];
export declare function normalizeClipboardPayload(
  input: unknown
): GraphDocument | GraphFragment | null;
