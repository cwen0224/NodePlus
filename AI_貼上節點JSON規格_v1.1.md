# 視覺小說節點化 JSON 規格（Canonical + Legacy）

本文件同時描述三個層級：
- Canonical full document: `vn-node-graph`
- Canonical fragment: `vn-node-graph.fragment`
- Legacy transport: `node-editor.clipboard`

正式型別契約位於 `./vn-node-graph.contract.d.ts`。

## 0. Canonical schema

`vn-node-graph` 是完整專案檔，適合畫布渲染、存檔與 VN runtime 編譯。

`vn-node-graph.fragment` 是選取片段格式，適合複製、貼上、資料托盤與外部 AI 生成。

`node-editor.clipboard` 保留給目前 MVP 的相容傳輸格式，未來可以逐步淘汰。

### 0.1 Full document

```json
{
  "schema": "vn-node-graph",
  "version": "1.0.0",
  "document": {
    "id": "story_ch1",
    "title": "第一章",
    "language": "zh-Hant",
    "description": "視覺小說劇情圖"
  },
  "entryNodeId": "n_start",
  "nodes": [],
  "edges": [],
  "assets": [],
  "variables": {},
  "editor": {
    "viewport": { "x": 0, "y": 0, "zoom": 1 }
  }
}
```

### 0.2 Fragment

```json
{
  "schema": "vn-node-graph.fragment",
  "version": "1.0.0",
  "bounds": {
    "x": 120,
    "y": 80,
    "width": 640,
    "height": 360
  },
  "nodes": [],
  "edges": []
}
```

### 0.3 Core rules

- `nodes[]` 代表節點，`edges[]` 代表有向連線。
- `edges.key` 是輸出給 VN runtime 的 JSON Key。
- `nodes[].position` 在 full document 中是世界座標，在 fragment 中是相對於 `bounds` 的局部座標。
- `selectedNodeIds`、dragging 狀態、hover 狀態、歷史紀錄都不屬於 canonical JSON。
- `assets[]` 只保留可序列化的資源資訊，不存檔案 handle。
- AI 生成時，優先輸出 `vn-node-graph.fragment`，除非明確要求完整專案檔。

### 0.4 Import / export pipeline

- Export selection:
  - 先收集被選節點與其內部連線。
  - 計算 `bounds` 與相對座標。
  - 產生 `vn-node-graph.fragment`。
- Paste fragment:
  - 解析 JSON。
  - 重新對映 node id。
  - 以滑鼠當前位置或視窗中心作為 anchor。
  - 還原節點與連線。
- Export project:
  - 使用 `vn-node-graph` full document。
  - 保留 `entryNodeId`、`assets`、`variables` 與 `editor.viewport`。

### 0.5 Compatibility note

現有 `node-editor.clipboard` 仍可被視為 legacy fragment transport。
如果輸入不是 canonical schema，仍可接受舊格式，再由匯入層正規化成 canonical node/edge 結構。

## 1. Legacy transport schema（`node-editor.clipboard`）

以下內容描述目前 MVP 仍支援的 legacy 剪貼簿格式，會先被正規化成 canonical fragment，再進入畫布。

```json
{
  "type": "node-editor.clipboard",
  "version": 1,
  "nodes": [
    {
      "id": "1",
      "offsetX": 0,
      "offsetY": 0,
      "title": "節點 1",
      "content": "說明文字",
      "metadata": {}
    }
  ],
  "connections": [],
  "width": 170,
  "height": 95
}
```

### 1.1 欄位說明

- `type`: 建議固定 `"node-editor.clipboard"`。
- `version`: 建議固定 `1`。
- `nodes`: 必填陣列，至少 1 個節點。
- `connections`: 可省略，省略時視為空陣列。
- `width` / `height`: 可省略，系統會自動計算。

`nodes[]` 內欄位：

- `id`: 建議使用字串且唯一（例如 `"1"`, `"2"`）。
- `offsetX` / `offsetY`: 相對於貼上群組左上角的位置（建議使用）。
- `x` / `y`: 也可用絕對座標。若未給 `offsetX/offsetY` 會用 `x/y` 換算。
- `title`: 可省略，預設 `節點 {id}`。
- `content`: 可省略，預設 `拖曳節點或拉線連接`。
- `metadata`: 可省略，預設 `null`。

`connections[]` 內欄位：

- `fromId` / `toId`: 來源與目標節點 `id`。
- `fromSide` / `toSide`: 連接點位置，合法值：
  - 單一點位：`"top" | "right" | "bottom" | "left"`
  - 多點位：`"top:2"`, `"right:3"` 這種 `side:index` 格式。

## 2. 快速格式 A（只貼一般 JSON 物件）

若貼上的 JSON 不是 `nodes` 結構，而是一般物件，例如：

```json
{
  "user_id": "TC-8829",
  "status": "active"
}
```

系統會自動轉成 1 個節點（通常標題為 `未定義結點`），並把內容摘要放在節點內。

## 3. 快速格式 B（功能性綁定 JSON）

若物件包含功能參數，會自動轉成「綁定節點」：

- 連接參數：`node_Right_1`, `node_Left_1`, `node_Top_1`, `node_Bottom_1`（可遞增 `_2`, `_3`）
- 勾選參數：`check_1`, `check_2`, ...
- 文字欄位：`text_1`, `text_2`, ...
- 日期欄位：`date_yyyy_mm_dd_1`, `date_yyyy_mm_dd_2`, ...

範例：

```json
{
  "__nodeTitle": "分歧節點",
  "__nodeDescription": "依條件選擇路徑",
  "next": null,
  "last": null,
  "prime": false,
  "node_Right_1": "next",
  "node_Right_2": "last",
  "check_1": "prime",
  "text_1": "備註",
  "date_yyyy_mm_dd_1": "2026-03-05"
}
```

### 3.1 綁定目標寫法

- 可以寫簡單鍵名：`"next"`、`"prime"`
- 也可寫路徑：`"$.profile.name"`、`"$.items[0].id"`

### 3.2 多個同方向連接點（重要）

若同一側需要多個連接點，請使用遞增編號：

- `node_Right_1`
- `node_Right_2`
- `node_Right_3`

這會在節點右側產生 3 個連接點，供分歧線使用（例如 Visual Novel 選項分支）。

Visual Novel 分歧節點範例（功能性 JSON）：

```json
{
  "__nodeTitle": "對話分歧",
  "__nodeDescription": "玩家選項",
  "choiceA_next": null,
  "choiceB_next": null,
  "node_Right_1": "choiceA_next",
  "node_Right_2": "choiceB_next",
  "text_1": "旁白或提示"
}
```

說明：
- `node_Right_1` 綁到 `choiceA_next`
- `node_Right_2` 綁到 `choiceB_next`
- 匯入後同一節點右側會有兩個可拉線端點

## 3.3 多分岐完整貼上範例（含 3 個節點與連線）

```json
{
  "type": "node-editor.clipboard",
  "version": 1,
  "nodes": [
    {
      "id": "1",
      "offsetX": 0,
      "offsetY": 0,
      "title": "對話分歧",
      "content": "玩家選擇",
      "metadata": {
        "__nodeTitle": "對話分歧",
        "__nodeDescription": "Visual Novel 分歧",
        "choiceA_next": null,
        "choiceB_next": null,
        "node_Right_1": "choiceA_next",
        "node_Right_2": "choiceB_next"
      }
    },
    {
      "id": "2",
      "offsetX": 320,
      "offsetY": -80,
      "title": "A 路線",
      "content": "進入 A 劇情",
      "metadata": {}
    },
    {
      "id": "3",
      "offsetX": 320,
      "offsetY": 80,
      "title": "B 路線",
      "content": "進入 B 劇情",
      "metadata": {}
    }
  ],
  "connections": [
    { "fromId": "1", "fromSide": "right", "toId": "2", "toSide": "left" },
    { "fromId": "1", "fromSide": "right:2", "toId": "3", "toSide": "left" }
  ],
  "width": 520,
  "height": 260
}
```

重點：
- 第一條線用 `fromSide: "right"`（等同 `right:1`）
- 第二條線用 `fromSide: "right:2"` 對應第二個右側端點
- 若有第三條分支就用 `right:3`，依此類推

## 4. AI 生成時的硬性規則

- 回傳內容必須是合法 JSON（不能有註解、尾逗號）。
- 若使用標準格式，`nodes` 一定要有至少 1 筆。
- `connections` 指向的節點 id 必須存在於 `nodes`。
- 禁止自連：`fromId` 不能等於 `toId`。
- `fromSide` / `toSide` 必須是合法 side 字串（含 `side:index`）。
- 若有 `node_Right_2`、`node_Top_3` 這類多端點定義，連線需使用對應的 `right:2`、`top:3`。

## 5. 建議給其他 AI 的指令模板

可直接貼這段給其他 AI：

```text
請輸出「純 JSON」且符合以下格式：
1) 使用 node-editor.clipboard v1
2) nodes 至少 1 筆，使用 id/offsetX/offsetY/title/content/metadata
3) 需要連線時，connections 使用 fromId/fromSide/toId/toSide
4) side 只能是 top/right/bottom/left 或 side:index（例如 right:2）
5) 不要輸出 Markdown，只輸出 JSON 本體
```

## 6. 三節點最小可用範例

```json
{
  "type": "node-editor.clipboard",
  "version": 1,
  "nodes": [
    { "id": "1", "offsetX": 0, "offsetY": 0, "title": "起點", "content": "開始", "metadata": {} },
    { "id": "2", "offsetX": 260, "offsetY": 0, "title": "處理", "content": "中間節點", "metadata": {} },
    { "id": "3", "offsetX": 520, "offsetY": 0, "title": "終點", "content": "結束", "metadata": {} }
  ],
  "connections": [
    { "fromId": "1", "fromSide": "right", "toId": "2", "toSide": "left" },
    { "fromId": "2", "fromSide": "right", "toId": "3", "toSide": "left" }
  ],
  "width": 700,
  "height": 120
}
```

## 7. AI 規畫建議（不同用途）

以下是給 AI 的建模建議，目標是讓貼入後更容易編輯與維護。

### 7.1 樹狀圖（決策樹 / 組織圖 / 技術分解）

建議：
- 每個節點至少保留 `node_Right_1` 作為主要子節點出口。
- 多子節點用 `node_Right_2`、`node_Right_3` 擴充。
- 資料欄位建議包含：
  - `node_id`（唯一識別）
  - `label`（節點名稱）
  - `category`（分類）
  - `priority`（優先級）
  - `status`（待辦/進行中/完成）
- 若要做父回指，可加 `node_Left_1` 綁到 `parent_id` 類欄位。

### 7.2 影片編輯檔（剪輯流程 / 分鏡 / 後製）

建議：
- 以「鏡頭片段」為節點，右側連到下一鏡頭。
- 常用欄位：
  - `clip_id`
  - `source_media`
  - `in_tc` / `out_tc`（可用 `HH:MM:SS:FF` 字串）
  - `duration_frames`
  - `track`（V1/A1）
  - `transition`
  - `note`
- 可加 `text_1` 放導演備註、`date_yyyy_mm_dd_1` 放交付日期。
- 分支敘事或平行版本可用 `node_Right_2`、`node_Right_3`。

### 7.3 遊戲任務安排（Quest / 流程腳本）

建議：
- 每個任務節點使用：
  - `quest_id`
  - `title`
  - `description`
  - `state`（locked/active/completed）
  - `reward`
  - `fail_condition`
- 主線用 `node_Right_1`，支線用 `node_Right_2`、`node_Right_3`。
- 若要回收前置任務，可用 `node_Left_1` 綁 `previous_quest_id`。
- 布林條件建議搭配：
  - `is_main_quest: false`
  - `check_1: "is_main_quest"`

### 7.4 通用結構建議（任何領域都適用）

- 固定保留一個穩定主鍵：`id` / `node_id` / `quest_id` 三選一。
- 標題用 `__nodeTitle`，摘要用 `__nodeDescription`。
- 若要讓 AI 穩定產生可連線節點，至少放一個：
  - `node_Right_1: "next"`（或對應欄位名）
- 大型資料避免整段塞在 `content`，改放 `metadata` 一般欄位。
- 若是批量生成，id 建議字串流水號（`"1"`, `"2"`, `"3"`）。

### 7.5 可直接貼給其他 AI 的規劃指令

```text
請先依用途（樹狀圖 / 影片剪輯 / 遊戲任務）建立節點資料模型，
再輸出 node-editor.clipboard v1 JSON。
每個節點 metadata 需包含可追蹤主鍵與業務欄位，
並至少定義一個可連線功能參數（例如 node_Right_1）。
若有分支，使用 node_Right_2、node_Right_3，
並在 connections 使用對應的 right:2、right:3。
```
