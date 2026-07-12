# 专项设计：ContextPack Schema 与上下文采集

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 评审通过（见 [`design-review-record.md`](./design-review-record.md)） |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §5.1 |
| 现有代码 | `Agent/Context/OfficeContext.vb`、`ChatContextBuilder`、三端 `CaptureOfficeContext`、`ChatContextTrace` |

---

## 1. 目标与非目标

### 1.1 目标

1. 定义 **唯一** 上下文档契约 `ContextPack`，供 Planner / Loop / UI 控制台 / RunTrace 共用。  
2. 规定 **先读后想**：Harness 在规划前必须拿到预算化快照。  
3. 三端（Word/Excel/PPT）结构字段可扩展，但外壳字段稳定。  
4. 支持 **执行前/后双快照**，为 Observe 差分提供输入（见 observation 专项）。  

### 1.2 非目标

- 不在本专项定义 LLM Prompt 全文模板（只定义可注入的结构化数据）。  
- 不做云端 Graph 索引。  
- 不规定 COM 实现细节（只规定宿主 Reader 输出形状）。  

### 1.3 设计原则

| 原则 | 说明 |
|---|---|
| 预算优先 | 任何字段都可被截断；截断必须留下 `Truncated=true` 与原因 |
| 可二次拉取 | 大内容用 `ref` + 读工具，而不是一次塞满 |
| 宿主无知 | ShareRibbon 只理解 `AppType` 与通用段；细节在 `Host` 对象 |
| 向后兼容 | 现有 `OfficeContext.ToPromptText()` 可视为 ContextPack 的降级渲染 |

---

## 2. 与现状的映射

| 现有 | 演进 |
|---|---|
| `OfficeContext` | 成为 `ContextPack.Host` 的简化投影，或由 Pack 生成 |
| `SelectionInfo` | → `ContextPack.Selection`（扩展语义字段） |
| `DocumentStructure` | → `ContextPack.Structure`（按 App 分型） |
| `ChatContextBuilder` 记忆块 | → `ContextPack.Memory` |
| `ChatContextTrace` | UI/调试视图；字段应对齐 Pack 摘要，不另起一套真相源 |
| 三端 `CaptureOfficeContext` | → 实现 `IHostContextReader.ReadAsync(options)` |

---

## 3. 顶层 Schema

### 3.1 JSON 逻辑结构（契约）

```json
{
  "schemaVersion": "1.0",
  "packId": "cp_...",
  "turnId": "turn_...",
  "capturedAt": "2026-07-11T12:00:00+08:00",
  "appType": "Word",
  "purpose": "plan | observe_before | observe_after | ui_trace",
  "budget": {
    "modelContextTokens": 128000,
    "allocatedTokens": 12000,
    "usedTokensEstimate": 8300,
    "strategy": "selection_first"
  },
  "document": { },
  "selection": { },
  "structure": { },
  "samples": { },
  "styles": { },
  "memory": { },
  "skills": { },
  "tools": { },
  "history": { },
  "risks": { },
  "host": { },
  "truncation": [],
  "readerErrors": []
}
```

### 3.2 字段字典（顶层）

| 字段 | 必填 | 说明 |
|---|---|---|
| schemaVersion | 是 | 语义版本，破坏变更升主版本 |
| packId | 是 | 本快照 id |
| turnId | 是 | 对齐日志 correlation / UserTurn |
| capturedAt | 是 | ISO-8601 |
| appType | 是 | `Excel` \| `Word` \| `PowerPoint` |
| purpose | 是 | 采集用途，影响预算策略 |
| budget | 是 | 预算与估算 |
| document | 是 | 文档元信息 |
| selection | 否 | 无选区可为 null |
| structure | 建议 | 大纲/表结构/幻灯片目录 |
| samples | 否 | 样例文本/行 |
| styles | 否 | 样式/主题摘要 |
| memory | 否 | 长期记忆与会话摘要 |
| skills | 否 | **仅元数据**，不含 SKILL 正文 |
| tools | 否 | 可见工具短列表 |
| history | 否 | 近 N 轮对话摘要 |
| risks | 否 | 保护、只读、宏、大表等 |
| host | 是 | 宿主扩展对象（见 §5） |
| truncation | 是 | 截断记录数组，可空 |
| readerErrors | 是 | 采集失败项，可空；**不得静默丢** |

### 3.3 purpose 与预算策略

| purpose | 目标 token 占比（建议） | 侧重 |
|---|---|---|
| plan | 100% 分配预算 | selection + structure + memory + tools |
| observe_before | 40% | 将变更范围指纹化，少 samples |
| observe_after | 40% | 同范围再采，供 Diff |
| ui_trace | 30% | 给前端控制台，更短 |

策略名 `selection_first`：有有效选区时，选区与近邻优先于全文大纲。

---

## 4. 通用段详细定义

### 4.1 document

```json
{
  "name": "季度报告.docx",
  "fullPath": "C:\\...\\季度报告.docx",
  "isSaved": true,
  "isReadOnly": false,
  "isProtected": false,
  "protectionType": null,
  "languageGuess": "zh-CN",
  "metrics": {
    "pageCount": 12,
    "wordCount": 5600,
    "paragraphCount": 180,
    "tableCount": 3
  }
}
```

| 字段 | Word | Excel | PPT |
|---|---|---|---|
| metrics.pageCount | 页 | — | 幻灯片数可放 host |
| metrics.wordCount | 字 | — | 文本框字数可选 |
| sheetCount / slideCount | — | host | host |

路径可脱敏：UI 显示名；日志可只保留 hash 路径。

### 4.2 selection

```json
{
  "hasSelection": true,
  "address": "§3 第2段" ,
  "addressCanonical": "Word:Para:12-15",
  "itemCount": 4,
  "dataType": "paragraphs",
  "semanticType": "body_text",
  "preview": "……最多 N 字……",
  "previewTruncated": false,
  "styleNames": ["正文"],
  "isEmpty": false,
  "isWholeDocument": false
}
```

**semanticType 枚举（跨端）**

| 值 | 含义 |
|---|---|
| empty | 无选区或插入点无内容 |
| body_text | 正文 |
| heading | 标题 |
| table_cells | 表内 |
| sheet_range | Excel 区域 |
| chart | 图表 |
| shape | 形状 |
| slide | 整页幻灯片 |
| mixed | 混合 |
| unknown | 未能识别 |

**不变量**

- `hasSelection=false` 时，Planner 默认范围策略：`infer_from_structure`（Excel→UsedRange/ListObject，Word→全文或当前节，PPT→当前页）。  
- `preview` 永远 ≤ `budget.selectionPreviewChars`（默认 1500）。  

### 4.3 structure（外壳）

```json
{
  "kind": "word_outline | excel_tables | ppt_deck | generic",
  "summary": "一句话人类可读",
  "nodeCount": 24,
  "truncated": false,
  "nodes": []
}
```

`nodes` 形态见 §5；外壳字段三端一致。

### 4.4 samples

```json
{
  "items": [
    { "ref": "Word:Para:1", "role": "heading", "text": "一、概述" },
    { "ref": "Excel:Sheet1!A1:F1", "role": "header", "text": "日期,金额,..." }
  ],
  "policy": "top_n_plus_neighbors"
}
```

### 4.5 styles

```json
{
  "theme": "Office Theme",
  "defaultFont": "宋体",
  "defaultFontSizePt": 12,
  "namedStyles": ["标题 1", "标题 2", "正文"],
  "note": "仅名称列表，细节按需 read_style"
}
```

### 4.6 memory

```json
{
  "userProfile": "偏好：公文语体；默认字体微软雅黑",
  "longTerm": [
    { "id": "m_1", "type": "preference", "content": "...", "score": 0.82 }
  ],
  "sessionSummaries": [
    { "sessionId": "s1", "title": "上周周报", "snippet": "..." }
  ],
  "policy": {
    "excludeShortTermFromRag": true,
    "topN": 5
  }
}
```

**不变量**：RAG 段不得注入 short_term（与现路线一致）。

### 4.7 skills（仅元数据）

```json
{
  "recalled": [
    {
      "name": "word-document-agent",
      "application": "Word",
      "description": "…",
      "score": 0.91,
      "allowedTools": ["ListParagraphs", "FormatText"],
      "detailLoaded": false
    }
  ]
}
```

正文加载后由 SkillRouter 另附 `SkillDetailBundle`，**不进**默认 ContextPack（避免爆窗）。

### 4.8 tools

```json
{
  "visibleCount": 36,
  "items": [
    { "id": "FormatText", "risk": "medium", "category": "format" }
  ],
  "omittedCount": 12
}
```

只含 Broker 过滤后可见工具。

### 4.9 history

```json
{
  "turns": [
    { "role": "user", "content": "…", "chars": 120 },
    { "role": "assistant", "content": "…", "chars": 400 }
  ],
  "maxTurns": 6
}
```

### 4.10 risks

```json
{
  "items": [
    { "code": "DOC_PROTECTED", "level": "high", "message": "文档限制编辑" },
    { "code": "LARGE_RANGE", "level": "medium", "message": "UsedRange 超过 5 万格" },
    { "code": "MACROS_PRESENT", "level": "medium", "message": "工作簿含 VBA 项目" }
  ]
}
```

供 SafetyGate 与 Planner 直接消费。

### 4.11 truncation / readerErrors

```json
"truncation": [
  { "path": "samples.items", "reason": "token_budget", "originalCount": 200, "kept": 20 }
],
"readerErrors": [
  { "path": "structure", "code": "COM_ERROR", "message": "获取标题树失败" }
]
```

**不变量**：`readerErrors` 非空时，Harness 仍可继续，但 Plan 复杂度至少 `exploratory`，并优先安排读工具。

---

## 5. 宿主扩展 host / structure.nodes

### 5.1 Word

```json
"host": {
  "view": "print",
  "trackRevisions": false,
  "currentSectionIndex": 1,
  "headingOutline": [
    { "level": 1, "text": "一、概述", "paraRef": "Word:Para:3", "page": 1 }
  ],
  "lists": { "hasNumbering": true, "restartRisk": true },
  "fields": { "tocPresent": true, "tocDirty": false }
},
"structure": {
  "kind": "word_outline",
  "nodes": [
    { "ref": "Word:Para:3", "type": "heading", "level": 1, "style": "标题 1", "text": "一、概述" },
    { "ref": "Word:Tbl:1", "type": "table", "rows": 5, "cols": 4, "preview": "..." }
  ]
}
```

**Word Reader 必采（plan purpose）**

1. 选区或插入点段落 ref、样式  
2. 标题大纲（深度 ≤ 3，节点 ≤ 50）  
3. 是否修订/保护  
4. 近邻 ±2 段 preview（若有选区）  

### 5.2 Excel

```json
"host": {
  "activeWorkbook": "销售.xlsx",
  "activeSheet": "Sheet1",
  "sheetNames": ["Sheet1", "汇总"],
  "calculationMode": "automatic",
  "hasFilter": false,
  "tables": [
    {
      "name": "Table1",
      "sheet": "Sheet1",
      "address": "A1:F200",
      "headers": ["日期", "区域", "金额"],
      "columnTypes": ["date", "text", "number"],
      "rowCount": 199,
      "hasTotals": false
    }
  ],
  "usedRange": "A1:F200",
  "formulaErrorSample": []
},
"structure": {
  "kind": "excel_tables",
  "nodes": [
    { "ref": "Excel:Sheet1!A1:F200", "type": "table", "name": "Table1" }
  ]
}
```

**Excel Reader 必采**

1. ActiveSheet + Selection 地址与值类型抽样  
2. ListObject 优先，否则 UsedRange 启发式  
3. 表头行、列类型猜测（number/text/date/formula）  
4. 公式错误样本（最多 10 个 `#REF!` 等）  
5. 行数/列数超阈值 → risks.LARGE_RANGE  

**大表规则**

| 单元格数 | 行为 |
|---|---|
| ≤ 5,000 | 可带 markdown 样例 30 行 |
| ≤ 50,000 | 仅表头 + 5 行样例 + 统计 |
| > 50,000 | 仅元数据 + 强制分块工具 |

### 5.3 PowerPoint

```json
"host": {
  "slideCount": 12,
  "currentSlideIndex": 3,
  "slideSize": { "widthIn": 13.33, "heightIn": 7.5 },
  "hasSlideMasterCustom": false,
  "slides": [
    {
      "index": 1,
      "layout": "Title Slide",
      "title": "项目汇报",
      "shapeCount": 4,
      "notesPreview": ""
    }
  ]
},
"structure": {
  "kind": "ppt_deck",
  "nodes": [
    { "ref": "Ppt:Slide:1", "type": "slide", "title": "项目汇报" }
  ]
}
```

**PPT Reader 必采**

1. 当前页 index + 标题 + 主要占位符文本 preview  
2. 全稿标题目录（≤ 40 页，超出截断）  
3. 选中形状类型与文本  

---

## 6. 引用体系（Ref）

跨 Observe / Tool 参数统一使用 **Canonical Ref**：

| App | 格式 | 示例 |
|---|---|---|
| Word | `Word:Para:{start}[-{end}]` | `Word:Para:12-15` |
| Word 表 | `Word:Tbl:{i}[:R{r}C{c}]` | `Word:Tbl:1:R2C3` |
| Excel | `Excel:{Sheet}!{A1}` | `Excel:Sheet1!B2:D10` |
| PPT | `Ppt:Slide:{i}[:Shape:{id}]` | `Ppt:Slide:3:Shape:4` |

**不变量**

- Tool 写操作参数优先使用 ref，而不是模糊「上面那段」。  
- Diff 的 `targetRefs[]` 必须可解析。  

---

## 7. ContextBudget 算法（设计）

### 7.1 输入

- 模型窗口 `W`（配置或画像）  
- 系统固定开销 `S`（工具 schema、角色提示）  
- 可用 `A = W - S - responseReserve`  

### 7.2 默认分配（plan）

| 段 | 占比 |
|---|---|
| selection + samples | 35% |
| structure | 20% |
| memory | 15% |
| history | 10% |
| tools + skills meta | 10% |
| document + risks + 其它 | 10% |

超限时丢弃顺序：history 细节 → samples → structure 深层节点 → memory 低分条。

### 7.3 估算

v0.1 可用 `chars/4` 粗估；后续可换 tiktoken 类库。`usedTokensEstimate` 必填。

---

## 8. 采集时序

```text
UserTurn 到达
  │
  ├─(1) HostReader.Read(purpose=plan)  ──► ContextPack_plan
  │         │
  │         ├ readerErrors? → 标记 exploratory
  │         └ risks → Safety 预热
  │
  ├─(2) Memory/Skills/Tools  enrich（ShareRibbon 纯逻辑）
  │
  ├─(3) Budget.Apply
  │
  ├─(4) Planner 使用 Pack
  │
  └─ 每步写工具前
        HostReader.Read(purpose=observe_before, focusRefs)
        Execute
        HostReader.Read(purpose=observe_after, focusRefs)
        Diff(before, after) → Observation
```

**线程**：HostReader 必须在 UI 线程调 COM；enrich/budget 在后台。

---

## 9. 接口契约（逻辑）

```text
Interface IHostContextReader
  Function ReadAsync(request As ContextReadRequest) As Task(Of ContextPack)

Class ContextReadRequest
  TurnId, AppType, Purpose
  FocusRefs As List(Of String)   ' observe 时缩小范围
  Options As ContextReadOptions  ' maxHeadings, maxSampleRows, ...

Interface IContextHub
  Function SnapshotAsync(request) As Task(Of ContextPack)
  Function DiffAsync(before As ContextPack, after As ContextPack) As DocumentDiff
  Function ToPromptSections(pack As ContextPack) As PromptSection[]
  Function ToUiTrace(pack As ContextPack) As ChatContextTrace
```

`IContextHub` 在 ShareRibbon；`IHostContextReader` 三端注入。

---

## 10. 渲染到 Prompt 的规则

| 段 | 是否默认进 system | 说明 |
|---|---|---|
| document + risks | 是 | 短 |
| selection | 是 | 有则必进 |
| structure.summary + 压缩 nodes | 是 | |
| samples | 是 | 受预算 |
| memory | 按配置 | |
| skills meta | 是 | 无正文 |
| tools | 是 | 短 id 列表；完整 schema 另通道 |
| host 原始大对象 | 否 | 只经压缩 |

保留 `OfficeContext.ToPromptText()` 作为 `ToPromptSections` 的兼容后端，直到迁移完成。

---

## 11. UI 控制台映射

现有「本轮上下文」面板字段建议绑定：

| UI 块 | ContextPack 路径 |
|---|---|
| 意图/任务 | 不在 Pack；来自 TaskSpec |
| Office 上下文 | document + selection + structure.summary |
| 记忆 | memory |
| Skills | skills.recalled |
| 工具 | tools.items |
| 风险 | risks |

---

## 12. 验收标准（设计冻结用）

1. 给定黄金 Word/Excel/PPT 文档，Reader 输出可通过 JSON Schema 校验。  
2. 无选区时 Pack 仍含可推断范围策略所需字段。  
3. 大表不导致 Pack 超过预算硬顶。  
4. `readerErrors` 非空可被 Planner 测试断言。  
5. `ToUiTrace` 与 Pack 同源，无第二套拼装逻辑。  

---

## 13. 开放问题（已冻结）

| # | 问题 | **冻结默认** |
|---|---|---|
| Q1 | 多工作簿/多文档是否进入 Pack？ | **v1 仅活动文档**；`scope=active` |
| Q2 | 是否采集图片 OCR？ | **否**；仅标记有图 |
| Q3 | 修订模式下是否包含修订气泡？ | **只标记** `trackRevisions` |
| Q4 | Canonical Ref 是否稳定跨保存？ | **不保证**；写后以 after Pack 为准 |

### 与 Safety / Excel 阈值

- 影响面 T1/T2 以 [`safety-policy.md`](./safety-policy.md) 为权威。  
- Excel 单元格分档（5k/50k）见 Context §5.2 与 Excel runtime，**不得与 Safety 冲突**：触发 Safety 上调后仍可分块执行。  

---

## 14. 决策摘要（评审）

- [x] 同意 ContextPack 为规划/观察唯一结构化上下文  
- [x] 同意 selection_first 预算策略  
- [x] 同意 Canonical Ref 体系  
- [x] 同意 v1 仅活动文档、无 OCR  
- [x] 同意 readerErrors 非空 → exploratory  

---

## 15. 落地顺序（实现时，本轮不做）

1. 定义 VB 模型 + JSON Schema 文件（fixtures）  
2. Word Reader 对齐  
3. Excel / PPT Reader  
4. ChatContextBuilder 改为消费 Pack  
5. 删除重复拼装路径  

---

*下一份关联：[`tool-result-observation.md`](./tool-result-observation.md)（使用 before/after Pack 做 Diff）。*
