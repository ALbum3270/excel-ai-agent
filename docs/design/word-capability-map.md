# 专项设计：Word Capability 地图与执行器映射

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 目标设计已评审；代码部分实现 |
| 实现状态 | **部分实现**：Word 是当前最接近样板的一端，已有 `WordActionHarness`、编号/格式/校对/语义排版 fast-path 结构化结果回灌；`ListParagraphs`/`GetParagraphInfo` 已通过 `ToolResult.Data` 返回给 Agent；`InsertText/FormatText/ReplaceText/DeleteText` 已返回 before/after/diff Observation，并通过 `ExecutionExplanation` 进入 Agent 卡片和轻量 RunTrace。快路径全量统一 `ToolResult`、翻译/续写纳入 Harness、Excel/PPT 同构仍待后续。 |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §6.1 |
| 现有 | `WordCapabilityRegistry`、`WordActionHarness`、`WordFormattingAgent`、`WordNumberingAgent`、校对/排版服务、`Tools/word/*.json`（24） |
| 关联 | Context / Observe / Safety / Skill / RunTrace / Golden |

---

## 1. 目标与非目标

### 1.1 目标

1. 给出 Word 端 **Capability ↔ Tool ↔ 服务类 ↔ UI** 的完整地图。
2. 规定快路径（ActionHarness）与 Loop 工具路径如何统一产出 `ToolResult` / Trace。
3. 明确每条 Capability 的 Input、Observe、Repair、风险与示例话术。
4. 标出缺口（读工具 Data、语义排版闭环、翻译/续写纳入 Harness）。

### 1.2 非目标

- 不重写校对算法细节。
- 不定义 Excel/PPT（见对应 runtime 专项）。

---

## 2. 架构位置

```text
OfficeHarness
  ├ Skill: word-document-agent (allowed-tools 见 SKILL.md)
  ├ ToolBroker → Tools/word/*.json
  └ Word Host
       ├ WordActionHarness          # 高置信确定性 Capability 快路径
       ├ WordCapabilityRegistry     # 元数据
       ├ *IntentCompiler / *Agent   # 计划编译与执行
       └ ChatControl.ExecuteJsonCommand  # 通用 tool 后端（应迁出）
```

**目标形态**：`IWordCapabilityExecutor` 注册表；`ExecuteJsonCommand` 仅作 ToolBroker 后端，不再被 NL 直接路由。

---

## 3. Capability 注册表（扩展现有 4 项）

### 3.1 已登记（代码现状）

| CapabilityId | Kind | 风险 | 预览 | 撤销 | 主执行 | Observe 要点 |
|---|---|---|---|---|---|---|
| `word.proofread` | Proofread | Low | 是 | 是 | Proofread 流程 / 侧栏 | 问题列表数量、是否应用 |
| `word.direct-formatting` | DirectFormatting | Medium | 否* | 是 | FormattingIntentCompiler → 应用 | 抽样 Range 字体/字号/对齐 |
| `word.numbering` | Numbering | Medium | 否 | 是 | WordNumberingAgent | ListString 连续、处理数量 |
| `word.semantic-reformat` | SemanticReformat | Medium/High | **是** | 是 | SmartFormatter / Reformat 管线 | 标题层级、样式名分布 |

\* 直接格式可后续加 ghost text 预览（`WordGhostTextManager` 可复用）。

### 3.2 建议新增登记（设计）

| CapabilityId | 说明 | 主要 Tools | 风险 | 优先级 |
|---|---|---|---|---|
| `word.read-structure` | 列段落/标题/样式 | ListParagraphs, GetParagraphInfo | safe | P0 |
| `word.replace` | 查找替换 | ReplaceText, Find 语义 | medium→risky 全文 | P0 |
| `word.insert-content` | 插入段落/表/图/页眉页脚 | Insert* | medium | P1 |
| `word.toc` | 目录生成/更新 | GenerateTOC | medium | P1 |
| `word.page-setup` | 页边距/行距/缩进 | SetPageMargins, SetLineSpacing, SetIndent | medium | P1 |
| `word.table-edit` | 表格式与行列 | FormatTable, InsertTable* | medium | P1 |
| `word.translate` | 文档/选区翻译 | 翻译服务 → 写回工具 | medium | P1 |
| `word.continue-write` | 续写/改写 | InsertText / 选区替换 | medium | P1 |
| `word.beautify` | 一键美化 | BeautifyDocument, ApplyStyle | medium | P1 |
| `word.vba-escape` | 仅缺口 | ExecuteVBA | risky | P2 |

---

## 4. Tool 全表映射（24）

| ToolId | 建议 Capability | 读写 | base risk | 备注 |
|---|---|---|---|---|
| ListParagraphs | word.read-structure | R | safe | 已回 Data，后续补稳定 ref |
| GetParagraphInfo | word.read-structure | R | safe | 已回 Data，后续补稳定 ref |
| FormatText | word.direct-formatting | W | medium | 已有 before/after/diff Observation |
| SetParagraphFormat | word.direct-formatting | W | medium | |
| ApplyStyle | word.semantic-reformat / beautify | W | medium | |
| BeautifyDocument | word.beautify | W | medium | |
| ReplaceText | word.replace | W | medium/risky | 已有 before/after/diff Observation；全文风险待 Safety |
| InsertText | word.insert-content / continue-write | W | medium | 已有 before/after/diff Observation |
| InsertParagraph | word.insert-content | W | medium | |
| DeleteText | word.replace / edit | W | risky | 已有 before/after/diff Observation；审批待 Safety |
| CopyPasteText | word.insert-content | W | medium | |
| InsertTable | word.table-edit | W | medium | |
| FormatTable | word.table-edit | W | medium | |
| InsertTableRow | word.table-edit | W | medium | |
| DeleteTableRow | word.table-edit | W | risky | |
| GenerateTOC | word.toc | W | medium | |
| InsertHeader / InsertFooter | word.insert-content | W | medium | |
| InsertPageNumber | word.page-setup | W | medium | |
| InsertImage | word.insert-content | W | medium | |
| SetIndent / SetLineSpacing / SetPageMargins | word.page-setup | W | medium | |
| ExecuteVBA | word.vba-escape | W | risky | 默认关 |

---

## 5. 服务类职责地图

| 服务/类 | 职责 | 不应再做 |
|---|---|---|
| WordActionHarness | Plan(kind)+置信度；调度 Capability | 堆业务 if 到 ChatControl |
| WordCapabilityRegistry | 元数据 | 执行 |
| FormattingIntentCompiler | 自然语言 → FormattingIntentPlan | COM 应用 |
| ProofreadIntentCompiler | → ProofreadIntentPlan | |
| WordFormattingAgent | 执行格式计划 + 观察摘要 | UI 路由 |
| WordNumberingAgent | 编号重排 | |
| SmartFormatter / Reformat* | 语义排版/模板 | 绕过 Trace |
| ParagraphService | 段落读写辅助 | |
| ContentLocator | 定位 | |
| WordGhostTextManager | 预览态 | 永久写入 |
| ChatControl | Surface + 注册 Executor | **新增业务分支** |

---

## 6. 快路径 vs Loop 路径

### 6.1 何时快路径

```text
WordActionHarness.Plan.ShouldHandle
  AND confidence ≥ 0.55
  AND kind ∈ {Proofread, DirectFormatting, Numbering, SemanticReformat}
  AND Safety 允许
```

### 6.2 统一回执（强制）

快路径结束后必须：

```text
WordCapabilityExecutionResult
  → 映射 ToolResult {
       toolId: capabilityId 或 主 tool,
       success, userMessage, observation, undoPointId
     }
  → RunTrace.AppendStep
  → UI 时间线
```

`Fallback` 状态：转 OfficeHarness 全量 Loop，不得静默失败。

### 6.3 与 Skill 关系

- primary skill 通常为 `word-document-agent`
- 快路径 tools 必须 ⊆ skill.allowed-tools
- 若 skill 未声明某 tool，快路径也不得调用

---

## 7. 各 Capability 契约卡片

### 7.1 word.proofread

| 项 | 内容 |
|---|---|
| 输入 | scope=selection\|document；issueTypes；applyMode |
| 前置读 | 选区文本或全文抽样 |
| 成功标准 | 侧栏有问题列表；或高置信项已应用且可撤销 |
| Observe | issueCount、appliedCount、scope refs |
| Repair | 读失败 → 换全文；空文档 → 友好结束 |
| UI | proofread-ui.js 面板 |
| Golden | W-proofread-basic |

### 7.2 word.direct-formatting

| 项 | 内容 |
|---|---|
| 输入 | FormattingIntentPlan |
| 成功标准 | 目标 Range 样式抽样匹配 |
| Observe | font/size/bold/align 前后 |
| Repair | 选区空 → 扩到段落/正文样式范围 |
| Golden | U2 子集「字号+2」 |

### 7.3 word.numbering

| 项 | 内容 |
|---|---|
| 输入 | scope；目标序列 1,2,3… |
| 约束 | **不**把普通文本「1.」强行变 List（可配置） |
| Observe | 前 K 个 ListString |
| Repair | 无自动编号 → 说明不可处理 |
| Golden | W-numbering-fix |

### 7.4 word.semantic-reformat

| 项 | 内容 |
|---|---|
| 输入 | 模板 id 或语义目标（公文/论文） |
| 必须预览 | 是（高风险结构变更） |
| Observe | 标题层级直方图、TOC 脏标记 |
| Repair | 用户拒绝预览 → cancelled；应用失败 → 回滚点 |
| Golden | U2-reformat-basic |

### 7.5 word.read-structure（新）

| 项 | 内容 |
|---|---|
| 纯读 | ListParagraphs / GetParagraphInfo |
| Data | items[{ref,style,text,listString}] |
| 用途 | exploratory plan 第一步；禁止只 Debug.WriteLine |

---

## 8. UI 映射

| UI | Capability / 消息 |
|---|---|
| 校对侧栏 | word.proofread |
| 排版模板选择 | word.semantic-reformat |
| Agent 时间线 | 所有 step |
| 上下文控制台 | ContextPack + skill |
| 修订接受/拒绝 | 若走修订模式（可选） |
| 续写面板 | word.continue-write |

---

## 9. 缺口与优先级

| ID | 缺口 | 优先级 |
|---|---|---|
| W-GAP-1 | 读工具结构化 Data 未回 AI | P0 |
| W-GAP-2 | 快路径未统一 RunTrace | P0 |
| W-GAP-3 | ExecuteJsonCommand 仍在巨型 ChatControl | P0 |
| W-GAP-4 | 翻译/续写未进 Capability 地图 | P1 |
| W-GAP-5 | 多级列表/跨节编号 | P1 |
| W-GAP-6 | 直接格式预览 | P2 |

---

## 10. 验收标准

1. Registry 中每个 Capability 有 Observe/Repair/Explain 非空。
2. 24 个 word tools 均映射到 Capability 或显式 `unassigned` 清单。
3. 快路径与 Loop 的 Trace step 字段同构。
4. Golden：U2、W-proofread-basic、W-numbering-fix、读工具 Data 断言。
5. 新增 NL 分支不得落在 ChatControl（架构评审检查表）。

---

## 11. 决策摘要（评审）

- [x] 同意现有 4 Capability 为快路径核心
- [x] 同意快路径必须 ToolResult + Trace（**D9**）
- [x] 同意读工具 Data 为 P0（W-GAP-1）
- [x] 同意 ChatControl 不再新增业务分支
- [x] 同意翻译/续写后续 Capability 化（P1）

---

## 12. 落地顺序

1. 已完成：读工具 Data 回传。
2. 已完成：基础写工具 before/after/diff Observation。
3. 已完成：轻量 Trace 埋点进入 `agent_run` / `agent_run_step`。
4. 下一步：快路径全量 ToolResult 适配器。
5. 下一步：迁出 Executor 出 ChatControl。
6. 下一步：登记 replace/toc/page-setup，翻译/续写 Capability 化。

---

*Word 是三端样板：Excel/PPT runtime 专项应复用同一 CapabilityDescriptor 形状。*
