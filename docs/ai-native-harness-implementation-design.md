# Office AI Agent · AI Native Harness 落地实施设计

> **文档类型**：工程落地设计  
> **状态**：实施入口；基于当前代码渐进改造，不做颠覆性重构  
> **适用范围**：后续编码、任务拆分、验收与回归  
> **上游目标设计**：[`ai-native-harness-design.md`](./ai-native-harness-design.md)  
> **专项合同**：[`design/`](./design/)

---

## 0. 结论

当前代码已经具备 AI Native 的可执行骨架：`ChatRoutingOrchestrator`、`AgentKernel`、`LoopEngine`、`ToolRegistry`、目录型 Skills、`ToolResult` 错误契约、Word fast-path 结构化结果回灌、P0 guardrails、`OfficeHarness` adapter、执行期 `allowed-tools` 硬门禁、最小 `SafetyGate`、轻量 RunTrace 落库。下一步不应推倒重写，而应继续沿现有链路补齐三端 Observation/Diff、Golden 回归和完整审批/回放能力。

落地原则：

1. **先收口，不重写**：`IOfficeHarness` 第一版只 adapter 现有 `ChatRoutingOrchestrator`/`AgentKernel`。
2. **先执行安全，再体验完美**：H1 先做 `allowed-tools` 执行期硬拒绝、读工具 Data、最小 Observation。
3. **先 Word 闭环，再复制到 Excel/PPT**：Word 已有样板，Excel/PPT 不抢先抽象大框架。
4. **RunTrace 先轻量，后回放**：`agent_run`/`agent_run_step` 已落库；后续补 `agent_run_event`、回放 UI 和 orphan 清理。
5. **旧路径标记并包裹，不立即删除**：避免破坏三端已有功能。

---

## 1. 当前代码事实

| 领域 | 当前事实 | 不能误读为 |
|---|---|---|
| 产品主路径 | `BaseChatControl` → `ChatRoutingOrchestrator` → `AgentKernelService` → `OfficeHarness` adapter → `AgentKernel` → `LoopEngine` | 最终 Harness 状态机、审批控制 API、ContextHub 尚未完整替换 |
| ToolResult | 已有 `Success/Message/Data/ErrorCode/UserMessage/DebugDetail/Recoverable/ToolId/Observation/UndoPointId/Artifacts`；Agent 原生工具已强制 ToolResult 回执 | 已有完整 `DocumentDiff` 与 UndoPoint 绑定 |
| Skill | 已有目录型 Skill、front matter、`allowed-tools` 元数据；执行期已按 primary Skill 硬拒绝 | 已有完整 Skill/Health/Trace 三维 Broker |
| Tool 校验 | `TryNormalizeToolCall` 校验工具存在与 appType；`ToolExecutionContext` 校验 primary Skill allowed-tools；`SafetyGate` 已做最小执行前裁决；Prompt 可见工具已收紧 | 已有完整 Safety/Skill/Health 三维 Broker |
| Word | fast-path 有结构化结果回灌；读工具已回 `ToolResult.Data`；基础写工具已有 before/after/diff Observation，并进入 `ExecutionExplanation`/RunTrace | 已覆盖翻译/续写/所有 Word 能力 Diff |
| Excel | 工具和直接操作能力较多 | 已有 `ExcelActionHarness` |
| PowerPoint | 工具和生成 Handler 存在 | 已有 `PptActionHarness` |
| Trace | `ChatContextTrace`、Agent 卡片、`AppLogger`；`agent_run`/`agent_run_step` 轻量落库已接 `OfficeHarness` | 已有完整事件回放、审批轨迹和崩溃 orphan 清理 |
| Golden | smoke/code checks 已有 | 已有场景黄金集 |

---

## 2. 本轮裁剪

下列目标暂不作为 H0/H1 必做，避免过早重构：

| 裁剪项 | 处理方式 | 原因 |
|---|---|---|
| 完整替换 `BaseChatControl` | 不做；只新增 Harness adapter 并逐步接入 | 三端 ChatControl 仍承载大量 UI/宿主桥接 |
| 完整 `IModelClient` + 原生 tool-calling | H3 后再做 | 当前问题在执行链路，不在模型网关 |
| 完整 RunTrace 回放 | H2 已有轻量表；`agent_run_event`、回放 UI、orphan 清理后续做 | 先把 ToolResult/Observation 做实 |
| 完整 DocumentDiff | H1 做最小 Observation；H2 扩展 Diff | 三端 COM 差分复杂，先保执行正确 |
| Excel/PPT 大 Harness 一次到位 | 不做；Word 先闭环，再复制 | 降低跨端风险 |
| PowerQuery 完整执行 | v1 不做 | 高风险且不影响 P0 主链路 |
| 删除旧 JSON Skill | 不立即删；标 deprecated，只读兼容 | 避免历史功能断裂 |

---

## 3. 目标状态

阶段目标不是“重写成理想架构”，而是把当前链路逐步收敛成下面的稳定形态：

```text
Surface(Chat/Ribbon)
  → IOfficeHarness.RunAsync(UserTurn)
      → Context snapshot
      → Skill recall/detail
      → VisibleTools = app tools ∩ skill.allowed-tools ∩ safety
      → AgentKernel/LoopEngine(adapter first)
      → ToolRegistry.ExecuteToolAsync(context, call)
      → Host executor(COM)
      → ToolResult + Observation
      → Step event / RunTrace
      → Explain / UI timeline
```

---

## 4. H0：Harness Adapter 收口

### 4.1 目标

建立正式入口，但不改变核心行为。H0 完成后，新代码只调用 `IOfficeHarness`，旧代码仍可通过 adapter 工作。

### 4.2 新增文件

| 文件 | 说明 |
|---|---|
| `ShareRibbon/Agent/Harness/IOfficeHarness.vb` | Harness 接口 |
| `ShareRibbon/Agent/Harness/OfficeHarness.vb` | adapter 实现，内部转现有 `AgentKernel` 或 `ChatRoutingOrchestrator` 能力 |
| `ShareRibbon/Agent/Harness/HarnessModels.vb` | `UserTurn`、`HarnessRunResult`、`HarnessStepEvent` 等轻量 DTO |
| `ShareRibbon/Agent/Harness/HarnessEvents.vb` | Plan/Step/Approval/Failed 事件模型 |

新增 `.vb` 必须加入 `ShareRibbon/ShareRibbon.vbproj`。

### 4.3 最小接口

```vbnet
Public Interface IOfficeHarness
    Event PhaseChanged As EventHandler(Of HarnessPhaseChangedEventArgs)
    Event StepChanged As EventHandler(Of HarnessStepChangedEventArgs)
    Event ContextReady As EventHandler(Of HarnessContextEventArgs)

    Function RunAsync(turn As UserTurn, cancellationToken As Threading.CancellationToken) As Task(Of HarnessRunResult)
End Interface
```

H0 暂不实现 `ApproveAsync/CancelAsync/ResumeAsync`，只在 DTO 中预留状态。

### 4.4 DTO 最小字段

| DTO | 字段 |
|---|---|
| `UserTurn` | `TurnId`, `SessionId`, `AppType`, `Text`, `Mode`, `References`, `HostContextText` |
| `HarnessRunResult` | `RunId`, `Status`, `UserMessage`, `DebugMessage`, `AgentSessionId`, `StartedAt`, `FinishedAt` |
| `HarnessStepEventArgs` | `RunId`, `StepIndex`, `ToolId`, `Description`, `Status`, `Message`, `ErrorCode` |

### 4.5 接入方式

当前 `BaseChatControl` 的智能路径：

```text
HandleSendMessage
  → ChatRoutingOrchestrator.RouteSmartModeAsync
  → StartAgentPlanningFlow
  → AgentKernelService / AgentKernel
```

H0 不改掉这条路径，而是新增：

```text
OfficeHarness.RunAsync
  → 复用 host 提供的 appType/context/executor 回调
  → 调 AgentKernel.ExecuteAsync
  → 转换 AgentResult 为 HarnessRunResult
```

H0 结束后，`ChatRoutingOrchestrator` 可以保留，但只作为 `OfficeHarness` 的内部适配依赖或过渡 host。

### 4.6 验收

1. `ShareRibbon` 构建通过。
2. Word/Excel/PPT 原有智能模式不回退。
3. 新增单元/脚本检查：能实例化 `OfficeHarness` 并返回 `HarnessRunResult`。
4. `BaseChatControl` 不新增业务关键词分支。

---

## 5. H1：ToolBroker 上下文与 Skill 硬门禁

### 5.1 目标

解决当前最关键的执行正确性问题：模型不能调用当前 Skill 未声明的工具，未知工具不能假成功，读工具必须把结构化数据返回给 Agent。

### 5.2 新增/修改类型

| 类型 | 动作 |
|---|---|
| `ToolExecutionContext` | 新增，承载 `AppType`, `RunId`, `PrimarySkillName`, `AllowedTools`, `CorrelationId` |
| `ToolRegistry.ExecuteToolAsync` | 增加 overload：`ExecuteToolAsync(context, toolId, params)` |
| `LoopEngine` | 从 `AgentSession.SelectedSkill/RequiredTools` 构造 `ToolExecutionContext` |
| `AgentSkill.RequiredTools` | 明确语义为 allowed set，后续重命名但先兼容 |

### 5.3 ToolExecutionContext

```vbnet
Public Class ToolExecutionContext
    Public Property AppType As String
    Public Property RunId As String
    Public Property CorrelationId As String
    Public Property PrimarySkillName As String
    Public Property AllowedTools As HashSet(Of String)
    Public Property EnforceAllowedTools As Boolean = True
End Class
```

### 5.4 执行期门禁

`ToolRegistry.ExecuteToolAsync(context, toolId, params)` 的顺序必须是：

1. 工具存在性校验。
2. appType 支持校验。
3. 若 `context.EnforceAllowedTools=True` 且 `AllowedTools` 非空：
   - `toolId ∈ AllowedTools`：继续。
   - 否则返回 `ToolResult.Failed(..., errorCode:="TOOL_NOT_ALLOWED", recoverable:=True)`，不调用 COM。
4. Safety 初版校验。
5. 执行工具。

无 Skill 命中时按 D14：允许全 app 工具，但仍过 Safety。

### 5.5 Prompt 可见工具收紧

当前 `PromptManager.BuildSystemPrompt(appType, tools, ...)` 接收 appType 工具列表。H1 改为由 Loop/Harness 传入 `VisibleTools`：

```text
VisibleTools = ToolRegistry.GetAvailableTools(appType)
  ∩ SelectedSkill.RequiredTools(if any)
  - unavailable/error tools
```

这样模型在规划阶段尽量看不到越权工具；执行期仍做硬拒绝，防止 repair 或幻觉绕过。

### 5.6 验收

1. 构造 Word primary Skill，调用 Excel/PPT 工具，返回 `TOOL_NOT_ALLOWED` 或 `HOST_UNSUPPORTED`，不得调用宿主 executor。
2. Skill 未声明 `ExecuteVBA` 时，模型调用 `ExecuteVBA` 必须拒绝。
3. 无 Skill 命中时，Word app 工具可用，但 Excel/PPT 工具不可用。
4. Repair prompt 中包含当前可用工具列表。

---

## 6. H1：Word 读工具 Data 回传

### 6.1 目标

修复 `WordAi/ChatControl.vb` 中 `ListParagraphs`、`GetParagraphInfo` 只 `Return True`、数据只写 Debug 的问题。读工具必须把结构化 `Data` 返回给 Agent。

### 6.2 不做的事

H1 不要求完整 `DocumentDiff`，只要求读工具 `Data` 可被 `ToolResult` 传回 Loop。

### 6.3 建议改法

旧 native tool 曾通过 `ExecuteCodeWithResult` 返回 Boolean，无法带数据；当前 Agent 主路径已改为强制 `ExecuteCodeWithToolResult`，Boolean 只保留给非 Agent 的手动代码执行兼容入口。

| 类型 | 动作 |
|---|---|
| `CodeExecutionService` | 新增 `ExecuteCodeWithToolResult` 委托，返回 `ToolResult` |
| `AgentKernelService` | 绑定 `ExecuteCodeWithToolResult`，不再回退 Boolean |
| `ToolRegistry` | native tool 执行时只消费 ToolResult |
| `WordAi.ChatControl.ExecuteJsonCommand` | 保留 Boolean override；新增内部 `ExecuteJsonCommandForToolResult` |

### 6.4 Data 结构

`ListParagraphs`：

```json
{
  "items": [
    { "index": 1, "style": "正文", "text": "...", "listString": "", "outlineLevel": 10 }
  ],
  "total": 120,
  "truncated": true
}
```

`GetParagraphInfo`：

```json
{
  "index": 3,
  "text": "...",
  "style": "标题 1",
  "fontName": "仿宋",
  "fontSize": 16,
  "alignment": "Justify",
  "listString": "一、"
}
```

### 6.5 验收

1. Word 中调用 `ListParagraphs`，Loop observation 中能看到 `data.total/items`。
2. 读工具成功不创建 Undo 点。
3. 读工具失败有 `ErrorCode` 和 `UserMessage`。

---

## 7. H1：最小 Observation

### 7.1 目标

写工具不再只返回“执行成功”，至少能告诉 Agent 和用户：改了哪个范围、是否可撤销、是否可能 noop。

### 7.2 ToolResult 扩展

在不破坏旧调用的前提下给 `ToolResult` 增加可选字段：

```vbnet
Public Property Observation As Object
Public Property UndoPointId As String
Public Property Artifacts As Object
```

H1 不强制所有工具都有复杂 diff，但写工具必须至少填：

```json
{
  "kind": "write",
  "summary": "在光标位置插入 1 段文本",
  "targetRefs": ["Word:Selection"],
  "changed": true,
  "warnings": []
}
```

### 7.3 执行器最小接入

| 宿主 | H1 必做 |
|---|---|
| Word | `InsertText`, `ReplaceText`, `DeleteText`, `FormatText` 返回最小 Observation |
| Excel | `WriteData`, `ApplyFormula`, `FormatRange` 返回 target range |
| PPT | `CreateSlides`, `InsertText`, `FormatSlide` 返回 slide index |

### 7.4 FormatObservation

`LoopEngine.FormatObservation` 优先使用：

1. `ToolResult.Observation.summary`
2. `ToolResult.ToObserveSummary()`
3. `ToolResult.Message`

### 7.5 验收

1. “帮我写个请假申请”执行后，Agent 卡片不再出现“0 个迭代成功”，并显示插入目标/结果。
2. 写工具失败时仍能返回错误 Observation。
3. smoke 能断言写工具成功时 `Observation` 非空。

---

## 8. H2：轻量 RunTrace

**当前状态**：已完成轻量版。`ShareRibbon/Storage/Migrations/011_agent_run_trace.sql`、`OfficeAiDbSchema.current.sql` 和 `OfficeAiDatabase` 已创建 schema version 11；`SqliteRunTraceStore` 已由 `OfficeHarness` 写入 `agent_run` 和 `agent_run_step`。当前仍不是完整回放系统：`agent_run_event`、审批事件、崩溃 orphan 清理和 UI 回放待后续。

### 8.1 目标

让一次 Agent run 可回放、可测试。H2 先做最小数据库表，不追求完整审计中台。

### 8.2 SQLite 表

新增迁移版本时必须幂等创建：

```sql
CREATE TABLE IF NOT EXISTS agent_run (
  run_id TEXT PRIMARY KEY,
  turn_id TEXT,
  session_id TEXT,
  app_type TEXT,
  status TEXT,
  user_text TEXT,
  started_at TEXT,
  finished_at TEXT,
  final_message TEXT,
  error_code TEXT
);

CREATE TABLE IF NOT EXISTS agent_run_step (
  step_id TEXT PRIMARY KEY,
  run_id TEXT NOT NULL,
  seq INTEGER NOT NULL,
  tool_id TEXT,
  status TEXT,
  message TEXT,
  error_code TEXT,
  observation_json TEXT,
  started_at TEXT,
  finished_at TEXT
);
```

`agent_run_event` 可 H2.5 再补。

### 8.3 新增类型

| 类型 | 位置 | 说明 |
|---|---|---|
| `IRunTraceStore` | `ShareRibbon/Agent/Harness/RunTraceStore.vb` | 已实现：抽象写入 |
| `SqliteRunTraceStore` | `ShareRibbon/Agent/Harness/RunTraceStore.vb` | 已实现：SQLite 轻量写入 |
| `NoopRunTraceStore` | `ShareRibbon/Agent/Harness/RunTraceStore.vb` | 已实现：测试/禁用兜底 |

### 8.4 写入时机

1. `OfficeHarness.RunAsync` 开始：insert `agent_run(status=running)`。
2. 每个 ToolResult 后：insert `agent_run_step`。
3. 完成/失败：update `agent_run.status/final_message/error_code`。
4. 崩溃遗留 run：下次启动可标记 `PROCESS_ORPHAN`。

### 8.5 验收

1. 任意智能执行有一条 `agent_run`。
2. 有 N 次 tool 调用则至少 N 条 `agent_run_step`。
3. 失败 step 有 `error_code`。
4. `scripts/smoke-db-schema-drift.ps1` 更新并通过。

---

## 9. H2：最小 DocumentDiff

**当前状态**：Word 基础写工具已完成最小 Diff。`InsertText`、`FormatText`、`ReplaceText`、`DeleteText` 会采集执行前后文档字符数、段落数、文档文本 hash、选区文本/格式 hash、目标 preview，并在 `ToolResult.Observation` 中输出 `before`、`after`、`diff`、`changed`、`warnings`。`LoopEngine` 会把 `ObservationJson` 和读工具 `DataSummaryJson` 写入 `ExecutionExplanation`，再进入 Agent 卡片和 RunTrace step。Excel/PPT 仍待补同构最小 Diff。

### 9.1 目标

不是做完美 diff，而是让 repair 和用户解释有证据。

### 9.2 最小采样

| 宿主 | before/after 指标 |
|---|---|
| Word | 已落地：段落数、文档字符数、文档 hash、选区文本 hash/preview、选区格式 hash、目标 preview、find match count |
| Excel | Range 值 hash、公式 hash、错误单元格计数、UsedRange 行列 |
| PPT | 幻灯片数、目标页标题、形状数、备注文本 hash |

### 9.3 接口

```vbnet
Public Interface IHostSnapshotProvider
    Function CaptureSnapshot(targetRefs As IList(Of String)) As HostSnapshot
    Function Diff(before As HostSnapshot, after As HostSnapshot) As HostDocumentDiff
End Interface
```

COM 具体实现放各宿主项目，不放 `ShareRibbon`。

### 9.4 验收

1. `InsertText` 后 Word diff.changed=true。
2. Excel 写公式后目标 Range formula hash 变化。
3. PPT 新增页后 slideCount delta > 0。
4. noop 能被标记 warning，不误报成功修改。

---

## 10. H2/H3：三端 Harness 化

### 10.1 Word 优先闭环

Word 的落地顺序：

1. `ListParagraphs`/`GetParagraphInfo` Data 回传。
2. `InsertText`/`ReplaceText`/`DeleteText` Observation。
3. fast-path `WordCapabilityExecutionResult` 适配成 `ToolResult`。
4. RunTrace step 同构。
5. 翻译/续写注册为 Capability。

### 10.2 Excel

新增 `ExcelActionHarness`，但先只覆盖三类高频：

| Capability | 工具 | 验收 |
|---|---|---|
| `excel.write-table` | `WriteData` | 写入新 Sheet 或目标 Range，有 Observation |
| `excel.apply-formula` | `ApplyFormula` | 填充后错误单元格可观察 |
| `excel.create-chart` | `CreateChart` | 返回 chart artifact/ref |

### 10.3 PowerPoint

新增 `PptActionHarness`，先覆盖：

| Capability | 工具 | 验收 |
|---|---|---|
| `ppt.create-deck` | `CreateSlides` | slideCount 增加，标题可读 |
| `ppt.insert-text` | `InsertText` | 目标页/shape 可定位 |
| `ppt.beautify` | `BeautifySlides`/`FormatSlide` | 返回页范围和 warnings |

---

## 11. Safety 最小落地

H1 先做执行前同步裁决，不做复杂审批 UI。

### 11.1 SafetyDecision

```vbnet
Public Enum SafetyAction
    Allow
    RequireApproval
    Deny
End Enum
```

字段：`Action`, `Reason`, `UserMessage`, `ErrorCode`, `RiskLevel`。

### 11.2 H1 规则

| 规则 | 默认 |
|---|---|
| `ExecuteVBA` | 若配置未开启，Deny `VBA_DISABLED` |
| `RiskLevel=risky` | RequireApproval；若 UI 未支持审批，返回 `SAFETY_NEEDS_APPROVAL` |
| 跨宿主工具 | Deny `HOST_UNSUPPORTED` |
| 批量删除/清空 | RequireApproval |
| 工具未知风险 | 按 medium/risky 处理 |

### 11.3 验收

1. 未开启 VBA 时 `ExecuteVBA` 不进入 COM。
2. 删除/清表类工具不自动执行。
3. Deny/RequireApproval 返回结构化 `ToolResult`。

---

## 12. Prompt 与模型约束

### 12.1 PromptManager 改动

Prompt 必须注入：

1. 当前 `appType`。
2. 当前 `primarySkill` 名称和摘要。
3. 当前 `VisibleTools`，不能注入全量工具。
4. 上一步 `ToolResult.ToObserveSummary()`。
5. 失败时的 `errorCode/userMessage/availableTools`。

### 12.2 禁止

- 不再让模型自由发明 `clear_document`、`replace_text` 这类未注册工具。
- 不让 repair 输出完整计划文本；repair 只输出一个合法 tool call。
- 不把无关历史会话塞入 system prompt。

---

## 13. 文档与代码同步规则

每完成一个阶段，必须同步改这些位置：

| 阶段 | 需要同步 |
|---|---|
| H0 | `design/office-harness-api.md` 实现状态；总纲热力图 |
| H1 Skill | `design/skill-runtime-and-gates.md` 实现状态；ErrorCode 表 |
| H1 Observation | `design/tool-result-observation.md` 实现状态 |
| H2 Trace | `design/run-trace-storage.md` 实现状态；DB smoke |
| H2 Diff | `design/context-pack-schema.md`、三端 runtime 专项 |
| Excel/PPT Harness | 对应 runtime 文档状态 |

---

## 14. 实施任务表

| ID | 阶段 | 任务 | 文件范围 | 验收 |
|---|---|---|---|---|
| H0-1 | H0 | 新增 `IOfficeHarness`/DTO/adapter | `ShareRibbon/Agent/Harness/*` | 已完成：`ShareRibbon` build 通过 |
| H0-2 | H0 | `BaseChatControl` 智能路径可经 Harness adapter | `ShareRibbon/Controls/*` | 已完成：`AgentKernelService.StartAgentAsync` 经 `OfficeHarness` 调用同一 `AgentKernel` |
| H1-1 | H1 | 新增 `ToolExecutionContext` | `ShareRibbon/Agent/*` | 已完成：工具执行可拿到 Skill/app/run |
| H1-2 | H1 | 执行期 `allowed-tools` 硬拒绝 | `ToolRegistry`, `LoopEngine` | 已完成：越权工具 `TOOL_NOT_ALLOWED` |
| H1-3 | H1 | VisibleTools 收紧 Prompt | `PromptManager`, `AgentKernel` | 已完成：Prompt/repair 使用可见工具 |
| H1-4 | H1 | Word 读工具 Data 回传 | `WordAi/ChatControl.vb`, `CodeExecutionService` | 已完成：`ListParagraphs/GetParagraphInfo` 返回 data |
| H1-5 | H1 | 最小 Observation 字段 | `ToolResult`, 三端 executor | 部分完成：字段已加，Word 基础写工具已回 before/after/diff observation；Excel/PPT 待补 |
| H1-6 | H1 | Safety 最小裁决 | `SafetyGate`, `ToolRegistry` | 已完成最小版：VBA 默认 `VBA_DISABLED`；risky/删除/全文替换返回 `SAFETY_NEEDS_APPROVAL`，不进入 COM |
| H2-1 | H2 | `agent_run`/`agent_run_step` 迁移 | DB 初始化/迁移 | 已完成：schema version 11，schema drift/empty DB smoke 通过 |
| H2-2 | H2 | `IRunTraceStore` | `ShareRibbon/Agent/Harness/RunTraceStore.vb` | 已完成轻量版：每次 Harness run 写 `agent_run`，step explanation 写 `agent_run_step` |
| H2-3 | H2 | Word 最小 Diff | `WordAi/ChatControl.vb`, `LoopEngine` | 已完成：基础写工具 Observation 含 before/after/diff，Explanation/RunTrace 可读取 |
| H2-4 | H2 | Excel/PPT 最小 Diff | `ExcelAi`, `PowerPointAi` | range/slide delta 可观察 |
| H3-1 | H3 | `ExcelActionHarness` | `ExcelAi/Services` | U3 跑通 |
| H3-2 | H3 | `PptActionHarness` | `PowerPointAi/Services` | U4 跑通 |
| H3-3 | H3 | Golden runner 初版 | `scripts`, `fixtures` | P0 场景可回归 |

---

## 15. 验证命令

日常代码变更：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run-code-checks.ps1 -Configuration Debug
```

文档变更至少跑：

```powershell
git diff --check
```

如果当前环境不是 git repo，则改用：

```powershell
Get-ChildItem docs -Recurse -Filter *.md | Select-String -Pattern "评审通过（见|Frozen|已全部实现"
```

---

## 16. 第一轮建议落地顺序

最小可交付闭环建议如下：

1. `H1-1` + `H1-2` + `H1-3`：已完成，执行期门禁和 Prompt 可见工具收紧已降低幻觉工具风险。
2. `H1-4`：已完成，Word 读工具 Data 可返回 Loop。
3. `H1-5`：部分完成，Word 基础写工具已有 before/after/diff Observation；Excel/PPT 仍待补。
4. `H0-1` + `H0-2`：已完成，入口经 `OfficeHarness` adapter 收口，未重写三端 ChatControl。
5. `H1-6`：已完成最小 SafetyGate，同步阻断 VBA/风险工具直接执行。
6. `H2-1` + `H2-2`：已完成轻量 RunTrace，为 Golden 做准备。
7. `H2-3`：已完成 Word 基础写工具最小 Diff；下一步补 Excel/PPT 最小 Diff，并把 Golden 场景覆盖“请假申请写入 Word”。

这组顺序比先做完整 `IOfficeHarness` 更稳，因为它优先修复当前用户可见失败：工具越权、读不到上下文、假成功、无法追踪。
