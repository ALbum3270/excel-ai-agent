# 专项设计：IOfficeHarness API 与 Surface 协议

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 目标设计已评审；代码部分实现 |
| 实现状态 | **H0 adapter 已落地并进入主路径**：`ShareRibbon/Agent/Harness` 已有 `IOfficeHarness`、`OfficeHarness`、轻量 DTO 和 `IRunTraceStore`；`AgentKernelService.StartAgentAsync` 已经通过 `OfficeHarness` adapter 调用现有 `AgentKernel`；轻量 RunTrace 已写 `agent_run`/`agent_run_step`。完整 `ApproveAsync/CancelAsync/ContextHub/agent_run_event` 和回放 API 仍待后续。 |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §3 / Phase H0 |
| 现有 | `ChatRoutingOrchestrator`、`AgentKernel`/`AgentKernelService`、WebView 消息、`ExecutionPathPolicy` |
| 关联 | 全部 design/* 专项；实现时 Surface 只依赖本 API |

---

## 1. 目标与非目标

### 1.1 目标

1. 冻结 **唯一产品入口** `IOfficeHarness` 的方法、事件、DTO。  
2. 规定 Chat / Ribbon / 快捷入口如何调用与订阅，**禁止**旁路进 Executor。  
3. 对齐 WebView 消息与宿主事件，便于前后端并行。  
4. 为 Phase H0 提供「先接口后实现」的编码合同。  

### 1.2 非目标

- 本专项不规定最终 VB 实现细节；当前 H0 已有 adapter。
- 不规定 COM 细节。  
- 不替代 Tool JSON schema。  

---

## 2. 设计原则

| 原则 | 说明 |
|---|---|
| 单一入口 | 所有 NL 智能执行走 `RunAsync` |
| 异步优先 | 返回 `Task`；进度靠事件 |
| 可取消 | `CancellationToken` |
| 可测试 | 接口可 mock；DTO JSON 可序列化 |
| UI 无知业务 | Surface 不解析 Tool 参数业务含义 |
| 兼容迁移 | 旧 `AgentKernelService` 可作第一版适配器 |

---

## 3. 顶层接口

```text
Interface IOfficeHarness
  ' 主入口
  Function RunAsync(turn As UserTurn, Optional ct As CancellationToken = Nothing) As Task(Of HarnessResult)

  ' 运行期控制
  Function ApproveAsync(runId As String, approval As ApprovalDecision, Optional ct As CancellationToken = Nothing) As Task
  Function CancelAsync(runId As String, Optional reason As String = Nothing) As Task
  Function SubmitUserInputAsync(runId As String, input As UserInputReply, Optional ct As CancellationToken = Nothing) As Task

  ' 查询
  Function GetRunAsync(runId As String) As Task(Of AgentRunDetail)
  Function GetActiveRunId(sessionId As String) As String

  ' 事件（.NET Event 或 callback 总线）
  Event OnPhaseChanged(e As HarnessPhaseEvent)
  Event OnPlanReady(e As PlanReadyEvent)
  Event OnStepStarted(e As StepEvent)
  Event OnStepFinished(e As StepEvent)
  Event OnApprovalRequired(e As ApprovalRequiredEvent)
  Event OnUserInputRequired(e As UserInputRequiredEvent)
  Event OnExplain(e As ExplainEvent)
  Event OnContextTrace(e As ContextTraceEvent)
  Event OnCompleted(e As HarnessCompletedEvent)
  Event OnFailed(e As HarnessFailedEvent)
```

**实现类建议名**：`OfficeHarness`（ShareRibbon），构造注入：

- `IContextHub`
- `ISkillRouter`
- `IToolBroker`
- `ISafetyGate`
- `IMemoryRuntime`（检索注入 + 回合提取入队）
- `IRunTraceStore`
- `IHostExecutor`（按 app 解析）
- `IModelClient`（gateway）

第一期可用 `OfficeHarnessAdapter` 包装现有 Kernel，逐步换心。

---

## 4. 核心 DTO

### 4.1 UserTurn

```json
{
  "turnId": "turn_...",
  "sessionId": "sess_...",
  "appType": "Word",
  "text": "用户原话",
  "mode": "smart",
  "references": {
    "files": [{ "path": "...", "name": "..." }],
    "selections": [{ "id": "...", "preview": "..." }]
  },
  "uiHints": {
    "proofreadPanelActive": false,
    "templateId": null
  },
  "options": {
    "enableMemory": true,
    "autoRunSafePlan": true
  }
}
```

| mode 值 | 含义 | Harness 行为 |
|---|---|---|
| smart | 默认智能 | 完整 Analyze+可能 Agent |
| agent | 强制 Agent | 跳过「仅追问闲聊」短路可配置 |
| chat | 强制纯聊 | 不跑工具（仍可记 trace 空 run） |
| proofread | 校对模式 | 绑定 proofread capability 偏好 |
| template | 模板渲染 | 专用 system，少工具 |

**mode 只是偏好，不是关键词路由表。**

### 4.2 HarnessResult

```json
{
  "runId": "run_...",
  "turnId": "turn_...",
  "status": "completed|failed|cancelled|needs_input|awaiting_approval",
  "finalUserMessage": "已完成……",
  "errorCode": "",
  "skillName": "word-document-agent",
  "stats": {
    "steps": 4,
    "repairs": 1,
    "durationMs": 12000
  }
}
```

注意：`awaiting_approval` / `needs_input` 时 `RunAsync` **可以**：

- **A)** 阻塞直到 Approve/Input（简单但占线程）  
- **B)** 立即返回中间状态，等 `ApproveAsync` 继续（推荐）  

**冻结 B（D2）**：`RunAsync` 在需审批时 **立即返回** `status=awaiting_approval`（不长期阻塞调用方线程）；内部 run 状态机挂起；`ApproveAsync` 恢复执行并继续发事件。

### 4.3 ApprovalDecision

```json
{
  "approved": true,
  "scope": "once",
  "note": ""
}
```

`scope`: `once` | `session_same_risk`（v2）

### 4.4 UserInputReply

```json
{
  "answers": [
    { "questionId": "q1", "text": "只处理第三章" }
  ]
}
```

---

## 5. 事件契约（Surface 订阅）

### 5.1 事件一览

| 事件 | 何时 | UI 动作 |
|---|---|---|
| OnPhaseChanged | Analyzing/Planning/Running… | 状态文案 |
| OnContextTrace | Pack 就绪 | 上下文控制台 |
| OnPlanReady | Plan 生成 | Plan 卡片；autoRun 则直接转圈 |
| OnApprovalRequired | Safety 需批 | 确认对话框 |
| OnUserInputRequired | 阻塞澄清 | 最小提问卡 |
| OnStepStarted/Finished | 每步 | 时间线 |
| OnExplain | 解释块 | Markdown/卡片 |
| OnCompleted/Failed | 终态 | 收尾、撤销入口 |

### 5.2 PlanReadyEvent

```json
{
  "runId": "...",
  "plan": {
    "summary": "将统一正文为宋体 12 磅并重排编号",
    "steps": [
      { "index": 1, "description": "读取段落结构", "risk": "safe" },
      { "index": 2, "description": "应用字体", "risk": "medium" }
    ],
    "autoRunEligible": true
  }
}
```

若 `autoRunEligible=false` 或用户设置总是确认 Plan → UI 显示「开始执行」按钮，对应消息 `harness:startExecution` → 内部等同 Approve plan。

### 5.3 StepEvent

```json
{
  "runId": "...",
  "seq": 2,
  "toolId": "FormatText",
  "status": "running|succeeded|failed",
  "userMessage": "已更新 3 段字体",
  "errorCode": "",
  "canUndo": true,
  "observationSummary": "..."
}
```

### 5.4 ApprovalRequiredEvent

```json
{
  "runId": "...",
  "requestId": "apr_...",
  "riskLevel": "risky",
  "userMessage": "将删除 3 行，是否继续？",
  "toolId": "DeleteRowCol",
  "reasons": ["LARGE_RANGE"]
}
```

---

## 6. WebView 消息协议（与事件双射）

### 6.1 浏览器 → Host

| type | 载荷 | Harness API |
|---|---|---|
| sendMessage / smart send | 文本+引用 | 组装 UserTurn → `RunAsync` |
| startAgent | request | `RunAsync` mode=agent |
| harness:approve | runId, requestId, approved | `ApproveAsync` |
| harness:cancel | runId | `CancelAsync` |
| harness:userInput | runId, answers | `SubmitUserInputAsync` |
| harness:startExecution | runId | Plan 级 approve |
| abortAgent / cancelLoop | | `CancelAsync`（兼容旧名） |

**废弃主路径**：`startLoop` 仅 remap（已有），新前端不发。

### 6.2 Host → 浏览器

| JS 函数/消息 | 来源事件 |
|---|---|
| showContextHints | OnContextTrace |
| showAgentPlanningStatus / plan card | OnPlanReady |
| agent step timeline APIs | OnStep* |
| showSafetyConfirm | OnApprovalRequired |
| showMinimalQuestions | OnUserInputRequired |
| completeAgent | OnCompleted/Failed |
| addExplainBlock | OnExplain |

命名可在实现时适配现有 `agent-card.js`；本专项要求 **语义对齐**，允许旧函数名包装。

---

## 7. Ribbon / 非 Chat 入口

| 入口 | 组装 UserTurn |
|---|---|
| Ribbon「智能处理」 | text=默认提示或剪贴板；selection 自动带 |
| 右键选区 | text 可空 → U1 行为 |
| 校对按钮 | mode=proofread |
| 模板按钮 | mode=template + uiHints.templateId |

全部 `RunAsync`，禁止直接调 `WordActionHarness`（Harness 内部可调）。

---

## 8. 与现有类型适配映射

| 现有 | Harness |
|---|---|
| ChatRoutingOrchestrator.RouteSmartModeAsync | RunAsync 前半 Analyze+路由 |
| AgentKernelService.StartAgentAsync | RunAsync 执行后端 |
| AgentKernel events | 翻译为 OnStep/OnPlan… |
| ExecutionPathPolicy | Harness 内部策略对象 |
| GlobalStatusStrip | Surface 订阅 Failed/Phase |

迁移期允许：

```text
BaseChatControl
  → IOfficeHarness (adapter)
       → ChatRouting + AgentKernelService
```

目标期：

```text
BaseChatControl
  → OfficeHarness
       → ContextHub + SkillRouter + Loop + Broker + HostExecutor
```

---

## 9. 并发与会话规则

| 规则 | 默认 |
|---|---|
| 同 session 同时 running | 1；新 Run 前取消或拒绝旧 run |
| 跨 session | 允许 |
| Cancel | 尽力停在下一步边界；已写变更保留 Undo |
| Dispose | 宿主窗格关闭 → Cancel active run |

---

## 10. 错误面

| 场景 | Result.status | errorCode |
|---|---|---|
| 无文档 | failed | DOC_MISSING |
| 用户取消 | cancelled | USER_CANCELLED |
| 安全拒绝 | failed 或 step fail | SAFETY_BLOCKED |
| 模型失败 | failed | NETWORK_ERROR / TIMEOUT |
| 无进展熔断 | failed | NO_PROGRESS |

`OnFailed.userMessage` 必须可展示；细节进 Trace。

---

## 11. 序列图（主路径）

```text
UI                Harness              Trace        HostExec
 │ RunAsync         │                   │              │
 │─────────────────>│ BeginRun          │              │
 │ ContextTrace     │──────────────────>│              │
 │<─────────────────│                   │              │
 │ PlanReady        │                   │              │
 │<─────────────────│                   │              │
 │ (auto or start)  │                   │              │
 │ StepStarted      │                   │              │
 │<─────────────────│ Execute           │              │
 │                  │─────────────────────────────────>│
 │                  │ ToolResult        │              │
 │                  │<─────────────────────────────────│
 │ StepFinished     │ AppendStep        │              │
 │<─────────────────│──────────────────>│              │
 │ Completed        │ CompleteRun       │              │
 │<─────────────────│──────────────────>│              │
```

审批分支在 Step 前插入 ApprovalRequired ↔ ApproveAsync。

---

## 12. 验收标准（H0）

1. 接口与 DTO 有独立文档/示例 JSON（本文件）。  
2. Surface 编译期仅依赖 `IOfficeHarness`（适配器阶段可放宽）。  
3. 契约测试：序列化 UserTurn/HarnessResult 往返。  
4. 旧 startLoop/abort 消息仍可用。  
5. 同 session 双开 Run 有明确拒绝或取消策略。  

---

## 13. 开放问题

| # | 问题 | 建议默认 |
|---|---|---|
| Q1 | RunAsync 是否在审批时挂起 Task？ | 否，返回中间状态（模式 B） |
| Q2 | 事件总线用 .NET Event 还是 IObservable？ | v1 Event + UI 线程投递 |
| Q3 | chat mode 是否创建 run？ | 创建 status=completed 无 step，便于统一会话 |
| Q4 | 多窗口同进程？ | run 绑定 sessionId+appType |

---

## 14. 落地顺序

1. 定义 DTO + 接口空实现  
2. Adapter 接 Kernel  
3. BaseChatControl 改为只调 Harness  
4. 事件 → JS 适配层  
5. 换心 Context/Skill/Loop  
6. 删除 Surface 内业务路由  

---

## 15. 决策摘要（评审勾选）

- [x] 同意唯一入口 `RunAsync`（**D1**）  
- [x] 同意审批模式 B（非阻塞返回）（**D2**）  
- [x] 同意 mode 仅为偏好  
- [x] 同意同 session 单 running（**D3**）  
- [x] 同意第一期 Adapter 包装 Kernel  
- [x] 同意注入列表含 `IMemoryRuntime`  

开放问题冻结：Q1=模式 B；Q2=v1 用 .NET Event + UI 投递；Q3=chat 可建无 step 的 completed run；Q4=run 绑定 sessionId+appType。

---

*本 API 是 Surface 与 Agent 平台的防腐层；冻结合同后再编码，可避免 ChatControl 继续膨胀。*
