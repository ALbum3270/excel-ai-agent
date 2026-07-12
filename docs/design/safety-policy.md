# 专项设计：SafetyGate 与办公安全策略

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 目标设计已评审；代码部分实现 |
| 实现状态 | **部分实现**：工具 JSON 有 `riskLevel` 字段；`Agent/Execution/SafetyGate.vb` 已接入 `ToolRegistry.ExecuteToolAsync`，在 COM 前同步裁决。当前最小版默认拒绝 VBA（`VBA_DISABLED`），risky/删除/全文替换返回 `SAFETY_NEEDS_APPROVAL` 且不执行；完整审批 UI、ContextPack 风险、影响面阈值和 RunTrace 事件仍待后续。 |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §5.6 |
| 现有代码 | `Agent/Execution/SafetyGate.vb`、`Agent/Execution/SafetyChecker.vb`（VBA/代码字符串子模块）、工具 JSON 的 `riskLevel` 字段、`ToolRegistry` 执行前裁决 |
| 关联 | [`tool-result-observation.md`](./tool-result-observation.md)、[`context-pack-schema.md`](./context-pack-schema.md) risks |

---

## 1. 目标与非目标

### 1.1 目标

1. 所有工具调用在执行前经过 **统一 SafetyGate**（不仅是 VBA 字符串扫描）。  
2. 风险分级与 **默认策略**（自动 / 确认 / 拒绝）对产品可配置、对用户可理解。  
3. 与 ContextPack.risks、Skill.allowed-tools、用户设置协同。  
4. 高风险操作可审计（谁、何时、什么工具、是否批准）。  
5. 从「关键词黑名单」演进到「**工具声明 + 参数影响面 + 文档状态**」三维评估。  

### 1.2 非目标

- 不实现完整企业 DLP / 租户合规中台（预留钩子即可）。  
- 不替代杀毒或 Office 宏安全中心。  
- 不在 v1 做云端策略下发。  

### 1.3 原则

| 原则 | 说明 |
|---|---|
| 默认安全 | 未知工具按 risky 处理 |
| 可逆优先 | 易 Undo 的 medium 可自动；难 Undo 的升级 |
| 最小惊讶 | 批量删除、清表、关文档必须确认或拒绝 |
| 透明 | 拒绝/确认必须有 userMessage，禁止静默失败 |
| 分层 | Gate 不执行 COM；只裁决 |

---

## 2. 现状与差距

| 现有 | 问题 |
|---|---|
| `SafetyChecker` 扫 VBA/代码子串 | 已降级为 VBA/代码子模块 |
| Tool JSON 有 riskLevel | 已在 `ToolRegistry` 执行前最小聚合 |
| Loop 对 risky 仅 Debug 日志 | 最小版已由 `SafetyGate` 返回 `SAFETY_NEEDS_APPROVAL`；审批 UX 待补 |
| 文档保护/大表 | 未系统进入裁决 |
| 快路径 Capability | 可能绕过检查 |

**目标**：一切进入 Executor 的调用，`SafetyGate.Evaluate(toolCall, context) → Decision`。

---

## 3. 风险等级

| Level | 定义 | 默认策略 | 示例 |
|---|---|---|---|
| **safe** | 只读或无持久副作用 | AutoAllow | ListParagraphs, DataAnalysis(只读模式), memory.search |
| **medium** | 可逆写、范围可控 | AutoAllow（可配置为 Confirm） | FormatText 选区, ApplyFormula 选区, 改主题色 |
| **risky** | 难逆、大范围、宏、结构破坏 | **RequireApproval** | DeleteRowCol 大范围, ExecuteVBA, 删幻灯片, ReplaceText 全文, Unprotect |
| **forbidden** | 绝对禁止 | **Deny** | 杀进程、写注册表、任意 Shell、未授权网络文件删除 |

### 3.1 等级上调规则（强制）

即使工具声明为 medium/safe，出现下列 **上下文因子** 时上调：

| 因子 | 来自 | 上调 |
|---|---|---|
| 目标单元格/段落数 > T1 | 参数或 selection | medium→risky |
| 目标 > T2 | | risky 保持并强制 Approval |
| 全文替换 | 参数 scope=document | → risky |
| 文档保护/只读 | ContextPack.risks | 写操作 → Deny 或 NeedsUnlock |
| 无 Undo | Host 声明 | medium→risky |
| VBA/宏 | toolId 或代码 | ≥ risky |
| 跨工作簿/另存覆盖 | 参数 | → risky |
| Skill 未声明该 tool | Broker | → Deny（`TOOL_NOT_ALLOWED`） |

**默认阈值（可配置）**

| 阈值 | Word | Excel | PPT |
|---|---|---|---|
| T1 | 50 段 | 2,000 格 | 3 页 |
| T2 | 200 段 | 20,000 格 | 10 页 |

---

## 4. SafetyDecision 契约

```json
{
  "decision": "allow | require_approval | deny",
  "riskLevel": "safe | medium | risky | forbidden",
  "reasons": [
    { "code": "LARGE_RANGE", "message": "将修改 25000 个单元格" }
  ],
  "userMessage": "此操作将删除 3 张幻灯片，是否继续？",
  "allowedToAutoRepair": true,
  "requiresUndoPoint": true,
  "policyVersion": "1.0"
}
```

| decision | RunController 行为 |
|---|---|
| allow | 直接执行 |
| require_approval | 状态 `AwaitingApproval`，UI 展示 userMessage；超时策略见 §8 |
| deny | 返回 ToolResult.Failed(SAFETY_BLOCKED)，**不调用 COM** |

`allowedToAutoRepair`：若因参数过大被拒，允许 Planner 改小范围后重试；若 forbidden，则 false。

---

## 5. 三维评估模型

```text
SafetyGate.Evaluate =
  BaseRisk(tool.descriptor.riskLevel)
  + ParamImpact(toolCall.parameters, selection)
  + DocumentState(context.risks, document flags)
  + PolicyOverrides(userSettings, skill constraints)
  → clamp to final risk → map to decision
```

### 5.1 BaseRisk

来自 Tool JSON / Capability 注册表，**必填**。缺失 = risky。

建议在每个 `ShareRibbon/Tools/**/*.json` 固化：

```json
{
  "id": "DeleteRowCol",
  "riskLevel": "risky",
  "sideEffect": "destructive",
  "undoReliable": true,
  "requiresSelection": false
}
```

### 5.2 ParamImpact 估算器（按工具族）

| 工具族 | 估算方法 |
|---|---|
| Range 写/删 | 解析 A1 高宽或用 selection.itemCount |
| Replace | scope + 预计匹配（若可知） |
| 幻灯片删/插 | 页数参数 |
| VBA | 代码长度 + SafetyChecker 子串结果 |
| 文件/另存 | 路径是否覆盖 |

估不出时：**宁可高估**（按 risky）。

### 5.3 DocumentState

消费 ContextPack.risks：

| risk code | 对写操作 |
|---|---|
| DOC_PROTECTED | Deny 或 RequireApproval+说明解锁 |
| LARGE_RANGE | 上调 |
| MACROS_PRESENT | VBA 相关更严 |
| TRACK_REVISIONS | 允许但提示 |

### 5.4 PolicyOverrides（用户设置，产品化预留）

| 设置键 | 含义 | 默认 |
|---|---|---|
| safety.autoMedium | medium 自动执行 | true |
| safety.vbaEnabled | 允许 VBA 工具 | false |
| safety.batchThreshold | 覆盖 T1 | 内置 |
| safety.alwaysConfirmDelete | 删除类总确认 | true |

---

## 6. 与现有 SafetyChecker 的关系

`SafetyChecker` **降级为 VBA/代码专用子模块**：

```text
If tool is ExecuteVBA or freeform code:
    sub = SafetyChecker.Check(code)
    if Blocked → decision=deny (forbidden)
    if NeedConfirm → decision=require_approval
Else:
    不走字符串黑名单，走工具声明模型
```

黑名单保留并扩展时 **只作用于代码字符串**，避免误伤 JSON 参数里的正常字段名。

### 6.1 建议保留/扩展的代码规则

**Deny（forbidden）**

- Shell / Kill / WScript / 注册表写删  
- 任意 DeleteFile/Folder  
- Application.Quit、强制 Close 不保存（可配置）  
- Environ 读敏感（可降为 approval）  

**RequireApproval**

- Delete/Clear 大范围  
- SaveAs 覆盖  
- 关闭工作簿/文档  

---

## 7. 审批 UX 契约

### 7.1 展示

```text
┌ 需要确认 ─────────────────────┐
│ 风险：高                                      │
│ 将执行：DeleteSlide × 3                       │
│ 原因：批量删除幻灯片                          │
│ 范围：第 5–7 页                               │
│ [拒绝]  [允许本次]  [允许本轮同类]（可选 v2） │
└───────────────────────────────────────────────┘
```

### 7.2 协议（WebView ↔ Host）

已有 `agent:approveStep` / `agent:rejectPlan` 等，统一语义：

| 消息 | 含义 |
|---|---|
| safety:approve | 批准当前 Decision |
| safety:reject | 拒绝；Run 标记该步失败或取消 Run |
| safety:approveSessionRiskClass | v2：本轮自动批准同 risk 类 |

### 7.3 超时

| 配置 | 默认 | 行为 |
|---|---|---|
| approvalTimeoutSec | 300 | 超时 = reject，errorCode=`SAFETY_NEEDS_APPROVAL` |

---

## 8. 与 Skill / Broker 的交叉门禁

执行前 **全部满足** 才 allow：

```text
1) tool ∈ Skill.allowed-tools（若本轮绑定了 Skill）
2) tool.application 匹配 appType 或 common
3) tool 健康度 available
4) SafetyGate.decision ≠ deny
5) require_approval 已获批
```

任一失败 → 不调用 Executor，直接结构化 ToolResult。

---

## 9. 审计日志

每次 Evaluate 记一条（AppLogger + 可选 DB）：

```json
{
  "turnId": "...",
  "toolId": "DeleteRowCol",
  "decision": "require_approval",
  "riskLevel": "risky",
  "reasons": ["LARGE_RANGE"],
  "approved": true,
  "approvedAt": "...",
  "userId": "local"
}
```

禁止记录：完整文档内容、API Key、单元格全量数据。

---

## 10. 工具风险登记表示例（节选）

| toolId | base | 备注 |
|---|---|---|
| ListParagraphs / GetParagraphInfo | safe | |
| DataAnalysis（只读） | safe | 若写回则 medium |
| FormatText / SetParagraphFormat | medium | |
| ApplyFormula / WriteData | medium | 大范围上调 |
| ReplaceText | medium/risky | scope=all → risky |
| DeleteText / DeleteRowCol / DeleteSlide | risky | |
| ProtectSheet / 取消保护 | risky | |
| ExecuteVBA（各端） | risky | 默认用户关 |
| memory.search | safe | |
| memory.promote | medium | |
| MCP 未标注 | risky | |

完整表应在实现时由脚本从 Tools JSON 生成，并在 PR 中检查 **无 risk 字段则失败**。

---

## 11. 与 Run 状态机的集成

```text
RunningStep
  → Evaluate
       allow → Execute → Observe
       require_approval → AwaitingApproval
            approve → Execute → Observe
            reject  → ToolResult(SAFETY_*) → Reflect/Stop
       deny → ToolResult(SAFETY_BLOCKED) → Reflect/Stop
```

**Repair 限制**：对 `SAFETY_BLOCKED` / `forbidden` 默认 **不自动 repair**（避免模型死循环改参撞墙），除非 reason 为 `RANGE_TOO_LARGE` 且 `allowedToAutoRepair=true`。

---

## 12. 产品文案原则

| 场景 | 文案要点 |
|---|---|
| 拒绝 | 说明「为什么不能做」+「可以怎么改」（缩小选区/关保护） |
| 确认 | 说明影响面数字，不恐吓 |
| VBA 关闭 | 提示可在设置开启，或改用原生工具 |

---

## 13. 验收标准

1. 无 Safety 评估的执行路径在架构测试中为 0（含快路径）。  
2. `ExecuteVBA` 在 `vbaEnabled=false` 时 100% deny。  
3. 超过 T2 的删除类 100% require_approval。  
4. Skill 外工具 100% deny（`TOOL_NOT_ALLOWED`）。  
5. 拒绝结果出现在 UI 时间线且含 userMessage。  
6. 审计日志无密钥与大段正文。  

---

## 14. 开放问题

| # | 问题 | 建议默认 |
|---|---|---|
| Q1 | medium 是否默认确认？ | 否，依赖 alwaysConfirmDelete 等细项 |
| Q2 | 「允许本轮同类」是否 v1？ | 否，放 v2 |
| Q3 | MCP 工具风险谁标注？ | 默认 risky，配置可覆盖 |
| Q4 | 企业策略文件格式？ | 预留 JSON policy 路径，v1 不读 |

---

## 15. 落地顺序（实现时）

1. 为全部 Tools JSON 补齐 risk/sideEffect/undoReliable  
2. 实现 SafetyGate.Evaluate + 单测矩阵  
3. Broker 接入 Skill 门禁  
4. Loop 接入 AwaitingApproval  
5. VBA 设置开关  
6. 快路径强制 Evaluate  
7. 审计落盘  

---

## 16. 决策摘要（供评审勾选）

- [x] 同意四档风险与默认策略表  
- [x] 同意 T1/T2 默认阈值（**影响面权威，D15**）  
- [x] 同意 VBA 默认关闭（**D4**）  
- [x] 同意 Skill allowed-tools 硬拒绝（**D5**）  
- [x] 同意 SafetyChecker 仅用于代码字符串  
- [x] 同意审批超时 = 拒绝  

### 与审批模式

需批准时走 Harness **模式 B**（`awaiting_approval` 立即返回，见 office-harness-api / **D2**），不在 UI 线程同步死等。

开放问题冻结：Q1 medium 默认不确认；Q2 session 同类批准 v2；Q3 MCP 默认 risky；Q4 企业策略文件 v1 不读。

---

*Safety 与 Context/Observe 共同构成「看见 → 裁决 → 证明」；T1/T2 为全局影响面权威阈值。*
