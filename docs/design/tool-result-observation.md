# 专项设计：ToolResult、文档观察（Observe）与修复输入

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 评审通过（见 [`design-review-record.md`](./design-review-record.md)） |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §4.5 / §5.4 / §5.5 |
| 现有代码 | `Agent/ToolRegistry.ToolResult`、`LoopEngine.FormatObservation`、`ExecutionExplanation`、`ExceptionClassifier`、`AppLogger` |
| 依赖 | [`context-pack-schema.md`](./context-pack-schema.md) 的 before/after 快照 |

---

## 1. 目标与非目标

### 1.1 目标

1. **ToolResult 成为唯一执行回执**：所有 Host/MCP/Memory/SkillScript 执行必须返回同一契约。  
2. **Observe 一等公民**：每次写操作后必须有可验证的文档观察，而不是仅 `"执行成功"`。  
3. **Repair 输入标准化**：修复提示只消费结构化失败信息 + Diff，禁止只喂原始异常字符串。  
4. **可撤销**：写操作绑定 `UndoPointId`，支持「撤销本轮/本步」。  
5. **可归类**：`ErrorCode` 覆盖率成为质量门禁。  

### 1.2 非目标

- 不实现完整 Word 修订（Track Changes）语义合并。  
- 不做像素级 PPT 视觉 diff（v1 以文本/形状计数为主）。  
- 不规定具体 COM 调用顺序（属 Executor 实现）。  

---

## 2. 现状与差距

| 能力 | 现状 | 差距 |
|---|---|---|
| ToolResult 基本字段 | Success/Message/Data/Elapsed/Error* | Observation/Undo/Artifacts 未强制 |
| FormatObservation | 成功/失败字符串 | 无文档差分 |
| Word 读工具 | ListParagraphs 等 TODO 未回传 AI | 违反读工具契约 |
| 修复循环 | 有，含 ErrorCode 摘要 | 缺 Diff、缺「重复调用」检测 |
| Undo | UndoManager 存在 | 未与 ToolResult 稳定关联 |
| 快路径 WordActionHarness | 直接改文档 | 未统一写 ToolResult/Trace |

---

## 3. ToolResult 完整契约

### 3.1 字段表

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| schemaVersion | string | 是 | `"1.0"` |
| toolId | string | 是 | 与 ToolBroker 一致 |
| success | bool | 是 | |
| message | string | 是 | 兼容旧逻辑的技术摘要 |
| userMessage | string | 是 | 面向用户，已脱敏 |
| errorCode | string | 失败时是 | 见 §6；成功为空 |
| debugDetail | string | 否 | 日志用，脱敏 |
| recoverable | bool | 是 | 默认 true；不可修复则 false |
| data | object | 否 | 结构化载荷（读工具主阵地） |
| observation | object | **写工具必填** | 见 §4 |
| elapsedMs | long | 是 | |
| undoPointId | string | 写工具建议 | 对应 UndoManager |
| artifacts | array | 否 | 生成的 sheet/slide/range 提示 |
| riskLevel | string | 是 | 执行时实际风险 |
| attempt | int | 否 | repair 第几次 |
| sideEffects | string[] | 否 | e.g. `["selection_changed"]` |

### 3.2 工厂方法语义（设计）

| 方法 | 行为 |
|---|---|
| `Succeed(toolId, message, data, observation?)` | success=true；写工具缺 observation → 开发断言/日志 Warn |
| `Failed(...)` | success=false；填 errorCode/userMessage |
| `FromException(toolId, ex)` | 走 ExceptionClassifier |
| `FromBlocked(toolId, safetyResult)` | errorCode=`SAFETY_BLOCKED`，recoverable=false 或需改参 |

### 3.3 data 约定（读工具）

读工具 **禁止** 只返回 true：

```json
{
  "toolId": "ListParagraphs",
  "success": true,
  "message": "listed 20 paragraphs",
  "data": {
    "items": [
      { "ref": "Word:Para:12", "style": "正文", "text": "……", "listString": "" }
    ],
    "total": 180,
    "truncated": true
  },
  "observation": {
    "kind": "read",
    "summary": "读取 20/180 段"
  }
}
```

---

## 4. Observation（文档观察）模型

### 4.1 顶层

```json
{
  "kind": "read | write | noop | error",
  "summary": "人类可读一句话",
  "targetRefs": ["Excel:Sheet1!B2:B100"],
  "beforeFingerprint": "fp_...",
  "afterFingerprint": "fp_...",
  "diff": { },
  "successCriteriaHints": [],
  "warnings": []
}
```

| kind | 何时 |
|---|---|
| read | 只读工具 |
| write | 任何可能改文档的工具 |
| noop | 参数导致无实际变更 |
| error | 执行抛错，仍尽量填 targetRefs |

### 4.2 DocumentDiff

```json
{
  "appType": "Word",
  "changed": true,
  "changeCount": 3,
  "areas": [
    {
      "ref": "Word:Para:12-14",
      "changeType": "text_replace | format | insert | delete | structure | value | formula | shape | slide",
      "beforePreview": "……",
      "afterPreview": "……",
      "metrics": {
        "charsDelta": -12,
        "fontChanged": true
      }
    }
  ],
  "globalMetricsDelta": {
    "paragraphCount": 0,
    "tableCount": 0,
    "slideCount": 0,
    "formulaErrorCount": -2
  },
  "truncated": false
}
```

### 4.3 三端 Diff 最小集

#### Word

| 检测项 | 方法（逻辑） | v1 是否必须 |
|---|---|---|
| 目标段文本前后 | 按 ref 取 Text | 是 |
| 样式名变化 | Style.NameLocal | 是 |
| 段落数量变化 | 文档级计数 | 是 |
| 列表/编号字符串 | ListFormat | 建议 |
| 表行列变化 | Tables 计数 | 建议 |

#### Excel

| 检测项 | 方法 | v1 |
|---|---|---|
| 目标 Range 值矩阵抽样 | Value2 哈希 | 是 |
| 公式文本变化 | Formula | 是 |
| 错误单元格计数 | SpecialCells 或抽样 | 是 |
| 行列数 UsedRange | | 是 |
| 图表个数 | ChartObjects.Count | 建议 |

#### PowerPoint

| 检测项 | 方法 | v1 |
|---|---|---|
| 幻灯片数量 | Slides.Count | 是 |
| 目标页标题文本 | | 是 |
| 形状数量 | Shapes.Count | 是 |
| 备注文本 | | 建议 |

### 4.4 Fingerprint 算法（建议）

```text
fingerprint = hash(
  appType + join(sorted(targetRefs)) +
  stable_serialize(sampled_content) +
  key_metrics
)
```

用途：

- 检测 **noop**（before==after）  
- 检测 **无进展**（连续两步 after 相同且 tool 失败或重复）  

v1 可用 SHA256 截断 16 hex；不要求密码学强度。

---

## 5. 执行与观察时序（单步）

```text
RunController 准备执行 step
  │
  ├─ SafetyGate.Check(toolCall) → allow | confirm | deny
  │
  ├─ UndoManager.CreatePoint() → undoPointId
  │
  ├─ ContextHub.Snapshot(observe_before, focusRefs=plan.targets)
  │     → packBefore
  │
  ├─ Executor.Execute(toolCall) → raw result / exception
  │     ※ 必须 UI 线程 COM
  │
  ├─ ContextHub.Snapshot(observe_after, focusRefs)
  │     → packAfter
  │
  ├─ ContextHub.Diff(packBefore, packAfter) → diff
  │
  ├─ Build ToolResult {
  │     success, data, observation{diff, fingerprints}, undoPointId
  │   }
  │
  ├─ Verify optional SuccessCriteria hooks
  │
  └─ Loop: FormatObservation(toolResult) → memory.lastObservation
        → 成功? next : Repair | Reflect
```

### 5.1 失败时仍要观察

COM 中途失败可能已部分修改文档：

- 仍尝试 after 快照  
- `success=false`，`observation.kind=error`，`diff.changed` 如实反映  
- `recoverable` 由分类器 + 是否部分变更共同决定  

### 5.2 读工具

- 可跳过 UndoPoint  
- observation.kind=read  
- diff 可省略，但 data 必填  

---

## 6. ErrorCode 目录（与 ExceptionClassifier 对齐并扩展）

| Code | 含义 | recoverable 默认 | 用户文案方向 |
|---|---|---|---|
| OK | 成功（空 code） | — | — |
| UNKNOWN | 未分类 | true | 请重试 |
| COM_ERROR | COM/跨线程/断开 | true | 检查文档是否可用 |
| NETWORK_ERROR | HTTP/API | true | 检查网络与 Key |
| TIMEOUT | 超时取消 | true | 重试或缩小范围 |
| JSON_ERROR | 解析失败 | true | 重试规划 |
| ARGUMENT_ERROR | 参数非法 | true | 修正参数 |
| NOT_FOUND | 工具/对象不存在 | false/ true | 换工具或范围 |
| IO_ERROR | 文件 IO | true | 路径权限 |
| SAFETY_BLOCKED | 安全拒绝 | false | 说明原因 |
| SAFETY_NEEDS_APPROVAL | 待批准 | true | 等待用户 |
| NO_PROGRESS | 重复无进展 | false | 停止或重规划 |
| PARTIAL_APPLY | 部分成功 | true | 继续 repair |
| VERIFY_FAILED | 成功标准未满足 | true | 再修 |
| HOST_UNSUPPORTED | 宿主无能力 | false | 换端或降级 |
| SELECTION_EMPTY | 需要选区 | true | 自动扩范围或问用户 |
| RANGE_TOO_LARGE | 超阈值 | true | 分块 |
| DOC_PROTECTED | 保护/只读 | false | 解锁 |
| VBA_DISABLED | 宏禁用 | false | 提示开启或改原生 tool |

**门禁**：生产环境 Agent 工具失败日志中 `errorCode` 为空视为缺陷。

---

## 7. Observation → 人类/模型文本

### 7.1 用户时间线（UI）

```text
✅ FormatText · 已更新 3 个段落的字体
   范围: 第 12–14 段 · 可撤销
```

```text
❌ ApplyFormula · 公式错误 (VERIFY_FAILED)
   范围: Sheet1!C2:C100 · 仍有 12 个 #DIV/0!
   正在尝试修复 (2/3)…
```

### 7.2 模型观察（Loop）

```text
[OBS]
tool=ApplyFormula success=false code=VERIFY_FAILED recoverable=true
summary=填充后仍有 12 个除零错误
diff.areas=Excel:Sheet1!C2:C100 formula
before.errors=0 after.errors=12
data.sampleErrors=["C2","C5",...]
undo=up_123
```

`FormatObservation` 必须优先使用 `ToObserveSummary()` + diff 压缩，而不是只拼 Message。

---

## 8. Repair 输入契约

### 8.1 何时 Repair

| 条件 | 动作 |
|---|---|
| success=false 且 recoverable=true 且 attempt < Max | Repair |
| success=false 且 recoverable=false | 跳过 Repair，进入 Reflect/Fail |
| success=true 但 Verify 失败 | Repair 或补步 |
| success=true 且 diff.changed=false 且期望变更 | 视为 NO_PROGRESS |
| 连续相同 toolId+params hash ≥ 2 | NO_PROGRESS 熔断 |

### 8.2 Repair Prompt 必含块

1. errorCode + userMessage + debugDetail（脱敏）  
2. observation.diff 压缩  
3. 原 toolId + parameters  
4. **当前可见 allowed-tools 列表**（防止幻觉工具）  
5. 可选：SuccessCriteria  

### 8.3 Repair 输出

仅允许：

```json
{ "toolId": "...", "parameters": { } }
```

经 ToolBroker 再校验；非法则记 `JSON_ERROR` 并计一次失败 attempt。

---

## 9. 与 SuccessCriteria 的关系

TaskSpec / Plan 可带：

```json
"successCriteria": [
  "标题层级仅使用 标题1-3",
  "C 列无公式错误"
]
```

Verifier 钩子（设计）：

```text
IStepVerifier.Verify(toolResult, packAfter, criteria) 
  → { ok, failedCriteria[], hints[] }
```

- ok=false → ToolResult 可标记 `VERIFY_FAILED`（即使 COM 未抛错）  
- Word 样板：编号连续、校对问题清零等可注册专用 verifier  

---

## 10. Undo 语义

| 层级 | 行为 |
|---|---|
| 步级 | 每写工具一个 `undoPointId` |
| 轮级 | Run 结束生成 `runUndoGroupId`，聚合本轮点 |
| UI | 「撤销上一步」「撤销本轮 Agent」 |

**不变量**

- 失败且 `diff.changed=true` 时仍保留 undoPoint，供用户回滚部分变更。  
- Host 若无法建点，observation.warnings 含 `UNDO_UNAVAILABLE`。  

---

## 11. 快路径（WordActionHarness）纳入规则

确定性 Capability 可短时绕过 LLM，但必须：

1. 构造等价 `toolId` 或 `capabilityId`  
2. 走同一套 before/after + ToolResult  
3. 写入 RunTrace  
4. 触发同一套 UI 时间线  

否则视为 **旁路违规**（总纲禁止项）。

---

## 12. RunTrace 中的步骤记录（摘要）

```json
{
  "stepId": "s3",
  "toolId": "FormatText",
  "params": { },
  "result": { "success": true, "observation": { "summary": "..." } },
  "startedAt": "...",
  "finishedAt": "...",
  "repairAttempt": 0
}
```

完整落库见未来 `run-trace-storage.md`。

---

## 13. 验收标准

1. 任意写工具单测：无 observation 的成功结果在 Debug 构建报警。  
2. 黄金 Word 场景：改字体后 diff.areas 非空且 fingerprint 变化。  
3. 故意错误公式：success 或 verify 能暴露错误码。  
4. 读工具：模型下一轮能引用 data 中的 ref。  
5. Repair 提示快照夹具可静态断言含 diff 与 allowed-tools。  
6. 快路径与 Loop 路径的 Trace 字段同构。  

---

## 14. 开放问题（已冻结）

| # | 问题 | **冻结默认** |
|---|---|---|
| Q1 | 超大 diff 如何截断？ | areas ≤ 20，preview ≤ 200 字/块 |
| Q2 | 是否对只读工具也做 fingerprint？ | **否** |
| Q3 | MCP 工具无文档 diff？ | 由 MCP 声明 kind；无则 summary only |
| Q4 | 部分 Excel 操作难 diff？ | UsedRange 哈希 + 目标 range 抽样 |

ErrorCode 权威表见 [`design-review-record.md`](./design-review-record.md) §4。

---

## 15. 决策摘要（评审）

- [x] 写工具必须 observation（**D8**）；缺省 Debug 断言/Warn  
- [x] 读工具必须结构化 `data`  
- [x] 快路径必须映射 ToolResult + Trace（**D9**）  
- [x] Repair 必含 errorCode + diff 摘要 + allowed-tools  
- [x] recoverable=false 跳过无效 repair  

---

## 16. 落地顺序（实现时）

1. 扩展 ToolResult 模型字段（兼容旧调用）  
2. Diff 服务（先 Word，后 Excel/PPT）  
3. LoopEngine 强制观察管线  
4. 修 Word 读工具 Data  
5. 快路径接入 Trace  
6. 评测夹具  

---

*关联：[`safety-policy.md`](./safety-policy.md) 决定能否执行；本文件决定执行后如何证明与修复。*
