# 专项设计：RunTrace 存储、回放与 conversation_event 关系

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 评审通过（见 [`design-review-record.md`](./design-review-record.md)） |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §4.6 / §9 |
| 关联 | Observe / Safety / Skill 专项；Memory 管线 |
| 现有 | `conversation_event`、`conversation`、`memory_*`、`AgentSession` 内存对象、AppLogger 文件日志 |

---

## 1. 目标与非目标

### 1.1 目标

1. 每一轮 Agent/Harness 执行有 **可回放** 的 `agent_run` 记录。  
2. 与现有 `conversation_event` **分工清晰**，不重复造聊天历史。  
3. 支持：失败归因、评测回放、用户「查看本轮步骤」、审计（脱敏）。  
4. 默认本地 SQLite，迁移版本化。  

### 1.2 非目标

- 不做分布式 tracing（OpenTelemetry 可后续）。  
- 不存全文文档快照（只存 ref + fingerprint + 短 preview）。  
- 不替代 AppLogger 文本日志。  

### 1.3 设计原则

| 原则 | 说明 |
|---|---|
| 事件源友好 | run 头 + 有序 step/event 子表 |
| 脱敏默认 | 不落 API Key、Authorization、超长正文 |
| 可截断 | preview/payload 有硬顶 |
| 可关联 | turnId / sessionId / correlation 贯通 |
| 可选详尽 | Debug 档可多写；Release 默认摘要 |

---

## 2. 与现有表的分工

| 存储 | 职责 |
|---|---|
| `conversation` | 旧版会话消息展示（可逐步弱化） |
| `conversation_event` | **对话层** 用户/助手消息与记忆管线输入 |
| **`agent_run`（新）** | **执行层** 一次 Harness 运行 |
| **`agent_run_step`（新）** | 步骤级 tool 调用与结果摘要 |
| **`agent_run_event`（新）** | 细粒度事件（审批、门禁、repair） |
| `memory_item` | 长期记忆，可通过 source 关联 run/event |
| AppLogger 文件 | 运维诊断，非产品回放主源 |

### 2.1 一次用户发送的数据流

```text
User 发送
  → conversation_event (role=user, event_type=user_message)
  → agent_run 创建 (status=running)
  → … steps / events …
  → conversation_event (role=assistant, event_type=assistant_message,
        metadata 含 run_id)
  → agent_run 结束 (status=completed|failed|cancelled)
  → memory_job 可选引用 run_id
```

**不变量**：有工具执行的智能模式必须有 `agent_run`；纯闲聊可无 run 或 run 无 step。

---

## 3. 逻辑模型

### 3.1 AgentRun

| 字段 | 类型 | 说明 |
|---|---|---|
| run_id | TEXT PK | `run_` + uuid |
| turn_id | TEXT | 与 AppLogger correlation / ContextPack.turnId |
| session_id | TEXT | 会话 |
| app_type | TEXT | Word/Excel/PowerPoint |
| document_id | TEXT | 可选，文档指纹或路径 hash |
| document_name | TEXT | 显示名 |
| user_text | TEXT | 用户原话（可截断 4KB） |
| status | TEXT | see §5 |
| skill_name | TEXT | primary skill |
| task_spec_json | TEXT | TaskSpec 摘要 JSON |
| plan_json | TEXT | ExecutionPlan 摘要（可截断） |
| final_message | TEXT | 用户可见总结（截断 8KB） |
| error_code | TEXT | 失败时 |
| error_summary | TEXT | |
| step_count | INT | |
| repair_count | INT | |
| approval_count | INT | |
| started_at | TEXT | |
| finished_at | TEXT | |
| duration_ms | INT | |
| model_name | TEXT | 可选 |
| schema_version | TEXT | RunTrace 契约版本 `1.0` |
| meta_json | TEXT | 扩展 |

### 3.2 AgentRunStep

| 字段 | 类型 | 说明 |
|---|---|---|
| step_id | TEXT PK | |
| run_id | TEXT FK | |
| seq | INT | 从 0 递增 |
| plan_step_number | INT | 可选 |
| tool_id | TEXT | |
| risk_level | TEXT | |
| params_json | TEXT | **脱敏截断**，默认 ≤ 4KB |
| success | INT | 0/1 |
| error_code | TEXT | |
| user_message | TEXT | |
| observation_summary | TEXT | |
| target_refs_json | TEXT | |
| before_fp | TEXT | |
| after_fp | TEXT | |
| diff_json | TEXT | 压缩 diff，≤ 8KB |
| undo_point_id | TEXT | |
| elapsed_ms | INT | |
| repair_attempt | INT | |
| data_summary_json | TEXT | 读工具结果摘要，非全文 |
| started_at / finished_at | TEXT | |

### 3.3 AgentRunEvent

用于非 tool 的里程碑：

| event_type 示例 | payload |
|---|---|
| run_started | { } |
| context_captured | { packId, purpose, truncated } |
| skill_recalled | { names[] } |
| skill_bound | { name } |
| plan_generated | { stepCount, autoRun } |
| safety_decision | { toolId, decision, risk } |
| approval_requested / approval_resolved | { approved } |
| repair_started / repair_finished | { attempt } |
| replan | { reason } |
| skill_gate_denied | { toolId } |
| run_completed / run_failed / run_cancelled | { } |
| memory_enqueued | { jobId } |

| 字段 | 类型 |
|---|---|
| event_id | TEXT PK |
| run_id | TEXT FK |
| seq | INT |
| event_type | TEXT |
| payload_json | TEXT |
| created_at | TEXT |

---

## 4. SQL 草案（实现时迁移）

> 版本号在实现时分配；以下为语义草案。

```sql
CREATE TABLE IF NOT EXISTS agent_run (
  run_id TEXT PRIMARY KEY,
  turn_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  app_type TEXT DEFAULT '',
  document_id TEXT DEFAULT '',
  document_name TEXT DEFAULT '',
  user_text TEXT,
  status TEXT NOT NULL,
  skill_name TEXT DEFAULT '',
  task_spec_json TEXT,
  plan_json TEXT,
  final_message TEXT,
  error_code TEXT DEFAULT '',
  error_summary TEXT DEFAULT '',
  step_count INTEGER DEFAULT 0,
  repair_count INTEGER DEFAULT 0,
  approval_count INTEGER DEFAULT 0,
  started_at TEXT NOT NULL,
  finished_at TEXT,
  duration_ms INTEGER DEFAULT 0,
  model_name TEXT DEFAULT '',
  schema_version TEXT DEFAULT '1.0',
  meta_json TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_session
  ON agent_run(session_id, started_at);
CREATE INDEX IF NOT EXISTS idx_agent_run_turn
  ON agent_run(turn_id);
CREATE INDEX IF NOT EXISTS idx_agent_run_status
  ON agent_run(status, started_at);

CREATE TABLE IF NOT EXISTS agent_run_step (
  step_id TEXT PRIMARY KEY,
  run_id TEXT NOT NULL,
  seq INTEGER NOT NULL,
  plan_step_number INTEGER,
  tool_id TEXT NOT NULL,
  risk_level TEXT DEFAULT '',
  params_json TEXT,
  success INTEGER NOT NULL,
  error_code TEXT DEFAULT '',
  user_message TEXT,
  observation_summary TEXT,
  target_refs_json TEXT,
  before_fp TEXT,
  after_fp TEXT,
  diff_json TEXT,
  undo_point_id TEXT,
  elapsed_ms INTEGER DEFAULT 0,
  repair_attempt INTEGER DEFAULT 0,
  data_summary_json TEXT,
  started_at TEXT,
  finished_at TEXT,
  FOREIGN KEY (run_id) REFERENCES agent_run(run_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_agent_run_step_run
  ON agent_run_step(run_id, seq);

CREATE TABLE IF NOT EXISTS agent_run_event (
  event_id TEXT PRIMARY KEY,
  run_id TEXT NOT NULL,
  seq INTEGER NOT NULL,
  event_type TEXT NOT NULL,
  payload_json TEXT,
  created_at TEXT NOT NULL,
  FOREIGN KEY (run_id) REFERENCES agent_run(run_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_agent_run_event_run
  ON agent_run_event(run_id, seq);
```

**conversation_event.metadata_json 约定**

```json
{
  "run_id": "run_xxx",
  "turn_id": "turn_xxx",
  "skill_name": "word-document-agent"
}
```

---

## 5. Run 状态机

```text
running
  → awaiting_approval → running
  → completed
  → failed
  → cancelled
```

| status | 含义 |
|---|---|
| running | 执行中 |
| awaiting_approval | 阻塞在 Safety |
| completed | 正常结束（可有部分步失败但整体成功策略另定） |
| failed | 不可恢复失败 |
| cancelled | 用户终止 |

**整体成功策略（建议）**

- 默认：最后用户目标达成（verifier 或无失败写步）→ completed  
- 任一步 `forbidden` 拒绝导致无法继续 → failed  
- 用户点终止 → cancelled  

---

## 6. 写入 API（逻辑）

```text
IRunTraceStore
  BeginRun(startInfo) → run_id
  AppendEvent(run_id, type, payload)
  AppendStep(run_id, stepRecord)
  UpdateRun(run_id, patch)          ' status, plan_json, counts...
  CompleteRun(run_id, result)
  GetRun(run_id) → AgentRunDetail
  ListRuns(session_id, limit) → summaries
  ExportRun(run_id) → json file DTO
```

### 6.1 写入时机

| 时机 | 调用 |
|---|---|
| Harness 入口 | BeginRun |
| Context 完成 | AppendEvent context_captured |
| Skill 绑定 | skill_bound |
| Plan 生成 | UpdateRun.plan_json + plan_generated |
| 每步前后 | AppendStep（结束时一次完整写，或 start/finish 两次） |
| Safety | safety_decision / approval_* |
| Repair | repair_* + step.repair_attempt |
| 结束 | CompleteRun + conversation_event assistant |

### 6.2 失败写入

进程崩溃可能导致 run 停在 running：

- 启动时 `ReconcileOrphanRuns`：超过 N 分钟的 running → failed/`error_code=PROCESS_ORPHAN`  

---

## 7. 脱敏与截断策略

| 数据 | 策略 |
|---|---|
| params | 删除 key/password/token 字段；字符串 > 500 截断 |
| user_text | > 4KB 截断 + truncated 标记 |
| observation preview | 每 area ≤ 200 字 |
| diff_json | 总 ≤ 8KB |
| data 读工具 | 只存 count + 前 N 条 ref，不存全表 |
| 路径 | 可存 hash 或文件名，完整路径可选 |

复用 `AppLogger.Redact` 规则。

---

## 8. 回放与产品 UI

### 8.1 回放 DTO

```json
{
  "run": { "run_id": "...", "status": "completed", ... },
  "timeline": [
    { "kind": "event", "type": "plan_generated", "at": "..." },
    { "kind": "step", "toolId": "FormatText", "success": true, "summary": "..." }
  ]
}
```

### 8.2 UI

- Agent 时间线：读 `List`/`Get` 当前 turn 的 run  
- 「导出本轮诊断」：`ExportRun` → 用户本地 JSON（已脱敏）  
- 不在 v1 做「自动重放执行」（危险）；回放 = **只读查看**  

### 8.3 评测消费

黄金集跑完后对比：

- step 序列 toolId  
- 最终 error_code  
- observation.changed / metrics  

见 [`eval-golden-set.md`](./eval-golden-set.md)。

---

## 9. 与 Memory 的关联

| 关系 | 方式 |
|---|---|
| 对话沉淀 | conversation_event.metadata.run_id |
| 记忆来源 | memory_item 可增 `source_run_id`（可选迁移） |
| 任务级记忆 | meta 中 `from_run` |

v1 最小：只保证 conversation_event 能跳到 run；memory 列可 H4 再加。

---

## 10. 保留与清理策略

| 项 | 默认 |
|---|---|
| 保留时长 | 30 天 |
| 单 session 最多 run | 200，超出删最旧 |
| 用户清除 | 「清除记忆」时可级联删 agent_run* |
| Debug 库 | `OfficeAiAppData-Debug` 分离（现有策略） |

后台 job：`agent_run_gc` 每日或启动时。

---

## 11. 详细级别

| 模式 | 写 step.diff_json | 写 params 全量 |
|---|---|---|
| release | summary only | 截断 |
| debug | 完整压缩 diff | 更长截断 |
| 配置 | `trace.detail=minimal\|standard\|debug` | |

---

## 12. 验收标准

1. 智能模式跑通一次后，`agent_run` 有且仅有一条对应 turn_id。  
2. 含 3 次 tool 调用则 `agent_run_step` 至少 3 行，seq 有序。  
3. 审批拒绝产生 event 且无 COM 副作用（配合 Safety 测）。  
4. Export JSON 无 `sk-` / Bearer。  
5. 迁移可在空库与旧库升级。  
6. orphan running 可被 reconcile。  

---

## 13. 开放问题（已冻结）

| # | 问题 | **冻结默认** |
|---|---|---|
| Q1 | 是否把 plan 全量存库？ | **摘要 + step**；完整 plan 仅 debug 档 |
| Q2 | 多 run 并行？ | **D3**：同 session 仅 1 running |
| Q3 | SQLite FTS 搜 run？ | **v2** |
| Q4 | conversation 旧表？ | **v1 双写**；长期弱化 conversation 展示表 |

### 与 conversation_event（写死）

| 存储 | 写什么 |
|---|---|
| conversation_event | 用户可见对话与记忆管线源事件 |
| agent_run* | 工具步骤、审批、门禁、观察摘要 |
| 关联 | assistant event 的 `metadata_json.run_id` |

---

## 14. 决策摘要（评审）

- [x] 同意 agent_run / step / event 三表  
- [x] 同意脱敏截断策略  
- [x] 同意回放只读、不自动重放执行  
- [x] 同意同 session 单 running（**D3**）  
- [x] 同意 PR 不强制 L1，但 H1 后埋点必须可测  

---

## 15. 落地顺序（实现时）

1. 迁移 DDL + Repository  
2. Harness Begin/Complete 埋点  
3. Step/Event 埋点  
4. UI 时间线读库  
5. Export + GC  
6. 评测读取  

---

*RunTrace 是「可证明的执行历史」；没有它，Observe/Safety/Skill 门禁都难以回归与客诉定位。*
