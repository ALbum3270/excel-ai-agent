# 专项设计：Memory 运行时（单写模型、注入、隐私）

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 评审通过（见 [`design-review-record.md`](./design-review-record.md)） |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §5.7 |
| 蓝图参考 | `openspec/changes/fix-memory-rag-pipeline/agent-memory-architecture.md` |
| 现有 | `conversation_event` / `memory_item` / `memory_embedding` / `memory_job`、`atomic_memory` 兼容、`MemoryService`/`UnifiedMemoryService`/`AgentMemoryRepository`、`MemoryConfig`、`memory.*` 工具 |
| 关联 | ContextPack.memory、RunTrace、Skill、Golden U7 |

---

## 1. 目标与非目标

### 1.1 目标

1. **单一写入真相源**：新记忆只走 `conversation_event → memory_job → memory_item → memory_embedding`。  
2. `atomic_memory` **只读兼容**并设定淘汰期，不再承接新写入主路径。  
3. 记忆服务 Agent：被动 RAG + 主动 `memory.*` 工具 + 任务级记忆。  
4. **隐私产品化**：清除 / 导出 / 禁用 / 范围隔离可配置。  
5. 与 Harness 注入点明确：ContextHub 组装 `ContextPack.memory`。  

### 1.2 非目标

- 不做企业级跨设备同步。  
- 不做完整知识图谱推理引擎（`memory_graph` 仅辅助关联/冲突）。  
- 不在本专项定义 embedding 模型训练。  

### 1.3 原则（沿用蓝图并硬化）

| 原则 | 说明 |
|---|---|
| 事件不可丢，记忆可重建 | conversation_event 是源 |
| 记忆 ≠ 聊天记录 | 必须提取、打分、去重 |
| 先过滤再向量再重排 | app/document/scope/status |
| 短期不进 RAG | 与现路线一致 |
| 异步重活 | embedding/提取走 job |
| 脱敏 | 日志与导出不含 Key、可截断正文 |

---

## 2. 现状与差距

| 现状 | 问题 |
|---|---|
| atomic + memory_item 双模型 | 写入/检索路径分叉 |
| MemoryService 同步桥接 | 已缓解死锁，仍应主走 async |
| Agent memory.* 工具 | 需 EnableAgenticSearch |
| 晋升/过期/冲突 | 有实现，需统一到 job 语义 |
| 用户配置 | MemoryConfigForm 有基础开关，缺「清除全部/导出」产品流 |
| 与 RunTrace | 弱关联（设计见 run-trace-storage） |

---

## 3. 数据模型（运行时视角）

### 3.1 主路径表

| 表 | 角色 |
|---|---|
| conversation_event | 原始交互（user/assistant/tool 摘要） |
| memory_job | 提取/晋升/过期/embedding 任务 |
| memory_item | 结构化记忆 |
| memory_embedding | 向量 |
| memory_graph | 相似/冲突边 |
| user_profile | 稳定偏好（可视为 scope=user 的特化） |
| session_summary | 会话摘要（短期展示用） |

### 3.2 memory_item 类型

| memory_type | 含义 | 默认 scope | RAG |
|---|---|---|---|
| preference | 用户偏好（字体、语体） | user | 是 |
| fact | 稳定事实 | user/document | 是 |
| format_rule | 排版/公文规范 | user/document | 是 |
| task | 当前任务约定 | session/document | **否**（任务记忆通道） |
| solution | 有效解法摘要 | user/document | 是 |
| skill_feedback | Skill 成败经验 | skill | 条件 |

### 3.3 status

| status | 含义 |
|---|---|
| active | 可检索 |
| dirty | 待重算 embedding |
| expired | 不参与 RAG |
| deleted | 软删 |
| superseded | 被新版本替代（冲突保留） |

### 3.4 atomic_memory 策略

| 阶段 | 行为 |
|---|---|
| 现在 | 禁止新代码 `Insert` 到 atomic 作为主路径 |
| 读取 | RAG 可双读合并，atomic 权重降权 |
| 迁移 | 后台 job 将高 importance 迁到 memory_item |
| 淘汰 | 标记 `data_migration_marker` 后只读 N 个版本；再下线读路径 |

---

## 4. 写入管线

### 4.1 对话回合

```text
Harness / Chat 结束一轮
  → conversation_event(user)
  → conversation_event(assistant, metadata.run_id?)
  → memory_job(type=extract_turn, target=event_ids)
  → Worker:
       RuleBased + LlmMemoryExtractor（可关）
       去重 / 重要性
       upsert memory_item
       memory_job(type=embed)
```

### 4.2 Agent 工具结果（可选增强）

```text
写工具成功且 high_value
  → conversation_event(event_type=tool_result_summary)
  → 提取 solution / format_rule
```

不存全量 COM 数据，只存摘要 + refs。

### 4.3 任务级记忆（Task Memory）

| 项 | 说明 |
|---|---|
| 生命周期 | 单 run 或单 session |
| 内容 | 「本章只改标题」「输出列在 G」 |
| 注入 | ContextPack 优先于长期记忆同主题 |
| 存储 | memory_type=task, scope=session, expires_at 短 |
| 也可 | AgentMemory 工作区内存 + 结束时选择性持久化 |

### 4.4 晋升 / 过期 / 冲突

| 策略 | 规则（建议默认） |
|---|---|
| 晋升长期 | importance≥0.75 或 access_count≥3 或 type∈preference/format_rule |
| 过期短期 | importance&lt;0.3 且 14 天未访问 |
| 冲突 | 偏好类新值 supersede 旧值，graph 记 potential_conflict |
| 文档级 | document_id 变则文档 scope 记忆不串线 |

具体阈值进 `MemoryConfig` 扩展项。

---

## 5. 检索与注入

### 5.1 被动 RAG（ContextHub）

```text
Query = userText + selection.preview 摘要
Filter: status=active, memory_type≠task 的短期, app_type match or empty
Optional: document_id boost
Vector if available else keyword
Rerank: score * w1 + importance * w2 + recency * w3
TopN = MemoryConfig.RagTopN
→ ContextPack.memory
```

**不变量**：short_term atomic / 未晋升内容默认不进 RAG。

### 5.2 主动工具（Agent）

| tool | 行为 | 门禁 |
|---|---|---|
| memory.search | 长期检索 | EnableAgenticSearch |
| memory.recent | 最近长期 | 同上 |
| memory.promote | 手动晋升 | 同上 + medium 风险 |
| memory.list_session | 会话内 | 同上 |

未开启时返回明确 Failed，不静默空列表装成功。

### 5.3 与 Skill 优先级

```text
注入顺序（高→低）：
  1. 安全/文档 risks
  2. 选区与结构
  3. task memory
  4. skill 正文（已 LoadDetail）
  5. long-term memory
  6. session summary
  7. history turns
```

冲突时：显式 task memory > 长期 preference。

---

## 6. MemoryConfig 产品化

### 6.1 现有

- UseContextBuilder  
- EnableUserProfile  
- RagTopN  
- AtomicContentMaxLength  
- SessionSummaryLimit  
- EnableAgenticSearch  

### 6.2 建议新增

| 键 | 默认 | 说明 |
|---|---|---|
| EnableMemoryWrite | true | 总写入开关 |
| EnableLlmExtraction | true | 失败回退规则 |
| RetentionDays | 365 | 长期保留 |
| AllowDocumentScoped | true | 文档级记忆 |
| ExportEnabled | true | 导出 |
| ClearRequiresConfirm | true | 清除确认 |

### 6.3 用户操作（UI）

| 操作 | 行为 |
|---|---|
| 禁用记忆 | 不写不读；UI 提示 |
| 清除全部 | soft delete memory_item + 可选 event 保留 |
| 清除当前文档 | filter document_id |
| 导出 | JSON（脱敏路径/截断） |
| 查看画像 | user_profile 列表 |

路径：配置窗体扩展或独立「隐私与记忆」页。

---

## 7. 与 Harness 的集成点

```text
RunAsync
  ContextHub.Snapshot
    → MemoryRuntime.RetrieveForPack(query, app, doc)
  ...
  CompleteRun
    → MemoryRuntime.EnqueueTurnExtraction(events, runId)
```

`IMemoryRuntime` 建议接口：

```text
RetrieveForPack(...) As MemoryPackSection
EnqueueTurnExtraction(...)
SearchAsync(...)          ' 工具用
PromoteAsync(...)
ClearAsync(scope)
ExportAsync(scope) As Stream
```

---

## 8. Worker / Job 类型

| job_type | 说明 |
|---|---|
| extract_turn | 从 event 抽记忆 |
| embed | 写 embedding |
| promote_batch | 批量晋升 |
| expire_batch | 过期 |
| migrate_atomic | atomic→item |
| reembed_all | 模型切换 |

调度：启动时 + 空闲轮询；失败指数退避；`attempt_count` 上限。

---

## 9. 隐私与合规

| 要求 | 设计 |
|---|---|
| 本地默认 | DB 在用户文档/AppData |
| Debug 分离 | 现有 `-Debug` 目录 |
| 日志 | 不写记忆全文到 AppLogger（可写 id+type） |
| 导出 | 用户显式触发 |
| 第三方 | 仅 LLM API 请求中的注入片段；用户可关记忆减少上传 |

---

## 10. 验收标准

1. 新安装路径：关闭 atomic 写入后，一轮对话仍产生 memory_item（job 完成后）。  
2. RAG 单测：短期内容不出现在 RetrieveForPack。  
3. EnableMemoryWrite=false：无新 item，功能可聊。  
4. Clear 后 Retrieve 为空。  
5. Export 无 api key。  
6. memory.search 在 Agentic 关闭时 Failed 带明确 code。  
7. smoke-memory-pipeline 扩展覆盖晋升/过期（实现阶段）。  

---

## 11. 开放问题

| # | 问题 | 建议默认 |
|---|---|---|
| Q1 | 助手全文是否进 event？ | 是，截断 32KB |
| Q2 | 多用户同机？ | v1 单用户本地 |
| Q3 | embedding 模型变更？ | dirty 全量 reembed job |
| Q4 | 画像与 preference item 是否合并？ | 双写过渡，读时合并 |

---

## 12. 落地顺序

1. 冻结 IMemoryRuntime 与「禁止 atomic 新写」门禁  
2. 全路径写入 event+job  
3. ContextHub 只读 memory_item  
4. 隐私 UI  
5. atomic 迁移 job  
6. 下线 atomic RAG  

---

## 13. 决策摘要（评审勾选）

- [x] 同意 memory_item 为唯一新写真相源（**D6**）  
- [x] 同意短期不进 RAG（**D7**）  
- [x] 同意 task memory 通道  
- [x] 同意 Agentic 工具总开关默认保持可配  
- [x] 同意清除/导出为 P0 产品能力  

开放问题冻结：Q1 助手全文进 event（截断 32KB）；Q2 单用户本地；Q3 模型切换 reembed job；Q4 画像双写过渡。

---

*Memory 让 Agent「越用越懂用户」；单写模型与隐私开关是产品化底线，否则双轨会持续污染检索质量。*
