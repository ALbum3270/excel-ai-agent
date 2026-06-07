# Office AI Agent Memory / Intent / Skills 落地蓝图

## 目标

Office AI Agent 的目标不是简单保存聊天记录，而是为 Word、Excel、PowerPoint 提供类似 Claude for Office 的上下文能力：

- 能识别用户当前意图，并决定是否进入工具执行、排版、改写、分析、生成等流程。
- 能从对话、文档、文件、执行结果中提取稳定记忆。
- 能按当前 Office 应用、当前文档、当前项目和用户画像加载最相关上下文。
- 能自动召回 Skills，但不把全部 Skills 塞进 system prompt。
- 能增量更新 embedding、记忆、技能索引，避免每次全量重算。

## 设计原则

1. **事件不可丢，记忆可重建**
   - 原始交互写入 `conversation_event`。
   - 记忆、画像、embedding 都是由事件增量生成的派生数据。

2. **记忆不是聊天记录**
   - 聊天记录用于审计和回放。
   - 记忆用于未来上下文增强，必须经过提取、去重、打分、过期策略。

3. **检索先过滤，再向量召回，再重排**
   - 先按 `app_type/document_id/project_id/scope/status` 缩小候选。
   - 再做 SQLite 向量召回。
   - 最后按相似度、重要性、时间、访问频次、当前会话关系综合 rerank。

4. **Skills 渐进披露**
   - 第一阶段只加载 Skills 目录摘要。
   - 命中后才加载 1-3 个 Skill 的完整内容。
   - 带脚本的 Skill 必须经过意图和置信度判断后才提示或执行。

5. **增量任务队列**
   - 所有重活都写入 job 表，由后台异步处理。
   - embedding 失败、模型切换、schema 升级都通过 job 重试或 dirty 标记处理。

## 分层架构

```text
Office UI / BaseChatControl
        |
        v
IntentClassifier  ----> SkillSelector
        |                    |
        v                    v
ContextAssembler <---- RetrievalService
        |
        v
LLM Request
        |
        v
ConversationEventStore
        |
        v
MemoryExtractionJob -> MemoryItem -> MemoryEmbedding
```

## 数据模型

### conversation_event

原始事件表。所有对话、文件引用、选区、工具调用、执行结果都写这里。

关键字段：

- `event_id`
- `session_id`
- `app_type`
- `document_id`
- `event_type`
- `role`
- `content`
- `metadata_json`
- `created_at`
- `processed_at`

### memory_item

结构化记忆表。由事件提取，不直接等于聊天文本。

关键字段：

- `memory_id`
- `source_event_id`
- `scope`: `session/document/project/user/global/skill`
- `app_type`
- `document_id`
- `project_id`
- `memory_type`: `preference/fact/task/solution/format_rule/skill_feedback`
- `content`
- `summary`
- `confidence`
- `importance`
- `status`: `active/dirty/expired/deleted`
- `expires_at`
- `last_verified_at`
- `created_at`
- `updated_at`

### memory_embedding

向量表。与 `memory_item` 分离，便于切换 embedding 模型和重建索引。

关键字段：

- `embedding_id`
- `memory_id`
- `embedding_model`
- `embedding_dim`
- `embedding_json`
- `vector_status`: `pending/ready/failed/dirty`
- `last_error`
- `updated_at`

后续接 `sqlite-vec` 时，可以替换或并行增加虚拟向量表。

### memory_job

增量任务队列表。

任务类型：

- `extract_memory`
- `embed_memory`
- `rebuild_embedding`
- `summarize_session`
- `update_user_profile`
- `index_skill`

状态：

- `pending`
- `processing`
- `completed`
- `failed`

### skills_registry

Skills 自动加载索引。

关键字段：

- `skill_name`
- `file_path`
- `app_scope`
- `intent_types`
- `trigger_keywords`
- `description`
- `embedding_json`
- `usage_count`
- `success_count`
- `last_indexed_at`
- `enabled`

## 请求时上下文装配

每次用户发送消息时：

1. 生成轻量上下文快照：
   - 当前 Office 应用
   - 当前文档 id/path/name
   - 选区摘要
   - 附件摘要

2. 意图识别：
   - `GENERAL_QUERY`
   - `DOCUMENT_REWRITE`
   - `OFFICIAL_FORMAT`
   - `DATA_ANALYSIS`
   - `PRESENTATION_CREATE`
   - `TOOL_EXECUTION`
   - `SKILL_INVOCATION`

3. Skill 召回：
   - 先按 intent/app/keywords 过滤。
   - 再按 embedding 或关键词得分排序。
   - 只加载 top 1-3 的 detail。

4. Memory 召回：
   - scope 优先级：session > document > project > user > global。
   - 当前文档相关记忆优先。
   - 过期和低置信度记忆不注入。

5. ContextAssembler 输出：
   - system: 角色、规则、当前任务约束。
   - developer-style section: 当前文档上下文、记忆证据、Skills。
   - user: 原始用户请求。

## 增量更新流程

### 对话结束

```text
assistant response complete
        |
        v
insert conversation_event(user)
insert conversation_event(assistant)
enqueue extract_memory
enqueue summarize_session
```

### 记忆提取

```text
memory_job(extract_memory)
        |
        v
LLM / heuristic extraction
        |
        v
upsert memory_item
enqueue embed_memory
```

### Embedding 更新

```text
memory_job(embed_memory)
        |
        v
EmbeddingService.GetEmbeddingAsync()
        |
        v
upsert memory_embedding
```

### 切换 embedding 模型

```sql
UPDATE memory_embedding
SET vector_status='dirty'
WHERE embedding_model <> @current_model;
```

然后增量 enqueue `rebuild_embedding`。

## SQLite 向量路线

### 当前阶段

使用 SQLite 自定义函数：

```sql
cosine_similarity_json(embedding_json, @query_embedding_json)
```

优点：

- 不新增 native 扩展。
- 可以立即把排序和 limit 下推到 SQLite。
- 安装包风险低。

限制：

- 仍是线性扫描，不是 ANN 索引。
- 数据量大后需要升级。

### 下一阶段

接入 `sqlite-vec` 或 `sqlite-vss`：

- 新增 native DLL 打包和加载策略。
- 新增虚拟向量表。
- 将 `memory_embedding` 同步到向量表。
- 查询走 ANN topK，再回表拿 `memory_item`。

## 阶段计划

### Phase 1: Schema 与契约

- 新增 `conversation_event`
- 新增 `memory_item`
- 新增 `memory_embedding`
- 新增 `memory_job`
- 新增 `skills_registry`
- 新增 `IMemoryEventStore / IMemoryExtractor / IMemoryRetriever / ISkillSelector / IIntentClassifier`

### Phase 2: 写入链路

- BaseChatControl 在一轮对话结束后写 `conversation_event`
- 入队 `extract_memory`
- 现有 `atomic_memory` 继续保留，作为兼容层

### Phase 3: 检索链路

- RetrievalService 先读 `memory_item + memory_embedding`
- 保留 `atomic_memory` fallback
- ContextBuilder 改成按 intent/scope 选择记忆

### Phase 4: Skills 自动加载

- 启动时扫描 Skills 目录
- 写入 `skills_registry`
- 请求时按 intent/app/keywords/embedding 召回

### Phase 5: sqlite-vec

- 增加 native 扩展打包
- 建虚拟向量表
- 增量同步 memory embedding

## 当前兼容策略

- 不删除 `atomic_memory`。
- 新表先旁路写入。
- RAG 可逐步从 `atomic_memory` 迁移到 `memory_item`。
- 任一新流程失败时，不阻塞主聊天。
