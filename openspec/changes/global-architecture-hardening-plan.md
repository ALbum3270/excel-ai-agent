# Office AI Agent 全局架构优化落地计划

## 目标

当前项目已经具备 Word/Excel/PowerPoint AI 助手、Memory、Skills、MCP、排版、翻译、续写等能力，但核心问题是复杂度开始集中在少数超大类和多套重复运行时中。此计划目标是：

- 降低 UI 线程、WebView2、Office Interop 的运行时异常风险。
- 统一 LLM 请求构建、Provider 适配、Reasoning 参数和错误处理。
- 收敛 Memory/Skills/Intent/Agent 的上下文组装路径。
- 将数据库迁移、Skills 加载、Memory pipeline 变成可验证、可回归的工程能力。

## 当前主要问题

1. **超大类承担过多职责**
   - `BaseChatControl.vb`、`WordAi/ChatControl.vb`、`ConfigApiForm.vb` 同时承担 UI、业务流程、HTTP、WebView2、Office Interop、异常处理。
   - 结果是线程问题、状态同步问题、局部修改难以评估影响面。

2. **LLM 请求路径重复**
   - `HttpStreamService`、`DefaultConversationRuntime`、`LLMUtil`、`IntentRecognitionService`、`LlmMemoryExtractor`、`EmbeddingService` 分别构建请求。
   - Reasoning、Anthropic、tools、stream/non-stream 容易行为不一致。

3. **UI 线程访问缺少统一边界**
   - 大量 `Task.Run`、`Invoke`、`BeginInvoke`、`.Result`、`.Wait()` 混用。
   - WebView2 和 Office COM 都要求线程边界清晰。

4. **Memory/Skills 新旧系统并存**
   - `atomic_memory` 与 `memory_item` 并存。
   - `SkillsService`、`SkillsIndexService`、旧 `skills_usage.json`、SQLite `skills_registry` 同时存在。

5. **数据库迁移内联过重**
   - `OfficeAiDatabase.vb` 内联 schema 和迁移越来越大，长期维护风险高。

6. **异常处理和日志不统一**
   - 大量宽泛 `Catch ex As Exception`。
   - 用户提示、Debug 日志、可恢复错误没有统一模型。

## 执行原则

- 先加基础设施，再迁移调用方，避免一次性重写主链路。
- 先迁移后台/非 UI 逻辑，再迁移 WebView2 和 Office Interop。
- 每阶段必须可编译、可 smoke、可回退。
- 新路径与旧路径短期并存，但新代码必须只走新路径。

## Phase 1：统一非流式 LLM 网关

### 范围

- 新增 `AiGateway`，统一非流式 chat 请求。
- 支持 OpenAI-compatible chat completions。
- 支持 Anthropic 非流式 `/v1/messages` 基础转换。
- 统一 Reasoning 参数应用。
- 统一响应文本提取和错误返回。

### 首批迁移

- `LlmMemoryExtractor` 迁移到 `AiGateway`。
- 后续再迁移 `IntentRecognitionService`、校对/排版中的非流式 LLM 调用。

### 验收

- `WordAi.vbproj` Debug 编译 0 错误。
- `scripts/smoke-memory-pipeline.ps1` 通过。
- LLM 记忆提取失败时仍能规则回退。

## Phase 2：UI Dispatcher

### 范围

- 新增统一 `UiDispatcher`。
- 封装 `InvokeRequired / BeginInvoke / Invoke`。
- 所有 WebView2 和 Office Interop 入口逐步改为显式 UI 调度。

### 首批迁移

- `BaseChatControl` 中 WebView2 初始化、脚本执行、HTML 导航。
- `ConfigApiForm` 中 WebView2 初始化和 BeginInvoke helper。

### 验收

- 不再出现“CoreWebView2 can only be accessed from the UI thread”。
- 不再出现“创建窗口句柄之前不能 BeginInvoke”。

## Phase 3：Memory/Skills 收敛

### 范围

- 新写入统一走 `conversation_event -> memory_job -> memory_item -> memory_embedding`。
- `atomic_memory` 进入兼容读取模式。
- Skills 使用统计从 `skills_usage.json` 收敛到 `skills_registry`。

### 验收

- 真实 Word UI 对话能写入结构化记忆并召回。
- Skills 第一次只加载元数据，命中后二次加载详情。

## Phase 4：数据库迁移文件化

### 范围

- 将内联 SQL 拆到 `ShareRibbon/Storage/Migrations/*.sql`。
- 增加迁移执行记录和 smoke。
- 保留 `RunSchemaHealthChecks` 作为老用户兜底。

### 验收

- 空库初始化通过。
- 旧版本库升级通过。
- 重复启动不重复执行迁移。

## Phase 5：主聊天控制器拆分

### 范围

- 从 `BaseChatControl` 拆出：
  - `ChatRequestOrchestrator`
  - `WebViewBridge`
  - `ChatCommandRouter`
  - `ToolCallController`
  - `MemoryTurnRecorder`

### 验收

- `BaseChatControl` 只保留 UI 生命周期和事件绑定。
- 新增业务服务可以单独 smoke。

## 当前执行状态

- 已完成 Memory sidecar schema、pipeline、structured RAG、Skills 渐进式加载。
- 已完成 `scripts/smoke-memory-pipeline.ps1`。
- 下一步执行 Phase 1：新增 `AiGateway` 并迁移 `LlmMemoryExtractor`。
