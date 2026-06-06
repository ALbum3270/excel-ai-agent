# Conversation Runtime Migration Plan

目标：在不重写 VSTO/WebView2 外壳的前提下，把对话、历史、RAG、Skills、MCP 和后续开源 Agent 框架接入点从 `BaseChatControl` 中拆出来，形成可替换的运行时边界。

## Checklist

- [x] 1. 生成迁移 Plan 文档，并在执行过程中逐项标注完成。
- [x] 2. 新增运行时契约与上下文组装器。
  - 定义 `IContextComposer`、`IToolBroker`、`IConversationRuntime` 等接口。
  - 定义 `ChatRequestContext`、`ChatContextCompositionResult`、`ChatRequestBuildResult` 等数据模型。
  - 将 `CreateRequestBody` 中的 messages/RAG/Skills/历史窗口组装迁移到 `DefaultContextComposer`。
- [x] 3. 新增 MCP 工具代理。
  - 将 MCP tools 构造逻辑迁移到 `McpToolBroker`。
  - `BaseChatControl` 只负责调用工具代理并组装最终请求 JSON。
- [x] 4. 新增前端统一 bridge。
  - 增加 `office-ai-bridge.js`，统一封装 `window.chrome.webview.postMessage` 和 `window.vsto` fallback。
  - 让 `message-sender.js`、`history-manager.js` 走统一 bridge。
  - 注册 HTML、`.vbproj`、`.resx`、`ResourceExtractor`，并 bump resource version。
- [x] 5. 构建验证。
  - 使用 VS2022 MSBuild 构建 `AiHelper.sln`。
  - 若发现 VB.NET API/项目注册问题，修复后重新构建。

## Later Phases

- [ ] 接入 Semantic Kernel 适配器：作为 `IConversationRuntime` 的一个实现，不直接污染 `BaseChatControl`。
- [ ] 评估官方 MCP C# SDK 替换现有 `StreamJsonRpcMCPClient` 的成本。
- [ ] 评估 .NET 8 sidecar 承载 Microsoft Agent Framework，用 IPC 与当前 .NET Framework VSTO 插件通信。
- [ ] 若需要 LangGraph，仅以 sidecar 形式接入，避免 Python 运行时进入 Office 插件进程。
