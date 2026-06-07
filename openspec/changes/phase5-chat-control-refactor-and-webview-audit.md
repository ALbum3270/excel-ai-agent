# Phase 5 Chat Control Refactor And WebView Audit

## 背景

本轮改动继续执行 `global-architecture-hardening-plan.md` 的 Phase 5：拆分 `BaseChatControl`，降低主聊天控件同时承担 WebView2、消息路由、发送校验、系统提示词、记忆落库等职责带来的维护风险。

`ShareRibbon` 被 Word/Excel/PowerPoint 三端共用，因此本轮所有验证都按三端共享影响处理。

## 已落地改动

### 1. WebView Command Router

新增：

- `ShareRibbon/Controls/Services/ChatCommandRouter.vb`

改动：

- `BaseChatControl.WebView2_WebMessageReceived` 不再维护大段 `Select Case`。
- WebView 消息通过 `ChatCommandRouter` 注册表分发。
- 新增旧协议兼容注册：
  - `startLoop` -> `HandleLegacyStartLoop`
  - `continueLoop` -> `HandleStartAgentExecution`
  - `replanLoop` -> `HandleAgentRefinePlan`
  - `cancelLoop` -> `HandleAbortAgent`
  - `getCurrentAppInfo` -> `HandleGetCurrentAppInfo`

原因：

- html/js 中仍存在旧 Ralph Loop 和应用信息查询入口。
- 后端已迁移到 `AgentKernelService` / `startAgent` 体系，兼容层可以避免旧前端入口掉入 unknown message。

### 2. WebView Bridge

新增：

- `ShareRibbon/Controls/Services/WebViewBridge.vb`

迁移：

- `ExecuteJavaScriptAsyncJS`
- `ExecuteJavaScriptAndWaitAsync`
- `WaitForRendererMapAsync`
- `LoadLocalHtmlFile`
- `GetFullHtmlContentAsync`

结果：

- `BaseChatControl` 保留原有方法签名，避免影响子类和既有调用。
- WebView2 脚本执行、HTML 加载、完整 HTML 导出清理逻辑集中到 `WebViewBridge`。

### 3. System Prompt Resolver

新增：

- `ShareRibbon/Controls/Services/ChatSystemPromptResolver.vb`

迁移：

- `Send` 中系统提示词 fallback 选择逻辑。

保留顺序：

- 调用方传入的 `systemPrompt`
- `PromptManager.Instance.GetCombinedPrompt(context)`
- `ConfigSettings.propmtContent`
- 最终 Office AI assistant fallback

### 4. Send Validator

新增：

- `ShareRibbon/Controls/Services/ChatSendValidator.vb`

迁移：

- `Send` 开头的 ApiKey、ApiUrl、问题为空校验。

保持行为：

- 原中文提示文案不变。
- `changeSendButton()` 恢复按钮状态不变。
- 校验失败仍在创建 requestUuid 前返回。

### 5. Memory Turn Recorder

新增：

- `ShareRibbon/Controls/Services/MemoryTurnRecorder.vb`

迁移：

- 原 `BaseChatControl.PersistAgentMemoryTurnAsync`。

职责：

- 写入 user / assistant `conversation_event`。
- 创建 `extract_memory` job。
- 触发 `AgentMemoryPipelineService.KickoffPendingJobs()`。

保持行为：

- 仍为 fire-and-forget 后台执行。
- 写入前捕获 user/assistant/session/responseMode/appType，避免异步执行时读取 UI 状态。

## 项目文件注册

`ShareRibbon/ShareRibbon.vbproj` 已注册：

- `Controls\Services\ChatCommandRouter.vb`
- `Controls\Services\WebViewBridge.vb`
- `Controls\Services\ChatSystemPromptResolver.vb`
- `Controls\Services\ChatSendValidator.vb`
- `Controls\Services\MemoryTurnRecorder.vb`

未新增 html/js/css 文件，因此不需要新增 `ResourceExtractor` 映射，也不需要修改 Virtual Server 静态资源投放列表。

## Html/Js 联动检查

扫描范围：

- `ShareRibbon/Resources/chat-template-refactored.html`
- `ShareRibbon/Resources/js/*.js`
- `ShareRibbon/Resources/css/*.css`
- `ShareRibbon/Resources\ResourceExtractor.vb`
- `ShareRibbon/ShareRibbon.vbproj`

重点检查：

- `chrome.webview.postMessage`
- `window.vsto.postMessage`
- `sendMessageToServer`
- `sendMessageToVB`
- `changeSendButton`
- `showContextHints`
- `type: '...'` WebView 消息协议

结论：

- 本轮新增的后端服务没有新增前端 API，也没有新增静态资源文件。
- `changeSendButton()` 和 `showContextHints()` 仍由既有 `message-sender.js` 提供，不需要修改。
- 发现旧前端协议 `startLoop/continueLoop/replanLoop/cancelLoop/getCurrentAppInfo` 在路由化后需要后端兼容注册，已补齐。
- 补齐后，按 `postMessage/sendMessageToServer/sendMessageToVB` 上下文筛选的前端消息类型已经全部能被 `ChatCommandRouter` 接住。

## 验证结果

编译：

- `WordAi/WordAi.vbproj` Debug / AnyCPU：通过，0 warning / 0 error
- `ExcelAi/ExcelAi.vbproj` Debug / AnyCPU：通过，0 warning / 0 error
- `PowerPointAi/PowerPointAi.vbproj` Debug / AnyCPU：通过，0 warning / 0 error

Smoke：

- `scripts/smoke-memory-pipeline.ps1`：通过
- `scripts/smoke-skills-registry.ps1`：通过
- `scripts/smoke-db-schema-drift.ps1`：通过，`Drift=False`
- `git diff --check`：无空白错误，仅有既有 LF/CRLF 提示

协议检查：

- WebView 前端发送类型与 `ChatCommandRouter` 注册项对照：未发现仍需补齐的前端发送协议。

## 仍需真实 UI 验收

编译和 smoke 无法覆盖 WebView2/Office Interop 的全部运行时行为，仍建议三端手工跑：

- 打开聊天面板。
- 发送普通聊天消息。
- 空 ApiKey / 空 ApiUrl / 空问题时确认按钮恢复。
- 打开 API 配置面板并保存。
- 触发 `/loop xxx`，确认进入 Agent 规划路径而不是 unknown message。
- 触发排版、校对、续写、模板相关入口。
- 保存/加载聊天 HTML 快照。

## 后续建议

- 下一步可继续拆 `CreateRequestBody`，把 request 构造和历史写入从 `BaseChatControl` 移到独立 request builder。
- `MemoryTurnRecorder` 后续可移动到 `Services\Memory`，当前放在 `Controls\Services` 是为了先贴近 stream finalization 调用点，降低本轮迁移风险。
