# 当前项目评估与逐步优化计划

## 背景

本项目是基于 `Visual Basic.NET + VSTO + .NET Framework 4.7.2 + WebView2` 的三端 Office AI 插件，覆盖 `WordAi`、`ExcelAi`、`PowerPointAi`，共享能力集中在 `ShareRibbon`。

当前项目已经具备 AI Native Office Agent 的雏形：上下文读取、Memory、Skills、MCP、Agent Loop、数据库迁移、WebView2 UI 和三端 Office 执行能力均已接入。但代码仍处于从“聊天式插件”向“可执行 Office Agent”迁移的中间态，主要风险集中在构建发布、超大控制器、新旧运行时并存、线程边界和验证体系。

## 当前验证基线

本轮评估时执行过以下验证：

- `devenv.com .\AiHelper.sln /Rebuild Debug`
  - `ShareRibbon`、`ExcelAi`、`WordAi`、`PowerPointAi` 构建成功。
  - `OfficeAgent.vdproj` 失败，原因是安装项目在 Debug 解决方案构建中仍寻找多个 `bin\Release` 产物和安装包目录 DLL。
- `scripts/smoke-memory-pipeline.ps1` 通过。
- `scripts/smoke-skills-registry.ps1` 通过。
- `scripts/smoke-db-schema-drift.ps1` 通过，`Drift=False`。
- `node --check` 检查关键前端文件通过：
  - `ShareRibbon/Resources/js/message-sender.js`
  - `ShareRibbon/Resources/js/agent-card.js`
  - `ShareRibbon/Resources/js/office-ai-bridge.js`
  - `ShareRibbon/Resources/js/proofread-ui.js`

## 问题清单

### P0：构建与安装包链路不清晰

现象：

- 四个代码项目可构建，但完整解决方案 Debug Rebuild 会在 `OfficeAgent.vdproj` 预验证阶段失败。
- `OfficeAgent.vdproj` 中大量 `SourcePath` 指向 `..\WordAi\bin\Release`、`..\ExcelAi\bin\Release`、`..\PowerPointAi\bin\Release`。
- 根目录旧 `.bat` 构建脚本硬编码旧机器路径 `F:\ai\code\AiHelper` 和 VS 2022 MSBuild 路径，在当前仓库不可移植。

影响：

- 开发者容易把安装项目失败误判为代码编译失败。
- Release 打包缺少明确的前置检查。
- 构建脚本不可复用，不利于 CI 或多人协作。

优化方向：

1. 区分“代码项目构建”和“安装包构建”。
2. 新增仓库相对路径的代码构建脚本。
3. Release 打包前先构建三端 Release 产物，再验证安装项目引用文件。
4. 短期保持 `.vdproj` 最小改动，中长期评估 WiX/MSIX。

### P0：超大类承担过多职责

高风险文件：

- `ShareRibbon/Controls/BaseChatControl.vb`
- `WordAi/ChatControl.vb`
- `PowerPointAi/ChatControl.vb`
- `ShareRibbon/Config/ConfigApiForm.vb`

现象：

- UI 生命周期、WebView2、消息路由、上下文构建、Agent 启动、HTTP streaming、Memory 落库、Office 执行回调仍集中在少数类中。
- 已经拆出 `ChatCommandRouter`、`WebViewBridge`、`ChatSendValidator`、`MemoryTurnRecorder`，但主控制器仍偏重。

影响：

- 局部修改影响面难判断。
- Office COM、WebView2、异步任务状态容易互相干扰。
- 新能力容易继续堆在入口层，背离 AI Native 架构。

优化方向：

1. 继续拆 `BaseChatControl.CreateRequestBody`。
2. 拆出 `ChatRequestOrchestrator`，统一请求构造和发送前上下文准备。
3. 拆出 `ToolCallController`，统一 tools / MCP / ReAct tool result。
4. 三端 `ChatControl` 只保留宿主上下文采集和宿主执行器注册。

### P0：UI 线程和同步等待风险

现象：

- 仍存在 `Task.Run`、`.Result`、`.Wait()`、`Invoke/BeginInvoke` 混用。
- 部分路径涉及 WebView2 或 Office COM，属于 STA/UI 线程敏感区域。

影响：

- 偶发 UI 卡死。
- WebView2 跨线程异常。
- Office COM 对象失效或跨线程访问异常。

优化方向：

1. 所有 WebView2 调用走 `UiDispatcher`。
2. 所有 Office COM 访问明确在宿主 UI 线程执行。
3. 禁止 UI 线程调用 `.Result/.Wait()` 等同步等待。
4. 后台线程只处理 HTTP、SQLite、纯计算和文件 IO。

### P1：新旧 Agent/Intent/Command 路径并存

现象：

- 已有 `AiNativeRuntime`、`AgentKernel`、`LoopEngine`、`SelfCheckLoopController`。
- 同时仍存在 `IntentRecognitionService`、各端 `JsonCommandSchema`、`ExecuteJsonCommand`、直接 `Select Case` 执行路径。

影响：

- 同一句用户请求可能在不同入口走不同路径。
- Skills 的 `allowed-tools` 难以形成真正执行边界。
- 错误修复和执行解释不一定覆盖旧路径。

优化方向：

1. 以 Word 为样板，把 `WordActionHarness` 完整升级为 capability registry。
2. 将 `JsonCommandSchema` 逐步迁移到 tool/capability schema。
3. 让 `IntentRecognitionService` 只输出任务规格，不直接决定最终执行路径。
4. 所有复杂任务统一进入 `plan -> act -> observe -> repair -> explain`。

### P1：LLM 请求路径重复

现象：

- `AiGateway` 已经存在，但 `HttpStreamService`、`LLMUtil`、`IntentRecognitionService`、翻译、补全、Memory 等仍有各自请求构建逻辑。

影响：

- Provider 适配不一致。
- `ReasoningMode`、Anthropic/OpenAI 转换、tools 参数和错误处理可能分叉。
- 难以统一限流、超时、日志和脱敏。

优化方向：

1. `AiGateway` 作为非流式标准入口。
2. 增加 streaming gateway 或 adapter，逐步迁移 `HttpStreamService` 中的 provider 适配。
3. 所有请求日志统一脱敏，避免输出 API Key 片段。
4. 翻译、补全、Intent、Memory 分批迁移。

### P1：测试体系偏 smoke，缺少契约测试

现状：

- 已有数据库、Memory、Skills 的 smoke 脚本。
- 没有独立测试项目或不依赖 Office COM 的单元测试层。

影响：

- Agent 行为越复杂，手工 UI 验收成本越高。
- Prompt/JSON schema/工具选择的回归难以定位。

优化方向：

1. 先增加 PowerShell smoke 和静态契约检查。
2. 后续新增 `.NET Framework` 测试项目或轻量控制台测试入口。
3. 优先覆盖：
   - `AiGateway` 请求转换。
   - `SkillsDirectoryService` front matter 解析。
   - `JsonCommandSchema` 校验。
   - `FormattingIntentCompiler` / `ProofreadIntentCompiler`。
   - DB 迁移和 schema drift。

### P1：数据库迁移仍有内联 fallback 复杂度

现状：

- 外部迁移文件 `ShareRibbon/Storage/Migrations/*.sql` 已存在。
- `OfficeAiDatabase.vb` 仍保留大量内联 fallback SQL 和 health check。

影响：

- 数据库逻辑文件过大。
- schema 维护仍有双份来源。

优化方向：

1. 保留 health check 作为老用户兜底。
2. 将内联 SQL 降级为最小 fallback 或 documented constants。
3. 继续用 `smoke-db-schema-drift.ps1` 保护外部迁移与当前 schema 一致。

### P2：用户可见占位和边角残留

例子：

- `ExcelAi/ChatControl.vb` 仍有 `PowerQuery代码执行功能正在开发中`。
- `PowerPointAi/ChatControl.vb` 对未知命令提示 `暂不支持的PPT命令`。
- Word 工具结果回传仍有 TODO。

影响：

- 产品承诺和真实体验存在细小落差。
- Agent 失败时可能回退到传统提示，而不是自动换策略。

优化方向：

1. 用户可见能力入口不再提示“开发中”。
2. 未支持命令应进入 Agent repair / fallback，而不是只显示 warning。
3. TODO 归档到明确 roadmap 或实现为结构化 ToolResult。

### P2：数据路径与安全策略需要产品化

现状：

- 调试数据库位于 `Documents\OfficeAiAppData-Debug\office_ai.db`。
- API Key 已有安全管理，但部分日志和 DeepSeek token 注入路径需要审计。

优化方向：

1. 评估正式版数据是否迁到 `LocalApplicationData`。
2. 增加隐私导出、清除、禁用 Memory 的明确流程。
3. 日志统一脱敏，不输出 Key 前缀、token、Authorization header。

## 优化路线

### Step 1：构建链路整理

目标：

- 让开发者用一个稳定入口构建四个代码项目。
- 不再依赖旧机器绝对路径。
- 明确安装项目需要 Release 产物。

交付：

- 新增 `scripts/build-code-projects.ps1`。
- 更新根目录 `.bat` 包装器。
- 增强 Release checks 的 MSBuild 自动发现。

验收：

- `powershell -File .\scripts\build-code-projects.ps1 -Configuration Debug` 通过。
- 根目录 `build-all.bat` 可在当前仓库执行。

### Step 2：安装包 Release 输入预检

目标：

- 在构建 MSI 前发现缺失 DLL。
- 避免 `.vdproj` 在 Visual Studio 里才暴露大量缺失文件。

交付：

- 新增安装项目 `SourcePath` 文件存在性审计脚本。
- 将该脚本接入 `build/RunReleaseChecks.ps1`。

验收：

- 未构建 Release 时能明确报告缺失输入。
- 构建 Release 后预检通过或给出可操作缺失清单。

当前落地：

- 新增 `build/AuditInstallerInputs.ps1`。
- `build/RunReleaseChecks.ps1` 在 Release 代码项目构建后运行安装项目输入审计。
- 默认只检查带目录的 `SourcePath`，避免把框架/GAC 依赖误判为本地文件；如需更严格检查可传入 `-IncludePlainFiles`。

### Step 3：主聊天请求构建拆分

目标：

- 从 `BaseChatControl` 拆出请求构造与上下文组装。

交付：

- `ChatRequestOrchestrator` 或等价服务。
- `BaseChatControl` 保留兼容方法签名。

验收：

- 三端 Debug 构建通过。
- 普通聊天、Agent、带文件消息行为不变。

当前落地：

- 新增 `ShareRibbon/Controls/Services/ChatRequestOrchestrator.vb`。
- `BaseChatControl.CreateRequestBody` 保持原方法签名，只委托给 `ChatRequestOrchestrator`。
- 请求上下文组装、`IConversationRuntime.BuildRequest` 调用、发送前 user/system 历史写入从 `BaseChatControl` 移出。
- 已在 `ShareRibbon/ShareRibbon.vbproj` 注册新增 `.vb` 文件。

### Step 4：LLM Gateway 收敛

目标：

- Intent、Memory、翻译、补全逐步统一到 `AiGateway`。

交付：

- 迁移一个低风险非流式调用方。
- 增加 provider 转换 smoke。

验收：

- 迁移路径失败时仍有旧行为兜底。

当前落地：

- `Services/Memory/LlmMemoryExtractor.vb` 已使用 `AiGateway`。
- `Controls/Services/AutocompleteService.vb` 的 Chat Completion 补全路径已迁移到 `AiGateway.SendChatAsync`。
- `Controls/Services/IntentRecognitionService.vb` 的两个非流式调用点已迁移到 `AiGateway.SendChatAsync`，并兼容新旧响应格式。
- `AiGateway.BuildProviderRequest` 提供无网络 provider payload 测试入口。
- `scripts/smoke-ai-gateway-provider.ps1` 已覆盖 OpenAI-compatible、Anthropic 请求转换及响应提取，并接入 Release checks。
- 自动补全 FIM 路径仍保留原专用接口调用，主聊天 streaming 路径仍保留 `HttpStreamService`。

### Step 5：线程边界治理

目标：

- 消除高风险 `.Result/.Wait()` 和跨线程 UI/COM 访问。

交付：

- 优先治理 `WordAi/WebDataCapturePane.vb`、`BaseDeepseekChat.vb`、`WordAi/ChatControl.vb` 中的同步等待。

验收：

- 编译通过。
- Word UI 手工打开聊天、网页采集、校对、排版无跨线程异常。

当前落地：

- `WordAi/WebDataCapturePane.vb` 的图片、视频预览图、容器图片下载路径已移除 `.Result`。
- 新增局部 `DownloadBytesAsync` helper，HTTP 下载使用 `Await client.SendAsync(...).ConfigureAwait(False)`。
- Word 插入图片逻辑仍通过 `Me.Invoke` 回到 UI 线程执行，未改变 COM 操作线程边界。
- `BaseDeepseekChat.vb` 的 Cookie 获取和 token 注入已改用 `UiDispatcher.InvokeAsync` + `Await`。
- `BaseChatControl.vb` 的 PostFlush 校验已迁移到 `RunPostFlushValidationAsync`。
- `WordAi/ChatControl.vb` 的校对分析已改为在 UI 线程异步等待 `AnalyzeAsync`，完成后再继续后置排版。

### Step 6：Word Agent 闭环样板

目标：

- Word 排版/校对/编号形成完整 capability 执行闭环。

交付：

- `WordActionHarness` 下沉更多执行细节到 capability。
- Observe/Repair/Explain 结果统一返回。

验收：

- `字体统一加大2号`、`校对全文`、`把前面的序号改为12345` 三个场景都有执行解释和可撤销说明。

## 当前执行记录

### 2026-07-11 Step 1 启动

计划：

- 新增仓库相对构建脚本。
- 更新旧 `.bat` 包装器。
- 增强 Release checks 的 MSBuild 查找逻辑。

结果：

- 已新增 `scripts/build-code-projects.ps1`。
- 已更新 `build-all.bat`、`build-excel-ppt.bat`、`build-shareribbon.bat`，移除旧机器绝对路径。
- 已增强 `build/RunReleaseChecks.ps1` 的 MSBuild 查找逻辑，当前可识别 VS 18 Community。
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-code-projects.ps1 -Configuration Debug` 通过。
- `cmd /c build-shareribbon.bat` 通过。

### 2026-07-11 Step 2 完成

计划：

- 增加 `.vdproj` 输入文件存在性审计。
- 将审计接入 Release checks。

结果：

- 已新增 `build/AuditInstallerInputs.ps1`。
- 未构建 Release 产物时，脚本可明确报告缺失的 `..\WordAi\bin\Release`、`..\ExcelAi\bin\Release`、`..\PowerPointAi\bin\Release` 输入文件。
- 构建 Release 代码项目后，`build/AuditInstallerInputs.ps1` 通过，检查 52 个带目录的安装项目文件引用。
- `build/AuditInstallerRelease.ps1` 通过。
- `build/RunReleaseChecks.ps1 -SkipBuild` 通过：
  - Version audit 通过。
  - Installer Release audit 通过。
  - Installer input audit 通过。
  - `scripts/smoke-memory-pipeline.ps1` 通过。
  - `scripts/smoke-skills-registry.ps1` 通过。
  - `scripts/smoke-db-schema-drift.ps1` 通过，`Drift=False`。
  - `git diff --check` 通过，仅有 LF/CRLF 提示。

### 2026-07-11 Step 3 部分完成

计划：

- 在不改变 streaming 行为和三端调用签名的前提下，先拆出请求构造编排层。
- 保留 `BaseChatControl.CreateRequestBody` 作为兼容入口，降低三端 `ChatControl` 的改动面。

结果：

- 已新增 `ShareRibbon/Controls/Services/ChatRequestOrchestrator.vb`。
- 已将 `ChatRequestContext` 组装、`ConversationRuntime.BuildRequest` 调用、发送前历史写入迁出 `BaseChatControl`。
- 已补充 `.gitignore` 例外，确保本评估文档和 `build/AuditInstallerInputs.ps1` 不被忽略。
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-code-projects.ps1 -Configuration Debug -Projects "ShareRibbon\ShareRibbon.vbproj"` 通过。

后续：

- 继续拆分 `ToolCallController` 或等价服务，收敛 tools / MCP / ReAct tool result。
- 在三端手工验收普通聊天、Agent、带文件消息后，再将 Step 3 标记为完全完成。

### 2026-07-11 Step 4 完成第一阶段

计划：

- 先迁移低风险、非流式、内部辅助调用点。
- 暂不触碰主聊天 streaming、MCP tool calling 和 FIM 专用补全接口。

结果：

- 已将 `ShareRibbon/Controls/Services/AutocompleteService.vb` 的 Chat Completion 补全路径迁移到 `AiGateway.SendChatAsync`。
- 已将补全 JSON 解析从 HTTP 请求流程中拆成 `AddCompletionsFromMessage`。
- 已将 `IntentRecognitionService` 的两个非流式调用点迁移到 `AiGateway.SendChatAsync`。
- 新增 `AiGateway.BuildProviderRequest` 与 `scripts/smoke-ai-gateway-provider.ps1`，provider smoke 已接入 Release checks。
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-code-projects.ps1 -Configuration Debug -Projects "ShareRibbon\ShareRibbon.vbproj"` 通过。

后续：

- 设计 streaming gateway/adapter，再迁移主聊天 streaming provider 适配。
- FIM 专用补全保持独立，除非 `AiGateway` 后续提供专用 FIM 契约。

### 2026-07-11 Step 5 部分完成

计划：

- 先治理已知最高风险的同步 HTTP 等待点。
- 不改动 Word COM 插入图片的 UI 线程执行方式。

结果：

- 已将 `WordAi/WebDataCapturePane.vb` 中三处 `client.SendAsync(...).Result` 和 `ReadAsByteArrayAsync().Result` 改为异步下载。
- 已增加 `DownloadBytesAsync`，返回类型使用 `System.Threading.Tasks.Task(Of Byte())`，避免与 Word Interop `Task` 类型冲突。
- `rg "\.Result\b|\.Wait\(|GetAwaiter\(\)\.GetResult\(" WordAi\WebDataCapturePane.vb` 无匹配。
- `BaseDeepseekChat.vb` 的 Cookie/token WebView2 路径已移除同步等待。
- `BaseChatControl.vb` 的 PostFlush 校验已异步化。
- `WordAi/ChatControl.vb` 的 `proofreadMode.AnalyzeAsync(...).Wait()` 已改为 UI 线程异步等待。
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-code-projects.ps1 -Configuration Debug -Projects "WordAi\WordAi.vbproj"` 通过。

后续：

- 评估 `BaseChatControl.RunUiActionSync` / `WebViewService` 的同步 UI helper；这些调用面依赖同步顺序，需单独设计异步 API，不能机械替换。
- 将 `Task.Run(Async Function() ...)` 按是否触碰 UI/COM 分组处理，纯 IO 保留后台执行，UI/COM 统一走 `UiDispatcher` 或宿主 UI 线程。

### 2026-07-11 综合验证

- 四个代码项目 Debug 构建通过。
- build/RunReleaseChecks.ps1 -SkipBuild 通过，包含安装包 52 项输入审计、五组 smoke 和 git diff --check。
- smoke-ai-gateway-provider.ps1 验证 OpenAI-compatible / Anthropic payload 与 Anthropic 响应提取通过。
- 数据库 schema drift 检查结果为 Drift=False。

### 2026-07-11 Step 6 第一阶段完成

计划：

- 先建立 Word capability 的注册与契约，不改动现有 Word COM 执行器主体。
- 保留 `WordActionHarness` 的安全/成本兜底规则，但让计划结果携带 capability descriptor、输入 schema、风险级别、Observe/Repair/Explain 契约。

结果：

- 已新增 `WordAi/Services/WordCapabilityRegistry.vb`，登记四类 Word 能力：`word.proofread`、`word.direct-formatting`、`word.numbering`、`word.semantic-reformat`。
- `WordActionPlan` 已增加 `Capability`、`CapabilitySummary`、`ProofreadPlan`，为后续统一执行结果和 UI 展示提供结构化元数据。
- `WordActionHarness.Plan` 已接入 `ProofreadIntentCompiler`，明确的 `校对全文`、`检查选区错别字` 等请求无需完全依赖外部意图识别服务即可进入校对能力。
- 已新增 `scripts/smoke-word-capability-registry.ps1`，静态检查 registry、`.vbproj` 注册、harness 接线和校对计划入口。
- `build/RunReleaseChecks.ps1` 已接入 Word capability registry smoke。
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-code-projects.ps1 -Configuration Debug` 通过；普通沙箱下曾在 VSTO manifest 证书访问处失败，提升权限后构建通过。
- `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\smoke-word-capability-registry.ps1` 通过。

后续：

- 将 `WordNumberingResult`、`WordFormattingAgentResult`、校对侧边栏结果逐步收敛到统一 `Plan/Act/Observe/Repair/Explain` 返回模型。
- 给 `字体统一加大2号`、`校对全文`、`把前面的序号改为12345` 三个场景补 UI 层统一说明：能力名、处理范围、观察结果、撤销方式。
- 继续把结构化排版的 repair loop 从摘要文本推进为可检查的失败项列表，避免只靠自然语言描述失败原因。
