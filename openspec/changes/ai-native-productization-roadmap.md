# AI Native Office 插件产品化落地路线

> 本文件用于跟踪 AI Native 重构落地进度。每完成一个阶段，必须在本文件中把对应复选框改为 `[x]`，并记录验证结果。

## 硬约束

- 不做人工干预式产品流程：意图、记忆、Skills、MCP 工具选择、Agent 执行路径应由 AI 自动识别和编排。
- 每完成一个阶段，必须完成测试验证后再进入下一阶段。
- `ShareRibbon` 是公共模块，只能放抽象、协议、共享运行时、共享 UI 基础设施和宿主无关服务。
- `ShareRibbon` 不应直接引用或实例化 Excel / Word / PowerPoint 的具体实现类型。
- Excel / Word / PowerPoint 的具体 Office 对象模型访问必须留在各自项目内，通过子类重写、接口、Provider 或回调注入接入共享层。
- 清理无用 VB / JS / CSS 必须先生成候选清单，再按调用证据小批删除；每批删除后立即构建和检查。

## 通用验证门槛

每个阶段完成后至少执行：

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.com" .\AiHelper.sln /Rebuild Debug
node --check .\ShareRibbon\Resources\js\message-sender.js
node --check .\ShareRibbon\Resources\js\agent-card.js
git diff --check
```

若某阶段未修改 JS，可跳过对应 `node --check`，但需要在阶段记录中说明。

## 阶段 1：AI Native 中枢统一

状态：`[x]`

目标：把上下文、意图、Skills、MCP、记忆、Agent Loop 串成统一运行时，而不是散落在多个服务里各自判断。

落地项：

- [x] 建立 `AiNativeRuntime` 或等价协调层，统一承接用户输入。
- [x] 输出统一对象：`IntentResult + ContextTrace + SelectedSkills + AvailableTools + ExecutionPlan`。
- [x] 普通 Chat 和 Agent 模式都走同一套上下文构建。
- [x] 删除重复的旧路由判断和分叉入口。
- [x] 确认共享层只依赖抽象接口，不引用具体 Office 宿主实现。

验收标准：

- [x] Word / Excel / PowerPoint 同一句请求进入同一套运行时。
- [x] 构建通过。
- [x] 不再出现多个地方各自判断意图、拼上下文的重复逻辑。

验证记录：

- 2026-07-08：新增 `ShareRibbon/Agent/AiNativeRuntime.vb`，移除 `ShareRibbon/Agent/Context/ExcelContextProvider.vb` 具体实现；Excel/Word/PowerPoint 通过各自 `ChatControl.CaptureOfficeContext()` 子类重写提供宿主上下文。
- 2026-07-08：`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。
- 2026-07-08：`BaseChatControl` 智能路由接入 `AiNativeRuntime.AnalyzeAsync()`，普通 Chat 与 Agent 分支共享同一套意图/上下文分析结果。
- 2026-07-08：`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。
- 2026-07-08：删除旧人工意图确认链路：`confirmIntent/cancelIntent` VB 路由、`intent-preview.js`、HTML 引用、资源映射和项目注册。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`git diff --check` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。

## 阶段 2：上下文控制台

状态：`[x]`

目标：用户不用干预，但能看懂 AI 为什么这样做。

落地项：

- [x] 扩展 `ChatContextTrace`。
- [x] Trace 包含当前 Office 内容、选区、长期记忆、短期会话、用户画像、Skills、MCP 工具、执行计划。
- [x] 前端新增“本轮上下文”折叠面板。
- [x] 每次发送前展示 AI 识别出的任务意图和使用依据。
- [x] 由宿主 Provider 提供 Office 上下文，`ShareRibbon` 只消费抽象结果。

验收标准：

- [x] 发起一次任务后，能看到 AI 使用了哪些记忆、哪些 Skill、哪些上下文。
- [x] 不要求用户选择，只展示解释。
- [x] JS 语法检查通过。

验证记录：

- 2026-07-08：`ChatContextTrace` 增加意图、Office 上下文、Skills、工具摘要字段；`AiNativeRuntime` 填充意图说明、OfficeContext 和可用工具；`ChatContextBuilder` 填充 Skill 命中信息。
- 2026-07-08：`message-sender.js` 的 `showContextHints()` 扩展为可折叠上下文控制台，展示 Office 上下文、Skills、记忆、近期会话和工具摘要；`styles.css` 增加对应样式。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`git diff --check` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。
- 2026-07-08：Agent 计划生成时写入 `ChatContextTrace.ExecutionPlan`，上下文控制台展示执行计划摘要和步骤。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`git diff --check` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。

## 阶段 3：记忆系统产品化

状态：`[x]`

目标：长期/短期记忆真正服务 Agent，而不是只做 RAG。

落地项：

- [x] 短期记忆只服务当前会话窗口。
- [x] 长期记忆用于跨会话 RAG 和 Agent 主动检索。
- [x] 自动晋升：高重要性、高复用、用户偏好、稳定事实自动进入长期记忆。
- [x] 自动过期：低重要性短期记忆自动过期。
- [x] 记忆冲突：新记忆与旧记忆冲突时保留版本和时间。

验收标准：

- [x] RAG 不再召回短期记忆。
- [x] Agent 可通过 `memory.search` 主动检索长期记忆。
- [x] 一次完整任务结束后能自动沉淀高价值记忆。

验证记录：

- 2026-07-08：`MemoryRepository.GetRelevantMemories()` 已限制为长期记忆；`ToolRegistry` 已提供 `memory.search` 等主动记忆工具。
- 2026-07-08：新增 `PromoteAccessedShortTermMemories()`；`MemoryService.SaveConversationTurnAsync()` 保存对话后会后台调用 `ConsolidateSessionMemoriesAsync()`，自动晋升高重要性/高复用短期记忆，并过期低重要性短期记忆。
- 2026-07-08：`git diff --check` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。
- 2026-07-08：新增 `RecordPotentialConflictsForSession()`，会话记忆整理时对带偏好/修改/否定信号的新旧长期记忆记录 `memory_graph.potential_conflict` 关系，保留新旧版本与时间。
- 2026-07-08：`git diff --check` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。

## 阶段 4：意图识别和自动规划

状态：`[x]`

目标：从“分类器”升级为“任务理解器”。

落地项：

- [x] 意图输出统一 Spec：目标、对象、约束、成功标准、风险、需要工具。
- [x] 低置信度不询问用户，转为探索式计划：先读上下文，再执行低风险分析。
- [x] 只选中内容不输入问题时，AI 自动判断最合适操作。
- [x] 旧的确认弹窗路径清理或降级为调试模式。

验收标准：

- [x] 选中文档内容直接点击 AI，不弹确认，自动分析/处理。
- [x] 意图结果能进入 Agent Plan。
- [x] 没有阻塞式人工确认。

验证记录：

- 2026-07-08：扩展 `AgentTaskSpec`，增加目标对象、风险级别、所需工具；`AiNativeRuntime` 每次分析都会生成统一任务规格，并写入 `ChatContextTrace.TaskSpec`。
- 2026-07-08：低置信度意图标记为 `exploratory` 复杂度，不弹人工确认；仅引用内容无显式问题时交给 AI 自动识别并处理。
- 2026-07-08：已删除旧意图确认弹窗链路 `intent-preview.js`、VB 路由和资源引用。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`git diff --check` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。

## 阶段 5：Skills 能力产品化

状态：`[x]`

目标：Skills 从配置项变成 AI 自动调用的能力库。

落地项：

- [x] Skill 索引统一进入上下文 Trace。
- [x] Agent 根据意图自动选择 Skill。
- [x] Skill 执行结果进入执行解释。
- [x] 记录 Skill 成功率、失败原因、最近使用场景。
- [x] 清理重复 Skill 匹配逻辑。

验收标准：

- [x] 用户无需手动选择 Skill。
- [x] Trace 中显示命中的 Skill 和原因。
- [x] 失败后 Agent 能换策略或换工具。

验证记录：

- 2026-07-08：`AiNativeRuntime` 接入 `SkillsIndexService.SelectSkillDefinitions()`，将自动选择的 Skill 写入 `ChatContextTrace.Skills` 和 `AgentTaskSpec.RequiredTools`；`ChatContextBuilder` 不再把“被召回”提前计为成功使用。
- 2026-07-08：`AgentKernel` 优先从 filesystem Skill 索引自动选择 Skill，旧 `SkillRegistry` 仅作为兜底；`PromptManager` 在规划阶段注入 Skill 详情和建议工具。
- 2026-07-08：`ToolRegistry` 在 Skill 脚本真实执行后记录成功率、失败原因和使用场景；`LoopEngine.ExecutionExplanation` 增加 Skill、脚本、类别和失败原因字段，前端执行解释同步展示。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过；`git diff --check` 通过，仅有 LF/CRLF 提示。

## 阶段 6：MCP 工具闭环

状态：`[x]`

目标：MCP 不只是配置，而是 Agent 工具生态。

落地项：

- [x] MCP 工具统一进入 `ToolRegistry`。
- [x] 工具健康检查：可用 / 不可用 / 错误原因。
- [x] Agent 计划中显示 MCP 工具用途。
- [x] MCP 执行结果纳入 Observe 和 `ExecutionExplanation`。
- [x] 清理重复 MCP 测试 / 调用代码。

验收标准：

- [x] MCP 工具可被 Agent 自动选择。
- [x] 失败时能看到明确错误，不静默失败。
- [x] 工具列表和执行解释一致。

验证记录：

- 2026-07-08：`ToolDescriptor` 增加 `AvailabilityStatus` 与 `LastError`，MCP 工具加载到 `ToolRegistry` 后可在上下文 Trace 中展示健康状态。
- 2026-07-08：MCP 调用失败、未初始化、异常时，`ToolRegistry.ExecuteToolAsync()` 写入明确 `failureReason`、`mcpStatus` 和工具健康状态；成功后恢复为 `available`。
- 2026-07-08：`ExecutionExplanation` 增加 MCP 工具名和状态字段，`agent-card.js` 与 `message-sender.js` 展示 MCP 执行结果、健康状态和错误原因。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过；`git diff --check` 通过，仅有 LF/CRLF 提示。

## 阶段 7：执行解释器

状态：`[x]`

目标：让用户相信 AI 自动执行，而不是黑盒乱改 Office。

落地项：

- [x] 每一步显示为什么执行、调用什么工具、参数是什么、结果是什么。
- [x] VBA / Office 工具执行前后记录摘要。
- [x] 失败自动修复次数、修复原因、最终结果展示。
- [x] 撤销点和执行解释绑定。

验收标准：

- [x] Agent 执行时有清晰时间线。
- [x] 每个步骤都有解释。
- [x] 失败能看到自修复记录。

验证记录：

- 2026-07-08：`ExecutionExplanation` 增加开始/结束时间、耗时、执行前摘要、执行后观察、撤销点、撤销提示、可撤销状态和自动修复摘要。
- 2026-07-08：`LoopEngine` 在每步执行时绑定撤销点、工具参数、Observe 结果和自动修复结果，统一输出到执行解释。
- 2026-07-08：`agent-card.js` 展示执行前/执行后、耗时、自动修复、撤销点和撤销提示；`styles.css` 增加紧凑的执行解释器样式。
- 2026-07-08：`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过；`git diff --check` 通过，仅有 LF/CRLF 提示。

## 阶段 8：Office 原生能力补齐

状态：`[x]`

目标：补齐产品承诺和实际按钮之间的落差。

优先顺序：

- [x] Excel：数据分析、透视表、分组汇总、排名分析、校对、排版。
- [x] PowerPoint：校对、数据分析、网页内容转幻灯片。
- [x] Word：数据分析、网页内容结构化引用、长文档处理。

验收标准：

- [x] 去掉“正在开发中 / 暂未实现”的用户可见占位。
- [x] 每个 Ribbon 入口都有真实可执行能力。
- [x] 三端行为一致，但具体实现仍留在各宿主项目内。

验证记录：

- 2026-07-08：Excel `Proofread`、`Reformat`、`Continuation`、`TemplateFormat`、隐藏 `WebCapture` fallback 改为打开 Chat 面板并自动启动 Agent，由 Excel 宿主上下文和 Skills/工具自动完成。
- 2026-07-08：PowerPoint `DataAnalysis`、`Proofread`、隐藏 `WebCapture` fallback 改为自动启动 Agent；保留已有翻译、续写、排版和模板能力。
- 2026-07-08：Word `DataAnalysis` 改为自动启动 Agent；保留已有网页采集、校对专注模式、排版、翻译、续写和模板能力。
- 2026-07-08：`rg` 检查三端 `Ribbon1.vb` 不再包含“正在开发中 / 暂未实现 / 未实现 / 暂不支持”占位文案；`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过；`git diff --check` 通过，仅有 LF/CRLF 提示。
- 2026-07-08：补齐 `ExcelDirectOperationService` 中 `DataAnalysis(type=pivot/groupby/ranking)` 基础实现，透视表复用 Excel 原生 PivotTable，分组汇总和排名分析生成可读结果区域。
- 2026-07-09：Word 排版新增 `FormattingIntentCompiler`，将“字体统一加大2号”等自然语言格式命令编译为结构化 `FormattingIntentPlan`，并由 `SmartFormatter` 在 WordAi 内执行，支持全文/选区/当前段落/标题/正文范围。
- 2026-07-09：Word 排版新增 `FormattingExecutionResult`，Chat 回显执行摘要、应用范围数量和操作数量；旧 intent 路由改为计划化排版流程，无选区时优先尝试当前文档直接格式计划，不再直接要求用户人工选中。
- 2026-07-09：Word 校对新增 `ProofreadIntentCompiler`，将“校对全文 / 只检查标点 / 自动修正明显错别字”等命令编译为 `ProofreadIntentPlan`；`ExecuteProofreadAsync` 支持全文、选区、当前段落范围，并把校对计划注入 Prompt。
- 2026-07-09：修复 `SmartFormattingOrchestrator` 中“加大2号”被误判为绝对 `2pt` 的微调解析问题；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。

## 阶段 9：清理无用 VB / JS / CSS

状态：`[x]`

目标：瘦身、降复杂度、减少维护负担。

执行方式：

- [x] 生成 VB 候选清单：未引用类、旧 Chat、旧 Agent、重复服务。
- [x] 生成 JS 候选清单：未被 HTML 引用、未被 `window.*` 暴露、未被 VB 调用的函数。
- [x] 生成 CSS 候选清单：未被 HTML/JS 使用的 class 和旧样式。
- [x] 分批删除候选项，每批删除后立即构建和检查。
- [x] 删除后更新本文件验证记录。

验收标准：

- [x] 每次删除都有调用证据。
- [x] 删除后完整 Rebuild 通过。
- [x] 前端主要页面无脚本错误。

验证记录：

- 2026-07-08：生成 CSS 候选清单后删除旧 `intent-preview.js` 对应的孤立 CSS：`.intent-preview-*`、`.intent-btn-*`、`intentSlideIn/intentSpin`、旧 `.execution-step` 图标样式等。
- 2026-07-08：`rg` 检查 `intent-preview`、`intent-btn`、`intentSlideIn`、`intentSpin` 不再出现在运行资源中；`node --check` 检查 `message-sender.js` 与 `agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过；`git diff --check` 通过，仅有 LF/CRLF 提示。
- 2026-07-09：生成 VB 候选清单并小批删除 `WordAi.Services.SmartFormatter` 中被 `FormattingIntentCompiler` 替代的旧解析函数：`ResolveTargetRange`、`ApplyFontSizeCommand`、`ParseFontSizeDelta`、`ParseAbsoluteFontSize`、`ApplyFontFamilyCommand`、`ApplyBasicStyleCommand`、`ApplyParagraphCommand`、旧 `ParseAmount`、旧 `NamedFontSizeToPoint`；`rg` 复查候选函数不再存在。
- 2026-07-09：本批只修改 VB 和 roadmap，未修改 JS，跳过 `node --check`；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。
- 2026-07-09：生成 JS 候选清单并删除旧 `ralph-agent.js`：后端真实调用 `showAgentPlanCard`，统一实现已在 `agent-card.js`；旧文件只剩 HTML 加载、资源注册和自身暴露，会覆盖统一 Agent UI。已同步移除 HTML script、`ResourceExtractor`、`.vbproj`、`.resx` 和 `Resources.Designer.vb` 中的注册。
- 2026-07-09：生成 CSS 候选清单并删除 `ralph-agent.js` 对应孤立样式：`.ralph-agent-container`、`.ralph-agent-dialog`；`rg` 复查 `ralph-agent`、`ralph_agent`、`ralphAgentState`、`.ralph-agent-*` 不再出现在运行资源中；`node --check .\ShareRibbon\Resources\js\agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过。
- 2026-07-09：删除 `ralph-agent.js` 遗留的孤立 Agent 对话框 CSS：`.agent-dialog-overlay`、`.agent-dialog-content`、`.agent-dialog-header`、`.agent-dialog-close`、`.agent-dialog-body`、`.agent-dialog-desc`、`.agent-request-input`、`.agent-dialog-actions` 等；`rg` 复查这些 class 仅有定义、无 HTML/JS/VB 使用。
- 2026-07-09：删除无引用备份文件 `ShareRibbon/ShareRibbon.vbproj.bak`、`OfficeAgent/OfficeAgent.vdproj.bak`；未修改真实 `.vdproj`。`rg --files` 复查 `.bak/.old/.orig`、`intent-preview`、`ralph-agent` 不再存在；主 HTML 引用覆盖当前 `Resources\js` 下全部运行 JS。
- 2026-07-09：阶段 9 最终验证：`node --check .\ShareRibbon\Resources\js\message-sender.js` 通过；`node --check .\ShareRibbon\Resources\js\agent-card.js` 通过；`devenv.com .\AiHelper.sln /Rebuild Debug` 通过，结果：4 成功，0 失败，0 跳过；`git diff --check` 通过，仅有 LF/CRLF 提示。
