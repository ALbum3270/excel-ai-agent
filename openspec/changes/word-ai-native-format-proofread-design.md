# Word AI Native 排版与校对设计方案

## 目标

把 Word 插件中的「排版」和「校对」从按钮式功能升级为 AI Native 工作流：

- 用户只需要用自然语言表达目标，不需要理解功能入口、模板、选区、参数。
- AI 自动理解当前文档、用户意图、约束和成功标准。
- 系统用 Agent Loop 执行：规划、预览、应用、观察、修复、解释。
- 每次修改都可解释、可撤销、可复查。
- 具体 Word COM 操作只留在 `WordAi`，共享层只提供抽象协议、计划模型、执行解释和 UI 基础设施。

## 用户视角原则

### 1. 不要求用户先懂功能

用户可以直接说：

- `全文按公文格式排版`
- `字体统一加大2号`
- `把标题层级整理一下`
- `校对全文，只改明显错别字和标点`
- `检查表达是否正式一点，但不要改变原意`
- `把这份合同排版成正式提交版本`

系统应该自动判断：

- 当前是否有选区。
- 应作用于选区、当前段落还是全文。
- 这是快速格式调整、语义排版、校对、润色，还是组合任务。
- 是否需要预览，是否可以直接低风险执行。

### 2. 用户看到的是结果，不是流程负担

界面上不要求用户选择「意图」「Skill」「工具」「范围」。这些由 AI 自动完成。

用户只需要看到：

- AI 准备做什么。
- 将影响哪些内容。
- 已经做了什么。
- 哪些地方不确定。
- 如何撤销或继续调整。

### 3. 默认保护文档

低风险操作可以直接执行，例如：

- 字号统一加大。
- 统一字体。
- 标题居中。
- 校对只展示建议。

高风险操作必须先预览，例如：

- 全文重排版。
- 修改标题层级。
- 自动应用大量校对建议。
- 涉及正文内容改写。

### 4. 失败时不甩锅给用户

不能只提示「失败」或「没有发现问题」。

系统必须区分：

- AI 确认没有问题。
- AI 返回格式异常。
- 找不到目标文本。
- Word COM 应用失败。
- 应用了但观察结果不符合预期。

## 当前问题

### 校对

当前校对链路已经具备完整雏形：

`用户命令 -> ProofreadIntentPlan -> 选取范围 -> Prompt -> AI JSON -> 解析 -> 侧边面板 -> 接受/忽略 -> Word 替换`

主要问题：

- 解析失败和真正无问题都可能显示「没有发现问题」。
- `proofread-side-panel` 在窄任务窗格里像遮盖层，用户不理解后面是什么。
- 结果 UI 是侧边面板，但没有清晰展示本轮校对计划、范围、模式。
- Word 标注只设置波浪下划线，没有保存和恢复原格式快照。
- 问题定位依赖文本 offset，复杂 Word 文档中不够稳。
- 接受全部没有逐项反馈成功/失败。

### 排版

当前排版有两条链路：

- 直接格式命令：`FormattingIntentCompiler -> FormattingIntentPlan -> SmartFormatter`
- 语义排版：模板/规范/映射驱动的 `ReformatJob -> SemanticMapping -> AI 标注 -> 渲染`

主要问题：

- 两条链路还没有统一成一个 Word Formatting Agent。
- 规则编译器覆盖了明确短命令，但复杂排版还没有统一 Plan JSON。
- 执行后缺少 Observe：没有系统验证字号、字体、行距、标题层级是否真的符合目标。
- 直接格式调整没有预览、差异、撤销点解释。
- 用户说「按公文格式排版」时，系统还不够像一个持续工作的 Agent。

## 目标体验

### 场景 A：直接格式调整

用户输入：

`字体统一加大2号`

系统行为：

1. 自动识别为格式调整。
2. 自动推断范围：`全文`。
3. 生成 `FormattingIntentPlan`。
4. 创建 Word UndoRecord。
5. 应用字号增量。
6. 观察抽样段落，确认字号变化。
7. Chat 回显：

```text
已完成：全文字号 +2pt。
影响范围：全文。
已检查：正文段落字号已更新。
可撤销：Ctrl+Z 或点击撤销。
```

### 场景 B：全文 AI 排版

用户输入：

`把这篇文章按正式公文风格排版`

系统行为：

1. 分析文档结构：标题、正文、表格、图片、编号。
2. 推断目标风格：正式公文。
3. 生成排版计划。
4. 展示预览卡片：
   - 标题格式。
   - 正文字体字号。
   - 行距。
   - 首行缩进。
   - 编号规则。
   - 表格处理策略。
5. 用户无需调整也可以直接应用。
6. 应用后观察文档，发现异常自动修复。
7. 输出执行解释。

### 场景 C：校对全文

用户输入：

`校对全文，只修明显错别字和标点`

系统行为：

1. 自动识别校对计划：
   - 范围：全文。
   - 类型：错别字、标点。
   - 模式：建议预览。
2. AI 返回问题列表。
3. 如果 JSON 解析失败，自动重试一次格式修复 Prompt。
4. 如果仍失败，显示「校对结果解析失败」，不显示「没有发现问题」。
5. 如果确实无问题，显示：

```text
没有发现明显错别字和标点问题。
检查范围：全文。
检查类型：错别字、标点。
```

### 场景 D：组合任务

用户输入：

`先校对错别字，再按正式报告格式排版`

系统行为：

1. 生成多步骤 Agent Plan：
   - Step 1：校对全文。
   - Step 2：应用高置信错别字修正。
   - Step 3：分析文档结构。
   - Step 4：生成排版计划。
   - Step 5：应用排版。
   - Step 6：观察验证。
2. 每步都有状态、结果和失败处理。

## 架构设计

## 核心对象

### WordTaskPlan

统一承载 Word AI Native 任务。

```text
WordTaskPlan
- Id
- OriginalUserRequest
- TaskKind: Format / Proofread / Rewrite / Mixed
- Scope
- RiskLevel
- RequiresPreview
- Steps: List(Of WordTaskStep)
- SuccessCriteria
- RollbackStrategy
```

### WordTaskStep

```text
WordTaskStep
- Id
- Kind: Analyze / Plan / Preview / Apply / Observe / Repair / Explain
- Description
- Input
- Output
- Status
- ToolName
- ErrorMessage
```

### FormattingPlan

保留当前 `FormattingIntentPlan`，继续扩展为可被 LLM 输出的 JSON。

```text
FormattingPlan
- Scope
- Operations
- Constraints
- PreviewRequired
- SuccessCriteria
```

### ProofreadPlan

保留当前 `ProofreadIntentPlan`，扩展校对结果状态。

```text
ProofreadPlan
- Scope
- IssueTypes
- ApplyMode
- OutputMode
- RetryPolicy
```

### ProofreadAnalysisResult

替代「空 List 表示无问题」。

```text
ProofreadAnalysisResult
- Status: NoIssues / HasIssues / ParseFailed / ModelFailed
- Issues
- ParseError
- RawResponsePreview
- Summary
```

## Agent Loop

Word 排版/校对统一走下面的循环。

```text
Think
  理解用户目标、当前文档、选区、历史偏好。

Plan
  生成结构化 WordTaskPlan。

Act
  调用 WordAi 内的具体执行器。

Observe
  读取 Word 文档现状，验证结果。

Repair
  如果观察结果不满足成功标准，自动小步修复。

Explain
  输出用户能理解的执行说明。
```

## 模块边界

### ShareRibbon

只放共享抽象：

- 计划模型基础类型。
- Agent Loop 状态。
- 执行解释 UI。
- JSON 解析工具。
- Prompt 构建基础设施。
- WebView2 UI 基础组件。

不能直接访问：

- `Microsoft.Office.Interop.Word`
- `Microsoft.Office.Interop.Excel`
- `Microsoft.Office.Interop.PowerPoint`
- `Globals.ThisAddIn`

### WordAi

放 Word 具体实现：

- Word 文档读取。
- Word Range 定位。
- Word 格式应用。
- Word 校对替换。
- Word 标注和撤销。
- Word 专属 plan compiler。

## 校对优化方案

### 第一阶段：修正「无问题」误判

目标：

- JSON 解析失败不再显示「没有发现问题」。

落地：

- 新增 `ProofreadAnalysisResult`。
- `ProofreadJsonParser.Parse()` 保留错误信息。
- `ProofreadPromptBuilder.ParseProofreadResponse()` 不再把失败吞成空列表。
- `SmartProofreadFocusMode.AnalyzeAsync()` 根据状态显示不同 UI。

验收：

- AI 返回非 JSON 时，显示解析失败和重试状态。
- AI 返回 `[]` 时，才显示没有发现问题。

### 第二阶段：校对结果 UI 改造

目标：

- 面板不再像遮盖层。
- 用户能看懂检查范围和检查类型。

落地：

- `proofread-side-panel` 改为可折叠 Drawer。
- 默认宽度跟随任务窗格宽度，避免固定 `380px` 挤压。
- 顶部展示：
  - 本轮校对计划。
  - 检查范围。
  - 检查类型。
  - 结果状态。
- 无问题状态提供「重新检查」「退出校对」按钮。

验收：

- 小宽度任务窗格下不遮挡输入框。
- 用户能一眼知道 AI 检查了什么。

### 第三阶段：Word 标注可靠性

目标：

- 标注可撤销、可恢复原格式。

落地：

- 每个 issue 保存原始格式快照。
- 退出校对时恢复原始下划线和颜色。
- 接受/忽略时只处理对应 range。
- 定位改为段落 Range 内查找，不再依赖全局 offset。

验收：

- 退出校对后原格式恢复。
- 表格/多段落文档中定位更稳定。

## 排版优化方案

### 第一阶段：统一排版计划

目标：

- 直接格式命令和语义排版共用 `WordTaskPlan`。

落地：

- `FormattingIntentPlan` 增加 JSON 序列化。
- LLM 可以输出同结构 JSON。
- 规则编译器作为兜底。
- Chat/Ribbon/Intent 路由都只调用一个 `WordFormattingAgent`。

验收：

- `字体统一加大2号` 和 `按公文格式排版` 都能生成结构化计划。

### 第二阶段：排版预览

目标：

- 高风险排版先预览。

落地：

- 新增 `FormattingPreviewResult`。
- 展示：
  - 影响范围。
  - 操作数量。
  - 样式变化。
  - 可能风险。
- 用户可以直接应用，也可以输入「标题再大一点」继续细化。

验收：

- 全文排版不会直接粗暴修改。
- 用户能理解将要发生什么。

### 第三阶段：Observe 验证

目标：

- 排版不是应用完就结束，而是自检。

落地：

- 应用后抽样读取 Word 文档：
  - 标题字号。
  - 正文字体。
  - 行距。
  - 首行缩进。
  - 对齐方式。
- 与 `SuccessCriteria` 比较。
- 不满足时进入 Repair。

验收：

- 执行解释里显示「已验证」或「已自动修复」。

### 第四阶段：Agent Loop 化

目标：

- 复杂排版支持多轮自修复。

落地：

- `WordFormattingAgent.RunAsync(plan)`：
  - AnalyzeDocument
  - GeneratePlan
  - Preview
  - Apply
  - Observe
  - Repair
  - Explain
- 最大修复次数默认 2 次。
- 每轮记录执行解释。

验收：

- 失败不静默。
- 用户能看到 AI 为什么又修了一次。

## UI 设计

### Chat 中的表达

用户输入后，不显示技术细节，显示自然语言摘要。

```text
我将按「正式报告」风格处理全文：
1. 识别标题和正文层级
2. 统一标题格式
3. 统一正文格式
4. 检查行距和缩进

预计影响：全文
风险：中等，需要预览
```

### 执行解释

执行后显示：

```text
已完成排版
- 标题：识别 5 处，已统一为黑体 16pt
- 正文：统一为宋体 12pt，1.5 倍行距
- 段落：首行缩进 2 字符
- 自检：通过
- 撤销：可用
```

### 校对面板

面板顶部固定展示：

```text
校对计划
范围：全文
类型：错别字、标点
模式：建议预览
状态：发现 3 处问题
```

无问题时：

```text
没有发现明显错别字和标点问题
检查范围：全文
检查类型：错别字、标点
[重新检查] [退出校对]
```

解析失败时：

```text
AI 返回格式异常，无法生成校对列表
已保留原始响应，可重新尝试
[重新校对] [查看原始响应] [退出]
```

## 落地顺序

### Step 1：校对结果状态化

- 新增 `ProofreadAnalysisResult`。
- 区分无问题和解析失败。
- 优化 `showProofreadNoIssues()`。
- 增加解析失败 UI。

### Step 2：校对面板体验优化

- Drawer 可折叠。
- 显示本轮校对计划。
- 无问题状态增加退出和重新检查按钮。

### Step 3：排版计划 JSON 化

- `FormattingIntentPlan` 支持 JSON。
- Prompt 引导 LLM 输出同结构 plan。
- 规则编译器作为 fallback。

### Step 4：WordFormattingAgent

- 新增 Word 专属 Agent。
- 统一直接格式和语义排版入口。
- 接入执行解释。

### Step 5：Observe + Repair

- 应用后读取 Word 格式。
- 对照成功标准。
- 自动修复最多 2 次。

### Step 6：组合任务

- 支持「先校对，再排版」。
- 多步骤计划进入 Agent Loop。

## 验收标准

### 用户体验

- 用户可以不选区，直接说「校对全文」或「全文排版」。
- 用户不需要手动选择意图、模板、Skill 或工具。
- 用户能看懂 AI 做了什么。
- 用户能撤销。
- 出错时知道是 AI 返回格式问题、Word 应用问题，还是目标找不到。

### 工程质量

- `ShareRibbon` 不引入 Word/Excel/PowerPoint 具体实现。
- Word COM 操作只在 `WordAi`。
- 每个计划都有结构化 JSON。
- 每个执行步骤都有 Observe。
- 每个高风险操作都有撤销点。

### 验证

- `node --check .\ShareRibbon\Resources\js\message-sender.js`
- `node --check .\ShareRibbon\Resources\js\agent-card.js`
- `devenv.com .\AiHelper.sln /Rebuild Debug`
- `git diff --check`

