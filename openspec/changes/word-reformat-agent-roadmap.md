# Word AI Native 排版 Agent 落地计划

## 目标

把 Word 的“排版”从多个分散入口升级为一个可解释、可预览、可观察、可修复的 AI Native 排版 Agent。用户点击 Ribbon 排版按钮或在 Chat 中输入自然语言需求时，系统应自动判断范围、理解意图、生成方案、展示可读预览，并在应用后验证真实 Word 文档状态。

## 设计原则

- 用户不需要理解技术标签、模板映射或 JSON 指令。
- 默认自动推断范围：有选区优先选区，用户明确说全文/整篇时作用全文，无选区时作用全文。
- “排版”是统一任务入口，标题编号重构、标准排版、局部格式调整都进入同一个用户心智模型。
- `ShareRibbon` 只承载共享模型、卡片 UI、语义映射和通用编排；Word COM 操作、范围解析、样式应用、执行观察必须留在 `WordAi`。
- Agent Loop 固定为 `Plan -> Preview -> Apply -> Observe -> Repair -> Explain`，允许先分阶段实现。

## 当前问题

- Ribbon 排版按钮只打开 Chat 输入框，且要求选区；与“一键排版”的按钮说明不一致。
- Chat 排版、直接格式命令、模板排版路径分裂，同一用户需求可能进入不同执行体系。
- 预览卡片展示 `body.normal`、`title.1` 等技术标签，用户不容易判断是否可信。
- “预览对比”目前只是文本列表，不是面向用户的真实前后效果预览。
- 语义排版主链路缺少完整 Observe/Repair，应用后没有稳定地读取 Word 状态确认结果。
- 直接格式命令对中文 Word 用户常说的“加大 2 号/小 1 号”理解不够细。
- `ShareRibbon` 的旧 Word 应用反射访问路径已清理；后续重点转为真实 Word 宿主中的端到端体验验证。

## 实施清单

## 后续重排优先级

### P1：产品体验优先

- [x] “换一种”改为真实多方案切换，而不是简单重新分析。
- [x] 每个备选方案明确说明采用的标准/风格，以及和当前方案的差异。
- [x] 预览卡片继续沿用同一 `WordFormattingTaskPlan` 摘要。

### P2：统一执行管线

- [x] 建立 `WordFormattingTaskPlan` 摘要模型。
- [x] 将语义排版应用也逐步封装为 `WordFormattingAgent` 的任务执行结果。
- [x] 让直接格式和语义排版共享同一套 `Plan -> Preview -> Apply -> Observe -> Repair -> Explain` 摘要表达。

### P3：边界和旧代码清理

- [x] 共享层旧 Word 反射入口改为兼容占位。
- [x] 清理确认无调用的旧前端函数、旧样式和旧服务方法。
- [x] 评估并移除 `ShareRibbon` 对 `Microsoft.Office.Interop.Word` 的直接引用。

### P4：宿主验收

- [ ] 在真实 Word 宿主中验证 Ribbon 排版、Chat 排版、换一种、微调、应用、撤销。
- [x] 记录失败日志和用户可见行为，补充自动化可覆盖的检查。

### P4 自动验收记录

- [x] 前端排版卡片动作 `applySmartReformat`、`refineSmartReformat`、`switchReformatTemplate`、`previewReformatCompare` 均已在 `BaseChatControl` 注册。
- [x] `WordAi.ChatControl` 已覆写应用、微调、换一种、预览对比等宿主动作入口。
- [x] `Ribbon1.ReformatButton_Click` 已接入 `ChatControl.TriggerSmartReformat()`。
- [x] `ShareRibbon` 中未再检出 `showFormattingPreview`、`onReformatApplied`、`onMirrorFormatReady`、`ProcessMirrorFormatInternal`、`GetWordApplication` 旧链路。
- [x] `ShareRibbon` 源码和项目文件中未再直接引用 `Microsoft.Office.Interop.Word`。
- [x] `node --check ShareRibbon\Resources\js\reformat-chat.js` 通过。
- [x] `node --check ShareRibbon\Resources\js\proofread-ui.js` 通过。
- [x] `git diff --check` 通过，仅有 LF/CRLF 提示。
- [x] `devenv.com .\AiHelper.sln /Rebuild Debug` 通过，`4 成功，0 失败，0 已跳过`。
- [x] Word 宿主烟测通过：可启动 Word、创建临时文档，`WordAi` 出现在 `COMAddIns` 且 `Connect=True`。

### P4 真实宿主验收清单

- [x] Word 宿主启动后能加载 `WordAi` 插件。
- [ ] 未选中文本点击 Ribbon “排版”：应自动分析全文并展示排版建议卡片。
- [ ] 选中文本点击 Ribbon “排版”：应分析选区，并在卡片中显示作用范围。
- [ ] Chat 输入“给我重构序号和标题”：应进入排版建议卡片，不应退回普通聊天澄清。
- [ ] Chat 输入“把全文字体统一加大2号”：应作用全文并按中文字号等级增大。
- [ ] 点击“换一种”：应切换到不同标准/风格，并展示差异说明。
- [ ] 点击“微调”：应基于当前方案继续调整，不应丢失活动排版上下文。
- [ ] 点击“应用排版”：应修改真实 Word 文档，并展示 Observe/Repair 后的执行摘要。
- [ ] 点击“撤销”：应恢复本次排版前的快照。

### 阶段 1：统一入口与第一感知

- [x] Ribbon 排版按钮改为“一键分析当前选区/全文”，直接生成排版建议卡片。
- [x] 去掉 Ribbon 对选区的强制要求，保留无有效文档/无段落提示。
- [x] Chat 中明确排版需求继续走同一套 `ReformatJob -> PreviewPlan`。
- [x] 状态栏提示明确说明作用范围、段落数和推荐标准。

### 阶段 2：卡片 UX 升级

- [x] 排版卡片使用用户可读名称，不直接展示语义标签 ID。
- [x] 卡片展示范围、文档类型、置信度、推荐标准、将修改的样式区和段落数。
- [x] 当计划变更数为 0 时，展示“已完成分析但暂未发现需要调整的样式区”，避免用户误解为失败。
- [x] “预览对比”输出更接近用户语言的变更摘要。

### 阶段 3：意图与范围推断增强

- [x] 对“重构序号和标题/整理标题编号/规范标题层级”做确定性结构意图识别。
- [x] 对“全文/整篇/所有/统一”优先选择全文，即使当前存在选区。
- [x] 对“选中/所选/当前选择”优先选择选区。
- [x] 使用 LLM 的 `scope` 作为辅助信号，但不让低置信度结果覆盖明确用户表达。

### 阶段 4：直接格式命令 Agent 化

- [x] 将“字体统一加大 2 号”解释为中文字号等级调整，而不是简单 `+2pt`。
- [x] 对 `pt/磅/点` 明确单位继续使用磅值增量。
- [x] 直接格式命令执行后保留 Observe 摘要，告诉用户实际影响了多少范围。
- [x] 把直接格式和语义排版收敛到统一 `WordFormattingTaskPlan` 摘要模型。

### 阶段 5：Observe/Repair 闭环

- [x] 语义排版应用后读取应用结果与标签分布，输出观察摘要。
- [x] 观察到结构段落样式未生效时，自动尝试一次保守修复。
- [x] 应用结果卡片展示修改段落数、观察结论、撤销提示。

### 阶段 6：边界清理与旧链路收敛

- [x] 梳理 `ShareRibbon.Controls.Services.ReformatService` 中旧 Word 反射路径。
- [x] 将具体 Word 应用访问迁移到 `WordAi` 或标记为废弃入口。
- [x] 删除证据明确的死代码、旧 JS/CSS 和未引用样式。
- [x] 保持 `.vbproj` / 前端资源注册一致。

## 本轮落地范围

- [x] 完成阶段 1。
- [x] 完成阶段 2 的卡片可读性和零变更文案。
- [x] 完成阶段 3 的确定性范围/结构意图增强。
- [x] 完成阶段 4 的中文字号等级调整基础能力。
- [x] 编译验证：`devenv.com .\AiHelper.sln /Rebuild Debug`。
- [x] 静态验证：`git diff --check`，JS 改动时执行 `node --check`。

## 验收用例

- 点击 Ribbon “排版”，未选中文本但有打开文档时，应自动分析全文并展示排版建议卡片。
- 选中一段内容后点击 Ribbon “排版”，应分析选区并展示作用范围。
- Chat 输入“给我重构序号和标题”，应进入排版建议卡片，不应退回普通聊天澄清。
- Chat 输入“把全文字体统一加大2号”，应作用全文并按中文字号等级增大。
- 卡片中不应把 `body.normal`、`heading.1` 作为主要用户文案展示。
