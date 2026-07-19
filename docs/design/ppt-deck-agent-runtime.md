# 专项设计：PowerPoint Professional Design Agent 运行时

| 项 | 内容 |
|---|---|
| 状态 | 当前实现合同 |
| 主入口 | `CreateSlides` |
| 主链 | `OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> PowerPoint host` |
| 专业设计实现 | `PowerPointAi/Design/` |
| Skill | `ShareRibbon/Skills/powerpoint-deck-agent/SKILL.md` |
| Observe 合同 | [`tool-result-observation.md`](./tool-result-observation.md) |
| 动态长尾能力 | [`office-object-operation-integration.md`](./office-object-operation-integration.md) |

---

## 1. 目标与非目标

### 1.1 目标

1. 从用户目标和 Office 上下文生成具有叙事结构、专业构图和统一视觉语言的多页 Deck。
2. 使用结构化 Scene 表达语义，由确定性布局引擎编译为可编辑 PowerPoint Shape。
3. 在写入前完成整套 Deck 预检，在写入后验证真实 PowerPoint COM 渲染结果。
4. 失败时提供结构化 Observation、视觉证据、安全回滚，并由现有 Loop 修复参数。
5. 保持唯一 Harness/Loop/ToolRegistry/Safety 主链，不按页面类型扩展 Tool。

### 1.2 非目标

- 不训练专用配色或版式模型。
- 不保证任意第三方模板的像素级复刻。
- 不在 Executor、Renderer 或 Verifier 内直接调用 LLM。
- 不为 cover、timeline、architecture、SmartArt 等场景新增页面专用 Tool。
- 不用位图替代所有内容；默认输出应保持 PowerPoint 对象可编辑。

---

## 2. 当前运行时总览

```text
UserTurn (PowerPoint)
  -> OfficeHarness 构建 ContextPack / 选择 powerpoint-deck-agent
  -> AgentKernel / LoopEngine 生成 CreateSlides 调用
  -> ToolRegistry / CodeExecutionService 进入 PowerPoint UI 线程
  -> ProfessionalDeckExecutor
       -> Parse DeckDesignSpec / SlideDesignSpec
       -> Resolve DesignSystem
       -> Compile all Scene plans
       -> Deck preflight + composition rhythm verification
       -> preview: return plan summary without COM write
       -> execute: render slide -> observe real COM result -> pixel verification
       -> pass: ToolResult.Succeed + structured Observation
       -> fail: capture ephemeral screenshot -> safe rollback -> ToolResult.Failed
  -> LoopEngine
       -> multimodal Repair when evidence is available
       -> text-only Repair fallback when image input is unavailable
       -> retry registered CreateSlides with corrected parameters
```

高频专业生成继续使用稳定工具 `CreateSlides`。SmartArt、艺术字等长尾对象能力通过 `DiscoverOfficeCapability + OfficeObjectOperation` 补充，不新增平行页面工具。

---

## 3. Scene 输入合同

### 3.1 Deck

```json
{
  "designSystem": "modern-tech",
  "designTokens": {},
  "slides": []
}
```

- `designSystem` 必须是已注册设计系统；未知名称只有在提供完整 `designTokens` 时才允许。
- 单批至少一页，最多 50 页。
- `preview=true` 只编译和预检，不创建再删除幻灯片。

### 3.2 Slide

```json
{
  "id": "slide-1",
  "slideType": "content",
  "variant": "feature-left",
  "eyebrow": "SECTION",
  "title": "结论式标题",
  "subtitle": "辅助说明",
  "keyMessage": "本页唯一核心结论",
  "items": [],
  "metrics": [],
  "chart": null,
  "table": null,
  "imagePath": "",
  "notes": "",
  "source": ""
}
```

- 嵌套 `scene` 与顶层字段合并，Scene 字段覆盖同名顶层字段，未覆盖元数据继续保留。
- 未识别的 `slideType`、`variant`、`designSystem` 显式返回可修复 schema 错误，不静默退化。
- chart、table、imagePath 是 content 页面的互斥主视觉。
- 长文本优先拆页、改变构图或压缩语义，不能无限缩小字号。

### 3.3 已注册 Scene 类型

| slideType | 主要用途 |
|---|---|
| `cover` | 封面和主题建立 |
| `section` | 章节过渡 |
| `statement` | 强结论与少量证据 |
| `content` | 洞察、图文、chart、table |
| `two-column` | 双对象并列 |
| `comparison` | 双对象或比较表 |
| `kpi` | 指标和业务结果 |
| `process` | 流程、阶段和时间序列 |
| `architecture` | 分层或中心辐射架构 |
| `matrix` | 显式双轴四象限 |
| `quote` | 引用和观点 |
| `closing` | 总结与行动建议 |

当前专业变体包括：

- `content: feature-left`
- `kpi: hero-left`
- `process: vertical`
- `architecture: hub-spoke`

新增变体应扩展 Scene 编译器和验证器，不新增 Tool。

---

## 4. 设计系统与专业构图

`DesignSystemCatalog` 提供颜色、字体、字号、表面、分隔线、正负语义色等 Token。所有组件从 Token 取值，不在页面实现中散落主题常量。

专业构图至少满足：

- 标题表达结论而非目录标签。
- 每页存在明确视觉焦点和信息层级。
- 同一 Deck 保持字体、颜色、间距和组件语言一致。
- 连续三页不得使用相同的非焦点构图签名。
- comparison 强调项允许非对称权重，但不能改变比较语义。
- chart 支持正负值和跨零轴，table 根据内容动态分配列宽。
- architecture hub-spoke 使用可编辑节点和连接线。
- notes、source、chart/table 标签等 Scene 字段必须被真实消费或明确拒绝。

---

## 5. 编译、预览与写入事务

### 5.1 Deck 级预编译

正式写入前先编译所有 `SlideRenderPlan` 并执行 Preflight：

- Scene node ID 完整且唯一。
- Bounds 合法，无不可修复溢出和冲突。
- 文本层级、局部对比度和视觉结构满足门槛。
- 整套 Deck 的构图节奏不连续重复。

任一页面预检失败时，不开始部分写入。

### 5.2 Preview

Preview 返回：

- `changed=false`
- `rendered=false`
- `createdSlides=0`
- Scene 编译与 Deck 预检报告

Preview 不允许创建幻灯片后再删除，因为那会引入 COM 副作用和错误回滚风险。

### 5.3 Execute

每页执行：

1. 编译/复核 Scene plan。
2. `PowerPointSceneRenderer` 创建可编辑 Shape。
3. 写入 notes 并回读验证。
4. `PowerPointVisualVerifier` 对真实 Slide 做结构、几何、层级和像素检查。
5. 把 target ref、issues、metrics、aesthetic score 写入结构化 Observation。

---

## 6. 真实 PowerPoint 视觉验证

当前验证覆盖：

- 缺失/重复 Scene node。
- 无效边界、越界、碰撞和文本溢出。
- 标题/正文层级与字号可读性。
- 颜色对比、视觉焦点和结构完整性。
- 计划节点与实际带标记 Shape 的对应关系。
- notes、image 等请求资产是否真实写入。
- `Slide.Export` 是否成功产生有效 PNG。
- 真实像素的量化颜色数量、亮度标准差、近纯色/空白检查。
- Deck 连续构图重复。

专业交付阈值和具体指标属于 Verifier 实现配置；调整门槛时必须同步回归 fixture 和文档，不在 Prompt 中硬编码另一套标准。

轻量像素统计用于发现空白、扁平和低变化输出，但不能单独证明设计专业；必须与 Scene 语义、结构和层级验证结合。

---

## 7. 临时视觉证据与多模态 Repair

当渲染后验证失败：

1. 在回滚前通过 `Slide.Export` 采集当前实际渲染截图。
2. 保持页面宽高比，限制导出像素和最大字节数。
3. 读取为 Data URL 后立即删除临时文件。
4. 只放入 `ToolResult.VisualEvidence` 内存字段；该字段排除普通 JSON 序列化。
5. 执行安全回滚并返回 `VERIFY_FAILED`。
6. `LoopEngine` 将错误合同、原参数和最多少量截图组合为多模态 Repair 请求。
7. `AiGateway` 转换 OpenAI-compatible `image_url` 与 Anthropic base64 image block。
8. Provider/模型不支持图片时，同一 Run 停止重复尝试视觉通道，降级为文本 Repair。

禁止把截图写入 `Observation`、`Data`、`Artifacts`、History、Memory、RunTrace、日志、Notes 或 WebView2 UI。

当前多模态证据用于增强“已经由确定性 Verifier 判失败”的诊断与参数修复；它不是独立的第二套视觉审核 Loop。

---

## 8. 安全回滚与失败合同

- 仅删除包含本批 `office-ai-design:` Shape 标记、能够可靠识别的生成页。
- 不按“初始页数之后全部删除”猜测回滚范围。
- 无法可靠识别全部变更时返回 `PARTIAL_APPLY`，保留真实 Observation，不伪装成完整回滚。
- 截图采集失败不阻止安全回滚。
- notes、image 或视觉结构缺失属于交付失败，不以 warning 假成功。
- 未知 Scene、不可恢复 COM 状态和文档缺失使用统一 `ExceptionClassifier` 合同。

---

## 9. COM 规则

- PowerPoint COM 实现只位于 `PowerPointAi`，共享合同位于 `ShareRibbon`。
- 所有 COM 操作经过既有宿主 UI 线程入口；Executor/Renderer 不创建线程。
- 临时 `Presentation`、`Slides`、`Slide`、`ShapeRange`、`FillFormat`、`TextFrame2`、`TextRange2` 等对象在 `Finally` 中显式释放。
- 不硬编码 `Placeholders(2)`；按 `PpPlaceholderType` 查找并回读验证。
- `Slide.Export`、字体度量、透明度、阴影和文本边界可能受 PowerPoint 版本、主题和字体环境影响，因此必须保留渲染后验证。

---

## 10. 回归与验收

Fixture：

- `PowerPointAi/Design/Fixtures/professional-command-regression.json`
- `PowerPointAi/Design/Fixtures/professional-visual-regression.json`

至少覆盖：

- 嵌套 Scene 合并和未知字段拒绝。
- signed chart、动态 table、comparison emphasis。
- `feature-left`、`hero-left`、`vertical`、`hub-spoke`。
- matrix 显式双轴语义。
- notes 回读和缺失失败。
- Deck 构图重复门禁。
- 像素空白/扁平检测。
- 回滚只删除安全标记页。
- 视觉证据不进入 ToolResult JSON、Trace、History 或日志。
- 多模态失败只降级一次并继续文本 Repair。

静态源码解析不能替代以下验证：

1. `ShareRibbon` 与 `PowerPointAi` 的真实 MSBuild。
2. 安装目标 Office 版本中的 COM 运行。
3. 不同字体、主题、分辨率和纵横比的真实导出。
4. OpenAI-compatible 与 Anthropic 视觉模型的请求兼容性。

---

## 11. 后续重点

1. 建立真实 PowerPoint 截图 golden 与人工专业设计评分集。
2. 增强跨页叙事节奏、图像资产选择和数据故事能力。
3. 在不建立第二套 Loop 的前提下研究“渲染、视觉审核、提交”的事务化两阶段方案。
4. 扩展更多 Scene variant 时优先复用组件和通用动态 Office API。
5. 持续验证不同 PowerPoint 版本的 COM 渲染差异和失败恢复。

---

*专业 Deck 的交付标准不是“生成了若干页”，而是 Scene 语义被真实消费、PowerPoint 对象可编辑、整套视觉一致、渲染结果经过验证，并且失败能够安全回滚和修复。*
