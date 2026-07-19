# AI Native TaskSpec、Skill 与 Tool 路由合同

## 目标链路

```text
用户任意表达
  -> Office Context + 会话上下文
  -> 开放式 TaskSpec
  -> Skill 召回与完整内容加载
  -> Skill 允许的 Tool
  -> plan / act / observe / repair / explain
```

`OfficeIntentType` 是兼容、统计和粗粒度召回标签，不是能力白名单。未被枚举覆盖的新需求仍须
保留原始用户目标并进入 Skill/Tool 规划。

## 权威数据

- 用户原始输入是 `AgentTaskSpec.Goal` 的权威来源。
- `AiNativeRuntimeResult.TaskSpec` 和 `SelectedSkills` 必须沿
  `ChatRoutingOrchestrator -> AgentKernelService -> UserTurn -> OfficeHarness -> AgentKernel`
  传递，不允许 Kernel 无条件重新分析并覆盖。
- 当前只允许一个 primary Skill 建立 `allowed-tools` 安全边界；其他召回结果用于 trace。
- Tool Registry 是可执行原子能力的权威来源。没有 Tool 时报告 `capability gap`。
- AiNative 分析阶段与 AgentKernel 执行阶段必须通过同一运行时目录解析入口加载 Tool，
  禁止前置分析看到空工具集、执行阶段再加载另一套工具。

## Skill 发现与召回

Skill 搜索根依次覆盖：

1. `Documents/OfficeAiAppData/Skills` 用户目录；
2. 当前程序集或宿主输出目录的 `Skills`；
3. 向上查找的 `ShareRibbon/Skills` 开发目录。

SQLite 索引用于加速和统计，不是 Skill 存在性的权威来源。首次启动、索引尚未完成或索引过期时，
必须直接扫描元数据目录召回。匹配后才加载 `SKILL.md` 正文、references 和 scripts。

每个宿主可以声明一个 `default_for_application: true` 的 baseline Skill。专用 Skill 语义命中优先；
完全没有词法命中时才选择 baseline Skill，从而覆盖无法穷举的新表达，而不是继续增加关键词或意图枚举。

日志分别报告 `discovered`、`active`、`updated`，避免把“本轮无需更新”误报成“发现 0 个 Skill”。

## 新需求扩展

- 现有 Tool 可组合：只由 Agent 生成新计划，不改枚举。
- 需要新的领域工作流：新增或更新目录型 Skill。
- 缺少 Office 原子能力：新增 Tool schema、宿主 executor、observer/verifier，再授权给 Skill。
- 禁止以扩大关键词 `if/else` 或默认回退 VBA 的方式模拟新能力。

## 问答与执行分流

意图名称只作为兼容标签。模型还必须返回独立的 `interactionMode`：

- `answer`：解释原因、用法或概念，进入普通回答；
- `clarify`：只请求阻塞执行的最小信息；
- `execute`：进入 TaskSpec、Skill、Tool 和 Observe 链路。

这避免把 SmartArt 用法问题等普通问答强行送入 Agent 规划，也避免新型执行需求因为不在枚举中就被当成问答。

## 终态与验收

- success、planning failure、tool failure、capability gap、exception、cancel 都必须进入同一个 UI 终态函数。
- 终态必须清理 planning card、替换残留 thinking 消息、解锁输入框、隐藏 `stop-button` 并恢复 `send-button`。
- Tool 返回成功只表示原子调用成功，不等于用户目标完成。
- `TaskSpec.ExpectedOutputs` 保存页数、图片等可验证产物；Loop 结束前必须检查实际成功的工具调用。
- 要求图片时必须有成功的 `InsertImage`；占位形状或纯文本幻灯片不能判定为完成。
