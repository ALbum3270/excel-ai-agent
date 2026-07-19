# PowerPoint 原生工具路由与失败语义

## 问题

PowerPoint 执行请求曾出现以下错误链：LLM 已识别为 `SLIDE_CREATE`，但路由仍显示
`GENERAL_QUERY`；工具清单使用 `ppt` 而运行时使用 `PowerPoint`，导致原生工具被过滤；
规划器随后只看到默认关闭的 `ExecuteVBA`；不可恢复失败后循环仍继续执行后续步骤，并在
最终失败后回退普通聊天。

## 运行时合同

1. `IntentResult.OfficeIntent` 是 Word、Excel、PowerPoint 共用的权威路由字段；
   `IntentType` 仅作为旧版 Excel 兼容字段。
2. 工具宿主名在匹配前规范化：`ppt`/`PowerPoint`、`xls`/`Excel`、
   `doc`/`Word` 分别视为同一宿主。
3. 当前安全策略禁用 VBA 时，`ExecuteVBA` 不注入规划提示词。原生 JSON 工具优先；
   VBA 只能在显式启用并经过安全裁决后作为回退。
4. `recoverable=False` 会终止整个计划，禁止继续执行依赖失败步骤的后续动作。
5. Agent 已产生结构化执行失败时，不再回退普通聊天生成第二份回答。

## 内容与图片任务

创建多页演示文稿应优先使用 `CreateSlides`，排版使用 `FormatSlide`、`FormatText`、
`InsertShape` 等原生工具。`InsertImage` 只能接收真实可访问的本地图片路径。没有图片来源或
图片获取工具时，Agent 必须说明图片未插入；占位形状不能被描述为已经完成配图。

## 回归场景

- “帮我写 2 页唐诗 PPT”应保持 `SLIDE_CREATE`，并能看到 `CreateSlides`。
- “给当前页插入本地图片”应能看到 `InsertImage`，参数缺失时只请求必要路径。
- Word/Excel 工具简称与运行时宿主名应继续正确匹配。
- VBA 关闭时规划提示词中不出现 `ExecuteVBA`。
- 任一步返回不可恢复错误时，后续步骤不执行，且不会自动转为普通聊天。
