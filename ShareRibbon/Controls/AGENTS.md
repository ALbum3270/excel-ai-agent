# SHARED UI AND REQUEST ORCHESTRATION

## OVERVIEW

`ShareRibbon/Controls/` 承载跨 Office 插件共享的 WebView2 Chat UI、请求编排、AgentKernel 接线和流式通信。这里是 UI/运行时适配层，不是宿主 COM 执行层。

## STRUCTURE

```text
Controls/
├── BaseChatControl.vb                 # 共享 Chat UI 与路由入口
├── Models/                            # History、Selection 等 UI 模型
├── Services/AgentKernelService.vb     # Harness/AgentKernel 与 UI 的接线
├── Services/ChatRequestOrchestrator.vb
├── Services/ChatRoutingOrchestrator.vb
├── Services/CodeExecutionService.vb   # 进入宿主 UI 线程执行入口
├── Services/HttpStreamService.vb      # 普通流式请求
└── Services/                          # 其他共享 UI/会话服务
```

## WHERE TO LOOK

| Task | Location | Notes |
|---|---|---|
| Chat UI / WebView2 | `BaseChatControl.vb` | 收集输入、上下文、展示状态，不堆业务路由 |
| Agent 启动与事件 | `Services/AgentKernelService.vb` | 绑定 AgentKernel、Harness、AI 请求和 UI 终态 |
| 普通 Chat 请求 | `Services/ChatRequestOrchestrator.vb` | 构建文本请求和历史 |
| 路由 | `Services/ChatRoutingOrchestrator.vb` | 选择既有执行路径，不复制 Agent 逻辑 |
| 宿主执行边界 | `Services/CodeExecutionService.vb` | 统一调度到宿主执行回调/UI 线程 |
| Streaming / MCP | `Services/HttpStreamService.vb` | 流式响应和兼容工具调用 |

## REQUEST BOUNDARIES

- `BaseChatControl` 只负责输入、上下文、路由和展示；专业 PowerPoint 构图、验证或修复逻辑不得写在这里。
- `AgentKernelService` 只负责把共享 AI 请求能力绑定到 `AgentKernel`，不得创建第二套 Agent/Loop。
- 普通聊天历史仍是文本模型；临时视觉证据不得写入 `HistoryMessage`、ChatState、Memory 或 WebView2 消息。
- 多模态 Repair 请求只能消费 `LoopEngine` 提供的临时 messages，并通过共享 `AiGateway` 发送；请求体不得记录或持久化。
- Provider 不支持图片时返回失败给 Loop 处理降级，不在 UI 服务层递归重试。

## CONVENTIONS

- UI 组件继承 `UserControl`，WebView2 资源位于 `ShareRibbon/Resources/`。
- 长任务必须异步；Office COM 调用只能通过既有宿主 UI 线程入口。
- 所有 Agent 终态都必须恢复输入框、发送/停止按钮和 planning card 状态。
- 修改服务构造函数或委托签名时，同时检查三个宿主的 `ChatControl` 接线。

## ANTI-PATTERNS

- 在共享 UI 控件中直接访问 PowerPoint、Word 或 Excel COM。
- 把专业设计、Scene 或页面类型判断堆进 Chat 路由。
- 把截图 Data URL/Base64 放入 History、JavaScript 消息、日志或 RunTrace。
- 在多模态失败时由 UI 层无限重试或静默吞掉错误。
- 阻塞 UI 线程等待网络请求或文件 IO。
- 硬编码 API Key、Provider URL 或面向用户的敏感调试信息。
