# SHARE RIBBON SHARED CORE

## OVERVIEW

`ShareRibbon/` 是 Excel、Word、PowerPoint 插件共享的 AI Native 核心库，负责 Harness/Agent/Loop 合同、工具注册、AI 请求、共享 UI、配置、Skill 与 MCP 基础设施。它不得承载任何宿主专属 COM 实现。

## STRUCTURE

```text
ShareRibbon/
├── Agent/                 # AgentKernel、LoopEngine、ToolRegistry、Harness、Safety、共享合同
├── Config/                # API、模型、运行时配置
├── Controls/              # WebView2 Chat UI 与共享服务
├── Mcp/                   # Model Context Protocol client
├── Services/Ai/           # AiGateway 与 Provider 请求适配
├── Skills/                # 目录型 Skill 与兼容 Skill 定义
├── Tools/                 # 工具 schema
└── Resources/             # HTML/JS/CSS 与嵌入资源
```

## WHERE TO LOOK

| Task | Location | Notes |
|---|---|---|
| Agent 主链 | `Agent/AgentKernel.vb`, `Agent/LoopEngine.vb` | 规划、执行、观察、修复和终态 |
| ToolResult / Tool 注册 | `Agent/ToolRegistry.vb` | 工具合同、临时视觉证据合同、执行适配 |
| Harness | `Agent/Harness/` | 唯一运行入口、上下文和 RunTrace |
| AI 请求与 Provider 适配 | `Services/Ai/AiGateway.vb` | OpenAI-compatible / Anthropic 请求转换 |
| Chat UI 与请求编排 | `Controls/` | `BaseChatControl`、`AgentKernelService`、streaming 服务 |
| 配置 | `Config/` | API Key、模型、Prompt 与运行时设置 |
| Skill | `Skills/` | 召回元数据、allowed-tools、references/scripts/assets |
| MCP | `Mcp/` | MCP client、连接与协议模型 |
| WebView2 资源 | `Resources/` | HTML/JS/CSS；修改后检查项目注册和引用 |

## OWNERSHIP BOUNDARIES

- `ShareRibbon` 定义宿主无关合同，例如 `ToolResult`、视觉证据、Office object ref、operation batch、context、capability 和 loop 状态。
- PowerPoint、Word、Excel 的 COM resolver、executor、observer、renderer 必须放在各自宿主项目。
- 共享能力沿 `OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> Host` 接入，不得建立平行 Harness/Loop/Safety。
- AI 请求统一经过共享请求层；宿主 Executor 不得自行创建模型客户端或发起 Repair 请求。

## VISUAL EVIDENCE RULES

- `AgentVisualEvidence` 只用于当前 Observe/Repair 轮次，必须保持宿主无关。
- 截图的 `DataUrl`/Base64 必须是内存临时载荷，并从普通 JSON 序列化中排除。
- 禁止把视觉载荷复制到 `Observation`、`Data`、`Artifacts`、History、Memory、RunTrace、日志或 UI。
- `LoopEngine` 负责限制证据数量和大小、构造多模态消息以及文本降级；Provider 格式转换属于 `AiGateway`。
- 图片通道失败后必须有限次降级，不能在同一 Run 内反复发送必然失败的多模态请求。

## CONVENTIONS

- 共享服务优先使用依赖注入或显式委托，避免从 UI 入口反向引用宿主实现。
- 新增 `.vb`、资源或 Skill 文件后同步检查 `ShareRibbon.vbproj`。
- 日志必须脱敏；不得记录 API Key、完整文档内容或视觉证据载荷。
- VB.NET/VSTO、构建和验证硬规则以根 `CLAUDE.md` 为准。

## ANTI-PATTERNS

- 在 `ShareRibbon` 中引用 PowerPoint/Word/Excel COM 类型。
- 在单个插件中复制 `AgentKernel`、`LoopEngine`、`AiGateway` 或共享 ToolResult 合同。
- 把 Base64 截图序列化进工具结果、Trace 或聊天历史。
- 在 Executor 内直接调用模型或实现第二套视觉修复循环。
- 绕过 `ToolRegistry`、`SafetyGate` 或 `CodeExecutionService` 执行宿主写操作。
- 硬编码 API Key、Provider URL 或模型能力假设。
