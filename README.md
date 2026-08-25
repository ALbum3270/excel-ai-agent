# Excel Agent

这是从原 Office 插件仓库中独立迁出的 Excel AI Agent。仓库只包含 Excel VSTO 宿主、共享 Agent 运行时、Excel 工具适配器和 WebView2 任务窗格，不包含 Word、PowerPoint、安装器或旧聊天功能。

## 核心链路

```text
ChatControl
  -> AgentRunner
  -> OfficeHarness
  -> AgentKernel
  -> LoopEngine (plan -> act -> observe -> repair -> complete)
  -> ToolRegistry / SafetyGate
  -> ExcelAgentHost
  -> Excel Runtime Adapters
```

关键边界：

- 一次工具调用只产生 `ToolResult`，不会直接完成目标。只有模型明确返回 `decision=complete`，并通过冻结目标和宿主证据校验后，任务才进入完成态。
- 写入权限由任务窗格中的“执行 / 只读”控件明确授予，不再从关键词、旧意图路由或模型标签推断。
- Excel 宿主工具均为可信清单能力。Skill 用于指导规划，不会把已注册的 Excel 工具错误裁剪成不可用。
- 长尾 Excel 对象操作统一走 `DiscoverOfficeCapability -> OfficeObjectOperation`；不提供 VBA、MCP 或任意脚本旁路。
- `PythonCompute` 只接收 JSON 并返回 JSON，不接触 Excel COM、文件或网络；计算结果需要再由 Excel 写入工具落表。
- 停止操作会同时取消 Harness、Agent Loop 和正在进行的 AI HTTP 请求。

## 目录

```text
ExcelAi/              Excel VSTO 宿主、任务窗格和 Excel 执行器
ExcelAgent.Core/      Excel Agent 核心、提示词、Skill 和工具清单
scripts/build.ps1     还原依赖并构建解决方案
tests/run.ps1         架构边界和运行时回归检查
```

## 构建

前置条件：Visual Studio 2022 Build Tools、.NET Framework 4.7.2、Office/VSTO 开发工具和 WebView2 Runtime。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

`scripts/build.ps1` 使用 `BuildCodeOnly=true`，只验证并产出程序集，因此不依赖本地私钥。要在 Excel 中调试或发布 VSTO 加载项，请使用 Visual Studio 的正常 VSTO 构建，并在本机配置清单证书；发布构建需显式启用 `GenerateManifests=true`、`SignManifests=true` 并提供证书。私钥文件已被 `.gitignore` 排除，不能提交到仓库。

## 验证

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run.ps1
```

API 配置保存在当前用户 `Documents\ExcelAiAgent\settings.json`，API Key 使用 Windows DPAPI 加密。
