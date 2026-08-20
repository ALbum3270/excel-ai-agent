# Excel Agent 基础版

## 1. 本版范围

Excel Agent 基础版沿用现有主链：

```text
OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> Excel Host
```

本版不新增平行 Harness，也不允许 Python 直接控制 Excel。它补齐以下最小闭环：

- 从当前选区、Excel 表、连续区域或 UsedRange 探测主要 `TableRegion`。
- 将表头、列类型、行列数和探测置信度写入 `ContextPack.host.tables`。
- 用 `ReadRange` 读取小范围值和可选公式，返回结构化 JSON。
- 优先使用原生 Excel 工具完成写入、公式、清洗、格式、图表和报表操作。
- 原生工具不足时，经用户审批后用 `PythonCompute` 做 JSON 到 JSON 的计算。
- 每次 Excel 写操作返回真实 `ToolResult` 和 Observation，包括目标区域、UsedRange、工作表数、图表数及公式错误变化。

## 2. 新增能力

### TableRegion

探测顺序为：

1. 有内容的多单元格选区；
2. 活动单元格所在的 Excel `ListObject`；
3. 活动单元格所在连续区域；
4. 当前工作表 `UsedRange`。

列类型最多抽样 200 行、64 列，输出 `formula`、`date`、`number`、`text` 或 `empty`。多区域选区使用最大的矩形区域，并返回 warning。

### ReadRange

`ReadRange` 是只读、安全工具，支持：

- `A1:F100`
- `Sheet1!A1:F100`
- 可选返回公式矩阵
- 默认最多 5,000 个单元格，调用侧最多可提高到 20,000 个

返回数据包括工作簿、工作表、实际地址、行列数、值矩阵和可选公式矩阵。范围过大时不会读取，会要求缩小或分块。

### PythonCompute

`PythonCompute` 只接收 JSON，代码从 `input_data` 读取输入，并必须给 `result` 赋值。其结果仍是 JSON，不会直接写入工作簿。

基础版限制：

- 属于 `risky` 工具，每次执行前必须经过现有 Safety/Approval 链路。
- 默认超时 20 秒，最大 60 秒。
- 禁止文件、网络、子进程、Excel COM 和反射类操作。
- 只允许导入 `math`、`statistics`、`datetime`、`decimal`、`collections`、`re`、`functools`、`itertools`。
- 输入最多 1,000,000 字符，输出最多 2,000,000 字符。
- 使用独立临时目录，完成或失败后尝试清理。

这是一层应用级约束，不是操作系统级强沙箱。生产环境如需执行不可信代码，应改为低权限独立进程或容器服务。

## 3. 开发环境

必需环境：

- Windows 10/11
- Microsoft Excel 2016 或更高版本
- Visual Studio 2022/2026，安装 Office/VSTO 开发工作负载
- .NET Framework 4.7.2 Developer Pack
- WebView2 Runtime

`PythonCompute` 还需要 Python 3。插件会先读取环境变量 `OFFICE_AI_PYTHON_PATH`，然后依次尝试 `python.exe`、`python3.exe`、`python` 和 `python3`。

PowerShell 示例：

```powershell
[Environment]::SetEnvironmentVariable(
    "OFFICE_AI_PYTHON_PATH",
    "C:\Python312\python.exe",
    "User"
)
```

设置后需要重启 Excel，使加载项读取新的用户环境变量。

## 4. 构建与自动检查

在 Visual Studio Developer PowerShell 中执行：

```powershell
.\build-code.bat
powershell -File .\scripts\run-golden-l0.ps1 -Configuration Debug
powershell -File .\scripts\audit-p0-guardrails.ps1
```

L0 Golden 覆盖：

- Python 未审批不得启动进程或进入 Excel Host。
- TableRegion 已进入 ContextPack。
- `ReadRange` 是 Excel safe 工具，并返回 read Observation。
- `PythonCompute` 是 Excel risky 工具，具备超时、解释器配置和导入限制。
- Excel 写后观察包含工作表、图表和公式错误 delta。

## 5. 手工验收

准备一个包含“日期、区域、销售员、销售额”四列且不少于 20 行的工作表，然后依次验证：

1. 选中数据并询问“这张表有多少行、哪些列是数值列”，回答应使用实际表头和区域。
2. 要求“把销售额列设置为红色字体”，只应修改目标列，Observation 应报告目标范围已变化。
3. 要求“新增一个汇总工作表”，工作表数量应增加 1。
4. 要求“按区域汇总销售额并写入新工作表”，应使用原生透视/分析工具或 `ReadRange -> PythonCompute -> WriteData`。
5. 第一次调用 `PythonCompute` 时应停在审批状态；拒绝后工作簿不得变化。
6. 批准一次 Python 计算后应返回 JSON；同一批准令牌不能重复使用。
7. Python 代码尝试 `import os`、打开文件或访问网络时应被拒绝。
8. Python 代码超过指定超时时间时，进程应终止并返回 `TIMEOUT`。
9. 要求生成销售趋势图，图表数量应增加，目标数据不得被覆盖。
10. 写入错误公式后，Observation 应报告公式错误数量变化，并停止不确定的自动重试。

## 6. 当前边界

基础版暂不包含：

- 超过 20,000 个单元格的自动分块读取和断点续跑；
- 操作系统级 Python 沙箱；
- Power Query 执行引擎；
- 多人协作冲突处理；
- 完整的 Excel 专用快路径 Harness；
- 无 Windows + Excel 的跨平台 VSTO 集成测试。

这些不影响小中型表格的基础闭环，但在正式分发前仍需完成 Windows 真机 smoke 和 MSI 内容审计。
