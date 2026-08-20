# 专项设计：Excel Table Agent 运行时

| 项 | 内容 |
|---|---|
| 版本 | **v0.3（基础版落地）** |
| 状态 | 基础闭环代码已实现；待 Windows + Excel 真机验收 |
| 实现状态 | **基础版已实现**：TableRegion 强类型探测已进入 `ContextPack.host.tables`；新增结构化只读 `ReadRange` 和审批型 JSON-to-JSON `PythonCompute`；原生 Excel 命令已有目标 Range、公式错误、UsedRange、工作表与图表 delta Observation。统一 `ExcelActionHarness`、大表自动分块和强隔离 Python 仍待后续。 |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §6.2 |
| Skill | `Skills/excel-table-agent/SKILL.md` |
| 现有 | `ExcelDirectOperationService`、`ExcelContextService`、`ExcelJsonCommandSchema`、Tools/excel（27）、ExcelDna UDF |
| 关联 | ContextPack Excel host、Observe Diff、Safety 大表阈值、Golden U3 |

---

## 1. 目标与非目标

### 1.1 目标

1. 定义 Excel 从 **表探测 → 规划 → 工具执行 → 观察 → 修复** 的标准运行时。
2. 对标 Copilot 表格分析 + Cursor 级工具精度（公式、图、透视、清洗）。
3. 规定大表分块、公式 repair、图表选型、结果落表策略。
4. 设计 `ExcelActionHarness` 与 Capability 列表（对齐 Word 样板）。

### 1.2 非目标

- v1 不做完整 Power Query 执行引擎（占位能力保持「明确不做」或独立后期）。
- 不把 ExcelDna 同步 UDF 并进 Agent 主路径（隔离策略见 §10）。
- 不做多人共编冲突解决。

---

## 2. 运行时总览

```text
UserTurn (Excel)
  → ContextHub: active sheet, selection, tables[], usedRange, dtypes
  → SkillRouter: excel-table-agent
  → Planner:
       detect TableRegion
       choose pipeline (clean | formula | pivot | chart | report | mixed)
  → Safety: range size, protect, VBA
  → Loop tools (allowed-tools only)
  → Observe: value/formula/error/chart metrics
  → Repair or Complete
```

### 2.1 ExcelActionHarness（建议）

高置信确定性场景可快路径（仍写 Trace）：

| CapabilityId | 触发例 | 置信 |
|---|---|---|
| `excel.detect-table` | 隐式，总是可先跑 | 1.0 |
| `excel.quick-sum` | 「求和」「合计」+ 数值列清晰 | ≥0.8 |
| `excel.quick-chart` | 「做个柱状图」+ 选区二维表 | ≥0.75 |
| `excel.clean-duplicates` | 「去重」 | ≥0.8 |

复杂分析、多步、不清晰列语义 → 全量 Loop。

---

## 3. 表探测（TableRegion）

### 3.1 探测优先级

```text
1) 选区非空且 itemCount>1 → Selection 为核
2) 选区在 ListObject 内 → 整个 ListObject
3) ActiveCell 落在 UsedRange → 扩展为连续块（启发式）
4) 否则 UsedRange（若过大则只取当前区域邻域）
```

### 3.2 TableRegion 模型

```json
{
  "sheet": "Sheet1",
  "address": "A1:F200",
  "hasHeader": true,
  "headerRow": 1,
  "headers": ["日期", "区域", "金额"],
  "columnTypes": ["date", "text", "number"],
  "rowCount": 199,
  "colCount": 6,
  "source": "list_object | selection | used_range",
  "listObjectName": "Table1",
  "confidence": 0.92,
  "warnings": []
}
```

### 3.3 列类型猜测

| 信号 | 类型 |
|---|---|
| 可解析日期比例高 | date |
| 数值比例高且非文本格式 | number |
| 公式单元格比例高 | formula |
| 其它 | text |

抽样最多 200 行；结果进 ContextPack.host.tables。

### 3.4 失败

- 空表 / 单格：TaskSpec exploratory，先 `DataAnalysis` 只读描述，不写。
- 多区域选区：拆分或取最大矩形，warnings 记录。

---

## 4. 任务管道（Pipeline）

Planner 输出 `excelPipeline` 字段（可进 TaskSpec.meta）：

| Pipeline | 典型步骤 | 主 Tools |
|---|---|---|
| **profile** | 只读画像 | DataAnalysis |
| **clean** | 去空、去重、替换、Trim | CleanData, RemoveDuplicates, FindReplace |
| **formula** | 生成/修复/填充 | ApplyFormula, WriteData |
| **transform** | 转置、拆列 | TransformData |
| **aggregate** | 汇总/透视 | CreatePivotTable, DataAnalysis, WriteData |
| **chart** | 选型+插入 | CreateChart |
| **report** | 新表输出结论 | CreateSheet, GenerateReport, WriteData, CreateChart |
| **format** | 美化 | FormatRange, ConditionalFormat, AutoFit |
| **mixed** | 上列组合 | 有序多段 |

### 4.1 默认组合启发式

| 用户意图信号 | Pipeline |
|---|---|
| 分析/统计/平均/排名 | profile → aggregate → (chart?) → report |
| 清洗/去重/空值 | clean → profile |
| 公式/计算列/修复 | profile → formula |
| 图/可视化 | profile → chart |
| 透视 | aggregate(pivot) |
| 转置 | transform |

---

## 5. 公式生成与 Repair

### 5.1 生成流程

```text
1. 锁定输入列与输出列（用户说 or 推断空列）
2. 写首行公式（相对引用正确）
3. ApplyFormula fillDown 到数据行
4. Observe: 错误单元格、抽样值
5. 失败 → Repair 循环
```

### 5.2 常见错误 → 修复策略

| 观察 | 修复 |
|---|---|
| #REF! | 重映射列字母 |
| #DIV/0! | 包 IF/IFERROR |
| #VALUE! | 类型转换 VALUE()/文本清理 |
| #N/A | 查找键列/近似匹配 |
| 全空 | 填充范围偏移错误 |
| 循环引用 | 改输出列 |

### 5.3 公式观察指标

```json
"metrics": {
  "formulaErrorCount": 12,
  "sampleErrors": ["C2", "C5"],
  "filledRows": 198
}
```

`formulaErrorCount>0` 且目标为「正确计算」→ `VERIFY_FAILED` 可触发 repair。

---

## 6. 图表选型

与 SKILL 一致，固化为运行时规则：

| 数据形态 | 图表 |
|---|---|
| 时间序列 + 1 数值 | line |
| 类别 + 1 数值 | column/bar |
| 少类别占比 | pie（类别 ≤ 8） |
| 两数值相关 | scatter |
| 多系列类别 | clustered column |

插入位置：默认选区右侧或新图表页；`CreateChart` 参数必须含数据区域 ref。

Observe：`chart_count` delta ≥ 1。

---

## 7. 大表与分块

### 7.1 阈值（与 Safety 对齐）

| 规模 | 策略 |
|---|---|
| ≤ 5k 格 | 单步可写 |
| ≤ 50k | 分块 5k；进度事件 |
| > 50k | 强制分块 + ApprovalApproval 或拒绝全表格式化 |

### 7.2 分块算法

```text
rows chunked by N (e.g. 500)
for each chunk:
  toolCall with subrange
  observe chunk errors
  if fail: repair chunk only, 不重做已成功块
```

RunTrace 记录 `chunkIndex/chunkTotal`。

### 7.3 UI 线程

每块之间 `DoEvents` 或让出 UI（实现细节）；禁止单次 COM 锁死 > 2s 无反馈。

---

## 8. 结果落盘约定

| 输出类型 | 默认位置 |
|---|---|
| 汇总表 | 新 Sheet `汇总` / `分析_时间戳` |
| 清洗结果 | 优先新 Sheet，避免覆盖源（可配置） |
| 图表 | 当前 Sheet 或汇总 Sheet |
| 报告段落 | GenerateReport 目标 Sheet |

**禁止**默认静默覆盖用户唯一数据表；覆盖需 risky 确认。

---

## 9. Capability 与 Tool 地图

### 9.1 Capability 列表

| Id | 说明 | Tools |
|---|---|---|
| excel.detect-table | 表探测 | （内部 Reader） |
| excel.read | 结构化只读 | ReadRange |
| excel.profile | 画像 | DataAnalysis |
| excel.compute | 受控 JSON 计算 | PythonCompute |
| excel.clean | 清洗 | CleanData, RemoveDuplicates, FindReplace, FilterData |
| excel.formula | 公式 | ApplyFormula, WriteData |
| excel.transform | 变换 | TransformData, SortData |
| excel.pivot | 透视 | CreatePivotTable |
| excel.chart | 图 | CreateChart |
| excel.report | 报告 | CreateSheet, GenerateReport, WriteData |
| excel.format | 格式 | FormatRange, ConditionalFormat, AutoFit, MergeCells |
| excel.structure | 表结构 | InsertRowCol, DeleteRowCol, HideRowCol, CreateSheet, … |
| excel.protect | 保护 | ProtectSheet |
| excel.vba | 逃逸 | ExecuteVBA |

### 9.2 Tool → Capability（27）

| Tool | Capability |
|---|---|
| ApplyFormula | excel.formula |
| ReadRange | excel.read / profile |
| PythonCompute | excel.compute |
| WriteData | excel.formula / report |
| DataAnalysis | excel.profile |
| CleanData / RemoveDuplicates / FindReplace | excel.clean |
| SortData / FilterData | excel.clean / transform |
| TransformData | excel.transform |
| CreatePivotTable | excel.pivot |
| CreateChart | excel.chart |
| GenerateReport / CreateSheet | excel.report |
| FormatRange / ConditionalFormat / AutoFit / MergeCells | excel.format |
| InsertRowCol / DeleteRowCol / HideRowCol | excel.structure |
| CopySheet / DeleteSheet / RenameSheet | excel.structure |
| ProtectSheet | excel.protect |
| ExecuteVBA | excel.vba |

---

## 10. 与 ExcelDna UDF 隔离

| 路径 | 用途 |
|---|---|
| Agent/Harness | 多步、可观察、可撤销 |
| ExcelDna UDF | 单元格内同步短调用 |

**规则**

- UDF 不走 Skill/Loop。
- UDF 使用 `SendHttpRequestSync`（已 P0-3 桥接），有超时。
- 文档声明：UDF 不保证与 Agent 记忆/Trace 一致。

---

## 11. PowerQuery 策略

| 选项 | 决策 |
|---|---|
| v1 | **明确不做**完整执行；工具入口隐藏或返回 `HOST_UNSUPPORTED` 友好说明 |
| 文案 | 「当前版本请用清洗/公式/透视工具完成同类需求」 |
| 未来 | 独立 Capability `excel.powerquery` 专项 |

避免「开发中」占位出现在 Agent 主路径。

---

## 12. Observe 清单（Excel 专用）

每次写后至少：

1. 目标 `Excel:Sheet!Range`
2. `formulaErrorCount` delta（若相关）
3. 抽样 3 个单元格 before/after
4. chartCount / pivotCount delta（若相关）
5. UsedRange 尺寸 delta

---

## 13. 验收与 Golden

| Case | 断言 |
|---|---|
| U3-summary-chart | 汇总 sheet + chart≥1；无 VBA |
| E-formula-repair | error_count=0 |
| E-large-range-chunk | 有分块 event 或 approval |
| E-clean-dedupe | 行数减少且表头保留 |
| S-vba-disabled | ExecuteVBA deny |

---

## 14. 缺口优先级

| ID | 项 | P |
|---|---|---|
| E-GAP-1 | ExcelActionHarness 快路径（TableRegion Reader 已完成） | P1 |
| E-GAP-2 | 公式错误地址定位和自动 repair（数量 delta 已完成） | P1 |
| E-GAP-3 | 大表分块执行器 | P0 |
| E-GAP-4 | 结果默认写新 Sheet | P0 |
| E-GAP-5 | 拆巨类 ExcelDirectOperationService | P1 |
| E-GAP-6 | 透视字段智能推荐 | P1 |

---

## 15. 决策摘要（评审）

- [x] 同意 TableRegion 探测优先级
- [x] 同意管道模型（profile/clean/formula/…）
- [x] 同意公式 error 观察驱动 repair
- [x] 同意分块服从 Safety T1/T2（**D15**）
- [x] 同意结果默认新 Sheet、不静默覆盖
- [x] 同意 PowerQuery v1 **不做**（**D12**）
- [x] 同意 UDF 与 Agent 路径隔离

---

## 16. 落地顺序

1. TableRegion 探测进 ContextPack
2. profile/formula/chart 三条闭环 + Observe
3. Safety 阈值联动分块
4. ActionHarness 快路径
5. 清洗/透视/报告
6. 服务拆分

---

*Excel 运行时的核心资产是可靠的 TableRegion + 公式观察；没有这两者，Agent 只能「碰运气改格子」。*
