# 声明式 Office 对象操作：现有主链接入实施设计

| 项 | 内容 |
|---|---|
| 状态 | 实施设计，作为编码拆分与验收入口 |
| 范围 | `ShareRibbon` 共享合同 + PowerPoint 首个垂直切片，随后迁移 Word/Excel |
| 目标 | 用动态发现和声明式对象操作覆盖长尾 Office 能力，不为每个需求新增意图或专用 Tool |
| 非目标 | 不新建第二套 Harness/Planner/Loop；不一次删除现有 Tool；不提供无约束任意 COM 反射 |

## 1. 现有代码基线

本方案必须复用当前已运行的唯一主链：

```text
BaseChatControl
  -> ChatRoutingOrchestrator
  -> AgentKernelService
  -> OfficeHarness
  -> AgentKernel
  -> LoopEngine
  -> ToolRegistry
  -> CodeExecutionService
  -> <Host>Ai.ChatControl.ExecuteJsonCommandWithToolResult
  -> Office COM
  -> ToolResult / Observation
  -> Loop observe / repair / explain
```

禁止在该链旁新增 `CapabilityHarness`、第二个 ToolBroker、第二个 ReAct Loop 或宿主直连入口。

### 1.1 可直接复用的现有能力

| 现有模块 | 本方案复用方式 |
|---|---|
| `IOfficeHarness` / `OfficeHarness` | 继续作为唯一产品执行入口和 run 状态机 |
| `AgentKernel` | 继续选择 Skill、装配 Prompt、启动 Loop |
| `LoopEngine` | 继续 plan/act/observe/repair/explain；消费发现结果和声明式执行结果 |
| `ToolRegistry` | 继续负责 appType、Skill allowed-tools、MCP/native 分发和 Safety 前置 |
| `CodeExecutionService` | 继续作为 native JSON 工具到宿主的唯一桥接 |
| `ExecuteJsonCommandWithToolResult` | 宿主唯一执行入口；新增通用命令但不复制通道 |
| `ToolResult` | 继续承载 Data、Observation、Artifacts、ErrorCode、Recoverable |
| `SafetyGate` | 扩展操作级规则，保持唯一安全裁决入口 |
| `ContextPack` | 继续承载规划/观察上下文；API Catalog 不整体塞入 Pack |
| `RunTraceStore` | 继续记录通用操作步骤和 Observation |

## 2. 冻结设计决策

### D16：只新增两个稳定工具

```text
DiscoverOfficeCapability   ' read-only，动态检索当前宿主 API/对象能力
OfficeObjectOperation      ' 执行声明式 OperationBatch
```

SmartArt、艺术字、组合形状、复杂表格等长尾需求通过这两个工具完成，不继续新增：

```text
SLIDE_SMARTART
InsertSmartArt
InsertWordArt
CreateTimeline
```

高频、强语义、需要专项质量算法的现有 Tool 继续保留，例如 `CreateSlides`、`ApplyFormula`、`GenerateTOC`。

### D17：发现结果按步返回，不注入全量 Catalog

Planner 只看两个稳定 Tool。需要长尾能力时计划为：

```text
step 1 DiscoverOfficeCapability
step 2 OfficeObjectOperation（使用 step 1 Observation/Data）
step 3 read/observe/verify
```

Catalog 每次最多返回 20 个相关成员。禁止把 Word/Excel/PPT 全 API 目录放进 system prompt。

### D18：Catalog 可发现不等于可执行

- Catalog 可以记录当前 Type Library 中的全部成员。
- Executor v1 只执行经过风险标注和参数绑定验证的成员。
- 未评级、外部调用、宏、进程控制成员默认不可执行。

### D19：通用执行必须走 Office UI 线程

`AgentKernel` 当前可在 `Task.Run` 中执行。所有宿主 COM 调用在进入 `ExecuteJsonCommandWithToolResult` 前必须统一 marshal 到宿主 UI 线程。该修复同时覆盖现有 native Tool，不为通用 Executor 建立私有线程通道。

### D20：现有 Tool 渐进转 adapter

旧 Tool 第一阶段保持原实现。通用 Executor 稳定后，逐个将其内部实现转换为 `OperationBatch` adapter；外部 Tool ID、Skill 和用户体验保持兼容。

## 3. 共享数据合同

新增目录：

```text
ShareRibbon/Agent/OfficeOperations/
├── OfficeObjectRef.vb
├── OfficeOperationModels.vb
├── OfficeCapabilityModels.vb
└── OfficeOperationValidation.vb
```

新增 `.vb` 后必须登记到 `ShareRibbon.vbproj`。

### 3.1 OfficeObjectRef

```vbnet
Public Class OfficeObjectRef
    Public Property AppType As String
    Public Property DocumentRef As String
    Public Property Path As String
End Class
```

Canonical 文本格式：

```text
Word:documents/active/paragraphs/12
Excel:workbooks/active/worksheets/销售/ranges/A1:F20
PowerPoint:presentations/active/slides/2/shapes/5
```

计划、Observation 和 RunTrace 中只能保存 Canonical Ref，不能跨 step 保存 COM RCW。

### 3.2 OfficeOperationBatch

```vbnet
Public Class OfficeOperationBatch
    Public Property SchemaVersion As String = "1.0"
    Public Property AppType As String
    Public Property Operations As New List(Of OfficeOperation)()
    Public Property Atomic As Boolean = True
    Public Property SuccessCriteria As New List(Of OperationCriterion)()
End Class

Public Class OfficeOperation
    Public Property Id As String
    Public Property TargetRef As String
    Public Property Action As String       ' get/set/invoke/create/delete/collection_item
    Public Property MemberId As String
    Public Property Arguments As JObject
    Public Property ExpectedEffects As JObject
End Class
```

### 3.3 CapabilitySearchResult

```vbnet
Public Class OfficeCapabilityMember
    Public Property MemberId As String
    Public Property DeclaringType As String
    Public Property MemberName As String
    Public Property MemberKind As String
    Public Property Parameters As New List(Of OfficeCapabilityParameter)()
    Public Property ReturnType As String
    Public Property RiskLevel As String
    Public Property Executable As Boolean
    Public Property UnsupportedReason As String
End Class
```

### 3.4 Observation

继续使用 `ToolResult.Observation`，通用操作约定：

```json
{
  "kind": "office_operation_batch",
  "summary": "创建 1 个 SmartArt，并写入 3 个节点",
  "changed": true,
  "targetRefs": ["PowerPoint:presentations/active/slides/1/shapes/6"],
  "operations": [
    {
      "id": "op-1",
      "status": "succeeded",
      "diff": {"shapeCountDelta": 1, "nodeCount": 3}
    }
  ],
  "warnings": []
}
```

## 4. ToolRegistry 接入

新增 Tool schema：

```text
ShareRibbon/Tools/ppt/DiscoverOfficeCapability.json
ShareRibbon/Tools/ppt/OfficeObjectOperation.json
```

后续 Word/Excel 使用同一 Tool ID 和不同 `appType`，由现有 `RegisterOrMergeTool` 合并宿主范围。

### 4.1 DiscoverOfficeCapability

参数：

| 参数 | 类型 | 必需 | 说明 |
|---|---|---:|---|
| query | string | 是 | 自然语言能力目标 |
| targetType | string | 否 | 已知对象类型，如 Shapes/Range/Paragraphs |
| includeReadOnly | boolean | 否 | 是否包含只读成员，默认 true |
| maxResults | integer | 否 | 默认 12，最大 20 |

风险固定为 `safe`，不得产生 COM 写副作用。

### 4.2 OfficeObjectOperation

参数：

| 参数 | 类型 | 必需 | 说明 |
|---|---|---:|---|
| batch | object | 是 | `OfficeOperationBatch` |

Tool JSON 风险基线为 `medium`；最终风险由 `SafetyGate` 根据 batch 内 action、member、影响范围上调。

### 4.3 Skill 接入

三个 baseline Skill 增加：

```yaml
allowed-tools: ..., DiscoverOfficeCapability, OfficeObjectOperation
```

专用 Skill 可只开放发现工具或只开放有限高层 Tool。`allowed-tools` 执行期硬门保持不变。

## 5. CodeExecutionService 与 UI 线程

现有路径：

```text
ToolRegistry.ExecuteToolAsync
  -> ExecuteCodeWithToolResult(command, "json", false)
  -> CodeExecutionService.ExecuteJsonCommandWithToolResult
  -> host ExecuteJsonCommandWithToolResult
```

保持不变。

### 5.1 必做线程修复

在 `BaseChatControl` 创建 `AgentKernelService`/`CodeExecutionService` 回调时，通过现有 `UiDispatcher` 增加返回值版本：

```vbnet
UiDispatcher.InvokeSync(Of Agent.ToolResult)(control, Function() ...)
```

所有宿主 COM JSON 命令统一从这里进入 UI 线程。禁止每个 Executor 自己 `Control.Invoke`。

验收：Debug 下记录执行线程 ID；同一次 run 的所有 COM 写操作均为宿主 UI 线程。

## 6. PowerPoint 首个垂直切片

新增：

```text
PowerPointAi/Runtime/
├── PowerPointApiCatalogProvider.vb
├── PowerPointObjectResolver.vb
├── PowerPointOperationExecutor.vb
└── PowerPointOperationObserver.vb
```

全部登记到 `PowerPointAi.vbproj`。

### 6.1 ChatControl 接入点

在 `PowerPointAi.ChatControl.ExecuteJsonCommandWithToolResult` 顶部解析 envelope：

```vbnet
Select Case toolId.ToLowerInvariant()
    Case "discoverofficecapability"
        Return PowerPointApiCatalogProvider.SearchAsToolResult(params)
    Case "officeobjectoperation"
        Return PowerPointOperationExecutor.Execute(params)
End Select

' 未命中继续现有 ExecuteJsonCommandCore
```

只增加这两个稳定入口。后续 API 成员不得继续增加 `Select Case` 分支。

### 6.2 PowerPointApiCatalogProvider

数据源：

1. `Microsoft.Office.Interop.PowerPoint` assembly reflection。
2. `Microsoft.Office.Core` assembly reflection。
3. 当前安装 Office/WPS Type Library（第二阶段）。

Provider 负责：

- 枚举类型、方法、属性、参数、枚举；
- 生成稳定 `MemberId`；
- 中文/英文别名索引；
- 风险基线；
- executable 状态；
- 内存缓存和版本 fingerprint。

第一版只给以下类型 `Executable=True`：

```text
Presentation, Slides, Slide, Shapes, Shape,
TextFrame, TextFrame2, TextRange, SmartArt, SmartArtNodes, SmartArtNode
```

其余成员可被发现，但返回 `Executable=False` 和原因。

### 6.3 PowerPointObjectResolver

职责：

- 解析 active presentation、slide、shape、collection item；
- 检查索引边界和对象类型；
- 返回 `ResolvedOfficeObject`，记录需释放的临时 COM 对象；
- operation 结束立即释放，禁止写入 Agent memory/session。

错误码：

```text
OBJECT_REF_INVALID
OBJECT_NOT_FOUND
OBJECT_TYPE_MISMATCH
DOC_MISSING
```

### 6.4 PowerPointOperationExecutor

固定支持：

```text
get
set
invoke
create
delete
collection_item
```

执行前顺序：

1. 反序列化并校验 batch schema。
2. 校验 `AppType=PowerPoint`。
3. Catalog 查找 MemberId。
4. 检查 `Executable=True`。
5. 参数名称、数量、类型和可选参数绑定。
6. SafetyGate operation-level 裁决。
7. before snapshot。
8. COM 调用。
9. after snapshot / diff。
10. 生成 ToolResult。

第一版不得支持任意字符串对象路径或 `CallByName` 绕过 Catalog。

### 6.5 SmartArt 首个 Golden

用户目标：

```text
在当前页创建一个三阶段 SmartArt，文字为需求分析、开发实施、测试上线
```

期望 Loop：

1. `DiscoverOfficeCapability(query=SmartArt...)`
2. `OfficeObjectOperation` 创建 SmartArt。
3. `OfficeObjectOperation` 写入三个节点文字，或同一 atomic batch 完成。
4. Observer 验证 shape type、node count、node text hash。
5. 不满足则最多 repair 两轮。

禁止新增 `SLIDE_SMARTART` 和 `InsertSmartArt`。

## 7. SafetyGate 扩展

仍修改现有 `ShareRibbon/Agent/Execution/SafetyGate.vb`，不新增平行安全服务。

| 操作 | 默认裁决 |
|---|---|
| get / 只读 property | Allow |
| set / create /普通 invoke | Allow 或 medium |
| delete / clear / remove | RequireApproval |
| Quit / Close / SaveAs 覆盖 | RequireApproval 或 Deny |
| VBProject / macro / Shell / 外部进程 | Deny |
| 未评级 MemberId | Deny |
| 无法计算影响范围 | RequireApproval |
| 跨宿主 ref/member | Deny `HOST_UNSUPPORTED` |

Safety decision 必须写入现有 RunTrace step。审批后只消费一次现有 `ApprovedTools` 授权。

## 8. Observe、Verify 与 Repair

### 8.1 Observer 复用

PowerPoint 第一版复用并迁移当前 `ChatControl` 中已有：

```text
CapturePowerPointCommandSnapshot
BuildPowerPointCommandObservation
GetPowerPointSlideText
GetPowerPointNotesText
```

先搬方法，不重写算法。新增 SmartArt 指标：

```text
shapeType
smartArtLayoutId
nodeCount
nodeTextHash
```

### 8.2 Verifier

成功必须同时满足：

- operation ToolResult.Success；
- Observation.changed；
- TaskSpec.SuccessCriteria 必需项通过；
- 无未处理的 partial apply；
- 用户要求的实际产物存在。

原子调用成功不等于 run 成功。

### 8.3 Repair

继续使用 `LoopEngine` 现有 repair/replan。新增错误码：

```text
CAPABILITY_NOT_FOUND
MEMBER_NOT_EXECUTABLE
OBJECT_REF_INVALID
OBJECT_NOT_FOUND
OBJECT_TYPE_MISMATCH
OPERATION_SCHEMA_INVALID
VERIFY_FAILED
PARTIAL_APPLY
```

可恢复：对象 ref、参数、布局、节点数、类型转换。

不可恢复：权限、受保护文档、Safety Deny、宿主不支持、同错误连续两次。

## 9. 迁移策略

### Phase O0：接缝与线程（不改变业务）

文件范围：

```text
ShareRibbon/Agent/OfficeOperations/*
ShareRibbon/Tools/*/DiscoverOfficeCapability.json
ShareRibbon/Tools/*/OfficeObjectOperation.json
ShareRibbon/Controls/BaseChatControl.vb
ShareRibbon/Controls/Services/CodeExecutionService.vb
三个 baseline SKILL.md
```

验收：新工具经过现有 Harness/ToolRegistry/Safety/ToolResult；COM 回调在 UI 线程；旧 Tool 回归通过。

### Phase O1：PPT 只读发现

实现 Catalog Provider 和 `DiscoverOfficeCapability`。不开放写操作。

验收：SmartArt 查询召回相关成员；普通问答不进入执行；全量 Catalog 不进入 Prompt。

### Phase O2：PPT SmartArt 垂直切片

实现 Resolver、Executor、Observer 首批类型。

验收：SmartArt Golden 端到端通过，无新增专用意图/Tool。

### Phase O3：PPT 扩面

顺序：

```text
Shape/Text -> Table -> Chart -> Animation -> Theme/SlideMaster
```

现有 `InsertShape`、`InsertText` 等逐步改成 Operation adapter。

### Phase O4：Word

新增：

```text
WordAi/Runtime/WordApiCatalogProvider.vb
WordAi/Runtime/WordObjectResolver.vb
WordAi/Runtime/WordOperationExecutor.vb
WordAi/Runtime/WordOperationObserver.vb
```

首批对象：Document、Selection、Range、Paragraph、Style、Table。

### Phase O5：Excel

新增：

```text
ExcelAi/Runtime/ExcelApiCatalogProvider.vb
ExcelAi/Runtime/ExcelObjectResolver.vb
ExcelAi/Runtime/ExcelOperationExecutor.vb
ExcelAi/Runtime/ExcelOperationObserver.vb
```

首批对象：Workbook、Worksheet、Range、ListObject、ChartObject、PivotTable。

### Phase O6：旧 Tool adapter 与收敛

- 统计专用 Tool 与通用 Operation 的成功率。
- 达到同等 Golden 质量后才迁移内部实现。
- 外部 Tool ID 保持兼容至少一个发布周期。
- 删除入口分支前先更新 Skill、Prompt、Golden 和迁移说明。

## 10. 测试与发布门禁

### L0：纯合同

- Operation JSON 往返。
- Canonical Ref 解析。
- Catalog MemberId 稳定性。
- 参数类型绑定。
- Safety 矩阵。
- 禁止成员列表。

### L1：Fake Host

新增内存对象图，验证：

```text
discover -> operation -> snapshot -> diff -> verify -> repair
```

### L2：真实 Office Smoke

PowerPoint 首批：

- SmartArt 创建和节点写入；
- 普通 Shape/Text；
- 删除必须审批；
- 错误 ref 自动 repair；
- UI 终态恢复；
- COM 操作线程正确。

### 发布指标

| 指标 | 门槛 |
|---|---:|
| 假成功率 | 0 |
| UI 终态恢复率 | 100% |
| 未审批危险副作用 | 0 |
| SmartArt Golden 成功率 | 100% |
| 可恢复参数/ref 错误自动恢复率 | ≥ 80% |
| 同错误无限循环 | 0 |
| 全量 API 注入 Prompt | 0 |

## 11. 提交拆分建议

| Commit | 内容 |
|---|---|
| O0-1 | shared Operation DTO + schema tests |
| O0-2 | 两个 Tool JSON + Skill allowed-tools |
| O0-3 | COM UI thread single entry |
| O1-1 | PPT Catalog Provider + discovery ToolResult |
| O2-1 | PPT Resolver + read operations |
| O2-2 | PPT write Executor + Safety |
| O2-3 | SmartArt Observer/Verifier/Golden |

每个 commit 必须可构建、可回滚，不允许一次同时改三个宿主。

## 12. Definition of Done

本方案第一里程碑完成的唯一判据：

> PowerPoint 用户以任意自然语言要求创建含三段文字的 SmartArt；系统通过 `DiscoverOfficeCapability` 动态发现成员，通过 `OfficeObjectOperation` 声明式执行，通过现有 Safety/ToolResult/Loop/RunTrace 完成观察和最多两轮修复；整个实现不新增 SmartArt 专用意图或 Tool，并且旧 PPT Tool 回归不受影响。
