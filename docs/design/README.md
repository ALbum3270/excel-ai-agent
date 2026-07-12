# AI Native Harness · 专项设计索引

> 本目录存放 **目标合同级专项设计**。
> 总纲：[`../ai-native-harness-design.md`](../ai-native-harness-design.md)
> **落地入口**：[`../ai-native-harness-implementation-design.md`](../ai-native-harness-implementation-design.md)
> **评审记录（含设计决策 D1–D15）**：[`design-review-record.md`](./design-review-record.md)
> 原则：契约可序列化、可测试、可回放；`ShareRibbon` 只放抽象，COM 在宿主项目。

> 重要说明：本目录的“评审通过”表示目标设计可作为实现约束，不表示当前代码已经实现。当前实现差距以各专项文头的“实现状态”和落地入口文档为准。

## 评审结论（2026-07-11）

| 项 | 结论 |
|---|---|
| 总体 | **有条件同意 → 整改已落实（v0.2）** |
| 驳回项 | 无 |
| 编码 | 文档冻结后可启动 **Phase H0**（仅接口/适配器） |
| 权威冲突 | 以 [`design-review-record.md`](./design-review-record.md) 设计决策为准；实现顺序以落地入口为准 |

### 设计决策速查（D1–D15，目标态）

| ID | 一句话 |
|---|---|
| D1 | 唯一 `RunAsync` |
| D2 | 审批模式 B（非阻塞返回） |
| D3 | 同 session 单 running |
| D4 | VBA 默认关 |
| D5 | allowed-tools 硬拒绝 |
| D6 | memory_item 单写 |
| D7 | 短期不进 RAG |
| D8 | 写必有 observation |
| D9 | 快路径必 Trace |
| D10 | Stream = UI Pump；ReAct 在 Harness |
| D11 | JSON Plan 默认；toolCalling=auto |
| D12 | PowerQuery v1 不做 |
| D13 | PR=L0；L1 Nightly |
| D14 | 无 Skill → 全 app 工具 + Safety |
| D15 | Safety T1/T2 权威 |

共享 ErrorCode 表见评审记录 §4。

---

## 已完成专项（12 + 评审记录）

这里的“已完成”指专项文档已完成评审，不指代码已完成。

### 平台闭环

| 文档 | 版本 | 解决什么 |
|---|---|---|
| [office-harness-api.md](./office-harness-api.md) | v0.2 | IOfficeHarness / 事件 / WebView |
| [context-pack-schema.md](./context-pack-schema.md) | v0.2 | ContextPack |
| [skill-runtime-and-gates.md](./skill-runtime-and-gates.md) | v0.2 | Skill 门禁 |
| [safety-policy.md](./safety-policy.md) | v0.2 | SafetyGate |
| [tool-result-observation.md](./tool-result-observation.md) | v0.2 | Observe |
| [memory-runtime.md](./memory-runtime.md) | v0.2 | Memory |
| [model-gateway-streaming.md](./model-gateway-streaming.md) | v0.2 | Gateway |
| [run-trace-storage.md](./run-trace-storage.md) | v0.2 | RunTrace |
| [eval-golden-set.md](./eval-golden-set.md) | v0.2 | Golden |
| [office-object-operation-integration.md](./office-object-operation-integration.md) | v1.0 | 现有主链接入动态 API 发现与声明式 Office 对象操作 |

### 宿主运行时

| 文档 | 版本 | 解决什么 |
|---|---|---|
| [word-capability-map.md](./word-capability-map.md) | v0.2 | Word |
| [excel-table-agent-runtime.md](./excel-table-agent-runtime.md) | v0.2 | Excel |
| [ppt-deck-agent-runtime.md](./ppt-deck-agent-runtime.md) | v0.2 | PPT |

### 可选再加深

| 文档 | 主题 |
|---|---|
| `prompt-profile-system.md` | 阶段提示装配 |

---

## 评审顺序（已完成）

| 序 | 文档 | 结论 |
|---:|---|---|
| 0 | 总纲 | 同意 |
| 1–7 | API → … → Gateway | 同意 / 有条件同意（已改） |
| 8–9 | Trace / Golden | 同意（已改） |
| 10–12 | 三端 | 同意（已改） |

历史阅读顺序仍适用新同学 onboarding；**裁决以评审记录为准**。

---

## 评审约定

1. 先定契约与设计决策，再按落地入口分阶段编码。
2. 专项与评审记录冲突 → **评审记录优先**，并回写专项。
3. 新增 ErrorCode / 阈值必须更新评审记录 §4 与 Safety。

## 版本

| 版本 | 日期 | 说明 |
|---|---|---|
| 0.4 | 2026-07-11 | 12 专项齐 |
| **0.5** | **2026-07-11** | **统一评审 + md 整改；v0.2 专项；评审记录** |
| **0.6** | **2026-07-12** | **实现审计追加：P0 guardrails、Word fast-path 回灌、VSTO shadow-copy 工具定位、0 迭代防误报** |
| **0.7** | **2026-07-12** | **拆分目标设计与实现状态；新增落地实施设计入口** |
| **0.8** | **2026-07-12** | **新增声明式 Office 对象操作现有主链接入实施设计** |
