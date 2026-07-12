# 专项设计：Skill 运行时、两阶段加载与工具门禁

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 评审通过（见 [`design-review-record.md`](./design-review-record.md)） |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §5.3 |
| 关联 | [`safety-policy.md`](./safety-policy.md)、[`context-pack-schema.md`](./context-pack-schema.md)、[`tool-result-observation.md`](./tool-result-observation.md) |
| 现有代码 | `Skills/*`、`SkillRegistry`（JSON）、`SkillsDirectoryService`、`skills_registry` 表、`AgentKernel.SelectSkill*`、`ToolRegistry` |

---

## 1. 目标与非目标

### 1.1 目标

1. Skill 成为 Harness 的 **操作手册 + 工具边界**，不是关键词路由表。  
2. **两阶段加载**：召回只用 front matter；命中后再加载正文与 references。  
3. **硬门禁**：执行期 `toolId ∈ allowed-tools ∩ appTools ∩ Safety`。  
4. **退役 JSON Skill 主路径**，统一目录型 `SKILL.md`。  
5. 与 `skills_registry` 统计、评测、ContextPack.skills 对齐。  

### 1.2 非目标

- 不做云端 Skill 市场（预留目录配置）。  
- 不在本专项定义每个业务 Skill 的业务算法（见 excel/word/ppt runtime 专项）。  
- 不强制原生 LLM function-calling（可与 JSON Plan 双模共存）。  

### 1.3 原则

| 原则 | 说明 |
|---|---|
| 元数据便宜 | 冷启动只读 front matter |
| 边界硬 | 未声明工具 = 不可见 + 不可执行 |
| 宿主隔离 | `application` 不匹配则永不召回 |
| 可观测 | 命中/加载/越权均记 Trace |
| 渐进退役 | JSON 只读兼容一版，禁止新增 |

---

## 2. 现状与差距

| 组件 | 现状 | 问题 |
|---|---|---|
| 目录 Skill | 4 个 `*/SKILL.md` | 规范有，运行时门禁未硬 |
| JSON Skill | 5 个 `*.json` + `triggerPatterns` | 关键词路由，与 AI Native 冲突 |
| `SkillRegistry` | 只扫 JSON | 与目录双轨 |
| `SkillsDirectoryService` | front matter / detail | 需成为唯一真相源 |
| `skills_registry` | 索引与 usage | 需与门禁事件打通 |
| Kernel 选 Skill | filesystem 优先 + JSON 兜底 | 兜底应标 deprecated |
| Tool 执行 | 基本不校验 allowed-tools | **P0 缺口** |

---

## 3. Skill 包规范（平台标准）

### 3.1 目录

```text
ShareRibbon/Skills/<skill-name>/
  SKILL.md                 # 必须
  references/              # 可选，按需加载
    *.md
  scripts/                 # 可选 → 可注册为 skill_script.* 工具
  assets/                  # 可选
```

命名：`kebab-case`，与 front matter `name` 一致。

### 3.2 Front matter（机器可读）

```yaml
---
name: excel-table-agent
description: >
  Use for Excel table understanding, cleanup, formulas, charts,
  pivot, and multi-step spreadsheet automation.
application: Excel          # Excel | Word | PowerPoint | Common
tags: [excel, formula, chart]
intent_types: [data_analysis, formula, chart, data_clean]
allowed-tools:
  - ApplyFormula
  - WriteData
  - CreateChart
  - CleanData
  - DataAnalysis
  # ...
risk_default: medium        # 默认写操作风险提示
max_steps: 12
version: "1.0"
enabled: true
---
```

| 字段 | 必填 | 说明 |
|---|---|---|
| name | 是 | 唯一 id |
| description | 是 | **给模型的召回描述**，写清 when to use |
| application | 是 | 宿主隔离 |
| allowed-tools | 是 | 非空；可含 `common.*` 映射 |
| intent_types | 建议 | 与 TaskSpec 对齐 |
| tags | 建议 | 辅助召回 |
| risk_default | 否 | 默认 medium |
| max_steps | 否 | 覆盖 Loop 上限（取 min） |
| version | 否 | 语义版本 |
| enabled | 否 | 默认 true |

### 3.3 正文（人/模型操作手册）

必须包含（模板章节）：

1. When to use / When not to use  
2. Operating rules（先读后写、选区优先等）  
3. Tool preferences  
4. Observe & repair 提示  
5. 示例任务（2–5 条）  

**禁止**：把正文写成用户话术关键词列表；禁止假设未声明工具存在。

### 3.4 references 加载策略

| 触发 | 行为 |
|---|---|
| 召回阶段 | 不读 references |
| 命中且 Plan 复杂度 ≥ medium | 按 SKILL 内链加载列出的 ref |
| 单文件上限 | 8KB 或 2k tokens（预算内截断） |

---

## 4. 两阶段运行时

### 4.1 阶段 A — Index / Recall（便宜）

```text
启动或 Skills 变更
  → Scan 目录
  → 解析 front matter only
  → Upsert skills_registry
  → 内存 SkillIndex（name, desc, app, tags, allowed-tools, embedding?）

UserTurn + ContextPack
  → SkillRouter.Recall(query, appType, intentHints, topK=3)
  → 仅返回 SkillMeta[]
  → 写入 ContextPack.skills.recalled（detailLoaded=false）
```

**召回信号（加权建议）**

| 信号 | 权重 |
|---|---|
| application 匹配 | 硬过滤，不匹配直接 0 |
| description 语义/关键词 | 高 |
| intent_types ∩ TaskSpec | 高 |
| tags | 中 |
| usage_count / success_rate | 中 |
| 最近同文档成功 | 低 |

v1 可先用：硬过滤 application + 描述/意图打分；embedding 可选。

### 4.2 阶段 B — Load Detail（贵）

```text
选定 primarySkill（通常 Top1，或 Planner 指定）
  → Load SKILL.md 正文
  → 解析 references 列表并按需加载
  → 校验 allowed-tools ⊆ 已注册 tools
  → 得到 SkillDetailBundle
  → detailLoaded=true 记入 Trace
  → 注入 Planner / Loop 的 system 附加段（预算内）
```

**校验失败**

- 未知 tool 名 → 索引时报 `SkillValidationError`，该 Skill `enabled=false` 或降级告警  
- 正文缺失 → 不可选为 primary  

### 4.3 多 Skill

| 策略 | v1 |
|---|---|
| 同时绑定 | 仅 **1 个 primary** |
| 其它 recalled | 仅元数据提示「可选」 |
| 切换 | 仅 replan 时允许更换 primary |

---

## 5. 工具门禁（硬）

### 5.1 可见集合

```text
VisibleTools =
  Registry.ByApp(appType)
  ∩ (primarySkill?.AllowedTools  ??  AllAppTools)   // 无 Skill 时：仅 app 工具，仍受 Safety
  ∩ Health(available)
  - DeniedBySafetyPolicy
```

**Planner 与模型只能看见 VisibleTools 的 schema。**

### 5.2 执行集合

```text
Execute(toolCall):
  if primarySkill != null AND toolId ∉ AllowedTools:
      return Failed(TOOL_NOT_ALLOWED)
  if toolId ∉ Registry.ByApp:
      return Failed(HOST_UNSUPPORTED or NOT_FOUND)
  SafetyGate.Evaluate(...)
  Executor.Run(...)
```

### 5.3 与 JSON `requiredTools` 映射

旧字段 `RequiredTools` ≡ 新 `allowed-tools`（允许集，不是「必须全部用到」）。

命名统一后废弃 `RequiredTools` 字面，避免「必用」误解。

### 5.4 特殊工具

| 工具 | 规则 |
|---|---|
| `ExecuteVBA` | 必须在 allowed-tools **显式列出** + `safety.vbaEnabled` |
| `memory.*` | 可作为 common；是否进 Skill 由 Skill 声明 |
| `mcp.*` | 默认不进 Skill；需 `allowed-tools: [mcp.xxx]` 或策略 `allowMcp: true` |
| `skill_script.*` | 仅本 Skill scripts 下文件可暴露 |

### 5.5 越权观测

每次 `TOOL_NOT_ALLOWED`：

- AppLogger.Warn  
- RunTrace.events += skill_gate_denied  
- skills_registry 可记 fail（可选）  

---

## 6. JSON Skill 退役计划

### 6.1 当前 JSON 清单 → 目录迁移映射

| 旧 JSON | 建议迁入 | 说明 |
|---|---|---|
| data-analysis.json | excel-table-agent | 合并进现有目录 Skill 正文 |
| data-clean.json | excel-table-agent | 同上 |
| chart-creation.json | excel-table-agent | 同上 |
| batch-process.json | excel-table-agent 或新建 excel-batch-agent | 评估后定 |
| document-format.json | word-document-agent | 合并 |

### 6.2 阶段

| 阶段 | 行为 |
|---|---|
| **S0 现在** | 禁止新增 JSON Skill；文档标明 deprecated |
| **S1** | Router 仅目录；JSON 只读 fallback 且日志 `legacy_json_skill` |
| **S2** | 删除 JSON 加载代码与文件；smoke 断言目录-only |

### 6.3 兼容加载器

```text
If file ends with .json AND feature.LegacyJsonSkills:
    map to AgentSkill-like meta
    allowed-tools = requiredTools
    application = infer from tools path heuristic or "Excel"
Else ignore
```

---

## 7. 与 Harness 的集成点

```text
OfficeHarness.RunAsync
  ├ ContextHub.Snapshot
  ├ SkillRouter.Recall → metas
  ├ Planner.BuildSpec (可用 metas)
  ├ SkillRouter.LoadDetail(primary)
  ├ ToolBroker.SetVisible(AllowedTools ∩ app)
  ├ Planner.BuildPlan (仅 visible tools)
  └ Loop.Execute → Gate 每次校验
```

### 7.1 无 Skill 命中

- 不绑定 primary  
- VisibleTools = 全 app 工具（仍 Safety）  
- Plan 复杂度至少 exploratory，优先读工具  
- Trace 标记 `skill=none`  

### 7.2 max_steps

```text
effectiveMaxSteps = Min(Loop.Default, skill.max_steps, userConfig)
```

---

## 8. skills_registry 扩展（设计）

现有字段保留，建议增量：

| 列 | 说明 |
|---|---|
| allowed_tools_json | front matter 快照 |
| format | `directory` \| `legacy_json` |
| last_error | 校验错误 |
| fail_count | 执行失败 |
| gate_deny_count | 越权拒绝次数 |

迁移版本：实现时 `schema_version+1`，本专项只定字段语义。

---

## 9. 作者体验与校验

### 9.1 静态校验（smoke / CI）

对每个 `Skills/*/SKILL.md`：

1. front matter 可解析  
2. name == 目录名  
3. application ∈ 枚举  
4. allowed-tools 非空且 ⊆ Tools 注册表  
5. description 长度 ≥ 40 字符  

### 9.2 与 office-skill-authoring

该 Skill 负责教人写 Skill；平台校验失败时应指向 authoring 规则章节。

---

## 10. 状态机（SkillRouter）

```text
Idle
  → Indexing → Ready
Ready
  → Recall → Recalled
  → LoadDetail → DetailReady | DetailFailed
  → BindPrimary → Bound
Bound
  → (replan) Unbind → Recalled
```

---

## 11. 验收标准

1. 单元：未知 tool 在 allowed 中 → 索引校验失败。  
2. 单元：Bound Skill 下调用未声明 tool → `TOOL_NOT_ALLOWED`，无 COM。  
3. 集成：Excel 用户请求不会加载 `word-document-agent` 正文。  
4. 性能：冷启动只解析 front matter，正文文件未读（可用文件读计数断言）。  
5. 退役：S2 后仓库无 `Skills/*.json` 主路径加载。  
6. Trace：含 `recalled[]`、`primary`、`detailLoaded`、`gateDenies`。  

---

## 12. 开放问题（已冻结）

| # | 问题 | **冻结默认** |
|---|---|---|
| Q1 | 无 Skill 时是否收紧工具集？ | **D14**：v1 全 app 工具 + Safety；v2 再考虑只读收紧 |
| Q2 | 用户自定义 Skill 目录？ | 配置项 `SkillsExtraPaths`（可后做） |
| Q3 | allowed-tools 支持 glob？ | **否**，显式列表 |
| Q4 | 多 primary？ | **否**，仅 1 个 primary |

### JSON 退役节奏（冻结语义）

| 阶段 | 含义 | 建议触发 |
|---|---|---|
| S0 | 禁止新增 JSON Skill | **立即**（文档与 CR 约定） |
| S1 | 目录优先；JSON 只读 fallback + `legacy_json_skill` 日志 | H1 Skill 门禁落地时 |
| S2 | 删除 JSON 加载与仓库 JSON 文件 | 迁移完成 + smoke 绿后 |

---

## 13. 决策摘要（评审）

- [x] 同意两阶段加载（元数据 → 正文）  
- [x] 同意 allowed-tools **硬拒绝**（**D5**）  
- [x] 同意 application 硬隔离  
- [x] 同意无 Skill 时全 app 工具 + Safety（**D14**）  
- [x] 同意 JSON S0 立即、S1/S2 随 H1  

---

## 14. 落地顺序（实现时）

1. Front matter 校验 smoke  
2. ToolBroker 执行门禁  
3. Recall 仅目录  
4. JSON fallback 打标  
5. 合并 JSON 内容进目录 Skill  
6. 删除 JSON 加载  

---

*门禁与 Safety 同时失败时优先返回更具体的 code：先 `TOOL_NOT_ALLOWED`，再 `SAFETY_*`。*
