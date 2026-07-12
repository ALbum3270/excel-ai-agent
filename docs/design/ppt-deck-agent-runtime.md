# 专项设计：PowerPoint Deck Agent 运行时

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 目标设计已评审；代码部分实现 |
| 实现状态 | **部分实现**：PPT 已有工具 JSON、生成 Handler、翻译/续写服务和 ChatControl 命令分发；尚无统一 `PptActionHarness`、幻灯片级 Observation/Diff、删除/结构变更审批闭环。 |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §6.3 |
| Skill | `Skills/powerpoint-deck-agent/SKILL.md` |
| 现有 | Tools/ppt（22）、生成 Handler、ChatControl 命令分发、翻译/续写服务 |
| 关联 | ContextPack ppt_deck、Observe 幻灯片 Diff、Safety 删页、Golden U4 |

---

## 1. 目标与非目标

### 1.1 目标

1. 定义 PPT 从 **大纲 → 版式 → 填内容 → 美化 → 备注** 的标准管道。
2. 对标「一句话生成演示」体验，同时保持工具级可观察与可撤销。
3. 设计 `PptActionHarness` 与 Capability 地图。
4. 消灭「暂不支持的 PPT 命令」式死胡同，改为 repair/换工具。

### 1.2 非目标

- v1 不做设计师级自动配色网络模型训练。
- 不做嵌入式视频智能剪辑。
- 不保证跨主题模板像素级还原。

---

## 2. 运行时总览

```text
UserTurn (PowerPoint)
  → Context: slideCount, currentIndex, titles[], selection shapes
  → Skill: powerpoint-deck-agent
  → Planner 选择 pipeline:
       generate | rewrite | beautify | notes | chart_table | structure | review
  → 逐步工具（CreateSlides / Insert* / Format* / Beautify…）
  → Observe: slideCount, titles, shapeCount, notes
  → Repair（版式错误、空标题、溢出文本）
```

---

## 3. 演示文稿结构模型（DeckModel）

### 3.1 逻辑结构

```json
{
  "title": "项目季度汇报",
  "locale": "zh-CN",
  "slides": [
    {
      "role": "title | section | content | two_column | agenda | ending | blank",
      "title": "...",
      "bullets": ["...", "..."],
      "notes": "...",
      "visual": null | { "type": "chart|table|image", "hint": "..." },
      "layoutHint": "Title Slide | Title and Content | ..."
    }
  ]
}
```

### 3.2 角色 → 版式映射（默认 16:9）

| role | 默认 layoutHint |
|---|---|
| title | Title Slide |
| agenda / section | Section Header |
| content | Title and Content |
| two_column | Two Content |
| ending | Title Only / 空白+结束语 |
| blank | Blank |

若母版无对应 layout，降级为 Title and Content 并 warning。

---

## 4. 任务管道

| Pipeline | 步骤概要 | 主 Tools |
|---|---|---|
| **generate** | 大纲 → DeckModel → CreateSlides/InsertSlide 循环 → 填文本 → 可选美化 | CreateSlides, InsertSlide, InsertText, FormatText, BeautifySlides |
| **rewrite** | 当前页/选中形状改写 | FormatText, InsertText |
| **beautify** | 对齐、主题、字体层级 | BeautifySlides, ApplyTheme, FormatSlide, FormatText |
| **notes** | 讲者备注 | AddSpeakerNotes |
| **chart_table** | 插图/表 | InsertChart, InsertTable, InsertImage |
| **structure** | 增删移复制页 | InsertSlide, DeleteSlide, MoveSlide, DuplicateSlide |
| **motion** | 切换/动画（谨慎） | ApplyTransition, AddAnimation |
| **review** | 只读检查建议 | （读上下文 + 聊天解释，少写） |

### 4.1 generate 详细算法

```text
1. 输入源：用户文本 | 粘贴大纲 | Word 纪要引用 | 选中备注
2. LLM → DeckModel JSON（校验：slides 1..30，title 非空）
3. Safety：页数 > 15 → RequireApproval
4. 若空演示：CreateSlides 批量；若已有内容：确认插入位置（当前后/末尾）
5. For each slide in model:
     ensure slide exists (InsertSlide)
     SetSlideLayout(layoutHint)
     InsertText/FormatText 标题与要点
     optional visual
     AddSpeakerNotes
6. Observe 全稿：slideCount、每页 title 非空率
7. 空标题率 > 20% → Repair 补标题
8. 可选 BeautifySlides(scope=all|created)
```

### 4.2 页数策略

| 用户说 | 页数 |
|---|---|
| 明确 N 页 | N（clamp 3..30） |
| 「简短」 | 5–7 |
| 「详细汇报」 | 10–15 |
| 未说 | 由内容块数决定，默认 6–10 |

---

## 5. 美化启发式（Beautify）

不依赖「黑盒一次成功」，拆成可观察子目标：

| 子目标 | 观察 |
|---|---|
| 标题字号 > 正文字号 | 抽样 shape 字号 |
| 同级要点缩进一致 | |
| 左右边距不过密 | 形状 Left/Width 相对 slide |
| 主题一致 | ApplyTheme 一次 |
| 过渡不过度 | 仅切换，不默认狂乱动画 |

`BeautifySlides` 若过于原子化不足，运行时可展开为多个 Format* 步（Plan 可见）。

**动画**：默认不添加；用户明确要求才 `AddAnimation`（medium/risky 视范围）。

---

## 6. PptActionHarness（建议快路径）

| CapabilityId | 场景 | 置信 |
|---|---|---|
| `ppt.add-notes` | 「加备注」+ 当前页 | 高 |
| `ppt.duplicate-slide` | 「复制当前页」 | 高 |
| `ppt.apply-theme` | 「换成 XX 主题」 | 中高 |
| `ppt.delete-slide` | 「删掉这页」 | 高但 **risky 确认** |
| `ppt.generate-outline` | 明确「生成 N 页」 | 中 → 常走全 Loop |

快路径同样输出 ToolResult + Trace。

---

## 7. Capability 与 Tool 地图

### 7.1 Capability

| Id | 说明 |
|---|---|
| ppt.read-deck | 读目录/当前页（Context 为主） |
| ppt.generate | 生成页 |
| ppt.rewrite | 改写文本 |
| ppt.beautify | 美化 |
| ppt.notes | 备注 |
| ppt.visual | 图/表/图示/视频 |
| ppt.structure | 增删移复制 |
| ppt.theme | 主题与母版 |
| ppt.motion | 切换动画 |
| ppt.review | 审阅建议 |
| ppt.vba | ExecuteVBA |

### 7.2 Tool 映射（22）

| Tool | Capability | risk 基线 |
|---|---|---|
| CreateSlides | ppt.generate | medium |
| InsertSlide | ppt.generate / structure | medium |
| DeleteSlide | ppt.structure | **risky** |
| DuplicateSlide / MoveSlide | ppt.structure | medium |
| SetSlideLayout | ppt.generate / beautify | medium |
| InsertText / FormatText | ppt.rewrite / generate | medium |
| FormatSlide / BeautifySlides | ppt.beautify | medium |
| ApplyTheme | ppt.theme | medium |
| EditSlideMaster | ppt.theme | risky |
| InsertTable / InsertChart / InsertImage / InsertShape | ppt.visual | medium |
| InsertVideo | ppt.visual | medium |
| AddSpeakerNotes | ppt.notes | safe/medium |
| ApplyTransition / AddAnimation | ppt.motion | medium |
| SetSlideShow | ppt.structure | medium |
| ExecuteVBA | ppt.vba | risky |

---

## 8. Observe（PPT）

### 8.1 全局

- slideCount before/after
- 标题列表 hash

### 8.2 页级

```json
{
  "ref": "Ppt:Slide:3",
  "titleBefore": "",
  "titleAfter": "市场进展",
  "shapeCountDelta": 2,
  "notesLenDelta": 120
}
```

### 8.3 成功标准示例

| 任务 | 标准 |
|---|---|
| 生成 8 页 | slideCount 增加 8 或 =8（空稿） |
| 美化 | 用户未删页；标题仍在 |
| 备注 | notes 非空 |
| 删页 | count-1 且需审批 |

---

## 9. Repair 策略

| 观察 | 修复 |
|---|---|
| 标题空 | 根据 bullets 生成短标题再 InsertText |
| 要点溢出 | 拆页或降字号（先拆页） |
| 版式不存在 | 换 Title and Content |
| 插图失败 | 去掉 visual 降级纯文本 |
| 未知 tool 幻觉 | 用 VisibleTools 重选 |
| 旧路径「暂不支持」 | **禁止**；改为 NOT_FOUND + replan |

---

## 10. 内容安全与风格

- 默认商务简报语体；用户可 Memory 偏好。
- 每页要点建议 ≤ 5 条，每条 ≤ 40 字（生成时约束）。
- 敏感删页/清稿：Safety RequireApproval。

---

## 11. 与 Word/Excel 跨宿主

| 场景 | 策略 |
|---|---|
| Word 纪要 → PPT | 用户在 PPT 打开或引用文件；Context 读文件摘要（非跨进程 COM） |
| Excel 图 → PPT | 导出图片插入 or 用户复制；v1 可用 InsertImage 文件路径 |
| 真跨 COM | v2 连接器，不在本专项 |

---

## 12. 验收与 Golden

| Case | 断言 |
|---|---|
| U4-deck-from-notes | slide_count_gte；标题非空率≥0.8 |
| P-beautify-align | 页数不变；shapeCount 合理 |
| P-delete-confirm | 无批准则不删 |
| P-notes | notes 非空 |
| 未知命令 | 无「暂不支持」用户死胡同；有 replan/tool |

---

## 13. 缺口优先级

| ID | 项 | P |
|---|---|---|
| P-GAP-1 | DeckModel 校验与 generate 管道 | P0 |
| P-GAP-2 | 幻灯片 Diff 观察 | P0 |
| P-GAP-3 | 删除/结构变更审批 | P0 |
| P-GAP-4 | PptActionHarness | P1 |
| P-GAP-5 | ChatControl 瘦身迁 Executor | P1 |
| P-GAP-6 | 母版/主题稳健降级 | P1 |
| P-GAP-7 | 动画默认关闭策略 | P1 |

---

## 14. 决策摘要（评审）

- [x] 同意 DeckModel 为一等公民
- [x] 同意 generate 管道与页数 clamp 3..30
- [x] 同意 DeleteSlide **risky + 审批**
- [x] 同意动画默认不添加
- [x] 同意禁止「暂不支持」死胡同，改为 replan/换工具
- [x] 同意快路径同样写 Trace（**D9**）

---

## 15. 落地顺序

1. Context 目录 + 当前页快照
2. generate 管道（小页数）+ Observe
3. Safety 删页
4. beautify/notes
5. visual
6. ActionHarness 快路径
7. 未知命令路径清理

---

*PPT 的关键体验是「结构正确的多页故事」而不是单页堆形状；DeckModel 是一等公民。*
