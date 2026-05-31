# Word智能排版系统改进设计文档

## 一、现状诊断：为什么当前排版"愚蠢"

### 1.1 根本原因

当前系统的核心架构思路是对的——LLM理解文档、规则应用格式。但执行层面存在三个致命问题：

**问题A：LLM被当分类器用，没有发挥理解能力**

当前提示词：
```
你是一个严格的JSON输出器。你必须只输出一个JSON数组，不要输出任何其他内容。
```

这等于告诉LLM："你是个打标签机器，不要思考。" 结果LLM只能机械匹配模式，遇到模糊段落就退化成 `body.normal`。

**问题B：LLM看不到完整文档，只能看到碎片**

当前做法：只发送每个段落的前120字符，不告诉LLM段落之间的关系。

LLM看到的：
```
[0] [样式:正文] [格式:16pt] ×政发〔2024〕15号
[1] [样式:正文] [格式:16pt 加粗] 关于加强安全生产工作的通知
```

LLM不知道：[0]是发文字号（因为它在标题前面），[1]是文件标题（因为它紧跟发文字号）。它只能靠文本模式匹配来猜。

**问题C：排版流程中LLM参与太晚、太浅**

当前流程：
```
用户说"帮我排版" → 纯规则分析 → 生成预览卡 → 用户点"应用" → 规则直接应用
                                                    ↓ (可选)
                                              AI语义标注 → 验证 → 应用
```

AI语义标注是可选的、后置的，而正文段落默认标记为 `""`（AI待标注），实际上经常直接退化成 `body.normal`。

### 1.2 具体问题清单

| 位置 | 问题 | 影响 |
|------|------|------|
| `SemanticPromptBuilder` | 提示词自称"严格JSON输出器"，抑制LLM推理 | 分类质量差 |
| `SemanticPromptBuilder` | 示例中包含格式参数（字号、字体），但要求不输出格式参数 | LLM混淆 |
| `SemanticPromptBuilder` | 中文排版规则（行首标点等）无法通过JSON标签实现 | 浪费token |
| `SemanticPromptBuilder` | 只发前120字符 | 丢失关键上下文 |
| `SemanticPromptBuilder` | 禁止输出理由/推理过程 | 无法解释为何如此标注 |
| `ProofreadPromptBuilder` | 没有强制JSON-only约束 | 输出格式不稳定 |
| `ProofreadPromptBuilder` | "的/在/得/地混用"是笔误 | 误导LLM |
| `ProofreadPromptBuilder` | 不了解文档的排版标准 | 校对只看文字不管格式 |
| `FormatMirrorService` | 只采样前80段落，只发送前20组 | 长文档格式丢失 |
| `FormatMirrorService` | 可用标签只有9个 | 公文等复杂格式无法克隆 |
| `FormatMirrorService` | grouping key忽略字体名和缩进 | 不同样式被合并 |
| `SmartFormattingOrchestrator` | 正文段落NewTag="" | 正文全变body.normal |
| `SmartFormattingOrchestrator` | 对话微调是规则匹配（关键词→属性调整） | 无法理解自然语言意图 |
| `DocumentAnalyzer` | `Math.Min(a,b) - 1 - 1` 循环边界多减了1 | 漏掉最后一段分析 |
| `UndoStack` | MaxSize裁剪逻辑删最新的而非最旧的 | 撤销行为错误 |

---

## 二、核心设计原则

### 2.1 LLM是文档分析师，不是标签机

**错误做法：** 让LLM做机械分类——"这个段落是什么标签？"

**正确做法：** 让LLM做文档理解——"这篇文档的结构是什么？每个段落扮演什么角色？为什么？"

LLM的真正价值不在于识别"第X章"是标题1，而在于理解：
- 一段没有明显标记的文字，在上下文中是署名还是正文
- "特此通知"是结束语还是正文的一部分
- 某段文字虽然在标题位置，但内容是日期，应该标为日期而非标题

### 2.2 规则定义"正确格式"，LLM决定"段落角色"

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  LLM理解层  │     │  规则验证层  │     │  格式应用层  │
│             │     │             │     │             │
│ 分析文档结构 │────→│ 映射：角色→  │────→│ 精确应用    │
│ 识别段落角色 │     │     格式规格 │     │ Word格式    │
│ 解释推理过程 │     │ 验证分类合理性│     │ 页面设置    │
│ 处理模糊情况 │     │ 修正明显错误 │     │ 撤销支持    │
└─────────────┘     └─────────────┘     └─────────────┘
```

- **LLM做它擅长的事**：理解文本、推理结构、处理歧义、解释原因
- **规则做它可靠的事**：存储格式规格（字号、字体、行距等精确值）、验证分类合理性、精确应用格式

### 2.3 对话是关于理解的对话，不是关于参数的对话

**当前：** 用户说"标题再大一点" → 关键词匹配"标题"+"大" → 属性调整字号+2pt

**目标：** 用户说"这个标题不够突出" → LLM理解"突出"可能意味着加大字号、加粗、或改颜色 → 结合当前格式和标准给出建议 → 用户确认后应用

---

## 三、具体改进方案

### 3.1 Bug修复（5项）

#### 3.1.1 DocumentAnalyzer循环边界bug

**文件：** `ShareRibbon/Services/Reformat/DocumentAnalyzer.vb`

**问题：**
```vb
For i = 0 To Math.Min(paragraphs.Count - 1, paragraphStyles.Count - 1) - 1
```
`Math.Min` 已经各减了1，再减1导致漏掉最后一个段落。

**修复：**
```vb
For i = 0 To Math.Min(paragraphs.Count, paragraphStyles.Count) - 1
```

#### 3.1.2 UndoStack裁剪逻辑bug

**文件：** `ShareRibbon/Undo/UndoStack.vb`

**问题：** 遍历Stack时是LIFO顺序，Skip(1)跳过的是最新操作而非最旧操作。

**修复方案：** 改用Queue或在Push时检查Count并移除底部元素：
```vb
Public Sub Push(operation As UndoableOperation)
    _undoStack.Push(operation)
    _redoStack.Clear()

    ' 超过最大容量时移除最旧的操作
    While _undoStack.Count > _maxSize
        Dim tempList = _undoStack.ToList()
        tempList.RemoveAt(tempList.Count - 1) ' 移除最旧（栈底）
        _undoStack.Clear()
        For Each item In tempList
            _undoStack.Push(item)
        Next
    End While
End Sub
```

#### 3.1.3 GB/T 9704页边距单位问题

**文件：** `ShareRibbon/Services/Reformat/FormattingStandardData.vb`

**问题：** 当前用cm值设置PageMargins，需确认Word对象模型期望的单位。Word的PageSetup margins属性使用磅(pt)为单位。1cm ≈ 28.35pt。

**验证并修复：** 检查 `SemanticRenderingEngine` 中应用页边距时的单位转换是否正确。如果直接把cm值赋给pt属性，则需乘以28.35。

#### 3.1.4 扩展FormattingStandardData模板

**文件：** `ShareRibbon/Services/Reformat/FormattingStandardData.vb`

在现有4套标准基础上，新增：
- 简历格式标准（含语义标签：`title.name`、`body.summary`、`section.education`、`section.experience`等）
- 合同格式标准（含语义标签：`title.contract`、`party.a`、`party.b`、`clause.main`等）
- 通用信函标准（含语义标签：`header.sender`、`header.addressee`、`body.letter`、`footer.closing`等）

保持 `SemanticStyleMapping` 数据模型不变，只增加标准和标签定义。

#### 3.1.5 SemanticRenderingEngine增强撤销支持

**文件：** `ShareRibbon/Services/Reformat/SemanticRenderingEngine.vb`

**当前：** 格式化通过Word原生UndoRecord包装，但快照恢复功能 `RestoreFormatSnapshot` 没有被调用。

**改进：**
1. 在 `ApplySemanticFormatting` 之前自动调用 `CaptureFormatSnapshot`
2. 提供明确的"撤销上次排版"入口，调用 `RestoreFormatSnapshot`
3. 快照数据存储到 `RefinementContext` 中，支持对话级别的撤销

---

### 3.2 提示词重设计（核心改进）

#### 3.2.1 语义标注提示词重设计

**设计理念：** 从"标签分类器"变为"文档结构分析师"

**新提示词架构：**

```
角色定义 → 文档全貌 → 分析步骤 → 输出格式 → 标签说明 → 段落列表
```

**关键改变：**

1. **角色重新定义**
   - 旧："你是一个严格的JSON输出器"
   - 新："你是一位文档结构分析专家，擅长识别中文文档的结构和语义角色"

2. **提供文档全貌**
   - 发送完整段落文本（不截断120字符）
   - 标注段落位置信息："文档开头"、"文档中部"、"文档末尾"
   - 标注相邻段落关系（上下文线索）

3. **分步思考**
   - 第一步：识别文档类型（公文/论文/报告/其他）
   - 第二步：识别文档整体结构
   - 第三步：逐段标注语义角色
   - 输出包含reasoning字段，记录关键判断依据

4. **移除无关指令**
   - 删除中文排版规则（行首标点等），这些不是语义标注的职责
   - 删除示例中的格式参数（字号、字体），避免混淆
   - 精简禁止列表，只保留最必要的

5. **示例重写**
   - 旧：`「XX市人民政府文件」 → header.org（发文机关标志，居中，方正小标宋22pt红色）`
   - 新：`「XX市人民政府文件」 → header.org（理由：位于文档最开头，符合发文机关标志的文本模式）`

**新提示词模板（核心部分）：**

```vb
Public Shared Function BuildSemanticTaggingPrompt(
    mapping As SemanticStyleMapping,
    paragraphs As List(Of String),
    Optional paragraphStyles As List(Of String) = Nothing,
    Optional originalParaIndices As List(Of Integer) = Nothing,
    Optional detectedHeadings As String = Nothing,
    Optional documentTypeContext As String = Nothing,
    Optional paragraphFontSizes As List(Of Single) = Nothing,
    Optional paragraphIsBold As List(Of Boolean) = Nothing) As String

    Dim sb As New StringBuilder()

    ' ===== 1. 角色定义 =====
    sb.AppendLine("你是一位文档结构分析专家，擅长识别中文文档的结构和语义角色。")
    sb.AppendLine("你的任务是分析文档内容，识别每个段落的语义角色，以便系统自动应用对应的标准格式。")
    sb.AppendLine()

    ' ===== 2. 任务说明 =====
    sb.AppendLine("【任务】")
    sb.AppendLine("请按以下步骤分析文档：")
    sb.AppendLine("步骤1：判断文档类型和整体结构")
    sb.AppendLine("步骤2：识别文档中的关键结构元素（标题、正文、署名、日期等）")
    sb.AppendLine("步骤3：为每个段落分配合适的语义标签")
    sb.AppendLine()

    ' ===== 3. 文档类型上下文 =====
    If Not String.IsNullOrEmpty(documentTypeContext) Then
        sb.AppendLine("【排版标准】")
        sb.AppendLine(documentTypeContext)
        sb.AppendLine()
    End If

    ' ===== 4. 输出格式 =====
    sb.AppendLine("【输出格式】")
    sb.AppendLine("请输出纯JSON数组，格式如下：")
    sb.AppendLine("[")
    sb.AppendLine("  {""paraIndex"":0, ""tag"":""header.org"", ""reason"":""位于文档开头，文本符合发文机关标志模式""},")
    sb.AppendLine("  {""paraIndex"":1, ""tag"":""header.refno"", ""reason"":""包含发文字号格式〔20XX〕X号""},")
    sb.AppendLine("  ...")
    sb.AppendLine("]")
    sb.AppendLine()
    sb.AppendLine("要求：")
    sb.AppendLine("- 只输出JSON数组，不要输出其他内容")
    sb.AppendLine("- 不要使用markdown代码块包裹")
    sb.AppendLine("- reason字段简短说明判断依据（不超过30字）")
    sb.AppendLine("- 每个段落必须且只能有一个标签")
    sb.AppendLine()

    ' ===== 5. 可用标签 =====
    sb.AppendLine("【可用语义标签】")
    For Each tag In mapping.SemanticTags
        sb.Append($"- {tag.TagId}: {tag.DisplayName}")
        If Not String.IsNullOrEmpty(tag.MatchHint) Then
            sb.Append($"。识别提示：{tag.MatchHint}")
        End If
        sb.AppendLine()
    Next
    sb.AppendLine()

    ' ===== 6. 文档结构识别指南 =====
    sb.AppendLine("【结构识别指南】")
    sb.AppendLine("判断段落角色时，请综合考虑以下线索：")
    sb.AppendLine("1. 文本内容：是否包含特定模式（发文字号、日期、编号等）")
    sb.AppendLine("2. 段落位置：在文档开头、中间还是末尾")
    sb.AppendLine("3. 上下文关系：与前后段落的关系（如标题后面通常跟正文）")
    sb.AppendLine("4. 格式线索：字号偏大且加粗的短段落通常是标题")

    ' 公文特殊规则
    If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
       (documentTypeContext.Contains("公文") OrElse documentTypeContext.Contains("GB/T 9704")) Then
        sb.AppendLine()
        sb.AppendLine("【公文结构识别】")
        sb.AppendLine("公文文档有固定的结构顺序，请按此顺序识别：")
        sb.AppendLine("发文机关标志(header.org) → 发文字号(header.refno) → 签发人(header.signer)")
        sb.AppendLine("→ 文件标题(title.main) → 主送机关(title.recipient) → 正文(body.normal)")
        sb.AppendLine("→ 附件说明(body.attachment) → 发文机关署名(footer.signature)")
        sb.AppendLine("→ 成文日期(footer.date) → 抄送机关(footer.cc)")
        sb.AppendLine()
        sb.AppendLine("注意：公文中的标题层级使用title.1/title.2/title.3，不要使用heading.*")
    End If
    sb.AppendLine()

    ' ===== 7. 自动检测结果（参考） =====
    If Not String.IsNullOrEmpty(detectedHeadings) Then
        sb.AppendLine("【系统自动检测到的标题结构（仅供参考，你可以修正）】")
        sb.AppendLine(detectedHeadings)
        sb.AppendLine()
    End If

    ' ===== 8. 文档段落（完整文本+上下文） =====
    sb.AppendLine("【文档段落】")
    Dim hasStyles = paragraphStyles IsNot Nothing AndAlso paragraphStyles.Count = paragraphs.Count
    Dim hasOrigIdx = originalParaIndices IsNot Nothing AndAlso originalParaIndices.Count = paragraphs.Count
    Dim hasFontSizes = paragraphFontSizes IsNot Nothing AndAlso paragraphFontSizes.Count = paragraphs.Count
    Dim hasBold = paragraphIsBold IsNot Nothing AndAlso paragraphIsBold.Count = paragraphs.Count

    ' 计算位置信息
    Dim nonEmptyCount = paragraphs.Count(Function(p) Not String.IsNullOrWhiteSpace(p))

    For i = 0 To paragraphs.Count - 1
        Dim origIdx = If(hasOrigIdx, originalParaIndices(i), i)
        Dim text = paragraphs(i)
        If String.IsNullOrWhiteSpace(text) Then Continue For

        ' 位置标签
        Dim positionLabel = ""
        If i = 0 Then
            positionLabel = " [文档开头]"
        ElseIf i >= paragraphs.Count - 3 Then
            positionLabel = " [文档末尾]"
        End If

        ' 样式提示（简洁）
        Dim styleHint As String = ""
        If hasStyles AndAlso Not String.IsNullOrEmpty(paragraphStyles(i)) Then
            styleHint = $" [样式:{paragraphStyles(i)}]"
        End If

        ' 格式线索（简洁）
        Dim formatHint As String = ""
        If hasFontSizes Then
            formatHint = $" {paragraphFontSizes(i):F0}pt"
        End If
        If hasBold AndAlso paragraphIsBold(i) Then
            formatHint &= " 加粗"
        End If
        If formatHint <> "" Then
            formatHint = $" [格式:{formatHint.Trim()}]"
        End If

        ' 上下文：显示前一段落的最后20字
        Dim contextBefore As String = ""
        If i > 0 AndAlso Not String.IsNullOrWhiteSpace(paragraphs(i - 1)) Then
            Dim prevText = paragraphs(i - 1).Trim()
            If prevText.Length > 20 Then prevText = "..." & prevText.Substring(prevText.Length - 20)
            contextBefore = $"  ↑上文: {prevText}" & vbCrLf
        End If

        ' 不截断段落文本，但超长段落只取前300字+后缀
        If text.Length > 300 Then
            text = text.Substring(0, 300) & $"...[全文{text.Length}字]"
        End If

        sb.Append(contextBefore)
        sb.AppendLine($"[{origIdx}]{positionLabel}{styleHint}{formatHint} {text}")
    Next

    Return sb.ToString()
End Function
```

**与旧版的关键差异：**

| 方面 | 旧版 | 新版 |
|------|------|------|
| 角色 | 严格JSON输出器 | 文档结构分析专家 |
| 输出 | `{paraIndex, tag}` | `{paraIndex, tag, reason}` |
| 段落长度 | 120字符 | 300字符+长度提示 |
| 上下文 | 无 | 显示上文片段+位置标签 |
| 示例 | 含格式参数 | 只含识别理由 |
| 排版规则 | 包含（但无法执行） | 移除 |
| 禁止列表 | 7条 | 3条（精简） |
| 文档类型 | 可选提示 | 强制步骤1先判断 |

#### 3.2.2 校对提示词重设计

**关键改进：**

1. 强化JSON-only约束
2. 修复"的/在/得/地混用"笔误
3. 增加文档标准上下文（如果已应用排版标准，按标准校对格式）
4. 增加最小修改原则
5. 精简issueType大小写一致性

**新提示词核心改动：**

```vb
' 角色定义
sb.AppendLine("你是专业的中文文档校对专家。请仔细检查以下文档，找出需要修正的问题。")
sb.AppendLine()
sb.AppendLine("【输出要求】")
sb.AppendLine("只输出JSON数组，不要输出任何其他内容（包括markdown代码块、解释说明等）。")
sb.AppendLine()

' 校对范围（修复笔误，精简描述）
sb.AppendLine("【校对范围】")
sb.AppendLine("1. 错别字和拼写错误")
sb.AppendLine("2. 词语使用错误（的/地/得混用、他/她/它混用等）")
sb.AppendLine("3. 标点符号错误（中英文标点混用、缺失、多余、不匹配）")
sb.AppendLine("4. 语法和语病问题")
sb.AppendLine("5. 表达不通顺或歧义")
sb.AppendLine()

' 最小修改原则（新增）
sb.AppendLine("【最小修改原则】")
sb.AppendLine("- 只修改确实有问题的内容，不要为了优化表达而改动正确文本")
sb.AppendLine("- suggestion必须保持原文含义，只修正错误部分")
sb.AppendLine("- original必须精确匹配原文，包含标点和空格")
sb.AppendLine()

' 格式规范
sb.AppendLine("【issueType可选值（统一小驼峰）】")
sb.AppendLine("- SpellingError: 拼写错误")
sb.AppendLine("- WordUsageError: 用词错误")
sb.AppendLine("- PunctuationError: 标点错误")
sb.AppendLine("- GrammaticalError: 语法错误")
sb.AppendLine("- ExpressionError: 表达问题")
sb.AppendLine()
sb.AppendLine("【severity可选值】")
sb.AppendLine("- High: 必须修改（错别字、严重语法错误）")
sb.AppendLine("- Medium: 建议修改（用词不当、轻微语病）")
sb.AppendLine("- Low: 可选优化（表达可更精炼，不确定是否需要改）")
sb.AppendLine()

' 文档格式标准上下文（新增，当文档已排版时）
If Not String.IsNullOrEmpty(documentStandardContext) Then
    sb.AppendLine("【文档排版标准】")
    sb.AppendLine($"本文档已按{documentStandardContext}排版，请同时检查格式是否符合该标准：")
    sb.AppendLine("- 标题层级是否合理")
    sb.AppendLine("- 字体字号是否符合标准要求")
    sb.AppendLine("- 段落缩进和间距是否规范")
    sb.AppendLine("- 标点符号使用是否符合标准")
    sb.AppendLine()
End If
```

#### 3.2.3 格式克隆提示词重设计

**关键改进：**

1. 扩大可用标签集（从9个到完整集）
2. 接收映射的完整标签列表（不硬编码）
3. 同时分析源文档和目标文档
4. LLM辅助段落角色匹配

**新提示词核心改动：**

```vb
Public Shared Function BuildClonePrompt(
    sourceFormats As List(Of ExtractedParagraphFormat),
    targetParagraphs As List(Of String),
    availableTags As List(Of SemanticTag)) As String

    Dim sb As New StringBuilder()
    sb.AppendLine("你是排版专家。需要将源文档的格式应用到目标文档。")
    sb.AppendLine()

    ' 可用标签（动态生成，不硬编码）
    sb.AppendLine("【可用语义标签】")
    For Each tag In availableTags
        sb.Append($"- {tag.TagId}: {tag.DisplayName}")
        If Not String.IsNullOrEmpty(tag.MatchHint) Then
            sb.Append($"（{tag.MatchHint}）")
        End If
        sb.AppendLine()
    Next
    sb.AppendLine()

    ' 源文档格式信息
    sb.AppendLine("【源文档提取的格式（按出现频率排序）】")
    For Each f In sourceFormats.Take(30) ' 增加到30组
        sb.AppendLine($"- 样式:{f.StyleName} | 出现{f.OccurrenceCount}次 | 样本:「{f.SampleText}」")
        sb.AppendLine($"  字体: CN={f.FontNameCN} EN={f.FontNameEN} {f.FontSize}pt Bold={f.Bold}")
        sb.AppendLine($"  段落: 对齐={f.AlignmentStr} 首行缩进={f.FirstLineIndentCm}cm 行距={f.LineSpacingPt}pt")
    Next
    sb.AppendLine()

    ' 目标文档段落（新增！让LLM看到目标内容）
    sb.AppendLine("【目标文档段落（需要被格式化的文档）】")
    For i = 0 To Math.Min(targetParagraphs.Count - 1, 99) ' 最多100段
        Dim text = targetParagraphs(i)
        If String.IsNullOrWhiteSpace(text) Then Continue For
        If text.Length > 100 Then text = text.Substring(0, 100) & "..."
        sb.AppendLine($"[{i}] {text}")
    Next
    sb.AppendLine()

    ' 任务说明
    sb.AppendLine("【任务】")
    sb.AppendLine("1. 分析源文档的格式规则，识别每种格式对应的语义角色")
    sb.AppendLine("2. 分析目标文档的结构，识别每个段落的语义角色")
    sb.AppendLine("3. 为目标文档的每个段落匹配源文档中对应角色的格式")
    sb.AppendLine()
    sb.AppendLine("【输出格式】只输出JSON，不要解释：")
    sb.AppendLine("{"")
    sb.AppendLine("  ""name"": ""克隆格式"",")
    sb.AppendLine("  ""paragraphTags"": [")
    sb.AppendLine("    {""paraIndex"":0, ""tag"":""title.main"", ""sourceRole"":""标题""},")
    sb.AppendLine("    {""paraIndex"":1, ""tag"":""body.normal"", ""sourceRole"":""正文""}")
    sb.AppendLine("  ],")
    sb.AppendLine("  ""semanticTags"": [")
    sb.AppendLine("    {""tagId"":""title.main"",""font"":{""fontNameCN"":""..."",""fontNameEN"":""..."",""fontSize"":22,""bold"":true},""paragraph"":{""alignment"":""center""}}")
    sb.AppendLine("  ]")
    sb.AppendLine("}")

    Return sb.ToString()
End Function
```

---

### 3.3 对话式排版流程重设计

#### 3.3.1 当前流程的问题

```
用户："帮我排版"
  ↓
SmartFormattingOrchestrator.AnalyzeAndRecommend (纯规则)
  ↓
生成预览卡（正文段落标记为"AI待标注"）
  ↓
用户点"应用排版" → 规则直接应用（正文全变body.normal）
  或
用户点"预览" → AI语义标注（可选步骤）→ 应用
```

**核心问题：** AI语义标注是可选的后置步骤，大部分用户会直接点"应用"，导致正文段落格式单一。

#### 3.3.2 新流程设计

```
用户："帮我排版"
  ↓
阶段1：规则快速分析（即时，0延迟）
  → DocumentAnalyzer.Analyze
  → FormattingKnowledgeEngine.GetStandardForDocumentType
  → 显示初步分析结果："检测到公文文档，推荐GB/T 9704-2012标准"
  ↓
阶段2：AI深度理解（异步，2-5秒）
  → 发送完整文档+格式线索给LLM
  → LLM返回：文档结构分析 + 每个段落的语义角色 + 理由
  → 系统验证LLM结果：检查标签有效性、段落完整性
  ↓
阶段3：格式方案展示
  → 展示LLM的文档分析摘要
  → 展示每个段落的标注结果和理由
  → 用户可以直接"应用"，也可以对话调整
  ↓
用户："第三段应该是落款不是正文"
  ↓
阶段4：对话式调整
  → LLM理解用户意图："段落2应该从body.normal改为footer.signature"
  → 系统更新标注 → 更新预览
  ↓
用户确认 → 应用排版
```

**关键改变：**

1. AI语义标注从"可选后置"变为"必选核心"
2. 用户先看到分析结果，再决定是否应用（而不是先应用再调整）
3. 对话调整走LLM理解，不走关键词匹配

#### 3.3.3 ChatFormatterAgent改进

**当前：** `HandleFormattingMessage` 直接调用 `ChatReformatAsync`，返回纯规则预览。

**改进：**

```vb
Public Async Function HandleFormattingMessage(
    userMessage As String,
    paragraphs As List(Of String),
    wordParagraphs As List(Of Object),
    responseUuid As String) As Task(Of Boolean)

    If Not IsFormattingRelated(userMessage) Then Return False

    Dim lastPlan = _orchestrator.RefinementContext.LastPlan

    If _orchestrator.HasActiveContext() AndAlso Not IsNewFormattingRequest(userMessage) Then
        ' 对话调整：用LLM理解用户意图
        Dim refinedPlan = Await HandleConversationalRefinement(userMessage, paragraphs)
        ' ... 展示调整结果
    Else
        ' 新排版请求：规则分析 + AI深度理解
        Dim quickResult = _orchestrator.AnalyzeAndRecommend(paragraphs)

        ' 立即展示初步分析
        Await ShowInitialAnalysis(quickResult, responseUuid)

        ' 异步触发AI深度分析
        Dim aiAnalysis = Await PerformDeepAnalysis(quickResult, paragraphs, wordParagraphs)

        ' 更新方案并展示
        Await UpdatePlanWithAIAnalysis(aiAnalysis, responseUuid)
    End If

    Return True
End Function
```

#### 3.3.4 对话调整改为LLM驱动

**当前：** `ApplyRefinement` 用关键词匹配解析用户意图（"标题"+"大"→字号+2pt）

**改进：** 用LLM理解用户意图并输出结构化调整指令

**新提示词（对话调整专用）：**

```
你是排版调整助手。用户对当前的排版方案提出了调整意见，请理解用户的意图并输出调整指令。

当前排版方案摘要：
- 文档类型：公文
- 使用标准：GB/T 9704-2012
- 标注结果：3个标题段落、5个正文段落、1个署名、1个日期

用户说：{userMessage}

请输出JSON格式的调整指令：
[
  {"action": "retag", "paraIndex": 2, "newTag": "footer.signature", "reason": "用户认为该段落是署名"},
  {"action": "adjust", "targetTag": "title.1", "property": "fontSize", "delta": 2, "reason": "用户希望标题更大"}
]

可用的action：
- retag：更改段落标签
- adjust：调整某个标签的格式属性（property可选：fontSize, lineHeight, indent, alignment, bold）
- changeStandard：更换排版标准（需要指定standardName）
```

这样用户可以说：
- "第三段应该是落款" → `retag paraIndex:2`
- "标题太大了" → `adjust title.1 fontSize delta:-2`
- "换成学术论文格式" → `changeStandard academic-paper`
- "正文行距太紧" → `adjust body.normal lineHeight delta:+0.5`
- "这段不对，应该是副标题不是正文" → `retag paraIndex:5 newTag:heading.2`

---

### 3.4 FormatMirrorService改进

#### 3.4.1 采样范围扩大

```vb
Private Const MaxSamplesToExtract As Integer = 200  ' 从80扩大到200
```

#### 3.4.2 Grouping Key改进

当前key忽略字体名和缩进，改进后包含：

```vb
Dim key As String = String.Join("|", {
    styleName,
    fontSize.ToString(),
    fontBold.ToString(),
    alignment.ToString(),
    fmt.FontNameCN,         ' 新增：区分不同中文字体
    fmt.FirstLineIndentCm.ToString("F1")  ' 新增：区分有无首行缩进
})
```

#### 3.4.3 selectionOnly参数实际实现

```vb
If selectionOnly Then
    Dim selection As Object = wordApp.GetType().InvokeMember(
        "Selection", Reflection.BindingFlags.GetProperty, Nothing, wordApp, Nothing)
    paragraphs = selection.GetType().InvokeMember(
        "Paragraphs", Reflection.BindingFlags.GetProperty, Nothing, selection, Nothing)
Else
    paragraphs = doc.GetType().InvokeMember(
        "Paragraphs", Reflection.BindingFlags.GetProperty, Nothing, doc, Nothing)
End If
```

#### 3.4.4 发送给LLM的标签集动态化

不再硬编码9个标签，而是从 `SemanticStyleMapping` 中动态获取。

---

### 3.5 校对功能改进

#### 3.5.1 校对与排版标准结合

当文档已按某个标准排版时，校对应该同时检查格式合规性：

```vb
Public Shared Function BuildFullDocumentPrompt(
    paragraphs As List(Of String),
    Optional appliedStandard As String = Nothing,
    Optional standardRules As String = Nothing) As String

    ' ... 基础校对提示词 ...

    If Not String.IsNullOrEmpty(appliedStandard) Then
        sb.AppendLine("【格式合规检查】")
        sb.AppendLine($"本文档应遵循{appliedStandard}标准，请同时检查以下格式问题：")
        sb.AppendLine(standardRules)
        sb.AppendLine()
    End If
End Function
```

#### 3.5.2 分段校对（长文档优化）

当前校对一次性发送全文，长文档会超出上下文窗口。改为分段校对：

```vb
Public Shared Function BuildSegmentedProofreadPrompt(
    paragraphs As List(Of String),
    segmentStart As Integer,
    segmentEnd As Integer) As String

    Dim sb As New StringBuilder()
    sb.AppendLine("你是专业的中文文档校对专家。请检查以下段落。")
    sb.AppendLine("只输出JSON数组，格式：[{""paragraphIndex"":0, ""original"":""..."", ""suggestion"":""..."", ""issueType"":""SpellingError"", ""severity"":""High"", ""explanation"":""...""}]")
    sb.AppendLine()

    For i = segmentStart To Math.Min(segmentEnd, paragraphs.Count - 1)
        If Not String.IsNullOrWhiteSpace(paragraphs(i)) Then
            sb.AppendLine($"[{i}] {paragraphs(i)}")
        End If
    Next

    Return sb.ToString()
End Function
```

#### 3.5.3 校对结果后处理

LLM返回的 `original` 字段经常无法精确匹配原文。增加模糊匹配：

```vb
''' <summary>
''' 在段落文本中模糊查找original文本的位置
''' </summary>
Public Shared Function FuzzyFindOriginal(
    paragraphText As String,
    original As String,
    Optional tolerance As Integer = 2) As FuzzyMatchResult

    ' 1. 精确匹配
    Dim exactPos = paragraphText.IndexOf(original)
    If exactPos >= 0 Then
        Return New FuzzyMatchResult With {.Found = True, .Position = exactPos, .Length = original.Length, .MatchType = "exact"}
    End If

    ' 2. 忽略首尾空白匹配
    Dim trimmedOriginal = original.Trim()
    Dim trimmedPos = paragraphText.IndexOf(trimmedOriginal)
    If trimmedPos >= 0 Then
        Return New FuzzyMatchResult With {.Found = True, .Position = trimmedPos, .Length = trimmedOriginal.Length, .MatchType = "trimmed"}
    End If

    ' 3. 标点等价匹配（中文标点↔英文标点）
    Dim normalizedOriginal = NormalizePunctuation(original)
    Dim normalizedParagraph = NormalizePunctuation(paragraphText)
    Dim normPos = normalizedParagraph.IndexOf(normalizedOriginal)
    If normPos >= 0 Then
        ' 映射回原始位置
        Return New FuzzyMatchResult With {.Found = True, .Position = normPos, .Length = original.Length, .MatchType = "punctuation-normalized"}
    End If

    ' 4. 编辑距离匹配（容错1-2个字符差异）
    ' ... 实现 Levenshtein 距离搜索 ...

    Return New FuzzyMatchResult With {.Found = False, .MatchType = "not-found"}
End Function
```

---

### 3.6 SmartFormattingOrchestrator改进

#### 3.6.1 正文段落不再留空标签

**当前：** `change.NewTag = ""`（AI待标注）

**改进：** 根据位置和上下文给出合理的默认标签

```vb
Private Function InferBodyTag(
    paraIndex As Integer,
    totalParagraphs As Integer,
    analysis As DocumentAnalysisResult,
    mapping As SemanticStyleMapping) As String

    ' 文档开头附近的非标题段落
    If paraIndex < 3 Then
        If mapping.FindTag("body.abstract") IsNot Nothing Then Return "body.abstract"
    End If

    ' 文档末尾附近的短段落
    If paraIndex >= totalParagraphs - 3 Then
        Dim text = analysis.Structure.ParagraphRanges.Where(
            Function(r) r.StartIndex = paraIndex).Select(Function(r) r.Text).FirstOrDefault()
        If text?.Length < 30 Then
            If mapping.FindTag("footer.signature") IsNot Nothing AndAlso
               IsLikelySignature(text) Then Return "footer.signature"
            If mapping.FindTag("footer.date") IsNot Nothing AndAlso
               IsLikelyDate(text) Then Return "footer.date"
        End If
    End If

    ' 默认正文
    Return "body.normal"
End Function
```

#### 3.6.2 对话调整从规则匹配改为LLM驱动

参见3.3.4节。

---

## 四、实施计划

### 阶段1：Bug修复 + 提示词优化（优先级：最高）

**工期估算：1-2天**

| 任务 | 文件 | 说明 |
|------|------|------|
| 修复DocumentAnalyzer循环边界 | `DocumentAnalyzer.vb` | 1行代码修改 |
| 修复UndoStack裁剪逻辑 | `UndoStack.vb` | Push方法重写 |
| 修复GB/T 9704页边距单位 | `FormattingStandardData.vb` + `SemanticRenderingEngine.vb` | 验证并添加单位转换 |
| 修复校对提示词笔误 | `ProofreadPromptBuilder.vb` | "的/在/得/地"→"的/地/得" |
| 重写语义标注提示词 | `SemanticPromptBuilder.vb` | 按3.2.1方案重写 |
| 重写校对提示词 | `ProofreadPromptBuilder.vb` | 按3.2.2方案重写 |

**验证方式：** 用相同文档对比新旧提示词的LLM返回质量。

### 阶段2：流程改进（优先级：高）

**工期估算：3-5天**

| 任务 | 文件 | 说明 |
|------|------|------|
| AI语义标注改为必选步骤 | `ChatFormatterAgent.vb` | 阶段1+2流程 |
| 新增对话调整LLM提示词 | 新文件 `RefinementPromptBuilder.vb` | 3.3.4方案 |
| ChatFormatterAgent流程重写 | `ChatFormatterAgent.vb` | 3.3.3方案 |
| 正文段落默认标签推断 | `SmartFormattingOrchestrator.vb` | 3.6.1方案 |
| 校对分段+模糊匹配 | `ProofreadPromptBuilder.vb` + 新文件 `ProofreadFuzzyMatcher.vb` | 3.5.2+3.5.3方案 |

**验证方式：** 端到端测试对话排版流程，确认LLM分析+对话调整可用。

### 阶段3：格式克隆+模板扩展（优先级：中）

**工期估算：2-3天**

| 任务 | 文件 | 说明 |
|------|------|------|
| FormatMirrorService采样扩大 | `FormatMirrorService.vb` | 80→200段，20→30组 |
| Grouping Key改进 | `FormatMirrorService.vb` | 加入字体名和缩进 |
| selectionOnly参数实现 | `FormatMirrorService.vb` | 实际支持选区采样 |
| 克隆提示词重写 | `FormatMirrorService.vb` | 3.2.3方案 |
| 新增排版标准 | `FormattingStandardData.vb` | 简历、合同、信函 |
| SemanticRenderingEngine撤销增强 | `SemanticRenderingEngine.vb` | 自动快照+恢复入口 |

**验证方式：** 用真实公文和论文文档测试克隆效果。

### 阶段4：校对与排版结合（优先级：中低）

**工期估算：1-2天**

| 任务 | 文件 | 说明 |
|------|------|------|
| 校对增加排版标准上下文 | `ProofreadPromptBuilder.vb` | 3.5.1方案 |
| 校对与排版系统联动 | `SmartProofreadFocusMode.vb` | 校对时参考排版标准 |

---

## 五、预期效果

| 指标 | 当前 | 改进后 |
|------|------|--------|
| 语义标注准确率 | ~60%（正文经常标错） | ~85%（LLM看到完整上下文+推理） |
| 校对结果可用率 | ~40%（original经常不匹配） | ~80%（模糊匹配+更强JSON约束） |
| 对话调整理解率 | ~30%（关键词匹配，无法理解自然语言） | ~80%（LLM理解意图+输出结构化指令） |
| 格式克隆效果 | ~50%（标签太少、采样太少） | ~75%（完整标签+目标文档分析） |
| 用户体验 | "排版后正文都一样" | "排版后能看到结构分析，可对话调整" |

---

## 六、风险与注意事项

1. **LLM调用延迟**：阶段2新增的AI深度分析需要2-5秒，需要用异步加载UI避免用户等待
2. **Token消耗**：新提示词发送完整段落文本（不截断），长文档token消耗增加。可通过分段发送控制
3. **reason字段解析**：新语义标注输出包含reason字段，需要更新 `TaggingValidator` 和解析逻辑
4. **对话调整LLM调用**：每次对话调整都需要LLM调用，需要做缓存避免重复分析
5. **VB.NET语法**：所有新增代码必须严格遵循VB.NET语法，避免C#污染
6. **.vbproj注册**：新增的 `.vb` 文件必须加入 `ShareRibbon.vbproj` 的 `<Compile Include>`
