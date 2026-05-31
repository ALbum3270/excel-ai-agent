' ShareRibbon\Controls\Services\ChatFormatterAgent.vb
' Chat格式化代理 - 处理Chat中的排版对话消息、生成排版卡片HTML、解析自然语言排版指令
' Phase 2改进：AI语义标注成为必经步骤，分析-确认模式，LLM驱动的微调

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 微调指令 - 从用户自然语言中解析出的排版微调操作
''' </summary>
Public Class RefinementCommand
    ''' <summary>目标区域: "title" / "body" / "page" / "heading" / "all"</summary>
    Public Property Target As String = ""
    ''' <summary>操作: "fontSize" / "alignment" / "color" / "spacing" / "fontFamily" / "indent"</summary>
    Public Property Action As String = ""
    ''' <summary>操作值: "+2pt" / "center" / "#FF0000" / "1.5"</summary>
    Public Property Value As String = ""
    ''' <summary>用户原始消息</summary>
    Public Property OriginalText As String = ""
End Class

''' <summary>
''' Chat格式化代理 - 在Chat对话中处理排版相关的用户交互
''' Phase 2改进：
''' 1. AI语义标注为必经步骤（无AI时明确告知用户）
''' 2. 先分析后确认：先展示AI分析结果，用户确认后才应用
''' 3. LLM驱动的微调：模糊指令通过RefinementPromptBuilder发送给LLM
''' </summary>
Public Class ChatFormatterAgent

    Private ReadOnly _orchestrator As SmartFormattingOrchestrator
    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _escapeJs As Func(Of String, String)
    Private ReadOnly _textAnalyzer As Func(Of String, String, Task(Of String))

    ''' <summary>
    ''' 存储最后AI标注的段落结果（用于应用排版时）
    ''' </summary>
    Private _lastTaggedParagraphs As List(Of TaggedParagraph) = Nothing

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(
        executeScript As Func(Of String, Task),
        escapeJs As Func(Of String, String),
        Optional textAnalyzer As Func(Of String, String, Task(Of String)) = Nothing,
        Optional orchestrator As SmartFormattingOrchestrator = Nothing)

        _executeScript = executeScript
        _escapeJs = escapeJs
        _textAnalyzer = textAnalyzer
        _orchestrator = If(orchestrator, New SmartFormattingOrchestrator())
    End Sub

    ''' <summary>访问编排器实例</summary>
    Public ReadOnly Property Orchestrator As SmartFormattingOrchestrator
        Get
            Return _orchestrator
        End Get
    End Property

    ''' <summary>是否有AI分析器可用</summary>
    Public ReadOnly Property HasAIAnalyzer As Boolean
        Get
            Return _textAnalyzer IsNot Nothing
        End Get
    End Property

    ''' <summary>
    ''' 处理Chat中的排版消息。
    ''' 根据消息内容自动判断是"首次排版请求"还是"微调指令"。
    ''' Phase 2改进：始终走AI标注流程，先展示分析结果再确认。
    ''' </summary>
    ''' <param name="userMessage">用户消息文本</param>
    ''' <param name="paragraphs">文档段落文本列表</param>
    ''' <param name="wordParagraphs">Word段落对象列表</param>
    ''' <param name="responseUuid">响应的UUID（用于推送HTML到前端）</param>
    ''' <returns>是否已处理（False表示非排版消息，应由其他处理器处理）</returns>
    Public Async Function HandleFormattingMessage(
        userMessage As String,
        paragraphs As List(Of String),
        wordParagraphs As List(Of Object),
        responseUuid As String) As Task(Of Boolean)

        ' 判断是否为排版相关消息
        If Not IsFormattingRelated(userMessage) Then
            Return False
        End If

        Try
            ' 通过ChatReformatAsync解析用户意图，获取排版方案
            Dim lastPlan = _orchestrator.RefinementContext.LastPlan
            Dim plan = Await _orchestrator.ChatReformatAsync(userMessage, paragraphs, wordParagraphs)

            If plan Is Nothing OrElse plan.Changes.Count = 0 Then
                ' 没有生成方案，展示提示
                Dim html = GenerateNoPlanCardHtml(userMessage)
                Await PushCardHtml(html, responseUuid)
                Return True
            End If

            ' 判断是否为微调请求
            Dim isRefinement = _orchestrator.HasActiveContext() AndAlso
                               Not IsNewFormattingRequest(userMessage) AndAlso
                               lastPlan IsNot Nothing

            If isRefinement Then
                ' 微调模式：AI标注结果已存在，展示微调对比卡片
                Dim html = GenerateRefinementCardHtml(lastPlan, plan)
                Await PushCardHtml(html, responseUuid)
            Else
                ' 首次排版：执行AI语义标注（必经步骤）
                If _textAnalyzer IsNot Nothing Then
                    ' AI可用：执行语义标注，展示带分析结果的卡片
                    Dim taggedParagraphs = Await PerformAISemanticTaggingAsync(
                        paragraphs,
                        plan.SemanticMapping,
                        documentTypeContext:=plan.StandardName)

                    ' 将AI标注结果合并到plan中
                    UpdatePlanWithAITags(plan, taggedParagraphs)

                    Dim html = GenerateAnalysisCardHtml(plan, taggedParagraphs)
                    Await PushCardHtml(html, responseUuid)
                Else
                    ' AI不可用：展示警告卡片，说明效果有限
                    Dim html = GenerateNoAICardHtml(plan)
                    Await PushCardHtml(html, responseUuid)
                End If
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 处理排版消息失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 使用LLM驱动的微调（替代纯正则的ApplyTweakToTags）
    ''' 当AI可用时，将模糊指令发送给LLM获取精确的映射调整
    ''' </summary>
    ''' <param name="userCommand">用户微调指令</param>
    ''' <param name="paragraphs">文档段落文本列表</param>
    ''' <param name="wordParagraphs">Word段落对象列表</param>
    ''' <param name="responseUuid">响应的UUID</param>
    Public Async Function HandleLLMRefinementAsync(
        userCommand As String,
        paragraphs As List(Of String),
        wordParagraphs As List(Of Object),
        responseUuid As String) As Task(Of Boolean)

        If _textAnalyzer Is Nothing Then
            ' 无AI，降级到正则微调
            Dim plan = _orchestrator.ApplyRefinement(userCommand)
            If plan IsNot Nothing Then
                Dim html = GenerateRefinementCardHtml(plan, plan)
                Await PushCardHtml(html, responseUuid)
            End If
            Return True
        End If

        Dim fallbackNeeded As Boolean = False

        Try
            ' 构建微调提示词
            Dim mapping = _orchestrator.RefinementContext.CurrentMapping
            If mapping Is Nothing Then
                ' 没有活动上下文，走完整排版流程
                Return Await HandleFormattingMessage(userCommand, paragraphs, wordParagraphs, responseUuid)
            End If

            Dim prompt = RefinementPromptBuilder.BuildRefinementPrompt(
                userCommand,
                mapping,
                _orchestrator.RefinementContext.ConversationHistory,
                _orchestrator.RefinementContext.CurrentStandard?.Name)

            ' 调用LLM
            Dim aiResponse = Await _textAnalyzer("refinement", prompt)

            If String.IsNullOrWhiteSpace(aiResponse) Then
                ' AI返回为空，降级到正则微调
                fallbackNeeded = True
            Else
                ' 解析LLM返回的调整后的映射
                Dim adjustedMapping = ParseRefinementResponse(aiResponse, mapping)

                ' 用调整后的映射重新生成预览方案
                Dim beforePlan = _orchestrator.RefinementContext.CurrentPreviewPlan
                _orchestrator.RefinementContext.CurrentMapping = adjustedMapping

                Dim analysis = _orchestrator.RefinementContext.CurrentAnalysis
                Dim standard = _orchestrator.RefinementContext.CurrentStandard
                Dim afterPlan As ReformatPreviewPlan = Nothing

                If analysis IsNot Nothing AndAlso standard IsNot Nothing Then
                    afterPlan = _orchestrator.GeneratePreviewPlan(analysis, standard, paragraphs)
                    afterPlan.SemanticMapping = adjustedMapping
                Else
                    afterPlan = _orchestrator.ApplyRefinement(userCommand)
                End If

                ' 记录对话历史
                _orchestrator.RefinementContext.AddConversation(userCommand)
                _orchestrator.RefinementContext.CurrentPreviewPlan = afterPlan

                Dim html2 = GenerateRefinementCardHtml(beforePlan, afterPlan)
                Await PushCardHtml(html2, responseUuid)

                Return True
            End If

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] LLM微调失败: {ex.Message}")
            fallbackNeeded = True
        End Try

        ' 降级到正则微调（在Try/Catch外部执行Await）
        If fallbackNeeded Then
            Dim plan = _orchestrator.ApplyRefinement(userCommand)
            If plan IsNot Nothing Then
                Dim html = GenerateRefinementCardHtml(plan, plan)
                Await PushCardHtml(html, responseUuid)
            End If
        End If

        Return True
    End Function

    ''' <summary>
    ''' 解析LLM微调响应，合并到当前映射
    ''' </summary>
    Private Function ParseRefinementResponse(aiResponse As String, originalMapping As SemanticStyleMapping) As SemanticStyleMapping
        ' 深拷贝原始映射
        Dim result As New SemanticStyleMapping()
        result.Name = originalMapping.Name
        result.SourceType = originalMapping.SourceType
        result.SourceId = originalMapping.SourceId
        result.PageConfig = originalMapping.PageConfig

        Try
            ' 清理响应
            Dim clean = aiResponse.Trim()
            If clean.StartsWith("```json") Then clean = clean.Substring(7)
            If clean.StartsWith("```") Then clean = clean.Substring(3)
            If clean.EndsWith("```") Then clean = clean.Substring(0, clean.Length - 3)
            clean = clean.Trim()

            ' 解析为JArray
            Dim tagsArray As JArray = Nothing
            If clean.StartsWith("[") Then
                tagsArray = JArray.Parse(clean)
            Else
                ' 尝试提取数组
                Dim firstBracket = clean.IndexOf("[")
                Dim lastBracket = clean.LastIndexOf("]")
                If firstBracket >= 0 AndAlso lastBracket > firstBracket Then
                    tagsArray = JArray.Parse(clean.Substring(firstBracket, lastBracket - firstBracket + 1))
                End If
            End If

            If tagsArray Is Nothing OrElse tagsArray.Count = 0 Then
                Debug.WriteLine("[ChatFormatterAgent] LLM微调响应解析为空，使用原始映射")
                Return originalMapping
            End If

            ' 构建tagId→新tag的映射
            Dim newTagsDict As New Dictionary(Of String, SemanticTag)()
            For Each item In tagsArray
                Dim tagId = item("tagId")?.ToString()
                If String.IsNullOrEmpty(tagId) Then Continue For

                Dim newTag As New SemanticTag()
                newTag.TagId = tagId

                ' 解析Font
                Dim fontObj = item("font")
                If fontObj IsNot Nothing Then
                    newTag.Font = New FontConfig()
                    newTag.Font.FontNameCN = If(fontObj("fontNameCN")?.ToString(), newTag.Font.FontNameCN)
                    newTag.Font.FontNameEN = If(fontObj("fontNameEN")?.ToString(), newTag.Font.FontNameEN)
                    Dim fontSizeVal = fontObj("fontSize")
                    If fontSizeVal IsNot Nothing Then newTag.Font.FontSize = Convert.ToDouble(fontSizeVal)
                    Dim boldVal = fontObj("bold")
                    If boldVal IsNot Nothing Then newTag.Font.Bold = Convert.ToBoolean(boldVal)
                    Dim italicVal = fontObj("italic")
                    If italicVal IsNot Nothing Then newTag.Font.Italic = Convert.ToBoolean(italicVal)
                    Dim underlineVal = fontObj("underline")
                    If underlineVal IsNot Nothing Then newTag.Font.Underline = Convert.ToBoolean(underlineVal)
                End If

                ' 解析Paragraph
                Dim paraObj = item("paragraph")
                If paraObj IsNot Nothing Then
                    newTag.Paragraph = New ParagraphConfig()
                    newTag.Paragraph.Alignment = If(paraObj("alignment")?.ToString(), newTag.Paragraph.Alignment)
                    Dim indentVal = paraObj("firstLineIndent")
                    If indentVal IsNot Nothing Then newTag.Paragraph.FirstLineIndent = Convert.ToDouble(indentVal)
                    Dim spacingVal = paraObj("lineSpacing")
                    If spacingVal IsNot Nothing Then newTag.Paragraph.LineSpacing = Convert.ToDouble(spacingVal)
                    Dim beforeVal = paraObj("spaceBefore")
                    If beforeVal IsNot Nothing Then newTag.Paragraph.SpaceBefore = Convert.ToDouble(beforeVal)
                    Dim afterVal = paraObj("spaceAfter")
                    If afterVal IsNot Nothing Then newTag.Paragraph.SpaceAfter = Convert.ToDouble(afterVal)
                End If

                ' 解析Color
                Dim colorObj = item("color")
                If colorObj IsNot Nothing Then
                    newTag.Color = New ColorConfig()
                    newTag.Color.FontColor = If(colorObj("fontColor")?.ToString(), newTag.Color.FontColor)
                End If

                newTagsDict(tagId) = newTag
            Next

            ' 合并：LLM返回的标签覆盖原始映射中对应的标签，其余保留
            For Each origTag In originalMapping.SemanticTags
                If newTagsDict.ContainsKey(origTag.TagId) Then
                    result.SemanticTags.Add(newTagsDict(origTag.TagId))
                Else
                    result.SemanticTags.Add(origTag)
                End If
            Next

            ' LLM可能新增了标签
            For Each kvp In newTagsDict
                If Not result.SemanticTags.Any(Function(t) t.TagId = kvp.Key) Then
                    result.SemanticTags.Add(kvp.Value)
                End If
            Next

            Return result

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 解析LLM微调响应失败: {ex.Message}")
            Return originalMapping
        End Try
    End Function

    ''' <summary>
    ''' 将AI标注结果更新到预览方案中
    ''' </summary>
    Private Sub UpdatePlanWithAITags(plan As ReformatPreviewPlan, taggedParagraphs As List(Of TaggedParagraph))
        If plan Is Nothing OrElse taggedParagraphs Is Nothing Then Return

        ' 更新plan.Changes中的NewTag为AI标注结果
        For Each change In plan.Changes
            Dim tagged = taggedParagraphs.FirstOrDefault(Function(t) t.ParaIndex = change.ParagraphIndex)
            If tagged IsNot Nothing AndAlso Not String.IsNullOrEmpty(tagged.TagId) Then
                change.NewTag = tagged.TagId
            End If
        Next
    End Sub

    ''' <summary>
    ''' 推送卡片HTML到前端
    ''' </summary>
    Private Async Function PushCardHtml(html As String, responseUuid As String) As Task
        Dim uuid = If(responseUuid, Guid.NewGuid().ToString())
        Dim jsonPayload As New JObject()
        jsonPayload("uuid") = uuid
        jsonPayload("html") = html
        Await _executeScript($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")
    End Function

    ''' <summary>
    ''' 生成AI分析结果卡片HTML（Phase 2新增：先分析后确认）
    ''' </summary>
    Public Function GenerateAnalysisCardHtml(plan As ReformatPreviewPlan, taggedParagraphs As List(Of TaggedParagraph)) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card formatting-card-analysis"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F50D;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">AI分析完成</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        ' 文档类型与标准
        If plan.DetectedType <> DocumentType.Unknown Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{plan.DocumentTypeName}</strong> (置信度{Math.Round(plan.TypeConfidence * 100)}%)</div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">排版标准: <strong>{plan.StandardName}</strong></div>")
        End If

        ' AI标注摘要
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">AI识别的文档结构:</div>")

        ' 按tagId分组
        Dim grouped = taggedParagraphs.GroupBy(Function(t) t.TagId).OrderByDescending(Function(g) g.Count()).ToList()
        For Each group In grouped
            Dim tagName = group.Key
            Dim count = group.Count()
            ' 取一条reason作为示例
            Dim sampleReason = group.FirstOrDefault()?.Reason
            Dim displayName = GetTagDisplayName(tagName, plan.SemanticMapping)

            sb.AppendLine($"      <div class=""formatting-change-item"">")
            sb.AppendLine($"        <span class=""formatting-change-section"">{System.Web.HttpUtility.HtmlEncode(displayName)}</span>")
            sb.AppendLine($"        <span class=""formatting-change-count"">({count}处)</span>")
            If Not String.IsNullOrEmpty(sampleReason) Then
                sb.AppendLine($"        <span class=""formatting-change-reason"">- {System.Web.HttpUtility.HtmlEncode(sampleReason)}</span>")
            End If
            sb.AppendLine($"      </div>")
        Next

        sb.AppendLine($"      <div class=""formatting-change-summary"">合计: {taggedParagraphs.Count}个段落, {grouped.Count}种样式</div>")
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">确认应用</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-secondary"" onclick=""previewReformat();"">预览对比</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-outline"" onclick=""alternateReformat();"">换一种</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-ghost"" onclick=""startRefinement();"">微调</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成无AI时的卡片HTML（告知用户AI不可用，效果有限）
    ''' </summary>
    Public Function GenerateNoAICardHtml(plan As ReformatPreviewPlan) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card formatting-card-warning"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x26A0;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版建议（基础模式）</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        sb.AppendLine("    <div class=""formatting-info-row formatting-info-warning"">AI分析器未启用，将使用基于规则的分析，精度有限。建议配置AI模型以获得更精准的排版效果。</div>")

        ' 文档类型与标准
        If plan.DetectedType <> DocumentType.Unknown Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{plan.DocumentTypeName}</strong></div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">排版标准: <strong>{plan.StandardName}</strong></div>")
        End If

        ' 变更列表
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">即将修改:</div>")
        Dim grouped = plan.Changes.GroupBy(Function(c) If(String.IsNullOrEmpty(c.NewTag), "__pending__", c.NewTag)).ToList()
        For Each group In grouped
            Dim tagName = If(group.Key = "__pending__", "待标注", group.Key)
            Dim count = group.Count()
            sb.AppendLine($"      <div class=""formatting-change-item"">")
            sb.AppendLine($"        <span class=""formatting-change-section"">{System.Web.HttpUtility.HtmlEncode(tagName)}</span>")
            sb.AppendLine($"        <span class=""formatting-change-count"">({count}处)</span>")
            sb.AppendLine($"      </div>")
        Next
        sb.AppendLine($"      <div class=""formatting-change-summary"">合计: {plan.TotalChanges}处段落</div>")
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">应用排版</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-outline"" onclick=""alternateReformat();"">换一种</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成无方案时的卡片HTML
    ''' </summary>
    Private Function GenerateNoPlanCardHtml(userMessage As String) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F4CB;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版建议</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")
        sb.AppendLine("    <div class=""formatting-info-row"">未能生成排版方案。请尝试更明确的指令，如""按公文排版""或""整理格式""。</div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成排版建议卡片HTML（兼容旧版调用）
    ''' </summary>
    Public Function GenerateFormattingCardHtml(plan As ReformatPreviewPlan) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F4CB;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版建议</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        ' 文档类型与标准
        If plan.DetectedType <> DocumentType.Unknown Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{plan.DocumentTypeName}</strong> (置信度{Math.Round(plan.TypeConfidence * 100)}%)</div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">推荐标准: <strong>{plan.StandardName}</strong></div>")
        End If

        ' 变更列表 — 按NewTag分组显示
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">即将修改:</div>")

        ' 按NewTag分组
        Dim grouped = plan.Changes.GroupBy(Function(c) If(String.IsNullOrEmpty(c.NewTag), "__pending__", c.NewTag)).ToList()
        For Each group In grouped
            Dim tagName = If(group.Key = "__pending__", "AI待标注", group.Key)
            Dim count = group.Count()
            Dim sampleDesc = group.FirstOrDefault()?.ChangeDescription
            sb.AppendLine($"      <div class=""formatting-change-item"">")
            sb.AppendLine($"        <span class=""formatting-change-section"">{System.Web.HttpUtility.HtmlEncode(tagName)}</span>")
            sb.AppendLine($"        <span class=""formatting-change-count"">({count}处)</span>")
            If Not String.IsNullOrEmpty(sampleDesc) Then
                sb.AppendLine($"        <span class=""formatting-change-desc"">: {System.Web.HttpUtility.HtmlEncode(sampleDesc)}</span>")
            End If
            sb.AppendLine($"      </div>")
        Next

        sb.AppendLine($"      <div class=""formatting-change-summary"">合计: {plan.TotalChanges}处段落, {grouped.Count}个样式区</div>")
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">应用排版</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-secondary"" onclick=""previewReformat();"">预览对比</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-outline"" onclick=""alternateReformat();"">换一种</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-ghost"" onclick=""startRefinement();"">微调</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成微调对比卡片HTML（显示变更前后对比）
    ''' </summary>
    Public Function GenerateRefinementCardHtml(
        before As ReformatPreviewPlan,
        after As ReformatPreviewPlan) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card formatting-card-refinement"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F504;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版已微调</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        ' 变更对比 — 按ParagraphIndex匹配
        sb.AppendLine("    <div class=""formatting-diff"">")
        For Each c In after.Changes
            Dim beforeChange = before.Changes.FirstOrDefault(Function(b) b.ParagraphIndex = c.ParagraphIndex)
            If beforeChange IsNot Nothing Then
                Dim oldDesc = If(String.IsNullOrEmpty(beforeChange.ChangeDescription), beforeChange.NewTag, beforeChange.ChangeDescription)
                Dim newDesc = If(String.IsNullOrEmpty(c.ChangeDescription), c.NewTag, c.ChangeDescription)
                If oldDesc <> newDesc Then
                    sb.AppendLine("      <div class=""formatting-diff-item"">")
                    sb.AppendLine($"        <span class=""formatting-diff-section"">{System.Web.HttpUtility.HtmlEncode(c.ParagraphPreview)}:</span>")
                    sb.AppendLine($"        <span class=""formatting-diff-old"">{System.Web.HttpUtility.HtmlEncode(oldDesc)}</span>")
                    sb.AppendLine($"        <span class=""formatting-diff-arrow"">&rarr;</span>")
                    sb.AppendLine($"        <span class=""formatting-diff-new"">{System.Web.HttpUtility.HtmlEncode(newDesc)}</span>")
                    sb.AppendLine("      </div>")
                End If
            End If
        Next
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">应用排版</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-ghost"" onclick=""startRefinement();"">继续微调</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 获取标签的显示名称（优先从mapping中获取，否则返回tagId）
    ''' </summary>
    Private Function GetTagDisplayName(tagId As String, mapping As SemanticStyleMapping) As String
        If mapping IsNot Nothing Then
            Dim tag = mapping.FindTag(tagId)
            If tag IsNot Nothing AndAlso Not String.IsNullOrEmpty(tag.DisplayName) Then
                Return tag.DisplayName
            End If
        End If

        ' 从注册表获取默认名称
        Dim layer2Tags = SemanticTagRegistry.GetCommonLayer2Tags()
        If layer2Tags.ContainsKey(tagId) Then Return layer2Tags(tagId)

        Dim layer1Tags = SemanticTagRegistry.GetLayer1TagDescriptions()
        Dim parentTag = SemanticTagRegistry.GetParentTag(tagId)
        If layer1Tags.ContainsKey(parentTag) Then Return $"{layer1Tags(parentTag)}.{tagId}"

        Return tagId
    End Function

    ''' <summary>
    ''' 解析用户消息中的微调指令
    ''' </summary>
    Public Shared Function ParseRefinementCommand(userMessage As String) As RefinementCommand
        Dim cmd As New RefinementCommand()
        cmd.OriginalText = userMessage

        If String.IsNullOrWhiteSpace(userMessage) Then Return cmd

        ' 简单文本解析微调指令
        Dim msg = userMessage.ToLower().Trim()
        If msg.Contains("大") OrElse msg.Contains("小") Then
            cmd.Action = "fontSize"
            cmd.Value = If(msg.Contains("大"), "+1pt", "-1pt")
        ElseIf msg.Contains("行距") Then
            cmd.Action = "spacing"
            cmd.Value = "1.5"
        ElseIf msg.Contains("红") OrElse msg.Contains("蓝") OrElse msg.Contains("颜色") Then
            cmd.Action = "color"
        ElseIf msg.Contains("居中") OrElse msg.Contains("对齐") Then
            cmd.Action = "alignment"
            cmd.Value = "center"
        End If

        cmd.Target = "all"
        Return cmd
    End Function

    ''' <summary>
    ''' 判断消息是否与排版相关
    ''' </summary>
    Public Shared Function IsFormattingRelated(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        Dim keywords As String() = {
            "排版", "格式", "样式", "字体", "字号", "行距", "对齐",
            "缩进", "页边距", "居中", "加粗", "红色", "标题", "正文",
            "仿宋", "宋体", "黑体", "楷体", "微软雅黑", "小标宋",
            "公文", "国标", "标准", "模板", "美化", "整理", "规范",
            "GB/T", "gbt", "克隆", "照这个", "参照", "按照"
        }

        Return keywords.Any(Function(k) message.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    ''' <summary>
    ''' 判断是否为全新的排版请求（而非微调）
    ''' 如果是全新的格式化请求，会重新做完整分析
    ''' </summary>
    Private Shared Function IsNewFormattingRequest(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        Dim newRequestKeywords As String() = {
            "重新", "再来", "换一种", "重新排", "重新排版",
            "换一个", "用另一个", "换成", "改用", "不要这个"
        }

        Return newRequestKeywords.Any(Function(k) message.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    ''' <summary>
    ''' 获取最后AI标注的段落结果（用于应用排版时）
    ''' </summary>
    Public Function GetLastTaggedParagraphs() As List(Of TaggedParagraph)
        Return _lastTaggedParagraphs
    End Function

    ''' <summary>
    ''' 使用AI进行语义标注（必经步骤）
    ''' 调用AI分析文档内容，返回每个段落对应的语义标签
    ''' </summary>
    ''' <param name="paragraphs">文档段落文本列表</param>
    ''' <param name="mapping">当前排版标准的语义样式映射</param>
    ''' <param name="paragraphStyles">段落样式名称列表（可选，用于增强AI判断）</param>
    ''' <param name="documentTypeContext">文档类型上下文描述（可选）</param>
    ''' <param name="detectedHeadings">已检测到的标题结构（可选）</param>
    Public Async Function PerformAISemanticTaggingAsync(
        paragraphs As List(Of String),
        mapping As SemanticStyleMapping,
        Optional paragraphStyles As List(Of String) = Nothing,
        Optional documentTypeContext As String = Nothing,
        Optional detectedHeadings As String = Nothing) As Task(Of List(Of TaggedParagraph))

        _lastTaggedParagraphs = Nothing

        ' AI不可用时返回规则推断结果（不再静默回退为body.normal）
        If _textAnalyzer Is Nothing Then
            Debug.WriteLine("[ChatFormatterAgent] 没有配置AI分析器，使用规则推断标注")
            Dim fallback As New List(Of TaggedParagraph)()
            For i = 0 To paragraphs.Count - 1
                ' 使用InferDefaultTag进行基于规则的推断（而非全部body.normal）
                Dim inferredTag = SmartFormattingOrchestrator.InferDefaultTagPublic(i, paragraphs.Count, paragraphs(i), Nothing)
                fallback.Add(New TaggedParagraph(i, inferredTag, "规则推断（无AI）"))
            Next
            _lastTaggedParagraphs = fallback
            Return fallback
        End If

        Try
            ' 构建AI提示词
            Dim originalIndices As New List(Of Integer)()
            For i = 0 To paragraphs.Count - 1
                originalIndices.Add(i)
            Next

            Dim prompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(
                mapping,
                paragraphs,
                paragraphStyles,
                originalIndices,
                detectedHeadings,
                documentTypeContext)

            ' 调用AI获取标注结果
            Debug.WriteLine("[ChatFormatterAgent] 正在调用AI进行语义标注...")
            Dim aiResponse = Await _textAnalyzer("semantic_tagging", prompt)

            If String.IsNullOrWhiteSpace(aiResponse) Then
                Debug.WriteLine("[ChatFormatterAgent] AI返回为空，使用规则推断标注")
                Return Await GetRuleBasedTaggingAsync(paragraphs, mapping)
            End If

            ' 解析AI响应
            Dim taggedParagraphs = ParseAITagResponse(aiResponse, paragraphs.Count)
            _lastTaggedParagraphs = taggedParagraphs

            Debug.WriteLine($"[ChatFormatterAgent] AI标注完成: {taggedParagraphs.Count}个段落")
            Return taggedParagraphs

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] AI标注失败: {ex.Message}")
        End Try

        ' 如果解析失败或结果为空，返回规则推断标注
        Return Await GetRuleBasedTaggingAsync(paragraphs, mapping)
    End Function

    ''' <summary>
    ''' 基于规则的标注（AI不可用时的替代方案，使用InferDefaultTag而非全部body.normal）
    ''' </summary>
    Private Function GetRuleBasedTaggingAsync(paragraphs As List(Of String), mapping As SemanticStyleMapping) As Task(Of List(Of TaggedParagraph))
        Dim result As New List(Of TaggedParagraph)()
        For i = 0 To paragraphs.Count - 1
            Dim inferredTag = SmartFormattingOrchestrator.InferDefaultTagPublic(i, paragraphs.Count, paragraphs(i), Nothing)
            result.Add(New TaggedParagraph(i, inferredTag, "规则推断"))
        Next
        _lastTaggedParagraphs = result
        Return Task.FromResult(result)
    End Function

    ''' <summary>
    ''' 获取默认标注（全部标记为body.normal）- 仅作为最后降级
    ''' </summary>
    Private Function GetDefaultTaggingAsync(paragraphs As List(Of String)) As Task(Of List(Of TaggedParagraph))
        Dim result As New List(Of TaggedParagraph)()
        For i = 0 To paragraphs.Count - 1
            result.Add(New TaggedParagraph(i, "body.normal", "默认降级"))
        Next
        _lastTaggedParagraphs = result
        Return Task.FromResult(result)
    End Function

    ''' <summary>
    ''' 解析AI标注响应
    ''' </summary>
    Private Function ParseAITagResponse(response As String, paragraphCount As Integer) As List(Of TaggedParagraph)
        Dim result As New List(Of TaggedParagraph)()

        Try
            ' 清理响应文本，移除可能的markdown代码块标记
            Dim cleanResponse = response.Trim()
            If cleanResponse.StartsWith("```json") Then
                cleanResponse = cleanResponse.Substring(7)
            ElseIf cleanResponse.StartsWith("```") Then
                cleanResponse = cleanResponse.Substring(3)
            End If
            If cleanResponse.EndsWith("```") Then
                cleanResponse = cleanResponse.Substring(0, cleanResponse.Length - 3)
            End If
            cleanResponse = cleanResponse.Trim()

            ' 尝试解析JSON数组
            Dim taggerd As List(Of TaggedParagraph) = Nothing
            Try
                taggerd = JsonConvert.DeserializeObject(Of List(Of TaggedParagraph))(cleanResponse)
            Catch ex As Exception
                ' JSON解析失败，尝试正则提取
                Debug.WriteLine($"[ChatFormatterAgent] JSON解析失败: {ex.Message}")
                taggerd = ParseTagResponseWithRegex(cleanResponse, paragraphCount)
            End Try

            If taggerd IsNot Nothing AndAlso taggerd.Count > 0 Then
                result.AddRange(taggerd)
            End If

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 解析标注响应失败: {ex.Message}")
        End Try

        ' 如果解析失败或结果为空，返回空列表（调用方会降级到规则推断）
        If result.Count = 0 Then
            For i = 0 To paragraphCount - 1
                result.Add(New TaggedParagraph(i, "body.normal", "解析失败降级"))
            Next
        End If

        Return result
    End Function

    ''' <summary>
    ''' 使用正则表达式解析标注响应（当JSON解析失败时）
    ''' </summary>
    Private Function ParseTagResponseWithRegex(response As String, paragraphCount As Integer) As List(Of TaggedParagraph)
        Dim result As New List(Of TaggedParagraph)()

        ' 匹配 {paraIndex:数字, tag:"标签", reason:"理由"(可选)} 模式
        Dim pattern = """paraIndex""\s*:\s*(\d+)\s*,\s*""tag""\s*:\s*""([^""]+)""(?:\s*,\s*""reason""\s*:\s*""([^""]*)"")?"
        Dim matches = System.Text.RegularExpressions.Regex.Matches(response, pattern)

        For Each match In matches
            Dim paraIndex = Integer.Parse(match.Groups(1).Value)
            Dim tag = match.Groups(2).Value
            Dim reason = If(match.Groups(3).Success, match.Groups(3).Value, "")
            If paraIndex >= 0 AndAlso paraIndex < paragraphCount Then
                result.Add(New TaggedParagraph(paraIndex, tag, reason))
            End If
        Next

        ' 如果正则也没有匹配到，返回空列表（将使用规则推断）
        Return result
    End Function

End Class
