' ShareRibbon\Services\Reformat\SemanticPromptBuilder.vb
' 统一构建语义标注提示词

Imports System.Text

''' <summary>
''' 语义提示词构建器 - 为AI构建语义标注的系统提示词
''' 模板排版和规范排版共用同一构建逻辑
''' </summary>
Public Class SemanticPromptBuilder

    ''' <summary>
    ''' 构建语义标注提示词（带样式上下文）
    ''' </summary>
    ''' <param name="mapping">语义样式映射（包含可用标签）</param>
    ''' <param name="paragraphs">段落文本列表（仅文本段落，非文本已过滤）</param>
    ''' <param name="paragraphStyles">段落样式名称（仅文本段落，与paragraphs一一对应）</param>
    ''' <param name="originalParaIndices">原文档中的段落索引（仅文本段落，用于映射回正确位置）</param>
    ''' <param name="detectedHeadings">DocumentAnalyzer检测到的标题信息</param>
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
        sb.AppendLine("只输出纯JSON数组，不要输出其他内容（不要使用markdown代码块，不要输出解释）。")
        sb.AppendLine("[")
        sb.AppendLine("  {""paraIndex"":0, ""tag"":""header.org"", ""reason"":""位于文档开头，文本符合发文机关标志模式""},")
        sb.AppendLine("  {""paraIndex"":1, ""tag"":""header.refno"", ""reason"":""包含发文字号格式""},")
        sb.AppendLine("  ...")
        sb.AppendLine("]")
        sb.AppendLine("要求：")
        sb.AppendLine("- reason字段简短说明判断依据（不超过30字）")
        sb.AppendLine("- 每个段落必须且只能有一个标签")
        sb.AppendLine("- paraIndex使用【文档段落】中给出的索引号")
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

        ' ===== 6. 结构识别指南 =====
        sb.AppendLine("【结构识别指南】")
        sb.AppendLine("判断段落角色时，请综合考虑以下线索：")
        sb.AppendLine("1. 文本内容：是否包含特定模式（发文字号、日期、编号等）")
        sb.AppendLine("2. 段落位置：在文档开头、中间还是末尾")
        sb.AppendLine("3. 上下文关系：与前后段落的关系（如标题后面通常跟正文）")
        sb.AppendLine("4. 格式线索：字号偏大且加粗的短段落通常是标题")
        sb.AppendLine("5. 不确定时使用最通用的body.normal标签")
        sb.AppendLine()

        ' ===== 公文特殊规则 =====
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("公文") OrElse documentTypeContext.Contains("GB/T 9704")) Then
            sb.AppendLine("【公文结构识别】")
            sb.AppendLine("公文文档有固定的结构顺序，请按此顺序识别：")
            sb.AppendLine("发文机关标志(header.org) → 发文字号(header.refno) → 签发人(header.signer)")
            sb.AppendLine("→ 文件标题(title.main) → 主送机关(title.recipient) → 正文(body.normal)")
            sb.AppendLine("→ 附件说明(body.attachment) → 发文机关署名(footer.signature)")
            sb.AppendLine("→ 成文日期(footer.date) → 附注(footer.note) → 抄送机关(footer.cc)")
            sb.AppendLine()
            sb.AppendLine("注意：")
            sb.AppendLine("- 文件标题使用title.main；正文内部的「一、」「（一）」「1.」分别使用title.1/title.2/title.3")
            sb.AppendLine("- 公文不要使用heading.*；heading.*只用于论文、报告等通用章节标题")
            sb.AppendLine("- 即使Word样式名为「标题1」，只要文本内容符合公文结构特征，必须使用公文专用标签")
            sb.AppendLine("- 文末短段落（机构名、日期）应使用footer.signature/footer.date，不要标为body.normal")
            sb.AppendLine("- 以「附注：」或圆括号联系人信息开头的段落使用footer.note；以「抄送：」开头的段落使用footer.cc")
            sb.AppendLine()
        End If

        ' ===== 7. 标注示例（按文档类型） =====
        sb.AppendLine("【标注示例】")
        Dim examples As String = GetExamplesByDocumentType(documentTypeContext, mapping)
        sb.Append(examples)
        sb.AppendLine()

        ' ===== 8. 严格要求 =====
        sb.AppendLine("【严格要求】")
        sb.AppendLine("1. 仅使用上述可用标签，禁止自创标签")
        sb.AppendLine("2. 返回纯JSON数组，不要包含markdown代码块标记")
        sb.AppendLine("3. 每个段落必须且只能有一个标签")
        sb.AppendLine("4. 层级合理：title.1后可接title.2或body，不能直接跳到title.3")
        sb.AppendLine()

        ' ===== 9. 自动检测结果 =====
        If Not String.IsNullOrEmpty(detectedHeadings) Then
            sb.AppendLine("【系统自动检测到的标题结构（仅供参考，你可以修正）】")
            sb.AppendLine(detectedHeadings)
            sb.AppendLine()
        End If

        ' ===== 10. 文档段落（完整文本+上下文） =====
        sb.AppendLine("【文档段落】")
        Dim hasStyles = paragraphStyles IsNot Nothing AndAlso paragraphStyles.Count = paragraphs.Count
        Dim hasOrigIdx = originalParaIndices IsNot Nothing AndAlso originalParaIndices.Count = paragraphs.Count
        Dim hasFontSizes = paragraphFontSizes IsNot Nothing AndAlso paragraphFontSizes.Count = paragraphs.Count
        Dim hasBold = paragraphIsBold IsNot Nothing AndAlso paragraphIsBold.Count = paragraphs.Count

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

    ''' <summary>
    ''' 根据文档类型获取对应的标注示例
    ''' </summary>
    ''' <param name="documentTypeContext">文档类型上下文（标准名称）</param>
    ''' <param name="mapping">语义样式映射</param>
    Private Shared Function GetExamplesByDocumentType(documentTypeContext As String, mapping As SemanticStyleMapping) As String
        Dim sb As New StringBuilder()

        ' 公文示例
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("公文") OrElse documentTypeContext.Contains("GB/T 9704")) Then
            sb.AppendLine("公文文档标注示例：")
            sb.AppendLine("「XX市人民政府文件」 → header.org（理由：位于文档开头，符合发文机关标志模式）")
            sb.AppendLine("「×政发〔2024〕15号」 → header.refno（理由：包含发文字号格式〔〕X号）")
            sb.AppendLine("「签发人：王××」 → header.signer（理由：包含签发人标识）")
            sb.AppendLine("「关于加强安全生产工作的通知」 → title.main（理由：公文标题，关于…的…格式）")
            sb.AppendLine("「各区县人民政府，市政府各部门：」 → title.recipient（理由：主送机关，以冒号结尾）")
            sb.AppendLine("「一、总体要求」 → title.1（理由：一级编号标题）")
            sb.AppendLine("「（一）基本原则」 → title.2（理由：二级编号标题）")
            sb.AppendLine("「1. 加强组织领导」 → title.3（理由：三级编号标题）")
            sb.AppendLine("「为进一步做好安全生产工作，根据...」 → body.normal（理由：正文段落）")
            sb.AppendLine("「附件：1. 工作方案」 → body.attachment（理由：附件说明）")
            sb.AppendLine("「XX市人民政府」 → footer.signature（理由：文末机构名称，落款）")
            sb.AppendLine("「2024年1月15日」 → footer.date（理由：文末日期格式）")
            sb.AppendLine("「（联系人：张三，电话：12345678）」 → footer.note（理由：公文附注）")
            sb.AppendLine("「抄送：市委各部门」 → footer.cc（理由：以""抄送""开头）")
            Return sb.ToString()
        End If

        ' 学术论文示例
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("学术") OrElse documentTypeContext.Contains("论文")) Then
            sb.AppendLine("学术论文文档标注示例：")
            sb.AppendLine("「基于深度学习的图像识别技术研究」 → title.main（理由：论文标题）")
            sb.AppendLine("「摘要」 → title.abstract（理由：摘要标题）")
            sb.AppendLine("「本文提出了一种新的...」 → body.abstract（理由：摘要正文）")
            sb.AppendLine("「关键词」 → title.keywords（理由：关键词标题）")
            sb.AppendLine("「深度学习；图像识别」 → body.keywords（理由：关键词内容）")
            sb.AppendLine("「第1章 引言」 → heading.1（理由：章节标题）")
            sb.AppendLine("「1.1 研究背景」 → heading.2（理由：二级编号标题）")
            sb.AppendLine("「近年来，随着人工智能技术...」 → body.normal（理由：正文段落）")
            sb.AppendLine("「参考文献」 → title.references（理由：参考文献标题）")
            Return sb.ToString()
        End If

        ' 商务报告示例
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("商务") OrElse documentTypeContext.Contains("报告")) Then
            sb.AppendLine("商务报告文档标注示例：")
            sb.AppendLine("「2024年度工作总结报告」 → title.main（理由：报告标题）")
            sb.AppendLine("「一、年度业绩回顾」 → heading.1（理由：一级标题）")
            sb.AppendLine("「（一）销售收入分析」 → heading.2（理由：二级标题）")
            sb.AppendLine("「2024年公司实现销售收入增长15%...」 → body.normal（理由：正文段落）")
            sb.AppendLine("「综上所述...」 → body.summary（理由：总结段落）")
            Return sb.ToString()
        End If

        ' 通用文档示例（默认）
        sb.AppendLine("通用文档标注示例：")
        sb.AppendLine("「第一章 总则」 → heading.1（理由：章节标题）")
        sb.AppendLine("「1.1 目的和依据」 → heading.2（理由：二级编号标题）")
        sb.AppendLine("「1.1.1 为规范...」 → heading.3（理由：三级编号标题）")
        sb.AppendLine("「本条例旨在...」 → body.normal（理由：正文段落）")

        ' 如果mapping中有自定义标签，也展示一下
        If mapping IsNot Nothing AndAlso mapping.SemanticTags.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("当前标准支持的特殊标签：")
            For Each tag In mapping.SemanticTags.Take(6)
                If tag.TagId.StartsWith("header.") OrElse tag.TagId.StartsWith("title.") OrElse tag.TagId.StartsWith("footer.") Then
                    sb.AppendLine($"- {tag.TagId}: {tag.DisplayName}")
                End If
            Next
        End If

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建带重试提示的标注提示词（当校验失败时使用）
    ''' </summary>
    Public Shared Function BuildRetryPrompt(
        mapping As SemanticStyleMapping,
        paragraphs As List(Of String),
        errors As List(Of String)) As String

        Dim sb As New StringBuilder()

        ' 原始提示词
        sb.Append(BuildSemanticTaggingPrompt(mapping, paragraphs))
        sb.AppendLine()

        ' 错误反馈
        sb.AppendLine("【上次输出存在以下错误，请修正】")
        For Each errMsg In errors
            sb.AppendLine($"- {errMsg}")
        Next
        sb.AppendLine()
        sb.AppendLine("请重新输出正确的JSON数组（包含paraIndex、tag、reason三个字段）。")

        Return sb.ToString()
    End Function
End Class
