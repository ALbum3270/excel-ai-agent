' ShareRibbon\Services\Proofread\ProofreadPromptBuilder.vb
' 校对Prompt构建器 - 构建AI校对Prompt，解析校对结果

Imports System.Collections.Generic
Imports System.Text

''' <summary>
''' 校对Prompt构建器
''' </summary>
Public Class ProofreadPromptBuilder

    ''' <summary>
    ''' 构建全文校对Prompt
    ''' </summary>
    Public Shared Function BuildFullDocumentPrompt(paragraphs As List(Of String)) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("你是专业的中文文档校对专家。请仔细检查以下文档，找出需要修正的问题。")
        sb.AppendLine()
        sb.AppendLine("只输出JSON数组，不要输出任何其他内容（不要使用markdown代码块，不要输出解释说明）。")
        sb.AppendLine()

        sb.AppendLine("【校对范围】")
        sb.AppendLine("1. 错别字和拼写错误")
        sb.AppendLine("2. 词语使用错误（包括但不限于）：")
        sb.AppendLine("   - 的/地/得混用（这是最常见的词语错误）")
        sb.AppendLine("   - 他/她/它在表示指代时的混用")
        sb.AppendLine("   - 其他常见用词错误")
        sb.AppendLine("3. 标点符号错误：")
        sb.AppendLine("   - 中英文标点混用（如中文句子里用了英文逗号）")
        sb.AppendLine("   - 标点缺失或多余")
        sb.AppendLine("   - 引号、括号不匹配")
        sb.AppendLine("4. 语法和语病问题")
        sb.AppendLine("5. 表达不通顺或容易引起歧义的地方")
        sb.AppendLine()

        sb.AppendLine("【最小修改原则】")
        sb.AppendLine("- 只修改确实有问题的内容，不要为了优化表达而改动正确文本")
        sb.AppendLine("- suggestion必须保持原文含义，只修正错误部分")
        sb.AppendLine("- original必须精确匹配原文，包含标点和空格")
        sb.AppendLine()

        sb.AppendLine("【文档内容】")
        For i = 0 To paragraphs.Count - 1
            Dim para = paragraphs(i)
            If Not String.IsNullOrWhiteSpace(para) Then
                sb.AppendLine($"[段落{i}] {para}")
            End If
        Next
        sb.AppendLine()

        sb.AppendLine("【输出格式】")
        sb.AppendLine("请输出纯JSON数组（不要使用markdown代码块包裹）：")
        sb.AppendLine("[")
        sb.AppendLine("  {")
        sb.AppendLine("    ""paragraphIndex"": 0,")
        sb.AppendLine("    ""original"": ""需要修正的原文片段"",")
        sb.AppendLine("    ""suggestion"": ""修正后的文本"",")
        sb.AppendLine("    ""issueType"": ""SpellingError"",")
        sb.AppendLine("    ""severity"": ""High"",")
        sb.AppendLine("    ""explanation"": ""简要说明修改原因""")
        sb.AppendLine("  }")
        sb.AppendLine("]")
        sb.AppendLine()

        sb.AppendLine("【issueType可选值（统一小驼峰格式）】")
        sb.AppendLine("- SpellingError: 拼写错误")
        sb.AppendLine("- WordUsageError: 用词错误")
        sb.AppendLine("- PunctuationError: 标点错误")
        sb.AppendLine("- GrammaticalError: 语法错误")
        sb.AppendLine("- ExpressionError: 表达问题")
        sb.AppendLine()

        sb.AppendLine("【severity可选值】")
        sb.AppendLine("- High: 必须修改（如错别字、严重语法错误）")
        sb.AppendLine("- Medium: 建议修改（如用词不当、轻微语病）")
        sb.AppendLine("- Low: 可选优化（如表达可以更精炼）")
        sb.AppendLine()

        sb.AppendLine("【注意事项】")
        sb.AppendLine("1. original必须精确匹配文档原文，包括标点和空格")
        sb.AppendLine("2. 同一段落有多处问题时，需要返回多个条目")
        sb.AppendLine("3. 只返回需要修改的内容，没问题的段落不要包含在结果中")
        sb.AppendLine("4. 如果文档没有需要修改的内容，请返回空数组：[]")
        sb.AppendLine("5. 不要遗漏明显的问题，但也不要过度修改")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 分析AI返回的校对结果，保留解析失败状态。
    ''' </summary>
    Public Shared Function AnalyzeProofreadResponse(
        aiResponse As String,
        Optional paragraphs As List(Of String) = Nothing) As ProofreadAnalysisResult

        Try
            Dim parseResult = ProofreadJsonParser.Parse(aiResponse)
            Dim analysis As New ProofreadAnalysisResult With {
                .RawResponsePreview = BuildRawResponsePreview(aiResponse),
                .Summary = parseResult.Summary,
                .FormatDetected = parseResult.FormatDetected
            }

            If Not parseResult.Success Then
                Debug.WriteLine($"[ProofreadPromptBuilder] 解析校对结果失败: {parseResult.ErrorMessage}")
                analysis.Status = ProofreadAnalysisStatus.ParseFailed
                analysis.ErrorMessage = parseResult.ErrorMessage
                Return analysis
            End If

            analysis.Issues = If(parseResult.Issues, New List(Of ProofreadIssue)())
            analysis.Status = If(analysis.Issues.Count > 0, ProofreadAnalysisStatus.HasIssues, ProofreadAnalysisStatus.NoIssues)
            Return analysis

        Catch ex As Exception
            Debug.WriteLine($"[ProofreadPromptBuilder] 解析校对结果异常: {ex.Message}")
            Return New ProofreadAnalysisResult With {
                .Status = ProofreadAnalysisStatus.ModelFailed,
                .ErrorMessage = ex.Message,
                .RawResponsePreview = BuildRawResponsePreview(aiResponse)
            }
        End Try
    End Function

    ''' <summary>
    ''' 解析AI返回的校对结果（兼容旧调用；失败时仍返回空列表）。
    ''' 新代码应使用 AnalyzeProofreadResponse 区分状态。
    ''' </summary>
    Public Shared Function ParseProofreadResponse(
        aiResponse As String,
        Optional paragraphs As List(Of String) = Nothing) As List(Of ProofreadIssue)

        Dim analysis = AnalyzeProofreadResponse(aiResponse, paragraphs)
        If analysis Is Nothing OrElse Not analysis.HasIssues Then
            Return New List(Of ProofreadIssue)()
        End If
        Return analysis.Issues
    End Function

    Private Shared Function BuildRawResponsePreview(aiResponse As String) As String
        If String.IsNullOrWhiteSpace(aiResponse) Then Return ""
        Dim normalized = aiResponse.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If normalized.Length <= 500 Then Return normalized
        Return normalized.Substring(0, 500) & "..."
    End Function

End Class
