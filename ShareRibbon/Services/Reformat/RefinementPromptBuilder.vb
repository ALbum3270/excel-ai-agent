' ShareRibbon\Services\Reformat\RefinementPromptBuilder.vb
' 对话式微调提示词构建器 - 将用户自然语言微调指令+当前映射发送给LLM，获取调整后的映射

Imports System.Text
Imports System.Linq

''' <summary>
''' 微调提示词构建器 - 为LLM构建对话式排版微调的提示词
''' 当前SmartFormattingOrchestrator.ApplyTweakToTags仅用正则匹配关键词，
''' 无法理解"标题再正式一点""正文稍微紧凑些"等模糊指令。
''' 本构建器将当前SemanticStyleMapping序列化为JSON，让LLM直接修改并返回。
''' </summary>
Public Class RefinementPromptBuilder

    ''' <summary>
    ''' 构建微调提示词（让LLM调整当前映射）
    ''' </summary>
    ''' <param name="userCommand">用户的微调指令（如"标题再大一点""正文行距1.5倍"）</param>
    ''' <param name="currentMapping">当前的语义样式映射</param>
    ''' <param name="conversationHistory">对话历史（之前的微调指令列表，用于上下文）</param>
    ''' <param name="documentTypeContext">文档类型上下文描述</param>
    Public Shared Function BuildRefinementPrompt(
        userCommand As String,
        currentMapping As SemanticStyleMapping,
        Optional conversationHistory As List(Of String) = Nothing,
        Optional documentTypeContext As String = Nothing) As String

        Dim sb As New StringBuilder()

        ' ===== 1. 角色 =====
        sb.AppendLine("你是一位排版微调专家。用户希望对当前排版方案进行微调，请根据用户的指令修改对应标签的格式参数。")
        sb.AppendLine()

        ' ===== 2. 当前映射 =====
        sb.AppendLine("【当前排版方案】")
        sb.AppendLine("以下JSON是当前的语义样式映射，每个标签定义了一类段落的格式：")
        sb.AppendLine()
        sb.AppendLine(SerializeMappingForPrompt(currentMapping))
        sb.AppendLine()

        ' ===== 3. 对话历史 =====
        If conversationHistory IsNot Nothing AndAlso conversationHistory.Count > 1 Then
            sb.AppendLine("【之前的微调指令】")
            For i = 0 To conversationHistory.Count - 2
                sb.AppendLine($"  {i + 1}. {conversationHistory(i)}")
            Next
            sb.AppendLine()
        End If

        ' ===== 4. 文档类型上下文 =====
        If Not String.IsNullOrEmpty(documentTypeContext) Then
            sb.AppendLine("【文档类型】")
            sb.AppendLine(documentTypeContext)
            sb.AppendLine()
        End If

        ' ===== 5. 用户指令 =====
        sb.AppendLine("【用户微调指令】")
        sb.AppendLine(userCommand)
        sb.AppendLine()

        ' ===== 6. 输出格式 =====
        sb.AppendLine("【输出要求】")
        sb.AppendLine("只输出修改后的完整semanticTags JSON数组（不要输出其他内容，不要使用markdown代码块）。")
        sb.AppendLine("格式如下：")
        sb.AppendLine("[")
        sb.AppendLine("  {""tagId"":""title.1"",""font"":{""fontNameCN"":""黑体"",""fontNameEN"":""Arial"",""fontSize"":22,""bold"":true,""italic"":false,""underline"":false},")
        sb.AppendLine("   ""paragraph"":{""alignment"":""center"",""firstLineIndent"":0,""lineSpacing"":1.5,""spaceBefore"":0,""spaceAfter"":0},")
        sb.AppendLine("   ""color"":{""fontColor"":""#000000""}},")
        sb.AppendLine("  ...")
        sb.AppendLine("]")
        sb.AppendLine()
        sb.AppendLine("规则：")
        sb.AppendLine("1. 只修改用户指令涉及的标签，其他标签保持原值不变")
        sb.AppendLine("2. 字号单位为pt，行距为倍数（如1.5表示1.5倍行距），缩进单位为字符数")
        sb.AppendLine("3. 字号范围8-72pt，行距范围0.5-3.0，缩进范围0-10字符")
        sb.AppendLine("4. 如果用户指令模糊（如""再正式一点""），请根据文档类型做合理推断")
        sb.AppendLine("5. 必须返回所有标签的完整数据，不要省略任何标签")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建AI语义标注+微调的联合提示词（当用户首次排版时）
    ''' 将文档分析和微调指令合并为一次LLM调用
    ''' </summary>
    ''' <param name="taggingPrompt">SemanticPromptBuilder构建的标注提示词</param>
    ''' <param name="userCommand">用户原始排版指令</param>
    Public Shared Function BuildTaggingWithRefinementPrompt(
        taggingPrompt As String,
        userCommand As String) As String

        Dim sb As New StringBuilder()
        sb.Append(taggingPrompt)
        sb.AppendLine()
        sb.AppendLine("【用户排版要求】")
        sb.AppendLine(userCommand)
        sb.AppendLine()
        sb.AppendLine("请在标注时考虑用户的排版要求。例如：")
        sb.AppendLine("- 如果用户说""按公文排版""，请优先使用公文专用标签（header.org, header.refno等）")
        sb.AppendLine("- 如果用户说""标题要醒目""，在reason中说明该段落应作为标题处理")
        sb.AppendLine("- 如果用户指定了具体格式（如""仿宋三号""），在reason中注明")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 将SemanticStyleMapping序列化为LLM可读的JSON格式
    ''' 只输出semanticTags部分，不包含页面设置等无关信息
    ''' </summary>
    Private Shared Function SerializeMappingForPrompt(mapping As SemanticStyleMapping) As String
        If mapping Is Nothing OrElse mapping.SemanticTags Is Nothing Then Return "[]"

        Dim sb As New StringBuilder()
        sb.AppendLine("[")

        For i = 0 To mapping.SemanticTags.Count - 1
            Dim tag = mapping.SemanticTags(i)
            sb.Append("  {")
            sb.Append($"""tagId"":""{EscapeJson(tag.TagId)}"",")

            ' Font
            If tag.Font IsNot Nothing Then
                sb.Append($"""font"":{{")
                sb.Append($"""fontNameCN"":""{EscapeJson(tag.Font.FontNameCN)}"",")
                sb.Append($"""fontNameEN"":""{EscapeJson(tag.Font.FontNameEN)}"",")
                sb.Append($"""fontSize"":{tag.Font.FontSize},")
                sb.Append($"""bold"":{tag.Font.Bold.ToString().ToLower()},")
                sb.Append($"""italic"":{tag.Font.Italic.ToString().ToLower()},")
                sb.Append($"""underline"":{tag.Font.Underline.ToString().ToLower()}")
                sb.Append("},")
            End If

            ' Paragraph
            If tag.Paragraph IsNot Nothing Then
                sb.Append($"""paragraph"":{{")
                sb.Append($"""alignment"":""{EscapeJson(tag.Paragraph.Alignment)}"",")
                sb.Append($"""firstLineIndent"":{tag.Paragraph.FirstLineIndent},")
                sb.Append($"""lineSpacing"":{tag.Paragraph.LineSpacing},")
                sb.Append($"""spaceBefore"":{tag.Paragraph.SpaceBefore},")
                sb.Append($"""spaceAfter"":{tag.Paragraph.SpaceAfter}")
                sb.Append("},")
            End If

            ' Color
            If tag.Color IsNot Nothing Then
                sb.Append($"""color"":{{")
                sb.Append($"""fontColor"":""{EscapeJson(tag.Color.FontColor)}""")
                sb.Append("}")
            End If

            sb.Append("}")
            If i < mapping.SemanticTags.Count - 1 Then
                sb.Append(",")
            End If
            sb.AppendLine()
        Next

        sb.AppendLine("]")
        Return sb.ToString()
    End Function

    ''' <summary>JSON字符串转义</summary>
    Private Shared Function EscapeJson(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
    End Function
End Class
