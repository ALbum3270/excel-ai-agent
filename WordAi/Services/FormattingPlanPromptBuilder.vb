' WordAi\Services\FormattingPlanPromptBuilder.vb
' 构建 Word 排版计划 JSON 的提示词；不访问 Word COM。

Imports System.Text

Namespace Services

    Public Class FormattingPlanPromptBuilder

        Public Shared Function BuildPlanPrompt(userRequest As String, documentSummary As String) As String
            Dim sb As New StringBuilder()

            sb.AppendLine("你是 Word 文档排版规划器。请把用户的自然语言需求转换为可执行的 FormattingIntentPlan JSON。")
            sb.AppendLine("只输出 JSON，不要输出 Markdown，不要解释。")
            sb.AppendLine()
            sb.AppendLine("【用户需求】")
            sb.AppendLine(If(userRequest, ""))
            sb.AppendLine()
            sb.AppendLine("【当前文档摘要】")
            sb.AppendLine(If(documentSummary, ""))
            sb.AppendLine()
            sb.AppendLine("【JSON 结构】")
            sb.AppendLine("{")
            sb.AppendLine("  ""OriginalText"": ""用户原始请求"",")
            sb.AppendLine("  ""Source"": ""llm"",")
            sb.AppendLine("  ""Confidence"": 0.0,")
            sb.AppendLine("  ""Scope"": ""Document"",")
            sb.AppendLine("  ""Operations"": [")
            sb.AppendLine("    {")
            sb.AppendLine("      ""Kind"": ""FontSizeDelta"",")
            sb.AppendLine("      ""Scope"": ""AutoScope"",")
            sb.AppendLine("      ""NumericValue"": 2,")
            sb.AppendLine("      ""TextValue"": """",")
            sb.AppendLine("      ""BooleanValue"": false,")
            sb.AppendLine("      ""HasBooleanValue"": false,")
            sb.AppendLine("      ""Explanation"": ""为什么执行这个操作""")
            sb.AppendLine("    }")
            sb.AppendLine("  ],")
            sb.AppendLine("  ""Notes"": [""规划备注""]")
            sb.AppendLine("}")
            sb.AppendLine()
            sb.AppendLine("【Scope 可选值】")
            sb.AppendLine("- AutoScope")
            sb.AppendLine("- Selection")
            sb.AppendLine("- Document")
            sb.AppendLine("- CurrentParagraph")
            sb.AppendLine("- Headings")
            sb.AppendLine("- Body")
            sb.AppendLine()
            sb.AppendLine("【Kind 可选值】")
            sb.AppendLine("- FontSizeDelta: 字号增减，NumericValue 为 pt 增量")
            sb.AppendLine("- FontSizeAbsolute: 设置绝对字号，NumericValue 为 pt")
            sb.AppendLine("- FontFamily: 设置字体，TextValue 为字体名")
            sb.AppendLine("- Bold: 加粗或取消加粗，BooleanValue 表示目标值，HasBooleanValue=true")
            sb.AppendLine("- Italic: 斜体或取消斜体")
            sb.AppendLine("- Underline: 设置下划线")
            sb.AppendLine("- FontColor: 设置字体颜色，TextValue 可用 red/blue/green/black")
            sb.AppendLine("- Alignment: 段落对齐，TextValue 可用 left/center/right/justify")
            sb.AppendLine("- LineSpacing: 行距倍数，NumericValue 如 1.5")
            sb.AppendLine("- FirstLineIndent: 首行缩进字符数，NumericValue 如 2")
            sb.AppendLine()
            sb.AppendLine("【规划规则】")
            sb.AppendLine("1. 如果用户说[统一][全文][全部]，Scope 优先为 Document。")
            sb.AppendLine("2. 如果用户说[标题]，Scope 优先为 Headings。")
            sb.AppendLine("3. 如果用户说[正文]，Scope 优先为 Body。")
            sb.AppendLine("4. 如果用户没有明确范围，但当前有选区，Scope 可为 Selection；否则为 Document。")
            sb.AppendLine("5. 不要生成当前执行器不支持的 Kind。")
            sb.AppendLine("6. 低置信度时也要给出保守计划，并在 Notes 说明不确定点。")

            Return sb.ToString()
        End Function

    End Class

End Namespace
