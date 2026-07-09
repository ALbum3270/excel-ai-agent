' WordAi\Services\FormattingIntentCompiler.vb
' 将自然语言排版/格式指令编译为结构化计划；不直接访问 Word COM。

Imports System.Text
Imports System.Text.RegularExpressions

Namespace Services

    Public Enum FormattingTargetScope
        AutoScope
        Selection
        Document
        CurrentParagraph
        Headings
        Body
    End Enum

    Public Enum FormattingOperationKind
        FontSizeDelta
        FontSizeAbsolute
        FontFamily
        Bold
        Italic
        Underline
        FontColor
        Alignment
        LineSpacing
        FirstLineIndent
    End Enum

    Public Class FormattingOperation
        Public Property Kind As FormattingOperationKind
        Public Property Scope As FormattingTargetScope = FormattingTargetScope.AutoScope
        Public Property NumericValue As Double
        Public Property TextValue As String
        Public Property BooleanValue As Boolean
        Public Property HasBooleanValue As Boolean
        Public Property Explanation As String
    End Class

    Public Class FormattingIntentPlan
        Public Property OriginalText As String
        Public Property Source As String = "rules"
        Public Property Confidence As Double = 0
        Public Property Scope As FormattingTargetScope = FormattingTargetScope.AutoScope
        Public Property Operations As New List(Of FormattingOperation)()
        Public Property Notes As New List(Of String)()

        Public ReadOnly Property HasOperations As Boolean
            Get
                Return Operations IsNot Nothing AndAlso Operations.Count > 0
            End Get
        End Property

        Public Function ToHumanReadableSummary() As String
            If Not HasOperations Then Return "未识别到可执行格式操作"

            Dim sb As New StringBuilder()
            sb.Append($"范围: {ScopeToText(Scope)}；操作: ")
            Dim parts As New List(Of String)()
            For Each op In Operations
                parts.Add(OperationToText(op))
            Next
            sb.Append(String.Join("，", parts))
            Return sb.ToString()
        End Function

        Private Shared Function ScopeToText(scope As FormattingTargetScope) As String
            Select Case scope
                Case FormattingTargetScope.Selection : Return "当前选区"
                Case FormattingTargetScope.Document : Return "全文"
                Case FormattingTargetScope.CurrentParagraph : Return "当前段落"
                Case FormattingTargetScope.Headings : Return "标题段落"
                Case FormattingTargetScope.Body : Return "正文段落"
                Case Else : Return "自动"
            End Select
        End Function

        Private Shared Function OperationToText(op As FormattingOperation) As String
            Select Case op.Kind
                Case FormattingOperationKind.FontSizeDelta
                    Return $"字号{If(op.NumericValue >= 0, "+", "")}{op.NumericValue}pt"
                Case FormattingOperationKind.FontSizeAbsolute
                    Return $"字号设为 {op.NumericValue}pt"
                Case FormattingOperationKind.FontFamily
                    Return $"字体设为 {op.TextValue}"
                Case FormattingOperationKind.Bold
                    Return If(op.BooleanValue, "加粗", "取消加粗")
                Case FormattingOperationKind.Italic
                    Return If(op.BooleanValue, "斜体", "取消斜体")
                Case FormattingOperationKind.Underline
                    Return "下划线"
                Case FormattingOperationKind.FontColor
                    Return $"颜色 {op.TextValue}"
                Case FormattingOperationKind.Alignment
                    Return $"对齐 {op.TextValue}"
                Case FormattingOperationKind.LineSpacing
                    Return $"行距 {op.NumericValue}"
                Case FormattingOperationKind.FirstLineIndent
                    Return $"首行缩进 {op.NumericValue} 字符"
                Case Else
                    Return op.Kind.ToString()
            End Select
        End Function
    End Class

    Public Class FormattingExecutionResult
        Public Property Plan As FormattingIntentPlan
        Public Property Success As Boolean
        Public Property AppliedRangeCount As Integer
        Public Property AppliedOperationCount As Integer
        Public Property ErrorMessage As String

        Public Function ToHumanReadableSummary() As String
            If Plan Is Nothing Then
                Return If(String.IsNullOrWhiteSpace(ErrorMessage), "未生成可执行格式计划", ErrorMessage)
            End If

            Dim summary = Plan.ToHumanReadableSummary()
            If Success Then
                Return $"{summary}；已应用到 {AppliedRangeCount} 个范围，执行 {AppliedOperationCount} 个操作"
            End If

            If Not String.IsNullOrWhiteSpace(ErrorMessage) Then
                Return $"{summary}；未应用：{ErrorMessage}"
            End If

            Return $"{summary}；未找到可应用目标"
        End Function
    End Class

    Public Class FormattingIntentCompiler

        Public Shared Function LooksLikeDirectFormattingCommand(message As String) As Boolean
            If String.IsNullOrWhiteSpace(message) Then Return False
            Dim text = message.Trim()
            Dim hasTarget = ContainsAny(text, {"字体", "字号", "行距", "缩进", "对齐", "颜色", "加粗", "粗体", "下划线", "宋体", "黑体", "仿宋", "楷体", "微软雅黑", "小标宋"})
            Dim hasAction = ContainsAny(text, {"统一", "改为", "改成", "设置", "设为", "加大", "增大", "调大", "放大", "减小", "缩小", "调小", "居中", "左对齐", "右对齐", "两端对齐"})
            Return hasTarget AndAlso hasAction
        End Function

        Public Function Compile(message As String, hasSelection As Boolean) As FormattingIntentPlan
            Dim plan As New FormattingIntentPlan With {
                .OriginalText = If(message, ""),
                .Scope = InferScope(message, hasSelection)
            }

            If String.IsNullOrWhiteSpace(message) Then Return plan

            AddFontSizeOperations(plan, message)
            AddFontFamilyOperation(plan, message)
            AddBasicStyleOperations(plan, message)
            AddParagraphOperations(plan, message)

            plan.Confidence = If(plan.HasOperations, 0.86, 0.1)
            If plan.Scope = FormattingTargetScope.Document AndAlso message.Contains("统一") Then
                plan.Notes.Add("检测到统一，默认作用于全文。")
            End If

            Return plan
        End Function

        Private Function InferScope(message As String, hasSelection As Boolean) As FormattingTargetScope
            Dim text = If(message, "")
            If ContainsAny(text, {"选中", "选区", "所选", "当前选择"}) Then Return FormattingTargetScope.Selection
            If ContainsAny(text, {"全文", "整篇", "整个文档", "全部", "所有", "统一"}) Then Return FormattingTargetScope.Document
            If ContainsAny(text, {"所有标题", "全部标题", "标题"}) Then Return FormattingTargetScope.Headings
            If ContainsAny(text, {"正文"}) Then Return FormattingTargetScope.Body
            If ContainsAny(text, {"当前段", "本段", "这一段"}) Then Return FormattingTargetScope.CurrentParagraph
            If hasSelection Then Return FormattingTargetScope.Selection
            Return FormattingTargetScope.Document
        End Function

        Private Sub AddFontSizeOperations(plan As FormattingIntentPlan, message As String)
            Dim inc = Regex.Match(message, "(加大|增大|调大|放大|变大|大)\s*([一二两三四五六七八九十\d]+)?\s*(号|磅|pt|点)?", RegexOptions.IgnoreCase)
            If inc.Success Then
                plan.Operations.Add(New FormattingOperation With {
                    .Kind = FormattingOperationKind.FontSizeDelta,
                    .NumericValue = Math.Max(0.5, ParseAmount(inc.Groups(2).Value, 1)),
                    .Explanation = "字号增量"
                })
                Return
            End If

            Dim dec = Regex.Match(message, "(减小|缩小|调小|变小|小)\s*([一二两三四五六七八九十\d]+)?\s*(号|磅|pt|点)?", RegexOptions.IgnoreCase)
            If dec.Success Then
                plan.Operations.Add(New FormattingOperation With {
                    .Kind = FormattingOperationKind.FontSizeDelta,
                    .NumericValue = -Math.Max(0.5, ParseAmount(dec.Groups(2).Value, 1)),
                    .Explanation = "字号减量"
                })
                Return
            End If

            Dim named = Regex.Match(message, "(设为|设置为|改为|改成|统一为).{0,8}(初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|六号|小六|七号|八号)")
            If named.Success Then
                Dim pt = NamedFontSizeToPoint(named.Groups(2).Value)
                If pt > 0 Then
                    plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.FontSizeAbsolute, .NumericValue = pt})
                    Return
                End If
            End If

            Dim absolute = Regex.Match(message, "(设为|设置为|改为|改成|统一为).{0,8}(\d+(?:\.\d+)?)\s*(pt|磅|点)")
            If absolute.Success Then
                Dim pt As Double
                If Double.TryParse(absolute.Groups(2).Value, pt) Then
                    plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.FontSizeAbsolute, .NumericValue = pt})
                End If
            End If
        End Sub

        Private Sub AddFontFamilyOperation(plan As FormattingIntentPlan, message As String)
            For Each fontName In New String() {"仿宋_GB2312", "楷体_GB2312", "方正小标宋简体", "微软雅黑", "宋体", "黑体", "仿宋", "楷体"}
                If message.Contains(fontName) Then
                    plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.FontFamily, .TextValue = fontName})
                    Return
                End If
            Next
        End Sub

        Private Sub AddBasicStyleOperations(plan As FormattingIntentPlan, message As String)
            If ContainsAny(message, {"取消加粗", "不加粗"}) Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Bold, .BooleanValue = False, .HasBooleanValue = True})
            ElseIf ContainsAny(message, {"加粗", "粗体"}) Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Bold, .BooleanValue = True, .HasBooleanValue = True})
            End If

            If message.Contains("下划线") Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Underline, .BooleanValue = True, .HasBooleanValue = True})
            End If

            Dim colors As New Dictionary(Of String, String) From {
                {"红色", "red"},
                {"蓝色", "blue"},
                {"黑色", "black"},
                {"绿色", "green"}
            }
            For Each kvp In colors
                If message.Contains(kvp.Key) Then
                    plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.FontColor, .TextValue = kvp.Value})
                    Exit For
                End If
            Next
        End Sub

        Private Sub AddParagraphOperations(plan As FormattingIntentPlan, message As String)
            If ContainsAny(message, {"居中", "居中对齐"}) Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Alignment, .TextValue = "center"})
            ElseIf ContainsAny(message, {"左对齐", "靠左"}) Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Alignment, .TextValue = "left"})
            ElseIf ContainsAny(message, {"右对齐", "靠右"}) Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Alignment, .TextValue = "right"})
            ElseIf message.Contains("两端对齐") Then
                plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.Alignment, .TextValue = "justify"})
            End If

            Dim spacing = Regex.Match(message, "行距.{0,4}(\d+(?:\.\d+)?)")
            If spacing.Success Then
                Dim value As Double
                If Double.TryParse(spacing.Groups(1).Value, value) Then
                    plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.LineSpacing, .NumericValue = value})
                End If
            End If

            Dim indent = Regex.Match(message, "首行缩进.{0,4}(\d+(?:\.\d+)?)")
            If indent.Success Then
                Dim value As Double
                If Double.TryParse(indent.Groups(1).Value, value) Then
                    plan.Operations.Add(New FormattingOperation With {.Kind = FormattingOperationKind.FirstLineIndent, .NumericValue = value})
                End If
            End If
        End Sub

        Private Shared Function ContainsAny(text As String, keywords As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return False
            For Each keyword In keywords
                If text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next
            Return False
        End Function

        Private Shared Function ParseAmount(value As String, defaultValue As Double) As Double
            If String.IsNullOrWhiteSpace(value) Then Return defaultValue
            Dim numeric As Double
            If Double.TryParse(value, numeric) Then Return numeric

            Select Case value.Trim()
                Case "一" : Return 1
                Case "二", "两" : Return 2
                Case "三" : Return 3
                Case "四" : Return 4
                Case "五" : Return 5
                Case "六" : Return 6
                Case "七" : Return 7
                Case "八" : Return 8
                Case "九" : Return 9
                Case "十" : Return 10
                Case Else : Return defaultValue
            End Select
        End Function

        Private Shared Function NamedFontSizeToPoint(sizeName As String) As Double
            Select Case sizeName
                Case "初号" : Return 42
                Case "小初" : Return 36
                Case "一号" : Return 26
                Case "小一" : Return 24
                Case "二号" : Return 22
                Case "小二" : Return 18
                Case "三号" : Return 16
                Case "小三" : Return 15
                Case "四号" : Return 14
                Case "小四" : Return 12
                Case "五号" : Return 10.5
                Case "小五" : Return 9
                Case "六号" : Return 7.5
                Case "小六" : Return 6.5
                Case "七号" : Return 5.5
                Case "八号" : Return 5
                Case Else : Return 0
            End Select
        End Function
    End Class

End Namespace
