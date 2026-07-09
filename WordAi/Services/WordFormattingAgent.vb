' WordAi\Services\WordFormattingAgent.vb
' Word 专属排版 Agent 外壳：承接计划、执行和解释；具体 Word COM 操作仍由 SmartFormatter 执行。

Imports Word = Microsoft.Office.Interop.Word
Imports System.Linq

Namespace Services

    Public Enum WordFormattingAgentStepStatus
        Pending
        Running
        Completed
        Failed
    End Enum

    Public Enum WordFormattingTaskKind
        DirectFormatting
        SemanticReformat
    End Enum

    Public Class WordFormattingTaskPlan
        Public Property TaskId As String = Guid.NewGuid().ToString("N")
        Public Property Kind As WordFormattingTaskKind = WordFormattingTaskKind.SemanticReformat
        Public Property OriginalRequest As String = ""
        Public Property ScopeSummary As String = ""
        Public Property TargetSummary As String = ""
        Public Property StandardName As String = ""
        Public Property RequiresPreview As Boolean = True
        Public Property Operations As New List(Of String)()

        Public ReadOnly Property KindLabel As String
            Get
                Select Case Kind
                    Case WordFormattingTaskKind.DirectFormatting
                        Return "直接格式调整"
                    Case Else
                        Return "语义智能排版"
                End Select
            End Get
        End Property

        Public Function ToHumanReadableSummary() As String
            Dim parts As New List(Of String)()
            parts.Add($"任务: {KindLabel}")
            If Not String.IsNullOrWhiteSpace(ScopeSummary) Then parts.Add($"范围: {ScopeSummary}")
            If Not String.IsNullOrWhiteSpace(StandardName) Then parts.Add($"标准: {StandardName}")
            If Not String.IsNullOrWhiteSpace(TargetSummary) Then parts.Add(TargetSummary)
            If Operations IsNot Nothing AndAlso Operations.Count > 0 Then
                parts.Add("计划: " & String.Join("；", Operations.Take(4)))
            End If
            Return String.Join("；", parts)
        End Function

        Public Shared Function FromDirectFormatting(userRequest As String,
                                                    plan As FormattingIntentPlan) As WordFormattingTaskPlan
            Dim task As New WordFormattingTaskPlan With {
                .Kind = WordFormattingTaskKind.DirectFormatting,
                .OriginalRequest = If(userRequest, ""),
                .RequiresPreview = False
            }

            If plan IsNot Nothing Then
                Dim summary = plan.ToHumanReadableSummary()
                Dim segments = summary.Split("；"c)
                If segments.Length > 0 Then
                    task.ScopeSummary = segments(0).Replace("范围: ", "")
                End If
                If segments.Length > 1 Then
                    task.Operations.Add(segments(1).Replace("操作: ", ""))
                Else
                    task.Operations.Add(summary)
                End If
            End If

            Return task
        End Function

        Public Shared Function FromSemanticReformat(userRequest As String,
                                                    scopeSummary As String,
                                                    standardName As String,
                                                    targetSummary As String,
                                                    changeCount As Integer) As WordFormattingTaskPlan
            Dim task As New WordFormattingTaskPlan With {
                .Kind = WordFormattingTaskKind.SemanticReformat,
                .OriginalRequest = If(userRequest, ""),
                .ScopeSummary = If(scopeSummary, ""),
                .StandardName = If(standardName, ""),
                .TargetSummary = If(targetSummary, ""),
                .RequiresPreview = True
            }

            If changeCount > 0 Then
                task.Operations.Add($"预览 {changeCount} 处结构/样式调整，确认后应用")
            Else
                task.Operations.Add("先完成结构分析，再根据用户微调继续生成调整")
            End If

            Return task
        End Function
    End Class

    Public Class WordFormattingAgentStep
        Public Property Name As String
        Public Property Description As String
        Public Property Status As WordFormattingAgentStepStatus = WordFormattingAgentStepStatus.Pending
        Public Property ResultSummary As String
        Public Property ErrorMessage As String
    End Class

    Public Class WordFormattingAgentResult
        Public Property Success As Boolean
        Public Property TaskPlan As WordFormattingTaskPlan
        Public Property FormattingResult As FormattingExecutionResult
        Public Property ExecutionSummary As String
        Public Property AppliedCount As Integer
        Public Property ExpectedCount As Integer
        Public Property RepairCount As Integer
        Public Property Steps As New List(Of WordFormattingAgentStep)()
        Public Property ErrorMessage As String

        Public Function ToHumanReadableSummary() As String
            If Not String.IsNullOrWhiteSpace(ExecutionSummary) Then
                If TaskPlan IsNot Nothing Then
                    Return TaskPlan.ToHumanReadableSummary() & "；执行: " & ExecutionSummary
                End If
                Return ExecutionSummary
            End If

            If FormattingResult IsNot Nothing Then
                Dim summary = FormattingResult.ToHumanReadableSummary()
                For Each stepInfo In Steps
                    If stepInfo.Name = "Observe" AndAlso
                       stepInfo.Status = WordFormattingAgentStepStatus.Completed AndAlso
                       Not String.IsNullOrWhiteSpace(stepInfo.ResultSummary) Then
                        Return summary & "；" & stepInfo.ResultSummary
                    End If
                Next
                Return summary
            End If
            If Not String.IsNullOrWhiteSpace(ErrorMessage) Then Return ErrorMessage
            Return If(Success, "排版任务已完成", "排版任务未完成")
        End Function

        Public Shared Function FromSemanticReformat(taskPlan As WordFormattingTaskPlan,
                                                    appliedCount As Integer,
                                                    expectedCount As Integer,
                                                    repairCount As Integer,
                                                    detailSummary As String) As WordFormattingAgentResult
            Dim result As New WordFormattingAgentResult With {
                .Success = appliedCount > 0 OrElse expectedCount = 0,
                .TaskPlan = taskPlan,
                .AppliedCount = appliedCount,
                .ExpectedCount = expectedCount,
                .RepairCount = repairCount
            }

            Dim parts As New List(Of String)()
            parts.Add($"已应用 {appliedCount} / {expectedCount} 个文本段落")
            If repairCount > 0 Then parts.Add($"自动修复 {repairCount} 个结构段落")
            If Not String.IsNullOrWhiteSpace(detailSummary) Then parts.Add(detailSummary)
            result.ExecutionSummary = String.Join("；", parts)
            Return result
        End Function
    End Class

    Public Class WordFormattingAgent

        Private ReadOnly _app As Word.Application

        Public Sub New(app As Word.Application)
            _app = app
        End Sub

        Public Function RunDirectFormatting(userRequest As String) As WordFormattingAgentResult
            Dim agentResult As New WordFormattingAgentResult()
            Dim planStep = AddStep(agentResult, "Plan", "理解用户排版需求并生成格式计划")
            Dim applyStep = AddStep(agentResult, "Apply", "把格式计划应用到当前 Word 文档")
            Dim observeStep = AddStep(agentResult, "Observe", "观察排版执行结果是否已作用到文档")
            Dim explainStep = AddStep(agentResult, "Explain", "生成用户可理解的执行摘要")

            Try
                planStep.Status = WordFormattingAgentStepStatus.Running
                Dim compiler As New FormattingIntentCompiler()
                Dim plan = compiler.Compile(userRequest, HasUsableSelection())
                agentResult.TaskPlan = WordFormattingTaskPlan.FromDirectFormatting(userRequest, plan)
                planStep.ResultSummary = If(plan Is Nothing, "未生成计划", plan.ToHumanReadableSummary())
                planStep.Status = If(plan IsNot Nothing AndAlso plan.HasOperations,
                                     WordFormattingAgentStepStatus.Completed,
                                     WordFormattingAgentStepStatus.Failed)

                If planStep.Status = WordFormattingAgentStepStatus.Failed Then
                    agentResult.ErrorMessage = "未识别到可执行格式操作"
                    Return agentResult
                End If

                applyStep.Status = WordFormattingAgentStepStatus.Running
                Dim formatter As New SmartFormatter(_app)
                Dim formattingResult = formatter.ApplyNaturalLanguageFormatDetailed(userRequest)
                agentResult.FormattingResult = formattingResult
                applyStep.ResultSummary = If(formattingResult Is Nothing, "没有执行结果", formattingResult.ToHumanReadableSummary())
                applyStep.Status = If(formattingResult IsNot Nothing AndAlso formattingResult.Success,
                                      WordFormattingAgentStepStatus.Completed,
                                      WordFormattingAgentStepStatus.Failed)

                If applyStep.Status = WordFormattingAgentStepStatus.Failed Then
                    agentResult.ErrorMessage = applyStep.ResultSummary
                    Return agentResult
                End If

                observeStep.Status = WordFormattingAgentStepStatus.Running
                Dim observeSummary As String = ""
                If ObserveFormattingResult(formattingResult, observeSummary) Then
                    observeStep.Status = WordFormattingAgentStepStatus.Completed
                    observeStep.ResultSummary = observeSummary
                Else
                    observeStep.Status = WordFormattingAgentStepStatus.Failed
                    observeStep.ErrorMessage = observeSummary
                    agentResult.ErrorMessage = observeSummary
                    Return agentResult
                End If

                explainStep.Status = WordFormattingAgentStepStatus.Completed
                explainStep.ResultSummary = formattingResult.ToHumanReadableSummary() & "；" & observeSummary
                agentResult.Success = True
                Return agentResult
            Catch ex As Exception
                agentResult.ErrorMessage = ex.Message
                MarkRunningStepsFailed(agentResult, ex.Message)
                Return agentResult
            End Try
        End Function

        Private Function AddStep(result As WordFormattingAgentResult, name As String, description As String) As WordFormattingAgentStep
            Dim stepInfo As New WordFormattingAgentStep With {
                .Name = name,
                .Description = description
            }
            result.Steps.Add(stepInfo)
            Return stepInfo
        End Function

        Private Function ObserveFormattingResult(formattingResult As FormattingExecutionResult, ByRef summary As String) As Boolean
            If formattingResult Is Nothing Then
                summary = "没有可观察的排版执行结果"
                Return False
            End If

            Try
                If _app Is Nothing OrElse _app.ActiveDocument Is Nothing Then
                    summary = "无法观察 Word 文档状态：没有活动文档"
                    Return False
                End If
            Catch ex As Exception
                summary = "无法观察 Word 文档状态：" & ex.Message
                Return False
            End Try

            If Not formattingResult.Success Then
                summary = If(String.IsNullOrWhiteSpace(formattingResult.ErrorMessage),
                             "排版执行未成功，无法进入观察确认",
                             formattingResult.ErrorMessage)
                Return False
            End If

            If formattingResult.AppliedRangeCount <= 0 OrElse formattingResult.AppliedOperationCount <= 0 Then
                summary = "执行器没有报告有效应用范围或操作数量"
                Return False
            End If

            Dim confirmations As New List(Of String)()
            Dim failures As New List(Of String)()
            Dim sampleRanges = ResolveObservationRanges(formattingResult.Plan)

            If sampleRanges.Count = 0 Then
                summary = $"执行器已应用 {formattingResult.AppliedRangeCount} 个范围、{formattingResult.AppliedOperationCount} 个操作；但未能读取可观察样本"
                Return True
            End If

            For Each op In formattingResult.Plan.Operations
                ObserveOperation(sampleRanges, op, confirmations, failures)
            Next

            If failures.Count > 0 Then
                summary = "观察到部分格式未达到预期：" & String.Join("；", failures.Take(3))
                Return False
            End If

            Dim observedText = If(confirmations.Count > 0, String.Join("；", confirmations.Distinct().Take(4)), "已读取文档样本")
            summary = $"已观察到排版应用结果：{formattingResult.AppliedRangeCount} 个范围，{formattingResult.AppliedOperationCount} 个操作；{observedText}"
            Return True
        End Function

        Private Function ResolveObservationRanges(plan As FormattingIntentPlan) As List(Of Word.Range)
            Dim result As New List(Of Word.Range)()
            If plan Is Nothing Then Return result

            Try
                Dim doc = _app.ActiveDocument
                If doc Is Nothing Then Return result

                Select Case plan.Scope
                    Case FormattingTargetScope.Selection
                        If HasUsableSelection() Then AddSampleRanges(result, _app.Selection.Range)
                    Case FormattingTargetScope.CurrentParagraph
                        If _app.Selection IsNot Nothing AndAlso _app.Selection.Paragraphs.Count > 0 Then
                            result.Add(_app.Selection.Paragraphs(1).Range)
                        End If
                    Case FormattingTargetScope.Headings
                        For Each para As Word.Paragraph In doc.Paragraphs
                            If IsHeadingParagraph(para) Then result.Add(para.Range)
                            If result.Count >= 20 Then Exit For
                        Next
                    Case FormattingTargetScope.Body
                        For Each para As Word.Paragraph In doc.Paragraphs
                            If Not IsHeadingParagraph(para) AndAlso HasVisibleText(para.Range.Text) Then result.Add(para.Range)
                            If result.Count >= 20 Then Exit For
                        Next
                    Case Else
                        AddSampleRanges(result, doc.Content)
                End Select

                If result.Count = 0 Then AddSampleRanges(result, doc.Content)
            Catch
            End Try

            Return result
        End Function

        Private Sub AddSampleRanges(result As List(Of Word.Range), sourceRange As Word.Range)
            If sourceRange Is Nothing Then Return

            Try
                If sourceRange.Paragraphs IsNot Nothing AndAlso sourceRange.Paragraphs.Count > 0 Then
                    For Each para As Word.Paragraph In sourceRange.Paragraphs
                        If HasVisibleText(para.Range.Text) Then result.Add(para.Range)
                        If result.Count >= 20 Then Exit For
                    Next
                Else
                    result.Add(sourceRange)
                End If
            Catch
                result.Add(sourceRange)
            End Try
        End Sub

        Private Function IsHeadingParagraph(para As Word.Paragraph) As Boolean
            Try
                If para.OutlineLevel <> Word.WdOutlineLevel.wdOutlineLevelBodyText Then Return True
                Dim styleName = para.Style?.NameLocal?.ToString()
                Return Not String.IsNullOrWhiteSpace(styleName) AndAlso
                       (styleName.Contains("标题") OrElse styleName.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0)
            Catch
                Return False
            End Try
        End Function

        Private Function HasVisibleText(text As String) As Boolean
            If text Is Nothing Then Return False
            Dim cleaned = text.Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
            Return cleaned.Length > 0
        End Function

        Private Sub ObserveOperation(sampleRanges As List(Of Word.Range),
                                     op As FormattingOperation,
                                     confirmations As List(Of String),
                                     failures As List(Of String))
            Dim checkedCount As Integer = 0
            Dim passedCount As Integer = 0

            For Each rng In sampleRanges
                If rng Is Nothing Then Continue For
                checkedCount += 1
                If IsOperationObserved(rng, op) Then passedCount += 1
            Next

            If checkedCount = 0 Then Return

            If op.Kind = FormattingOperationKind.FontSizeDelta OrElse
               op.Kind = FormattingOperationKind.FontSizeGradeDelta Then
                confirmations.Add("字号增量已由执行器应用")
                Return
            End If

            If passedCount = checkedCount Then
                confirmations.Add(OperationObservedText(op))
            Else
                failures.Add($"{OperationObservedText(op)} 仅 {passedCount}/{checkedCount} 个样本符合")
            End If
        End Sub

        Private Function IsOperationObserved(rng As Word.Range, op As FormattingOperation) As Boolean
            Try
                Select Case op.Kind
                    Case FormattingOperationKind.FontSizeAbsolute
                        Return Math.Abs(NormalizeFontSize(rng.Font.Size) - op.NumericValue) <= 0.2
                    Case FormattingOperationKind.FontFamily
                        Dim fontName = If(rng.Font.NameFarEast, rng.Font.Name)
                        Return String.Equals(fontName, op.TextValue, StringComparison.OrdinalIgnoreCase) OrElse
                               String.Equals(rng.Font.Name, op.TextValue, StringComparison.OrdinalIgnoreCase)
                    Case FormattingOperationKind.Bold
                        Return (CInt(rng.Font.Bold) <> 0) = op.BooleanValue
                    Case FormattingOperationKind.Italic
                        Return (CInt(rng.Font.Italic) <> 0) = op.BooleanValue
                    Case FormattingOperationKind.Underline
                        Return rng.Font.Underline <> Word.WdUnderline.wdUnderlineNone
                    Case FormattingOperationKind.FontColor
                        Return IsFontColorObserved(rng, op.TextValue)
                    Case FormattingOperationKind.Alignment
                        Return IsAlignmentObserved(rng, op.TextValue)
                    Case FormattingOperationKind.LineSpacing
                        Return Math.Abs(CDbl(rng.ParagraphFormat.LineSpacing) - (12 * op.NumericValue)) <= 1.0
                    Case FormattingOperationKind.FirstLineIndent
                        Return Math.Abs(CDbl(rng.ParagraphFormat.FirstLineIndent)) > 0.1
                    Case Else
                        Return True
                End Select
            Catch
                Return False
            End Try
        End Function

        Private Function IsFontColorObserved(rng As Word.Range, colorName As String) As Boolean
            Select Case If(colorName, "").ToLowerInvariant()
                Case "red"
                    Return rng.Font.Color = Word.WdColor.wdColorRed
                Case "blue"
                    Return rng.Font.Color = Word.WdColor.wdColorBlue
                Case "green"
                    Return rng.Font.Color = Word.WdColor.wdColorGreen
                Case "black"
                    Return rng.Font.Color = Word.WdColor.wdColorBlack
                Case Else
                    Return True
            End Select
        End Function

        Private Function IsAlignmentObserved(rng As Word.Range, alignment As String) As Boolean
            Select Case If(alignment, "").ToLowerInvariant()
                Case "center"
                    Return rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Case "right"
                    Return rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                Case "justify"
                    Return rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                Case Else
                    Return rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End Select
        End Function

        Private Function NormalizeFontSize(value As Single) As Double
            If value > 0 AndAlso value < 200 Then Return value
            Return 12
        End Function

        Private Function OperationObservedText(op As FormattingOperation) As String
            Select Case op.Kind
                Case FormattingOperationKind.FontSizeGradeDelta
                    Return $"字号等级{If(op.NumericValue >= 0, "+", "")}{op.NumericValue}"
                Case FormattingOperationKind.FontSizeAbsolute
                    Return $"字号={op.NumericValue}pt"
                Case FormattingOperationKind.FontFamily
                    Return $"字体={op.TextValue}"
                Case FormattingOperationKind.Bold
                    Return If(op.BooleanValue, "加粗已生效", "取消加粗已生效")
                Case FormattingOperationKind.Italic
                    Return If(op.BooleanValue, "斜体已生效", "取消斜体已生效")
                Case FormattingOperationKind.Underline
                    Return "下划线已生效"
                Case FormattingOperationKind.FontColor
                    Return $"颜色={op.TextValue}"
                Case FormattingOperationKind.Alignment
                    Return $"对齐={op.TextValue}"
                Case FormattingOperationKind.LineSpacing
                    Return $"行距={op.NumericValue}"
                Case FormattingOperationKind.FirstLineIndent
                    Return $"首行缩进={op.NumericValue}"
                Case Else
                    Return op.Kind.ToString()
            End Select
        End Function

        Private Sub MarkRunningStepsFailed(result As WordFormattingAgentResult, errorMessage As String)
            For Each stepInfo In result.Steps
                If stepInfo.Status = WordFormattingAgentStepStatus.Running OrElse
                   stepInfo.Status = WordFormattingAgentStepStatus.Pending Then
                    stepInfo.Status = WordFormattingAgentStepStatus.Failed
                    stepInfo.ErrorMessage = errorMessage
                End If
            Next
        End Sub

        Private Function HasUsableSelection() As Boolean
            Try
                Dim sel = _app.Selection
                If sel Is Nothing OrElse sel.Range Is Nothing Then Return False
                Dim selectedText = If(sel.Range.Text, "").Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
                Return selectedText.Length > 0
            Catch
                Return False
            End Try
        End Function

    End Class

End Namespace
