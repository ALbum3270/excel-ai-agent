Imports System.IO
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace Design

    Public NotInheritable Partial Class ProfessionalDeckExecutor
        Private Const GeneratedShapeTagPrefix As String = "office-ai-design:"

        Private Sub New()
        End Sub

        Public Shared Function ExecuteAsToolResult(params As JObject,
                                                   Optional preview As Boolean = False) As ToolResult
            Const toolId As String = "CreateSlides"
            Dim presentation As PowerPoint.Presentation = Nothing
            Try
                Dim previewRequested = preview
                If params IsNot Nothing AndAlso params("preview") IsNot Nothing Then
                    If params("preview").Type <> JTokenType.Boolean Then
                        Return ToolResult.Failed(toolId,
                                                 "CreateSlides preview must be a boolean",
                                                 errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                                 userMessage:="preview 必须是 true 或 false",
                                                 recoverable:=True)
                    End If
                    previewRequested = previewRequested OrElse params.Value(Of Boolean)("preview")
                End If
                presentation = Globals.ThisAddIn.Application.ActivePresentation
                If presentation Is Nothing Then
                    Return ToolResult.Failed(toolId,
                                             "No active PowerPoint presentation",
                                             errorCode:=ExceptionClassifier.CodeDocMissing,
                                             userMessage:="请先打开或新建一个 PowerPoint 演示文稿",
                                             recoverable:=False)
                End If
                Return ExecuteInternal(params, presentation, previewRequested)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            Finally
                ComObjectHelper.ReleaseComObject(presentation)
            End Try
        End Function

        Private Shared Function ExecuteInternal(params As JObject,
                                                presentation As PowerPoint.Presentation,
                                                preview As Boolean) As ToolResult
            Const toolId As String = "CreateSlides"
            Dim spec = DeckDesignSpec.Parse(params)
            If spec.Slides.Count = 0 Then
                Return ToolResult.Failed(toolId, "CreateSlides requires at least one slide spec",
                                         errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                         userMessage:="没有可创建的幻灯片设计规格", recoverable:=True)
            End If
            If spec.Slides.Count > 50 Then
                Return ToolResult.Failed(toolId, "CreateSlides supports at most 50 slides per run",
                                         errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                         userMessage:="单次最多创建 50 张幻灯片", recoverable:=True)
            End If
            For index = 0 To spec.Slides.Count - 1
                If String.IsNullOrWhiteSpace(spec.Slides(index).Title) Then
                    Return ToolResult.Failed(toolId, $"Slide {index + 1} is missing a title",
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:=$"第 {index + 1} 张幻灯片缺少标题", recoverable:=True)
                End If
                Dim sceneError = ValidateSceneSpec(spec.Slides(index))
                If Not String.IsNullOrWhiteSpace(sceneError) Then
                    Return ToolResult.Failed(toolId,
                                             $"Slide {index + 1}: {sceneError}",
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:=$"第 {index + 1} 张幻灯片的 Scene 信息不足：{sceneError}",
                                             recoverable:=True)
                End If
            Next

            If Not DesignSystemCatalog.IsSupported(spec.DesignSystem) AndAlso spec.DesignTokens Is Nothing Then
                Return ToolResult.Failed(toolId,
                                         $"Unknown designSystem '{spec.DesignSystem}' without designTokens",
                                         errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                         userMessage:="未知设计系统必须提供完整 designTokens，或改用已注册的设计系统",
                                         recoverable:=True)
            End If

            Dim tokens = DesignSystemCatalog.Resolve(spec.DesignSystem, spec.DesignTokens)
            Dim pageSetup As PowerPoint.PageSetup = Nothing
            Dim slideWidth As Single
            Dim slideHeight As Single
            Try
                pageSetup = presentation.PageSetup
                slideWidth = pageSetup.SlideWidth
                slideHeight = pageSetup.SlideHeight
            Finally
                ComObjectHelper.ReleaseComObject(pageSetup)
            End Try
            Dim initialCount = GetRequiredSlideCount(presentation)
            Dim targetRefs As New List(Of String)()
            Dim slideResults As New JArray()
            Dim warnings As New List(Of String)()
            Dim createdCount As Integer = 0

            If preview Then
                Return ExecutePreview(spec, tokens, slideWidth, slideHeight, initialCount)
            End If

            Dim deckPlans As New List(Of SlideRenderPlan)()
            Dim deckPreflightReports As New List(Of VisualVerificationReport)()
            For preflightIndex = 0 To spec.Slides.Count - 1
                Dim preflightSpec = spec.Slides(preflightIndex)
                Try
                    Dim precompiledPlan = SlideLayoutEngine.Compile(preflightSpec, tokens, slideWidth, slideHeight,
                                                                    preflightIndex, spec.Slides.Count)
                    Dim precompiledReport = PowerPointVisualVerifier.PreflightAndRepair(precompiledPlan,
                                                                                        slideWidth, slideHeight)
                    deckPlans.Add(precompiledPlan)
                    deckPreflightReports.Add(precompiledReport)
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", preflightIndex + 1}, {"title", preflightSpec.Title},
                        {"slideType", preflightSpec.SlideType}, {"status", "failed"},
                        {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "deck_compile_preflight"}
                    })
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        ex.Message, classified.ErrorCode)
                End Try
                If Not deckPreflightReports(preflightIndex).Passed Then
                    slideResults.Add(BuildSlideResult(preflightIndex, preflightSpec, "failed",
                                                       deckPreflightReports(preflightIndex), Nothing,
                                                       "LAYOUT_VERIFY_FAILED"))
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        "专业布局预检未通过", ExceptionClassifier.CodeVerifyFailed)
                End If
            Next

            VerifyDeckComposition(deckPlans, deckPreflightReports)
            If deckPreflightReports.Any(Function(report) Not report.Passed) Then
                For preflightIndex = 0 To spec.Slides.Count - 1
                    Dim report = deckPreflightReports(preflightIndex)
                    slideResults.Add(BuildSlideResult(preflightIndex, spec.Slides(preflightIndex),
                                                       If(report.Passed, "not_rendered", "failed"),
                                                       report, Nothing,
                                                       If(report.Passed, "", "DECK_COMPOSITION_VERIFY_FAILED")))
                Next
                Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                    "整套演示文稿的构图变化节奏未通过",
                                    ExceptionClassifier.CodeVerifyFailed)
            End If

            For index = 0 To spec.Slides.Count - 1
                Dim slideSpec = spec.Slides(index)
                Dim plan As SlideRenderPlan = Nothing
                Dim preflight As VisualVerificationReport = Nothing
                Try
                    plan = SlideLayoutEngine.Compile(slideSpec, tokens, slideWidth, slideHeight, index, spec.Slides.Count)
                    preflight = PowerPointVisualVerifier.PreflightAndRepair(plan, slideWidth, slideHeight)
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", index + 1}, {"title", slideSpec.Title}, {"slideType", slideSpec.SlideType},
                        {"status", "failed"}, {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "compile_preflight"}
                    })
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        ex.Message, classified.ErrorCode)
                End Try
                If Not preflight.Passed Then
                    slideResults.Add(BuildSlideResult(index, slideSpec, "failed", preflight, Nothing, "LAYOUT_VERIFY_FAILED"))
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        "专业布局预检未通过", ExceptionClassifier.CodeVerifyFailed)
                End If

                Dim renderResult As SceneRenderResult = Nothing
                Try
                    renderResult = PowerPointSceneRenderer.Render(presentation, slideSpec, plan,
                                                                  initialCount + createdCount + 1, tokens)
                    createdCount += 1
                    Dim slideIndex = renderResult.Slide.SlideIndex
                    Dim targetRef = $"PowerPoint:presentations/active/slides/{slideIndex}"
                    targetRefs.Add(targetRef)
                    warnings.AddRange(renderResult.Warnings)
                    Dim renderedReport = PowerPointVisualVerifier.VerifyAndRepairRenderedSlide(
                        renderResult.Slide, slideWidth, slideHeight, plan, tokens)
                    If Not String.IsNullOrWhiteSpace(slideSpec.ImagePath) AndAlso
                       renderResult.Warnings.Any(Function(warning)
                           Return warning.IndexOf("Image skipped", StringComparison.OrdinalIgnoreCase) >= 0
                       End Function) Then
                        renderedReport.Issues.Add(New VisualIssue With {
                            .Code = "IMAGE_ARTIFACT_MISSING",
                            .Severity = "error",
                            .Message = $"Requested image could not be rendered: {slideSpec.ImagePath}"
                        })
                    End If
                    If Not String.IsNullOrWhiteSpace(slideSpec.Notes) AndAlso
                       renderResult.Warnings.Any(Function(warning)
                           Return warning.IndexOf("Speaker notes skipped", StringComparison.OrdinalIgnoreCase) >= 0
                       End Function) Then
                        renderedReport.Issues.Add(New VisualIssue With {
                            .Code = "NOTES_ARTIFACT_MISSING",
                            .Severity = "error",
                            .Message = "Requested speaker notes could not be rendered"
                        })
                    End If
                    MergeReport(preflight, renderedReport)
                    slideResults.Add(BuildSlideResult(index, slideSpec,
                                                      If(preflight.Passed, "succeeded", "failed"),
                                                      preflight, targetRef,
                                                      If(preflight.Passed, "", "VISUAL_VERIFY_FAILED")))
                    If Not preflight.Passed Then
                        Return BuildFailureWithVisualEvidence(CaptureVisualEvidence(renderResult.Slide, index + 1,
                                                                                     slideWidth, slideHeight),
                                                              presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                            "渲染后的视觉质量检查未通过", ExceptionClassifier.CodeVerifyFailed)
                    End If
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", index + 1}, {"title", slideSpec.Title}, {"slideType", slideSpec.SlideType},
                        {"status", "failed"}, {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "render_verify"}
                    })
                    Dim visualEvidence As AgentVisualEvidence = Nothing
                    If renderResult IsNot Nothing AndAlso renderResult.Slide IsNot Nothing Then
                        visualEvidence = CaptureVisualEvidence(renderResult.Slide, index + 1, slideWidth, slideHeight)
                    End If
                    Return BuildFailureWithVisualEvidence(visualEvidence,
                                                          presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        ex.Message, classified.ErrorCode)
                Finally
                    If renderResult IsNot Nothing Then ComObjectHelper.ReleaseComObject(renderResult.Slide)
                End Try
            Next

            Try
                Dim observation = BuildObservation(spec, initialCount, createdCount, targetRefs, slideResults, warnings, True)
                Dim data As New JObject From {
                    {"designSystem", tokens.Name}, {"createdSlides", createdCount},
                    {"targetRefs", JArray.FromObject(targetRefs)}, {"slideResults", slideResults.DeepClone()}
                }
                Return ToolResult.Succeed("CreateSlides",
                                          $"已使用 {tokens.Name} 设计系统创建 {createdCount} 张专业幻灯片",
                                          data:=data, observation:=observation,
                                          artifacts:=New JObject From {{"slides", JArray.FromObject(targetRefs)}})
            Catch ex As Exception
                Dim classified = ExceptionClassifier.Classify(ex)
                slideResults.Add(New JObject From {
                    {"index", spec.Slides.Count}, {"status", "failed"},
                    {"errorCode", classified.ErrorCode}, {"message", ex.Message}, {"phase", "finalize_result"}
                })
                Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                    ex.Message, classified.ErrorCode)
            End Try
        End Function



        Private Shared Function ValidateSceneSpec(spec As SlideDesignSpec) As String
            If spec Is Nothing Then Return "Scene 为空"
            If Not spec.SlideTypeRecognized Then
                Return $"Unsupported slideType '{spec.RequestedSlideType}'; select a registered Scene archetype"
            End If
            Dim variantError = ValidateVariant(spec)
            If Not String.IsNullOrWhiteSpace(variantError) Then Return variantError
            If Not String.IsNullOrWhiteSpace(spec.ImagePath) AndAlso
               spec.SlideType <> "cover" AndAlso spec.SlideType <> "content" Then
                Return $"{spec.SlideType} 页面尚未消费 imagePath；请改用 cover/content 图文构图，或移除该页未使用的 imagePath"
            End If
            If spec.Chart IsNot Nothing Then
                If spec.SlideType <> "content" Then Return "chart 当前必须用于 content 页面"
                If Not String.IsNullOrWhiteSpace(spec.ImagePath) OrElse spec.Table IsNot Nothing Then
                    Return "同一 content 页面只能声明 imagePath、chart、table 中的一种主视觉"
                End If
                Dim chartError = ValidateChartSpec(spec.Chart)
                If Not String.IsNullOrWhiteSpace(chartError) Then Return chartError
            End If
            If spec.Table IsNot Nothing Then
                If spec.SlideType <> "content" Then Return "table 当前必须用于 content 页面"
                If Not String.IsNullOrWhiteSpace(spec.ImagePath) OrElse spec.Chart IsNot Nothing Then
                    Return "同一 content 页面只能声明 imagePath、chart、table 中的一种主视觉"
                End If
                Dim tableError = ValidateTableSpec(spec.Table)
                If Not String.IsNullOrWhiteSpace(tableError) Then Return tableError
            End If
            Select Case spec.SlideType
                Case "statement"
                    If spec.Items.Count > 3 Then Return "statement 页面最多容纳 3 条证据；请拆页或改用 content"
                Case "content"
                    If spec.Items.Count < 1 Then Return "content 页面至少需要 1 个 item"
                    If spec.Items.Count > 6 Then Return "content 页面最多容纳 6 个 items；请拆分页面"
                    If spec.Chart IsNot Nothing AndAlso spec.Items.Count > 4 Then
                        Return "带 chart 的 content 页面最多容纳 4 个洞察 items"
                    End If
                    If spec.Table IsNot Nothing AndAlso spec.Items.Count > 4 Then
                        Return "带 table 的 content 页面最多容纳 4 个洞察 items"
                    End If
                    If Not String.IsNullOrWhiteSpace(spec.ImagePath) AndAlso spec.Items.Count > 4 Then
                        Return "带 imagePath 的 content 页面最多容纳 4 个 items"
                    End If
                    If (String.Equals(spec.LayoutVariant, "feature-left", StringComparison.OrdinalIgnoreCase) OrElse
                        spec.Items.Any(Function(item) item.Emphasis)) AndAlso spec.Items.Count > 5 Then
                        Return "feature-left content 页面最多容纳 5 个 items"
                    End If
                Case "two-column"
                    If spec.Items.Count <> 2 Then Return "two-column 页面必须恰好包含 2 个 items"
                Case "comparison"
                    Dim isTable = spec.Items.Count >= 3 AndAlso
                                  spec.Items.All(Function(item) item.Features IsNot Nothing AndAlso item.Features.Count >= 2)
                    If isTable AndAlso spec.Items.Count > 5 Then Return "comparison 表格最多容纳 5 行；请拆分页面"
                    If isTable AndAlso (spec.ColumnHeaders Is Nothing OrElse spec.ColumnHeaders.Count <> 3 OrElse
                                        spec.ColumnHeaders.Any(Function(header) String.IsNullOrWhiteSpace(header))) Then
                        Return "comparison 表格必须提供 3 个非空 columnHeaders：[比较维度, 左侧方案, 右侧方案]"
                    End If
                    If Not isTable AndAlso spec.Items.Count <> 2 Then Return "comparison 页面需要恰好 2 个对比对象，或 3-5 行双列 features"
                Case "kpi"
                    If spec.Metrics.Count < 2 AndAlso spec.Items.Count < 2 Then Return "kpi 页面至少需要 2 个 metrics"
                    If spec.Metrics.Count > 4 OrElse spec.Items.Count > 4 Then Return "kpi 页面最多容纳 4 个指标；请拆分页面"
                    If spec.Metrics.Count >= 2 AndAlso
                       spec.Metrics.Any(Function(metric) String.IsNullOrWhiteSpace(metric.Value) OrElse String.IsNullOrWhiteSpace(metric.Label)) Then
                        Return "kpi metrics 必须同时包含 value 和 label；不要用占位符伪造指标"
                    End If
                    If spec.Metrics.Count = 0 AndAlso
                       spec.Items.Any(Function(item) String.IsNullOrWhiteSpace(item.Value) OrElse String.IsNullOrWhiteSpace(item.Title)) Then
                        Return "kpi items 必须同时包含 value 和 title；没有可靠数值时应改用 content/statement"
                    End If
                Case "process"
                    If spec.Items.Count < 3 Then Return "process 页面至少需要 3 个步骤"
                    If spec.Items.Count > 6 Then Return "process 页面最多容纳 6 个步骤；请拆分流程"
                Case "architecture"
                    If spec.Items.Count < 2 Then Return "architecture 页面至少需要 2 个层级"
                    If spec.Items.Count > 5 Then Return "architecture 页面最多容纳 5 个层级；请拆分架构"
                Case "matrix"
                    If spec.Items.Count <> 4 Then Return "matrix 页面必须恰好包含 4 个象限 items"
                    If String.IsNullOrWhiteSpace(spec.XAxisLabel) OrElse String.IsNullOrWhiteSpace(spec.YAxisLabel) Then
                        Return "matrix 页面必须提供 xAxisLabel 和 yAxisLabel，不能假设固定业务维度"
                    End If
            End Select
            Return ""
        End Function

        Private Shared Function ValidateVariant(spec As SlideDesignSpec) As String
            If spec Is Nothing OrElse String.IsNullOrWhiteSpace(spec.LayoutVariant) Then Return ""
            Dim layoutVariant = spec.LayoutVariant.Trim().ToLowerInvariant()
            If layoutVariant = "default" OrElse layoutVariant = "standard" Then Return ""
            Select Case spec.SlideType
                Case "content"
                    If layoutVariant = "feature-left" Then Return ""
                Case "kpi"
                    If layoutVariant = "hero-left" Then Return ""
                Case "process"
                    If layoutVariant = "vertical" Then Return ""
                Case "architecture"
                    If layoutVariant = "hub-spoke" AndAlso spec.Items.Count >= 3 Then Return ""
            End Select
            If spec.SlideType = "architecture" AndAlso layoutVariant = "hub-spoke" Then
                Return "architecture hub-spoke requires one core item and at least two spoke items"
            End If
            Return $"variant '{spec.LayoutVariant}' is not supported by the {spec.SlideType} Scene"
        End Function

        Private Shared Function ValidateChartSpec(chart As DesignChart) As String
            If chart Is Nothing Then Return ""
            Dim chartType = If(chart.ChartType, "column").Trim().ToLowerInvariant()
            If chartType <> "column" AndAlso chartType <> "line" Then
                Return "chart.chartType 仅支持 column 或 line"
            End If
            If chart.Categories.Count < 2 OrElse chart.Categories.Count > 8 Then
                Return "chart.categories 必须包含 2-8 个分类"
            End If
            If chart.Categories.Any(Function(category) String.IsNullOrWhiteSpace(category)) Then
                Return "chart.categories 不能包含空标签"
            End If
            If chart.Series.Count < 1 OrElse chart.Series.Count > 3 Then
                Return "chart.series 必须包含 1-3 个序列"
            End If
            If chart.Series.Count > 1 AndAlso chart.Series.Any(Function(series) String.IsNullOrWhiteSpace(series.Name)) Then
                Return "多序列 chart 的每个 series 都必须包含 name"
            End If
            Dim hasNonZeroValue As Boolean = False
            For Each series In chart.Series
                If series.Values.Count <> chart.Categories.Count Then
                    Return "每个 chart series 的 values 数量必须与 categories 一致"
                End If
                For Each value In series.Values
                    If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                        Return "chart values 必须是有限数字"
                    End If
                    If Math.Abs(value) > 0.000001R Then hasNonZeroValue = True
                Next
            Next
            If Not hasNonZeroValue Then Return "chart 至少需要一个非零数值"
            Return ""
        End Function

        Private Shared Function ValidateTableSpec(table As DesignTable) As String
            If table Is Nothing Then Return ""
            If table.Headers.Count < 2 OrElse table.Headers.Count > 5 Then
                Return "table.headers 必须包含 2-5 列"
            End If
            If table.Headers.Any(Function(header) String.IsNullOrWhiteSpace(header)) Then
                Return "table.headers 不能包含空标题"
            End If
            If table.Rows.Count < 1 OrElse table.Rows.Count > 6 Then
                Return "table.rows 必须包含 1-6 行"
            End If
            If table.Rows.Any(Function(row) row Is Nothing OrElse row.Count <> table.Headers.Count) Then
                Return "table 每行单元格数量必须与 headers 一致"
            End If
            If table.HighlightColumn < -1 OrElse table.HighlightColumn >= table.Headers.Count Then
                Return "table.highlightColumn 必须是有效的零基列索引"
            End If
            Return ""
        End Function
    End Class

End Namespace
