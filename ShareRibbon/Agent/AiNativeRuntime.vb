Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Interface IAiNativeRuntime
        Function AnalyzeAsync(request As AiNativeRequest) As Task(Of AiNativeRuntimeResult)
    End Interface

    Public Class AiNativeRequest
        Public Property UserInput As String
        Public Property AppType As String
        Public Property SystemPrompt As String
        Public Property RequestUuid As String
        Public Property OfficeContext As Context.OfficeContext
        Public Property ContextSnapshot As JObject
        Public Property HistoryMessages As List(Of HistoryMessage)
        Public Property PreviousTaskSpec As AgentTaskSpec
        Public Property EnableMemory As Boolean = True
        Public Property UseContextBuilder As Boolean = True
    End Class

    Public Class AiNativeRuntimeResult
        Public Property Intent As IntentResult
        Public Property ContextTrace As ChatContextTrace
        Public Property Messages As List(Of HistoryMessage)
        Public Property AvailableTools As List(Of ToolDescriptor)
        Public Property SelectedSkills As List(Of SkillFileDefinition)
        Public Property RagCount As Integer
        Public Property UsedContextBuilder As Boolean
        Public Property TaskSpec As AgentTaskSpec
        Public Property ExecutionPlan As ExecutionPlan
    End Class

    ''' <summary>
    ''' AI Native 统一分析运行时。只依赖共享抽象，不直接访问具体 Office 对象模型。
    ''' </summary>
    Public Class AiNativeRuntime
        Implements IAiNativeRuntime

        Private ReadOnly _toolRegistry As ToolRegistry

        Public Sub New(toolRegistry As ToolRegistry)
            _toolRegistry = If(toolRegistry, New ToolRegistry())
        End Sub

        Public Async Function AnalyzeAsync(request As AiNativeRequest) As Task(Of AiNativeRuntimeResult) Implements IAiNativeRuntime.AnalyzeAsync
            If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))

            Dim appType = If(String.IsNullOrWhiteSpace(request.AppType), "Excel", request.AppType)
            Dim input = If(request.UserInput, "")
            Dim contextSnapshot = BuildContextSnapshot(request, appType)
            Dim tools = _toolRegistry.GetAvailableTools(appType)
            AttachRuntimeToolFacts(contextSnapshot, tools)

            Dim intentService As New IntentRecognitionService(appType)
            Dim intent = Await intentService.IdentifyIntentAsync(input, contextSnapshot)
            NormalizeIntentForObservedContext(intent, input, request.OfficeContext, appType)
            NormalizeIntentForTaskContinuity(intent, request, appType)
            If intent IsNot Nothing Then
                intent.OriginalInput = input
                If String.IsNullOrWhiteSpace(intent.UserFriendlyDescription) Then
                    intentService.GenerateUserFriendlyDescription(intent)
                End If
                intentService.BuildExecutionPlanPreview(intent)
            End If

            Dim ragCount As Integer = 0
            Dim messages As List(Of HistoryMessage)
            Dim usedContextBuilder = request.UseContextBuilder
            If usedContextBuilder Then
                messages = ChatContextBuilder.BuildMessages(
                    appType.ToLowerInvariant(),
                    appType,
                    input,
                    NormalizeHistory(request.HistoryMessages),
                    input,
                    request.SystemPrompt,
                    New Dictionary(Of String, String)(),
                    request.EnableMemory,
                    ragCount)
            Else
                messages = BuildFallbackMessages(request)
            End If

            Dim selectedSkills = SelectSkills(input, intent, appType)
            Dim taskSpec = BuildTaskSpec(request, intent, tools, selectedSkills, appType)
            Dim resolvedSkills = SkillCapabilityResolver.ResolvePrimarySkill(selectedSkills,
                                                                             taskSpec.RequiredTools,
                                                                             appType)
            Dim primaryChanged = resolvedSkills.Count > 0 AndAlso
                (selectedSkills.Count = 0 OrElse
                 Not String.Equals(resolvedSkills(0).Name,
                                   selectedSkills(0).Name,
                                   StringComparison.OrdinalIgnoreCase))
            If primaryChanged Then
                selectedSkills = resolvedSkills
                taskSpec = BuildTaskSpec(request, intent, tools, selectedSkills, appType)
            End If
            Dim trace = If(ChatContextBuilder.LastTrace, New ChatContextTrace With {.Query = input, .AppType = appType})
            EnrichTrace(trace, request, intent, taskSpec, selectedSkills, appType)
            For Each tool In tools.Take(12)
                trace.Tools.Add(New ChatContextToolTrace With {
                    .Id = tool.Id,
                    .Name = tool.Name,
                    .Category = tool.Category,
                    .RiskLevel = tool.RiskLevel,
                    .AvailabilityStatus = tool.AvailabilityStatus,
                    .LastError = tool.LastError
                })
            Next

            Return New AiNativeRuntimeResult With {
                .Intent = intent,
                .ContextTrace = trace,
                .Messages = messages,
                .AvailableTools = tools,
                .SelectedSkills = selectedSkills,
                .RagCount = ragCount,
                .UsedContextBuilder = usedContextBuilder,
                .TaskSpec = taskSpec,
                .ExecutionPlan = Nothing
            }
        End Function

        ''' <summary>
        ''' Makes capability claims in the intent pass evidence-based. The registry is the
        ''' authority; model-generated descriptions must not infer tool absence from a narrow
        ''' retrieved Skill or from the active-cell preview.
        ''' </summary>
        Private Shared Sub AttachRuntimeToolFacts(snapshot As JObject,
                                                  tools As IEnumerable(Of ToolDescriptor))
            If snapshot Is Nothing Then Return
            Dim facts As New JArray()
            For Each tool In If(tools, Enumerable.Empty(Of ToolDescriptor)()).
                Where(Function(item) item IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(item.Id)).
                OrderBy(Function(item) item.Id, StringComparer.OrdinalIgnoreCase)

                facts.Add(New JObject From {
                    {"id", tool.Id},
                    {"name", If(tool.Name, "")},
                    {"accessMode", If(tool.AccessMode, "")},
                    {"availability", If(tool.AvailabilityStatus, "available")}
                })
            Next
            snapshot("availableTools") = facts
        End Sub

        Private Function BuildTaskSpec(request As AiNativeRequest,
                                       intent As IntentResult,
                                       tools As List(Of ToolDescriptor),
                                       selectedSkills As List(Of SkillFileDefinition),
                                       appType As String) As AgentTaskSpec
            Dim spec As New AgentTaskSpec()
            Dim input = If(request.UserInput, "").Trim()
            Dim intentDescription = If(intent?.UserFriendlyDescription, "")
            Dim isDestinationCorrection = IsDestinationOnlyCorrection(input)
            Dim priorTask = If(isDestinationCorrection, request.PreviousTaskSpec, Nothing)
            Dim semanticInput = input & vbCrLf & intentDescription
            If priorTask IsNot Nothing Then semanticInput &= vbCrLf & priorTask.Goal
            Dim interactionMode = If(intent?.ResponseMode, "").Trim().ToLowerInvariant()
            Dim isReadOnlyDataAnswer = IsReadOnlyExcelDataAnswer(intent, request.OfficeContext, appType)
            Dim isNonExecution = interactionMode = "clarify" OrElse
                (interactionMode = "answer" AndAlso Not isReadOnlyDataAnswer)
            Dim isPlainWorksheetCreation = IsOnlyWorksheetCreationRequest(input, appType)
            Dim isExplicitPythonCompute = NormalizeAppType(appType) = "excel" AndAlso
                (IsExplicitPythonComputeRequest(semanticInput) OrElse TaskRequiresTool(priorTask, "PythonCompute"))
            Dim isObservedChartCreation = NormalizeAppType(appType) = "excel" AndAlso
                HasObservedExcelTable(request.OfficeContext) AndAlso
                IsExplicitChartCreationRequest(input)

            ' 用户原始目标是开放式 TaskSpec 的权威来源；枚举意图只作为检索/观测标签，
            ' 不能把未穷举的新需求压缩成 GENERAL_QUERY。
            If priorTask IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(priorTask.Goal) Then
                spec.Goal = priorTask.Goal & "；本轮仅修正输出目标：" & input
            Else
                spec.Goal = If(String.IsNullOrWhiteSpace(input), "自动分析当前 Office 上下文并选择合适处理方式", input)
            End If
            Dim authoritativeRawRequest = input
            If priorTask?.GoalContract IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(priorTask.GoalContract.RawUserRequest) Then
                authoritativeRawRequest = priorTask.GoalContract.RawUserRequest & vbCrLf & input
            ElseIf priorTask IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(priorTask.RawUserRequest) Then
                authoritativeRawRequest = priorTask.RawUserRequest & vbCrLf & input
            End If
            If String.IsNullOrWhiteSpace(authoritativeRawRequest) Then authoritativeRawRequest = spec.Goal
            spec.CaptureRawUserRequest(authoritativeRawRequest)

            If isNonExecution Then
                spec.TargetObject = "用户问题"
            ElseIf isReadOnlyDataAnswer Then
                Dim observedSource = GetObservedExcelTableAddress(request.OfficeContext)
                spec.TargetObject = If(String.IsNullOrWhiteSpace(observedSource),
                                       InferTargetObject(request, intent),
                                       "表区域 " & observedSource)
            ElseIf priorTask IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(priorTask.TargetObject) Then
                spec.TargetObject = priorTask.TargetObject
            ElseIf isExplicitPythonCompute Then
                Dim observedSource = GetObservedExcelTableAddress(request.OfficeContext)
                spec.TargetObject = If(String.IsNullOrWhiteSpace(observedSource), InferTargetObject(request, intent), "表区域 " & observedSource)
            Else
                spec.TargetObject = InferTargetObject(request, intent)
            End If
            spec.Complexity = InferComplexity(intent)
            spec.RiskLevel = If(isExplicitPythonCompute, "medium", InferRiskLevel(intent))
            spec.Constraints.Add($"宿主应用: {appType}")
            If isNonExecution Then
                spec.Constraints.Add($"交互模式: {interactionMode}；不得调用 Office 写入工具")
            ElseIf isReadOnlyDataAnswer Then
                spec.MutationPolicy = "read_only"
                spec.Constraints.Add("只读数据问答：可以读取完成回答所需的当前工作簿数据，但不得调用任何修改工作簿的工具")
                spec.Constraints.Add("上下文中的数据预览可能被截断；必须以 ReadRange 返回的完整结构化数据为证据，数据不足时明确失败，禁止抽样外推、估算或编造")
            Else
                spec.Constraints.Add("不要求用户手动确认意图，优先由 AI 自动识别和执行")
            End If
            If Not isNonExecution AndAlso selectedSkills IsNot Nothing AndAlso selectedSkills.Count > 0 Then
                spec.Constraints.Add("已自动选择 Skill: " & String.Join(", ", selectedSkills.Select(Function(s) s.Name)))
            End If

            If intent IsNot Nothing Then
                spec.Constraints.Add($"兼容意图标签: {intent.OfficeIntent}")
                If Not String.IsNullOrWhiteSpace(intentDescription) Then
                    spec.Constraints.Add($"模型理解摘要: {intentDescription}")
                End If
                If intent.Confidence < 0.4 Then
                    spec.Constraints.Add("低置信度任务采用探索式计划，先分析上下文再选择低风险动作")
                End If
                For Each kv In intent.ExtractedEntities
                    spec.Constraints.Add($"{kv.Key}: {kv.Value}")
                Next
                If intent.RequestedOutputs IsNot Nothing AndAlso Not isReadOnlyDataAnswer Then
                    For Each output In intent.RequestedOutputs
                        AddExpectedOutput(spec, output)
                    Next
                End If
            End If

            If isNonExecution Then
                If interactionMode = "clarify" Then
                    spec.SuccessCriteria.Add("只询问阻塞任务继续所必需的最小问题")
                Else
                    spec.SuccessCriteria.Add("直接、简洁地回答用户问题")
                End If
                Return spec
            End If

            If isReadOnlyDataAnswer Then
                spec.SuccessCriteria.Add("完整读取回答所需的工作簿范围；禁止估算或使用截断预览外推")
                spec.Constraints.Add("只在聊天中返回结果；答案中的数值必须能由工具返回的数据逐项验证")
            ElseIf isExplicitPythonCompute Then
                spec.Constraints.Add("用户显式指定 Python；计算引擎必须是 PythonCompute，不得替换为 DataAnalysis、公式、透视表或 GenerateReport；读取和写入动作根据最新观察自适应选择")
                spec.Constraints.Add("PythonCompute 只接收 ReadRange 返回的 JSON，不直接操作 Excel、文件、网络或子进程；受控计算默认无需审批")
                If isDestinationCorrection Then
                    spec.Constraints.Add("本轮只修正输出工作表/位置；必须保留上一任务的数据源、分组字段、聚合方式和 Python 计算方法，不得把当前输出表改为数据源")
                End If
                spec.SuccessCriteria.Add("用 Python 完成计算，并将返回的结构化结果写入指定工作表")
            ElseIf isPlainWorksheetCreation Then
                spec.Constraints.Add("纯工作表创建任务只允许使用 CreateSheet；工作表名称不代表需要生成报表内容")
                spec.SuccessCriteria.Add("创建用户指定名称的空白工作表，不附加报表、摘要或图表")
            ElseIf isObservedChartCreation Then
                spec.Constraints.Add("插件已观察到可用表区域；直接从表头推断图表字段，不得要求用户重复提供列范围")
                spec.SuccessCriteria.Add("使用已识别的数据区域创建用户要求的图表")
            End If

            ' 确定性输出门只用于验收，不参与意图路由。
            Dim inputLower = input.ToLowerInvariant()
            If inputLower.Contains("图片") OrElse inputLower.Contains("配图") OrElse inputLower.Contains("image") OrElse inputLower.Contains("picture") Then
                AddExpectedOutput(spec, "images")
            End If
            If NormalizeAppType(appType) = "powerpoint" Then
                Dim slideMatch = System.Text.RegularExpressions.Regex.Match(inputLower, "(\d+)\s*(页|张)")
                If slideMatch.Success Then
                    Dim parsedSlideCount As Integer
                    If Integer.TryParse(slideMatch.Groups(1).Value, parsedSlideCount) Then spec.ExpectedSlideCount = parsedSlideCount
                End If
                If spec.ExpectedSlideCount > 0 Then AddExpectedOutput(spec, "slides")
            End If

            spec.Constraints.Add("输出结果必须与当前 Office 上下文相关")
            spec.Constraints.Add("执行过程必须可解释，并记录使用的上下文、工具和记忆")
            If spec.ExpectedSlideCount > 0 Then spec.SuccessCriteria.Add($"实际创建至少 {spec.ExpectedSlideCount} 张幻灯片")
            If spec.ExpectedOutputs.Contains("images") Then spec.SuccessCriteria.Add("实际插入可访问的图片，不得用占位形状冒充")
            If spec.RiskLevel <> "safe" Then
                spec.Constraints.Add("执行后必须保留可观察结果或失败原因")
            End If

            ' 当前执行合同使用一个 primary Skill 作为工具安全边界；其余召回结果只用于
            ' trace/候选解释，不能把次要 Skill 的工具混入 RequiredTools 后再被执行门拒绝。
            Dim primarySkill = If(selectedSkills?.FirstOrDefault(), Nothing)
            If isReadOnlyDataAnswer Then
                AddRequiredTool(spec, "ReadRange")
                AddRequiredCapability(spec, "ReadRange")
            ElseIf isExplicitPythonCompute Then
                AddRequiredTool(spec, "ReadRange")
                AddRequiredTool(spec, "PythonCompute")
                ' CreateSheet is part of the plan only when the user actually asks for a new
                ' worksheet. Merely naming an existing destination must not invite the planner
                ' to recreate it or report a false capability gap.
                If RequestsNewWorksheet(semanticInput) Then AddRequiredTool(spec, "CreateSheet")
                AddRequiredTool(spec, "WriteData")
                AddRequiredCapability(spec, "PythonCompute")
            ElseIf primarySkill IsNot Nothing Then
                If isPlainWorksheetCreation Then
                    spec.RequiredTools.Add("CreateSheet")
                ElseIf isObservedChartCreation Then
                    spec.RequiredTools.Add("CreateChart")
                ElseIf primarySkill.AllowedTools IsNot Nothing Then
                    For Each toolId In primarySkill.AllowedTools
                        If Not String.IsNullOrWhiteSpace(toolId) AndAlso Not spec.RequiredTools.Contains(toolId) Then
                            spec.RequiredTools.Add(toolId)
                        End If
                    Next
                End If
                If primarySkill.Scripts IsNot Nothing Then
                    For Each script In primarySkill.Scripts
                        Dim scriptToolId = $"skill_script.{primarySkill.Name}.{script.FileName}"
                        If Not spec.RequiredTools.Contains(scriptToolId) Then spec.RequiredTools.Add(scriptToolId)
                    Next
                End If
            End If

            Return spec
        End Function

        Private Shared Function IsReadOnlyExcelDataAnswer(intent As IntentResult,
                                                           officeContext As Context.OfficeContext,
                                                           appType As String) As Boolean
            If intent Is Nothing OrElse NormalizeAppType(appType) <> "excel" Then Return False
            If Not String.Equals(If(intent.ResponseMode, "").Trim(), "answer", StringComparison.OrdinalIgnoreCase) Then Return False
            If Not HasObservedExcelTable(officeContext) Then Return False

            ' DATA_ANALYSIS is the semantic signal that the requested answer depends on workbook
            ' values. GENERAL_QUERY remains the plain-chat path for explanations and how-to help.
            Return intent.OfficeIntent = OfficeIntentType.DATA_ANALYSIS OrElse
                intent.IntentType = ExcelIntentType.DATA_ANALYSIS
        End Function

        Private Shared Sub AddRequiredTool(spec As AgentTaskSpec, toolId As String)
            If spec Is Nothing OrElse String.IsNullOrWhiteSpace(toolId) Then Return
            If Not spec.RequiredTools.Contains(toolId) Then spec.RequiredTools.Add(toolId)
        End Sub

        Private Shared Sub AddRequiredCapability(spec As AgentTaskSpec, toolId As String)
            If spec Is Nothing OrElse String.IsNullOrWhiteSpace(toolId) Then Return
            If Not spec.RequiredCapabilities.Contains(toolId) Then spec.RequiredCapabilities.Add(toolId)
        End Sub

        Private Shared Sub AddExpectedOutput(spec As AgentTaskSpec, output As String)
            If spec Is Nothing OrElse String.IsNullOrWhiteSpace(output) Then Return
            Dim normalized = output.Trim().ToLowerInvariant()
            If Not spec.ExpectedOutputs.Contains(normalized) Then spec.ExpectedOutputs.Add(normalized)
        End Sub

        Private Shared Function IsOnlyWorksheetCreationRequest(input As String, appType As String) As Boolean
            If NormalizeAppType(appType) <> "excel" OrElse String.IsNullOrWhiteSpace(input) Then Return False

            Dim normalized = input.Trim().ToLowerInvariant()
            Dim hasCreateRequest = System.Text.RegularExpressions.Regex.IsMatch(
                normalized,
                "(新建|新增|添加|创建|插入).{0,40}(工作表|worksheet|sheet)|(?:create|add|insert).{0,40}(?:worksheet|sheet)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If Not hasCreateRequest Then Return False

            ' Scope only a single, empty-sheet request. Any conjunction or explicit content/data
            ' operation keeps the full Excel tool set available for a compound task.
            Dim compoundOrContentPattern =
                "(并且|并|以及|同时|然后|生成报表|报表内容|汇总数据|摘要|统计数据|分析|图表|透视|公式|写入|填充|复制.{0,8}数据|计算|排序|筛选|清洗|格式|\b(?:and|then|report|summary|chart|pivot|formula|write|copy|calculate|format)\b)"
            Return Not System.Text.RegularExpressions.Regex.IsMatch(
                normalized,
                compoundOrContentPattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End Function

        Friend Shared Function IsExplicitPythonComputeRequest(input As String) As Boolean
            If String.IsNullOrWhiteSpace(input) Then Return False
            Const engine As String = "(?:PythonCompute|Python)"
            Dim explicitInvocation = $"(?:(?:use|using|via|run|execute|with)\s+|(?:使用|用|通过|调用|运行|执行|基于)\s*){engine}(?![A-Za-z0-9_])"
            Dim explicitOperation = $"{engine}\s*(?:(?:to|for)\s+|(?:来|进行|用于)?\s*)(?:calculate|analyze|process|summarize|aggregate|compute|计算|分析|处理|汇总|统计|运行|执行)"
            Return System.Text.RegularExpressions.Regex.IsMatch(
                input,
                $"(?:{explicitInvocation})|(?:{explicitOperation})",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End Function

        Private Shared Function RequestsNewWorksheet(input As String) As Boolean
            If String.IsNullOrWhiteSpace(input) Then Return False
            Dim referencesExistingSheet = System.Text.RegularExpressions.Regex.IsMatch(
                input,
                "(之前创建|之前新建|已经创建|已创建|已有|已存在|现有|存在的)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase) AndAlso
                System.Text.RegularExpressions.Regex.IsMatch(
                    input,
                    "(工作表|worksheet|sheet)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If referencesExistingSheet Then
                Return False
            End If
            Return System.Text.RegularExpressions.Regex.IsMatch(
                input,
                "(新的工作表|新工作表|新建.{0,20}工作表|新增.{0,20}工作表|创建.{0,20}工作表|new\s+(?:worksheet|sheet))",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End Function

        Friend Shared Function IsDestinationOnlyCorrection(input As String) As Boolean
            If String.IsNullOrWhiteSpace(input) Then Return False
            ' A destination noun plus a write verb describes a perfectly valid new task.
            ' Inherit the previous task only when the user explicitly refers back to or
            ' corrects it; otherwise unrelated Python/formula state leaks across turns.
            Dim hasCorrection = System.Text.RegularExpressions.Regex.IsMatch(
                input,
                "((?:我.{0,4})?(?:刚刚|刚才).{0,12}(?:说|提到|要求)|上一步|上一个(?:任务|请求|操作)|上一轮|我说的|指的是|前面(?:说|提到|要求)的?|之前(?:说|提到|要求|执行)的?|原(?:任务|请求)中|\b(?:i\s+meant|previous\s+(?:task|request|step)|last\s+(?:task|request|step)|as\s+i\s+said|instead)\b)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            Dim hasDestination = System.Text.RegularExpressions.Regex.IsMatch(
                input,
                "(工作表|worksheet|sheet|单元格|位置|范围)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            Dim restartsCalculation = System.Text.RegularExpressions.Regex.IsMatch(
                input,
                "(改用|不要用|重新计算|另外计算|分组改为|汇总方式|聚合方式)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            Return hasCorrection AndAlso hasDestination AndAlso Not restartsCalculation
        End Function

        Private Shared Function TaskRequiresTool(spec As AgentTaskSpec, toolId As String) As Boolean
            Return spec IsNot Nothing AndAlso spec.RequiredTools IsNot Nothing AndAlso
                spec.RequiredTools.Any(Function(candidate) String.Equals(candidate, toolId, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Shared Sub NormalizeIntentForTaskContinuity(intent As IntentResult,
                                                            request As AiNativeRequest,
                                                            appType As String)
            If intent Is Nothing OrElse request Is Nothing OrElse NormalizeAppType(appType) <> "excel" Then Return
            Dim input = If(request.UserInput, "")
            Dim destinationCorrection = IsDestinationOnlyCorrection(input) AndAlso request.PreviousTaskSpec IsNot Nothing
            Dim previousTask = If(destinationCorrection, request.PreviousTaskSpec, Nothing)
            Dim explicitPython = IsExplicitPythonComputeRequest(input & vbCrLf & If(previousTask?.Goal, "")) OrElse
                TaskRequiresTool(previousTask, "PythonCompute")
            If Not explicitPython OrElse (Not HasObservedExcelTable(request.OfficeContext) AndAlso Not destinationCorrection) Then Return

            intent.OfficeIntent = OfficeIntentType.DATA_ANALYSIS
            intent.IntentType = ExcelIntentType.DATA_ANALYSIS
            intent.ResponseMode = "execute"
            intent.Confidence = Math.Max(intent.Confidence, 0.9R)
            If intent.RequestedOutputs Is Nothing Then intent.RequestedOutputs = New List(Of String)()
            If Not intent.RequestedOutputs.Contains("worksheet") Then intent.RequestedOutputs.Add("worksheet")
            If destinationCorrection Then
                intent.UserFriendlyDescription = "本轮仅修正上一个 Python 汇总任务的输出工作表，保留原数据源和计算方法"
            End If
        End Sub

        ''' <summary>
        ''' Explicit Excel chart requests are executable when the host has already observed a
        ''' usable table. This prevents a model-produced clarify mode from overriding facts in
        ''' the Office context merely because the active cell is outside the table.
        ''' </summary>
        Private Shared Sub NormalizeIntentForObservedContext(intent As IntentResult,
                                                             input As String,
                                                             officeContext As Context.OfficeContext,
                                                             appType As String)
            If intent Is Nothing OrElse NormalizeAppType(appType) <> "excel" Then Return
            If Not HasObservedExcelTable(officeContext) OrElse String.IsNullOrWhiteSpace(input) Then Return

            Dim normalized = input.Trim().ToLowerInvariant()
            Dim isChartIntent = intent.OfficeIntent = OfficeIntentType.CHART_GEN OrElse
                intent.IntentType = ExcelIntentType.CHART_GEN OrElse
                System.Text.RegularExpressions.Regex.IsMatch(normalized, "(图表|折线图|柱状图|条形图|饼图|散点图|chart|graph)")
            If Not isChartIntent OrElse Not IsExplicitChartCreationRequest(input) Then Return

            intent.OfficeIntent = OfficeIntentType.CHART_GEN
            intent.IntentType = ExcelIntentType.CHART_GEN
            intent.ResponseMode = "execute"
            intent.Confidence = Math.Max(intent.Confidence, 0.9R)
            If intent.RequestedOutputs Is Nothing Then intent.RequestedOutputs = New List(Of String)()
            If Not intent.RequestedOutputs.Contains("chart") Then intent.RequestedOutputs.Add("chart")
            intent.UserFriendlyDescription = "已识别当前工作表中的数据区域，将直接使用已观察到的字段生成图表"
        End Sub

        Private Shared Function IsExplicitChartCreationRequest(input As String) As Boolean
            If String.IsNullOrWhiteSpace(input) Then Return False
            Dim normalized = input.Trim().ToLowerInvariant()
            If System.Text.RegularExpressions.Regex.IsMatch(
                normalized,
                "(如何|怎么|怎样|为什么|教程|操作方法|能否介绍|请介绍|how\s+to|why)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                Return False
            End If

            Return System.Text.RegularExpressions.Regex.IsMatch(
                normalized,
                "(生成|创建|制作|绘制|画|做).{0,16}(图表|折线图|柱状图|条形图|饼图|散点图|chart|graph)|(?:create|make|draw|plot).{0,16}(?:chart|graph)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End Function

        Private Shared Function HasObservedExcelTable(officeContext As Context.OfficeContext) As Boolean
            If officeContext Is Nothing Then Return False

            Dim tables = TryCast(officeContext.HostData?("tables"), JArray)
            If tables IsNot Nothing AndAlso tables.Count > 0 Then
                For Each table In tables
                    If Not String.IsNullOrWhiteSpace(table?("address")?.ToString()) Then Return True
                Next
            End If

            Dim summary = officeContext.DocStructure?.Summary
            Return Not String.IsNullOrWhiteSpace(summary) AndAlso
                (summary.IndexOf("表区域:", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                 summary.IndexOf("使用区域:", StringComparison.OrdinalIgnoreCase) >= 0)
        End Function

        Private Shared Function NormalizeAppType(value As String) As String
            Dim normalized = If(value, "").Trim().ToLowerInvariant()
            If normalized = "ppt" OrElse normalized = "power point" Then Return "powerpoint"
            Return normalized
        End Function

        Private Function InferTargetObject(request As AiNativeRequest, intent As IntentResult) As String
            If NormalizeAppType(request?.AppType) = "excel" AndAlso
                intent IsNot Nothing AndAlso intent.OfficeIntent = OfficeIntentType.CHART_GEN Then
                Dim tableAddress = GetObservedExcelTableAddress(request.OfficeContext)
                If Not String.IsNullOrWhiteSpace(tableAddress) Then Return "表区域 " & tableAddress
            End If

            If request.OfficeContext IsNot Nothing AndAlso request.OfficeContext.Selection IsNot Nothing Then
                Dim selection = request.OfficeContext.Selection
                Return $"{If(selection.DataType, "选区")} {If(selection.Address, "")}".Trim()
            End If

            If intent IsNot Nothing AndAlso intent.ExtractedEntities IsNot Nothing AndAlso intent.ExtractedEntities.Count > 0 Then
                Return String.Join("; ", intent.ExtractedEntities.Select(Function(kv) $"{kv.Key}={kv.Value}"))
            End If

            Return "当前 Office 文档/工作区"
        End Function

        Private Shared Function GetObservedExcelTableAddress(officeContext As Context.OfficeContext) As String
            If officeContext Is Nothing OrElse officeContext.HostData Is Nothing Then Return ""
            Dim tables = TryCast(officeContext.HostData("tables"), JArray)
            If tables Is Nothing Then Return ""

            For Each table In tables
                Dim address = table?("address")?.ToString()
                If String.IsNullOrWhiteSpace(address) Then Continue For
                Dim sheet = table?("sheet")?.ToString()
                Return If(String.IsNullOrWhiteSpace(sheet), address, $"{sheet}!{address}")
            Next
            Return ""
        End Function

        Private Function InferComplexity(intent As IntentResult) As String
            If intent Is Nothing Then Return "exploratory"
            Dim interactionMode = If(intent.ResponseMode, "").Trim().ToLowerInvariant()
            If interactionMode = "answer" OrElse interactionMode = "clarify" Then Return "simple"
            If intent.Confidence < 0.4 Then Return "exploratory"
            If intent.CanUseDirectCommand AndAlso Not intent.RequiresVBA Then Return "simple"
            Return "medium"
        End Function

        Private Function InferRiskLevel(intent As IntentResult) As String
            If intent Is Nothing Then Return "safe"
            Dim interactionMode = If(intent.ResponseMode, "").Trim().ToLowerInvariant()
            If interactionMode = "answer" OrElse interactionMode = "clarify" Then Return "safe"
            If intent.RequiresVBA Then Return "medium"
            Return "safe"
        End Function

        Private Sub EnrichTrace(trace As ChatContextTrace,
                                request As AiNativeRequest,
                                intent As IntentResult,
                                taskSpec As AgentTaskSpec,
                                selectedSkills As List(Of SkillFileDefinition),
                                appType As String)
            If trace Is Nothing Then Return

            trace.AppType = appType
            trace.Query = If(request.UserInput, "")

            If intent IsNot Nothing Then
                trace.IntentType = intent.OfficeIntent.ToString()
                trace.IntentDescription = intent.UserFriendlyDescription
            End If

            If request.OfficeContext IsNot Nothing Then
                trace.OfficeContext = request.OfficeContext.ToPromptText()
            End If

            If taskSpec IsNot Nothing Then
                trace.TaskSpec = New ChatContextTaskSpecTrace With {
                    .Goal = taskSpec.Goal,
                    .TargetObject = taskSpec.TargetObject,
                    .Complexity = taskSpec.Complexity,
                    .RiskLevel = taskSpec.RiskLevel,
                    .Constraints = taskSpec.Constraints,
                    .SuccessCriteria = taskSpec.SuccessCriteria,
                    .RequiredTools = taskSpec.RequiredTools
                }
            End If

            If selectedSkills IsNot Nothing Then
                For Each skill In selectedSkills
                    If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then Continue For
                    If trace.Skills.Any(Function(s) String.Equals(s.Name, skill.Name, StringComparison.OrdinalIgnoreCase)) Then Continue For
                    trace.Skills.Add(New ChatContextSkillTrace With {
                        .Name = skill.Name,
                        .Source = "runtime",
                        .Reason = "AI Native runtime 自动选择"
                    })
                Next
            End If
        End Sub

        Private Function SelectSkills(input As String, intent As IntentResult, appType As String) As List(Of SkillFileDefinition)
            Dim selected As New List(Of SkillFileDefinition)()
            Dim intentType = If(intent Is Nothing, Nothing, intent.OfficeIntent.ToString())

            Try
                selected = SkillsIndexService.SelectSkillDefinitions(input, intentType, appType, 3)
            Catch ex As Exception
                Debug.WriteLine($"[AiNativeRuntime] SkillsIndexService selection failed: {ex.Message}")
            End Try

            If selected Is Nothing OrElse selected.Count = 0 Then
                Dim matches = SkillsService.MatchSkills(input, 3)
                If matches IsNot Nothing Then
                    selected = matches.
                        Where(Function(m) m.MatchScore >= 10).
                        Select(Function(m) SkillsDirectoryService.LoadSkillDetail(m.Skill)).
                        Where(Function(s) s IsNot Nothing).
                        ToList()
                End If
            End If

            Return If(selected, New List(Of SkillFileDefinition)())
        End Function

        Private Function BuildContextSnapshot(request As AiNativeRequest, appType As String) As JObject
            Dim snapshot As JObject
            If request.ContextSnapshot IsNot Nothing Then
                snapshot = CType(request.ContextSnapshot.DeepClone(), JObject)
            Else
                snapshot = New JObject()
            End If

            snapshot("appType") = appType
            snapshot("userInput") = If(request.UserInput, "")

            If request.OfficeContext IsNot Nothing Then
                snapshot("officeContext") = request.OfficeContext.ToPromptText()
            End If

            If request.HistoryMessages IsNot Nothing AndAlso request.HistoryMessages.Count > 0 Then
                Dim history As New JArray()
                For Each msg In request.HistoryMessages.Take(8)
                    history.Add(New JObject From {
                        {"role", If(msg.role, "")},
                        {"content", If(msg.content, "")}
                    })
                Next
                snapshot("conversationHistory") = history
            End If

            Return snapshot
        End Function

        Private Function NormalizeHistory(history As List(Of HistoryMessage)) As List(Of HistoryMessage)
            Dim result As New List(Of HistoryMessage)()
            If history Is Nothing Then Return result

            For Each msg In history
                If msg IsNot Nothing AndAlso msg.role <> "system" AndAlso Not String.IsNullOrWhiteSpace(msg.content) Then
                    result.Add(msg)
                End If
            Next

            Return result
        End Function

        Private Function BuildFallbackMessages(request As AiNativeRequest) As List(Of HistoryMessage)
            Dim result As New List(Of HistoryMessage)()
            If Not String.IsNullOrWhiteSpace(request.SystemPrompt) Then
                result.Add(New HistoryMessage With {.role = "system", .content = request.SystemPrompt})
            End If
            For Each msg In NormalizeHistory(request.HistoryMessages)
                result.Add(msg)
            Next
            result.Add(New HistoryMessage With {.role = "user", .content = If(request.UserInput, "")})
            Return result
        End Function
    End Class

End Namespace
