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

            Dim intentService As New IntentRecognitionService(appType)
            Dim intent = Await intentService.IdentifyIntentAsync(input, contextSnapshot)
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

            Dim tools = _toolRegistry.GetAvailableTools(appType)
            Dim selectedSkills = SelectSkills(input, intent, appType)
            Dim taskSpec = BuildTaskSpec(request, intent, tools, selectedSkills, appType)
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

        Private Function BuildTaskSpec(request As AiNativeRequest,
                                       intent As IntentResult,
                                       tools As List(Of ToolDescriptor),
                                       selectedSkills As List(Of SkillFileDefinition),
                                       appType As String) As AgentTaskSpec
            Dim spec As New AgentTaskSpec()
            Dim input = If(request.UserInput, "").Trim()
            Dim intentDescription = If(intent?.UserFriendlyDescription, "")

            spec.Goal = If(Not String.IsNullOrWhiteSpace(intentDescription), intentDescription, If(String.IsNullOrWhiteSpace(input), "自动分析当前 Office 上下文并选择合适处理方式", input))
            spec.TargetObject = InferTargetObject(request, intent)
            spec.Complexity = InferComplexity(intent)
            spec.RiskLevel = InferRiskLevel(intent)
            spec.Constraints.Add($"宿主应用: {appType}")
            spec.Constraints.Add("不要求用户手动确认意图，优先由 AI 自动识别和执行")
            If selectedSkills IsNot Nothing AndAlso selectedSkills.Count > 0 Then
                spec.Constraints.Add("已自动选择 Skill: " & String.Join(", ", selectedSkills.Select(Function(s) s.Name)))
            End If

            If intent IsNot Nothing Then
                spec.Constraints.Add($"识别意图: {intent.OfficeIntent}")
                If intent.Confidence < 0.4 Then
                    spec.Constraints.Add("低置信度任务采用探索式计划，先分析上下文再选择低风险动作")
                End If
                For Each kv In intent.ExtractedEntities
                    spec.Constraints.Add($"{kv.Key}: {kv.Value}")
                Next
            End If

            spec.SuccessCriteria.Add("输出结果与当前 Office 上下文相关")
            spec.SuccessCriteria.Add("执行过程可解释，并记录使用的上下文、工具和记忆")
            If spec.RiskLevel <> "safe" Then
                spec.SuccessCriteria.Add("执行后保留可观察结果或失败原因")
            End If

            If selectedSkills IsNot Nothing Then
                For Each skill In selectedSkills
                    If skill.AllowedTools IsNot Nothing Then
                        For Each toolId In skill.AllowedTools
                            If Not String.IsNullOrWhiteSpace(toolId) AndAlso Not spec.RequiredTools.Contains(toolId) Then
                                spec.RequiredTools.Add(toolId)
                            End If
                        Next
                    End If
                    If skill.Scripts IsNot Nothing Then
                        For Each script In skill.Scripts
                            Dim scriptToolId = $"skill_script.{skill.Name}.{script.FileName}"
                            If Not spec.RequiredTools.Contains(scriptToolId) Then spec.RequiredTools.Add(scriptToolId)
                        Next
                    End If
                Next
            End If

            If tools IsNot Nothing Then
                For Each tool In tools.Take(8)
                    If Not spec.RequiredTools.Contains(tool.Id) Then spec.RequiredTools.Add(tool.Id)
                Next
            End If

            Return spec
        End Function

        Private Function InferTargetObject(request As AiNativeRequest, intent As IntentResult) As String
            If request.OfficeContext IsNot Nothing AndAlso request.OfficeContext.Selection IsNot Nothing Then
                Dim selection = request.OfficeContext.Selection
                Return $"{If(selection.DataType, "选区")} {If(selection.Address, "")}".Trim()
            End If

            If intent IsNot Nothing AndAlso intent.ExtractedEntities IsNot Nothing AndAlso intent.ExtractedEntities.Count > 0 Then
                Return String.Join("; ", intent.ExtractedEntities.Select(Function(kv) $"{kv.Key}={kv.Value}"))
            End If

            Return "当前 Office 文档/工作区"
        End Function

        Private Function InferComplexity(intent As IntentResult) As String
            If intent Is Nothing Then Return "exploratory"
            If intent.Confidence < 0.4 Then Return "exploratory"
            If intent.CanUseDirectCommand AndAlso Not intent.RequiresVBA Then Return "simple"
            Return "medium"
        End Function

        Private Function InferRiskLevel(intent As IntentResult) As String
            If intent Is Nothing Then Return "safe"
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
