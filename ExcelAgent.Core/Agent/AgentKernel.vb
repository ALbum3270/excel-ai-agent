Imports System.IO
Imports System.Threading
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Single Excel Agent entry point: goal compilation, adaptive ReAct, registered tools,
    ''' host observations, repair, and grounded completion all share one session.
    ''' </summary>
    Public Class AgentKernel
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _loopEngine As LoopEngine
        Private _session As AgentSession

        Public Property PromptsDirectory As String
        Public Property ToolsDirectory As String
        Public Property SkillsDirectory As String
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, Task(Of String))
        Public Property ExecuteCodeWithToolResult As Func(Of String, String, Boolean, ToolResult)
        Public Property CaptureContextPack As Func(Of Context.ContextPack)

        Public Event OnStatusChanged(status As String)
        Public Event OnIterationUpdate(iteration As ReActIteration)
        Public Event OnStepCompleted(stepIndex As Integer, success As Boolean, message As String)
        Public Event OnExecutionExplained(explanation As ExecutionExplanation)
        Public Event OnRequestApproval(message As String, callback As Action(Of Boolean))
        Public Event OnPlanGenerated(plan As ExecutionPlan)
        Public Event OnCompleted(result As AgentResult)

        Public Sub New()
            Dim baseDirectory = ResolveRuntimeBaseDirectory()
            PromptsDirectory = Path.Combine(baseDirectory, "Prompts")
            ToolsDirectory = Path.Combine(baseDirectory, "Tools")
            SkillsDirectory = Path.Combine(baseDirectory, "Skills")
            _promptManager = New PromptManager(PromptsDirectory)
            _toolRegistry = New ToolRegistry()
            _memory = New AgentMemory()
            _loopEngine = New LoopEngine(_toolRegistry, _memory, _promptManager)
        End Sub

        Public Sub Initialize()
            Dim added = _toolRegistry.LoadFromRuntimeDirectories(ToolsDirectory)
            If added = 0 Then AppLogger.Warn("AgentKernel", "No Excel tool manifests were loaded")

            _loopEngine.SendAIRequest = Function(prompt, systemPrompt, history)
                                            If SendAIRequest Is Nothing Then Throw New InvalidOperationException("AI request callback is not configured")
                                            Return SendAIRequest(prompt, systemPrompt, history)
                                        End Function
            _loopEngine.SendAIRequestWithMessages = Function(messages)
                                                        If SendAIRequestWithMessages Is Nothing Then Return Task.FromResult(Of String)(Nothing)
                                                        Return SendAIRequestWithMessages(messages)
                                                    End Function
            _loopEngine.CaptureContextPack = Function()
                                                 Return If(CaptureContextPack Is Nothing, Nothing, CaptureContextPack.Invoke())
                                             End Function
            _loopEngine.OnPlanGenerated = Sub(plan) RaiseEvent OnPlanGenerated(plan)
            _loopEngine.OnStatusChanged = Sub(status) RaiseEvent OnStatusChanged(status)
            _loopEngine.OnIterationUpdate = Sub(iteration) RaiseEvent OnIterationUpdate(iteration)
            _loopEngine.OnStepCompleted = Sub(index, success, message) RaiseEvent OnStepCompleted(index, success, message)
            _loopEngine.OnExecutionExplained = Sub(explanation) RaiseEvent OnExecutionExplained(explanation)
            _loopEngine.OnRequestApproval = Async Function(message)
                                                If OnRequestApprovalEvent Is Nothing Then
                                                    Throw New InvalidOperationException("No approval handler is registered")
                                                End If
                                                Dim signal As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                                                RaiseEvent OnRequestApproval(message, Sub(approved) signal.TrySetResult(approved))
                                                Return Await signal.Task
                                            End Function
            _memory.SendAIRequest = SendAIRequest
        End Sub

        Public Async Function ExecuteAsync(userRequest As String,
                                           appType As String,
                                           currentContent As String,
                                           Optional officeContext As Context.OfficeContext = Nothing,
                                           Optional contextPack As Context.ContextPack = Nothing,
                                           Optional taskSpec As AgentTaskSpec = Nothing,
                                           Optional selectedSkills As List(Of SkillFileDefinition) = Nothing,
                                           Optional executionMode As String = "read_only",
                                           Optional cancellationToken As CancellationToken = Nothing) As Task(Of AgentResult)
            _session = New AgentSession(userRequest, "Excel", currentContent) With {
                .Spec = taskSpec,
                .RequestedMutationPolicy = If(String.Equals(executionMode, "execute", StringComparison.OrdinalIgnoreCase),
                                              "allow_mutation",
                                              "read_only")
            }
            Try
                cancellationToken.ThrowIfCancellationRequested()
                If Not String.Equals(If(appType, "Excel"), "Excel", StringComparison.OrdinalIgnoreCase) Then
                    Return AgentResult.Failed(_session.Id, "此独立项目只支持 Excel", taskFatal:=True, errorCode:=ExceptionClassifier.CodeHostUnsupported)
                End If

                _memory.ClearWorking()
                _memory.AddSessionMessage("user", userRequest)
                If officeContext Is Nothing Then officeContext = New Context.OfficeContext With {.AppType = "Excel"}
                If contextPack Is Nothing Then contextPack = Context.ContextPack.FromOfficeContext(officeContext, currentContent)
                _memory.SetWorking("lastOfficeContext", officeContext)
                _memory.SetWorking("lastContextPack", contextPack)
                _toolRegistry.ExecuteCodeWithToolResult = ExecuteCodeWithToolResult

                Dim skill = SelectExcelSkill(selectedSkills)
                _session.Skill = skill
                Dim executionContext = ToolExecutionContext.FromSession(_session, skill)
                executionContext.ReadOnlyTask = String.Equals(
                    _session.RequestedMutationPolicy,
                    "read_only",
                    StringComparison.OrdinalIgnoreCase)
                Dim visibleTools = _toolRegistry.GetVisibleTools("Excel", executionContext)
                Dim systemPrompt = _promptManager.BuildSystemPrompt("Excel", visibleTools, _memory)
                Dim result = Await _loopEngine.RunAsync(_session, systemPrompt, skill, cancellationToken)

                _memory.AddSessionMessage("assistant", If(result?.Message, ""))
                RaiseEvent OnCompleted(result)
                Return result
            Catch ex As OperationCanceledException
                Dim cancelled = AgentResult.Failed(_session.Id, "已取消", errorCode:=ExceptionClassifier.CodeCancelled)
                RaiseEvent OnCompleted(cancelled)
                Return cancelled
            Catch ex As Exception
                Dim classified = ExceptionClassifier.Classify(ex)
                Dim failed = AgentResult.Failed(_session.Id,
                                                $"执行异常: [{classified.ErrorCode}] {classified.UserMessage}",
                                                classified.TaskFatal,
                                                classified.SessionFatal,
                                                classified.ErrorCode)
                AppLogger.Error("AgentKernel", "ExecuteAsync failed", ex)
                RaiseEvent OnCompleted(failed)
                Return failed
            End Try
        End Function

        Public Sub AddHistoryMessage(role As String, content As String)
            _memory.AddSessionMessage(role, content)
        End Sub

        Public ReadOnly Property ToolCount As Integer
            Get
                Return _toolRegistry.ToolCount
            End Get
        End Property

        Public ReadOnly Property SkillCount As Integer
            Get
                Return SkillsDirectoryService.GetSkillsCatalog().Count
            End Get
        End Property

        Private Function SelectExcelSkill(selectedSkills As List(Of SkillFileDefinition)) As AgentSkill
            Dim definition As SkillFileDefinition = Nothing
            If selectedSkills IsNot Nothing AndAlso selectedSkills.Count > 0 Then definition = selectedSkills(0)
            If definition Is Nothing Then definition = SkillsDirectoryService.GetSkillsCatalog().FirstOrDefault()
            If definition Is Nothing Then Return Nothing
            If Not definition.IsContentLoaded Then definition = SkillsDirectoryService.LoadSkillDetail(definition)
            If definition Is Nothing Then Return Nothing
            _session.SelectedSkill = definition

            Dim prompt = New Text.StringBuilder()
            prompt.AppendLine("# " & definition.Name)
            prompt.AppendLine(definition.Description)
            prompt.AppendLine()
            prompt.AppendLine(If(definition.Content, ""))
            Dim skill As New AgentSkill With {
                .Id = "filesystem." & definition.Name,
                .Name = definition.Name,
                .Description = definition.Description,
                .PromptTemplate = prompt.ToString(),
                .MaxSteps = 12,
                .AutoApprove = False
            }
            For Each tag In definition.Tags
                skill.TriggerPatterns.Add(tag)
            Next
            For Each toolId In definition.AllowedTools
                skill.RequiredTools.Add(toolId)
            Next
            Return skill
        End Function

        Private Shared Function ResolveRuntimeBaseDirectory() As String
            For Each candidate In {Path.GetDirectoryName(GetType(AgentKernel).Assembly.Location), AppDomain.CurrentDomain.BaseDirectory}
                If String.IsNullOrWhiteSpace(candidate) Then Continue For
                If Directory.Exists(Path.Combine(candidate, "Tools")) OrElse Directory.Exists(Path.Combine(candidate, "Prompts")) Then Return candidate
            Next
            Return AppDomain.CurrentDomain.BaseDirectory
        End Function
    End Class
End Namespace
