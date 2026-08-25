' ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb
' Extracts smart-mode intake (analyze → precheck → adaptive Agent) from BaseChatControl.

Imports System.Diagnostics
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' Host callbacks required by smart-mode chat routing.
''' Implemented by BaseChatControl so the orchestrator stays free of UI inheritance.
''' </summary>
Public Interface IChatRoutingHost
    Function GetHistoryMessages() As List(Of HistoryMessage)
    Function IsFollowUpQuestionAsync(question As String, history As List(Of HistoryMessage)) As Task(Of Boolean)
    Function GetContextSnapshot() As JObject
    Sub EnrichContextForIntent(snapshot As JObject,
                               originalQuestion As String,
                               filePaths As List(Of String),
                               selectedContents As List(Of SendMessageReferenceContentItem))
    Function GetApplicationType() As String
    Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
    Function AnalyzeAiNativeAsync(request As Agent.AiNativeRequest) As Task(Of Agent.AiNativeRuntimeResult)
    Function BuildExecutionContextAsync(originalQuestion As String,
                                        filePaths As List(Of String),
                                        selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ExecutionContext)
    Function PreSendCheckAsync(execContext As ExecutionContext) As Task(Of ContextCheckResult)
    Sub ShowContextHints(ragCount As Integer, intentDescription As String, contextTrace As ChatContextTrace)
    Sub ShowWarning(message As String)
    Sub CompleteBlockedRequest(message As String)
    Sub ShowIdentifyingStatus()
    Sub StartAgentPlanningFlow(message As String, intent As IntentResult, analysis As Agent.AiNativeRuntimeResult)
    Sub SetCurrentIntentResult(intent As IntentResult)
End Interface

''' <summary>
''' Decision outcome of smart-mode routing (for tests / diagnostics).
''' </summary>
Public Enum ChatRouteDecision
    BlockedByPreCheck
    AgentKernel
    RoutingFailed
End Enum

''' <summary>
''' Coordinates AI Native analysis and the smart-mode route to AgentKernel.
''' Product primary path (P0-2): Analyze → AgentKernel/Loop → Tool/Capability Executor.
''' </summary>
Public Class ChatRoutingOrchestrator
    Private ReadOnly _host As IChatRoutingHost
    Private _lastExecutionTaskSpec As Agent.AgentTaskSpec

    Public Sub New(host As IChatRoutingHost)
        If host Is Nothing Then Throw New ArgumentNullException(NameOf(host))
        _host = host
    End Sub

    ''' <summary>
    ''' Smart mode: follow-up detection → AiNative analyze → pre-check → AgentKernel.
    ''' </summary>
    Public Async Function RouteSmartModeAsync(finalMessageToLLM As String,
                                              originalQuestion As String,
                                              filePaths As List(Of String),
                                              selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ChatRouteDecision)

        Dim routeClock = Stopwatch.StartNew()
        AppLogger.Info("ChatRoutingOrchestrator", "Smart route started")
        Try
            Dim hasReferences As Boolean =
                (filePaths IsNot Nothing AndAlso filePaths.Count > 0) OrElse
                (selectedContents IsNot Nothing AndAlso selectedContents.Count > 0)

            Dim history = _host.GetHistoryMessages()
            If history Is Nothing Then history = New List(Of HistoryMessage)()

            Dim isFollowUp As Boolean = False
            If history.Count >= 2 AndAlso Not String.IsNullOrWhiteSpace(originalQuestion) Then
                isFollowUp = Await _host.IsFollowUpQuestionAsync(originalQuestion, history)
                Debug.WriteLine($"[ChatRoutingOrchestrator] follow-up check: isFollowUp={isFollowUp}")
            End If

            Dim contextSnapshot = _host.GetContextSnapshot()
            If contextSnapshot Is Nothing Then contextSnapshot = New JObject()
            _host.EnrichContextForIntent(contextSnapshot, originalQuestion, filePaths, selectedContents)

            Dim recentHistory = history.
                Where(Function(m) m.role <> "system" AndAlso Not String.IsNullOrEmpty(m.content)).
                ToList()
            If recentHistory.Count > 0 Then
                Dim historyArray As New JArray()
                Dim takeCount = Math.Min(6, recentHistory.Count)
                For i = recentHistory.Count - takeCount To recentHistory.Count - 1
                    Dim hMsg = recentHistory(i)
                    Dim content = hMsg.content
                    If content IsNot Nothing AndAlso content.Length > 300 Then
                        content = content.Substring(0, 300) & "..."
                    End If
                    historyArray.Add(New JObject From {
                        {"role", hMsg.role},
                        {"content", content}
                    })
                Next
                contextSnapshot("conversationHistory") = historyArray
            End If

            Dim appType = _host.GetApplicationType()
            Dim officeContext = _host.CaptureOfficeContext(appType)
            Dim aiNativeRequest As New Agent.AiNativeRequest With {
                .UserInput = originalQuestion,
                .AppType = appType,
                .SystemPrompt = "",
                .RequestUuid = Guid.NewGuid().ToString(),
                .OfficeContext = officeContext,
                .ContextSnapshot = contextSnapshot,
                .HistoryMessages = recentHistory,
                .PreviousTaskSpec = If(isFollowUp OrElse Agent.AiNativeRuntime.IsDestinationOnlyCorrection(originalQuestion),
                                       _lastExecutionTaskSpec,
                                       Nothing),
                .EnableMemory = MemoryConfig.EnableUserProfile OrElse MemoryConfig.RagTopN > 0,
                .UseContextBuilder = MemoryConfig.UseContextBuilder
            }

            Dim analyzeClock = Stopwatch.StartNew()
            Dim aiNativeResult = Await _host.AnalyzeAiNativeAsync(aiNativeRequest)
            AppLogger.Info("ChatRoutingOrchestrator", $"Analyze completed elapsedMs={analyzeClock.ElapsedMilliseconds}")
            Dim intent = If(aiNativeResult?.Intent, New IntentResult())
            intent.OriginalInput = originalQuestion
            intent.IsFollowUp = isFollowUp
            _host.SetCurrentIntentResult(intent)

            If aiNativeResult IsNot Nothing AndAlso aiNativeResult.ContextTrace IsNot Nothing Then
                _host.ShowContextHints(
                    aiNativeResult.RagCount,
                    If(intent.UserFriendlyDescription, ""),
                    aiNativeResult.ContextTrace)
            End If

            If ShouldRunLegacyPreCheck(appType, intent) Then
                Try
                    Dim preCheckClock = Stopwatch.StartNew()
                    Dim execContext = Await _host.BuildExecutionContextAsync(originalQuestion, filePaths, selectedContents)
                    If execContext IsNot Nothing Then
                        execContext.IntentResult = intent
                        Dim preCheck = Await _host.PreSendCheckAsync(execContext)
                        AppLogger.Info("ChatRoutingOrchestrator", $"Legacy Word pre-check completed elapsedMs={preCheckClock.ElapsedMilliseconds}")
                        If preCheck IsNot Nothing AndAlso Not preCheck.IsValid Then
                            Dim errors = String.Join(";", preCheck.Errors)
                            Debug.WriteLine($"[ChatRoutingOrchestrator] PreSendCheck blocked: {errors}")
                            _host.ShowWarning($"请求未通过预检: {errors}")
                            _host.CompleteBlockedRequest($"请求未执行：{errors}")
                            Return ChatRouteDecision.BlockedByPreCheck
                        End If
                        If preCheck IsNot Nothing AndAlso preCheck.Warnings IsNot Nothing AndAlso preCheck.Warnings.Count > 0 Then
                            Debug.WriteLine($"[ChatRoutingOrchestrator] PreSendCheck warnings: {String.Join(";", preCheck.Warnings)}")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ChatRoutingOrchestrator] PreSendCheck exception: {ex.Message}")
                End Try
            End If

            If hasReferences AndAlso String.IsNullOrWhiteSpace(originalQuestion) Then
                intent.UserFriendlyDescription = "已根据引用内容自动识别处理意图"
                _host.SetCurrentIntentResult(intent)
            End If

            ' Smart mode has one execution model. The adaptive Agent decides on every turn
            ' whether to answer, clarify, read, compute or mutate; a model-produced routing
            ' label must never switch the request into a separate batch/chat protocol.
            Dim routeDecision = DecidePostAnalysisRoute(
                isFollowUp,
                intent,
                IntentAcceptancePolicy.HasExecutableToolRequirement(aiNativeResult?.TaskSpec))
            AppLogger.Info("ChatRoutingOrchestrator", $"Route decision={routeDecision} appType={appType} elapsedMs={routeClock.ElapsedMilliseconds}")

            ' The only smart-mode product path is AgentKernel.
            Debug.WriteLine($"[ChatRoutingOrchestrator] primary path AgentKernel intent={intent.OfficeIntent}, confidence={intent.Confidence:F2}")
            If aiNativeResult?.TaskSpec IsNot Nothing Then _lastExecutionTaskSpec = aiNativeResult.TaskSpec
            _host.ShowIdentifyingStatus()
            _host.StartAgentPlanningFlow(finalMessageToLLM, intent, aiNativeResult)
            AppLogger.Info("ChatRoutingOrchestrator", $"Agent planning dispatched elapsedMs={routeClock.ElapsedMilliseconds}")
            Return ChatRouteDecision.AgentKernel

        Catch ex As Exception
            Debug.WriteLine($"[ChatRoutingOrchestrator] analyze/route failed: {ex.Message}")
            AppLogger.Error("ChatRoutingOrchestrator", "Smart route failed", ex)
            Dim routeError = ExceptionClassifier.ToUserMessage(ex)
            _host.CompleteBlockedRequest($"请求路由失败：{routeError}")
            Return ChatRouteDecision.RoutingFailed
        End Try
    End Function

    ''' <summary>
    ''' The legacy pre-check validates Word document selections/templates and must not gate
    ''' native Excel/PowerPoint Agent tools. In Excel an empty active cell is common even
    ''' when OfficeContext contains a valid table region.
    ''' </summary>
    Public Shared Function ShouldRunLegacyPreCheck(appType As String, intent As IntentResult) As Boolean
        If Not String.Equals(If(appType, "").Trim(), "Word", StringComparison.OrdinalIgnoreCase) Then Return False
        If intent Is Nothing Then Return False
        Return intent.OfficeIntent = OfficeIntentType.FORMAT_STYLE OrElse
               intent.OfficeIntent = OfficeIntentType.TEXT_FORMAT
    End Function

    ''' <summary>
    ''' Smart mode deliberately has one route. Follow-up state and interaction mode remain
    ''' input facts for the adaptive model, not selectors for a different execution engine.
    ''' </summary>
    Public Shared Function DecidePostAnalysisRoute(isFollowUp As Boolean,
                                                    intent As IntentResult,
                                                    taskSpecRequiresExecution As Boolean) As ChatRouteDecision
        Return ChatRouteDecision.AgentKernel
    End Function

End Class
