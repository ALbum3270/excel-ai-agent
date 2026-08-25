Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent
Imports ExcelAgent.Core.Agent.Harness

Public Class ChatControl
    Inherits UserControl

    Private Const VirtualHost As String = "excelagent.local"
    Private ReadOnly _webView As New WebView2()
    Private ReadOnly _host As ExcelAgentHost
    Private ReadOnly _runner As New AgentRunner()
    Private _initialized As Boolean

    Public Sub New(host As ExcelAgentHost)
        _host = host
        _webView.Dock = DockStyle.Fill
        Controls.Add(_webView)
        Dock = DockStyle.Fill

        _runner.ExecuteHostTool = AddressOf _host.ExecuteTool
        _runner.CaptureContextPack = AddressOf _host.CaptureContextPack
        _runner.SendAIRequest = AddressOf SendAgentRequestAsync
        _runner.SendAIRequestWithMessages = AddressOf SendMessagesAsync
        AddHandler _runner.PhaseChanged, AddressOf HandlePhaseChanged
        AddHandler _runner.StepChanged, AddressOf HandleStepChanged
        AddHandler _runner.PlanGenerated, AddressOf HandlePlanGenerated
        AddHandler _runner.IterationUpdated, AddressOf HandleIterationUpdated
        AddHandler _runner.ExecutionExplained, AddressOf HandleExecutionExplained
        AddHandler _runner.Completed, AddressOf HandleCompleted
        AddHandler Load, AddressOf HandleLoad
    End Sub

    Public Async Function SubmitPromptAsync(prompt As String, Optional executionMode As String = "execute") As Task
        Await EnsureInitializedAsync()
        Await RunPromptAsync(prompt, executionMode)
    End Function

    Public Sub FocusInput()
        Post(New JObject From {{"type", "focusInput"}})
    End Sub

    Private Async Sub HandleLoad(sender As Object, e As EventArgs)
        Await EnsureInitializedAsync()
    End Sub

    Private Async Function EnsureInitializedAsync() As Task
        If _initialized Then Return
        Await _webView.EnsureCoreWebView2Async()
        Dim webRoot = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Web")
        If Not Directory.Exists(webRoot) Then Throw New DirectoryNotFoundException("Web UI directory is missing: " & webRoot)
        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(VirtualHost, webRoot, CoreWebView2HostResourceAccessKind.DenyCors)
        AddHandler _webView.CoreWebView2.WebMessageReceived, AddressOf HandleWebMessage
        _webView.Source = New Uri("https://" & VirtualHost & "/index.html")
        _initialized = True
    End Function

    Private Async Sub HandleWebMessage(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim message = JObject.Parse(e.WebMessageAsJson)
            Select Case If(message("type")?.ToString(), "")
                Case "send"
                    Await RunPromptAsync(If(message("text")?.ToString(), ""), If(message("mode")?.ToString(), "read_only"))
                Case "cancel"
                    Await _runner.CancelAsync()
                Case "approve"
                    Await _runner.ApproveAsync(True)
                Case "reject"
                    Await _runner.ApproveAsync(False)
                Case "settings"
                    Globals.ThisAddIn.ShowSettings()
            End Select
        Catch ex As Exception
            PostError(ExceptionClassifier.Classify(ex).UserMessage)
        End Try
    End Sub

    Private Async Function RunPromptAsync(prompt As String, executionMode As String) As Task
        prompt = If(prompt, "").Trim()
        If String.IsNullOrWhiteSpace(prompt) Then Return
        If Not AgentSettingsStore.IsConfigured() Then
            PostError("请先配置 API URL、API Key 和模型。")
            Globals.ThisAddIn.ShowSettings()
            Return
        End If

        Post(New JObject From {{"type", "appendMessage"}, {"role", "user"}, {"content", prompt}})
        Post(New JObject From {{"type", "busy"}, {"value", True}})
        Try
            Dim snapshot = _host.CaptureSnapshot()
            Dim result = Await _runner.RunAsync(prompt,
                                                snapshot.PromptText,
                                                snapshot.OfficeContext,
                                                snapshot.ContextPack,
                                                executionMode)
            If result.Status = HarnessRunStatus.AwaitingApproval Then
                PostApproval(result.UserMessage)
            ElseIf result.Status = HarnessRunStatus.Failed AndAlso result.ErrorCode = "RUN_ALREADY_ACTIVE" Then
                PostError(result.UserMessage)
            End If
        Catch ex As Exception
            PostError(ExceptionClassifier.Classify(ex).UserMessage)
            Post(New JObject From {{"type", "busy"}, {"value", False}})
        End Try
    End Function

    Private Async Function SendAgentRequestAsync(prompt As String,
                                                 systemPrompt As String,
                                                 history As List(Of HistoryMessage),
                                                 cancellationToken As CancellationToken) As Task(Of String)
        Dim messages As New JArray()
        If Not String.IsNullOrWhiteSpace(systemPrompt) Then
            messages.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})
        End If
        For Each item In If(history, New List(Of HistoryMessage)()).TakeLast(16)
            If String.IsNullOrWhiteSpace(item?.content) Then Continue For
            messages.Add(New JObject From {{"role", If(item.role, "user")}, {"content", item.content}})
        Next
        messages.Add(New JObject From {{"role", "user"}, {"content", If(prompt, "")}})
        Return Await SendMessagesAsync(messages, cancellationToken)
    End Function

    Private Async Function SendMessagesAsync(messages As JArray, cancellationToken As CancellationToken) As Task(Of String)
        Dim options As New AiRequestOptions With {
            .ApiUrl = ConfigSettings.ApiUrl,
            .ApiKey = ConfigSettings.ApiKey,
            .ModelName = ConfigSettings.ModelName,
            .Platform = ConfigSettings.platform,
            .ReasoningMode = ConfigSettings.ReasoningMode,
            .Messages = messages,
            .TimeoutSeconds = 90,
            .CancellationToken = cancellationToken
        }
        Dim response = Await AiGateway.SendChatAsync(options)
        AiGateway.ThrowIfFailed(response)
        Return response.Content
    End Function

    Private Sub HandlePhaseChanged(args As HarnessPhaseChangedEventArgs)
        If args Is Nothing Then Return
        Post(New JObject From {{"type", "phase"}, {"phase", args.Phase}, {"message", args.Message}})
        If String.Equals(args.Phase, "awaiting_approval", StringComparison.OrdinalIgnoreCase) Then PostApproval(args.Message)
    End Sub

    Private Sub HandleStepChanged(args As HarnessStepChangedEventArgs)
        If args Is Nothing Then Return
        Post(New JObject From {
            {"type", "step"},
            {"index", args.StepIndex},
            {"status", args.Status},
            {"message", args.Message}
        })
    End Sub

    Private Sub HandlePlanGenerated(plan As ExecutionPlan)
        Dim steps As New JArray()
        For Each item In If(plan?.Steps, New List(Of PlanStep)())
            steps.Add(New JObject From {{"index", item.StepNumber}, {"description", item.Description}})
        Next
        Post(New JObject From {{"type", "plan"}, {"understanding", If(plan?.Understanding, "")}, {"steps", steps}})
    End Sub

    Private Sub HandleIterationUpdated(iteration As ReActIteration)
        If iteration Is Nothing Then Return
        Post(New JObject From {
            {"type", "iteration"},
            {"index", iteration.Index},
            {"tool", If(iteration.Action?.ToolId, "")},
            {"observation", If(iteration.Observation, "")}
        })
    End Sub

    Private Sub HandleExecutionExplained(explanation As ExecutionExplanation)
        If explanation Is Nothing Then Return
        Post(New JObject From {
            {"type", "execution"},
            {"tool", explanation.ToolId},
            {"success", explanation.Success},
            {"message", If(explanation.ExplanationText, explanation.Message)}
        })
    End Sub

    Private Sub HandleCompleted(result As HarnessRunResult)
        If result Is Nothing Then Return
        Post(New JObject From {
            {"type", "appendMessage"},
            {"role", "assistant"},
            {"content", result.UserMessage},
            {"success", result.Status = HarnessRunStatus.Succeeded}
        })
        Post(New JObject From {{"type", "approval"}, {"visible", False}})
        Post(New JObject From {{"type", "busy"}, {"value", False}})
    End Sub

    Private Sub PostApproval(message As String)
        Post(New JObject From {{"type", "approval"}, {"visible", True}, {"message", If(message, "该操作需要确认")}})
    End Sub

    Private Sub PostError(message As String)
        Post(New JObject From {{"type", "appendMessage"}, {"role", "error"}, {"content", If(message, "操作失败")}})
    End Sub

    Private Sub Post(payload As JObject)
        If payload Is Nothing OrElse IsDisposed Then Return
        If InvokeRequired Then
            BeginInvoke(New Action(Of JObject)(AddressOf Post), payload)
            Return
        End If
        If _webView.CoreWebView2 IsNot Nothing Then _webView.CoreWebView2.PostWebMessageAsJson(payload.ToString(Newtonsoft.Json.Formatting.None))
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _webView.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class

Friend Module EnumerableCompatibility
    <System.Runtime.CompilerServices.Extension>
    Public Function TakeLast(Of T)(source As IEnumerable(Of T), count As Integer) As IEnumerable(Of T)
        Dim values = If(source, Enumerable.Empty(Of T)()).ToList()
        Return values.Skip(Math.Max(0, values.Count - Math.Max(0, count)))
    End Function
End Module
