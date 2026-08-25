Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core.Agent.Context
Imports ExcelAgent.Core.Agent.Harness

Namespace Agent

    ''' <summary>
    ''' Small UI-facing coordinator. Each request gets a fresh Kernel/Harness while the chat
    ''' history remains bounded in memory. Cancellation and approval always target the live run.
    ''' </summary>
    Public Class AgentRunner
        Private ReadOnly _history As New List(Of HistoryMessage)()
        Private _kernel As AgentKernel
        Private _harness As OfficeHarness
        Private _cancellation As CancellationTokenSource
        Private _currentRunId As String = ""

        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), CancellationToken, Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, CancellationToken, Task(Of String))
        Public Property ExecuteHostTool As Func(Of String, String, Boolean, ToolResult)
        Public Property CaptureContextPack As Func(Of ContextPack)

        Public Event PhaseChanged(args As HarnessPhaseChangedEventArgs)
        Public Event StepChanged(args As HarnessStepChangedEventArgs)
        Public Event ContextReady(args As HarnessContextEventArgs)
        Public Event PlanGenerated(plan As ExecutionPlan)
        Public Event IterationUpdated(iteration As ReActIteration)
        Public Event ExecutionExplained(explanation As ExecutionExplanation)
        Public Event Completed(result As HarnessRunResult)

        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _cancellation IsNot Nothing AndAlso Not _cancellation.IsCancellationRequested
            End Get
        End Property

        Public Async Function RunAsync(userText As String,
                                       hostContextText As String,
                                       officeContext As OfficeContext,
                                       contextPack As ContextPack,
                                       Optional executionMode As String = "read_only") As Task(Of HarnessRunResult)
            If IsRunning Then
                Return New HarnessRunResult With {
                    .Status = HarnessRunStatus.Failed,
                    .UserMessage = "已有任务正在执行",
                    .ErrorCode = "RUN_ALREADY_ACTIVE"
                }
            End If
            If String.IsNullOrWhiteSpace(userText) Then
                Return New HarnessRunResult With {.Status = HarnessRunStatus.Failed, .UserMessage = "请输入 Excel 操作目标"}
            End If

            _cancellation = New CancellationTokenSource()
            _kernel = New AgentKernel With {
                .ExecuteCodeWithToolResult = ExecuteHostTool,
                .CaptureContextPack = CaptureContextPack
            }
            _kernel.SendAIRequest = Function(prompt, systemPrompt, history)
                                        If SendAIRequest Is Nothing Then Throw New InvalidOperationException("AI gateway is not configured")
                                        Return SendAIRequest(prompt, systemPrompt, MergeHistory(history), _cancellation.Token)
                                    End Function
            _kernel.SendAIRequestWithMessages = Function(messages)
                                                    If SendAIRequestWithMessages Is Nothing Then Return Task.FromResult(Of String)(Nothing)
                                                    Return SendAIRequestWithMessages(messages, _cancellation.Token)
                                                End Function
            For Each message In _history
                _kernel.AddHistoryMessage(message.role, message.content)
            Next
            AddHandler _kernel.OnPlanGenerated, Sub(plan) RaiseEvent PlanGenerated(plan)
            AddHandler _kernel.OnIterationUpdate, Sub(iteration) RaiseEvent IterationUpdated(iteration)
            AddHandler _kernel.OnExecutionExplained, Sub(explanation) RaiseEvent ExecutionExplained(explanation)
            _kernel.Initialize()

            _harness = New OfficeHarness(_kernel)
            AddHandler _harness.PhaseChanged, AddressOf HandlePhaseChanged
            AddHandler _harness.StepChanged, Sub(sender, args) RaiseEvent StepChanged(args)
            AddHandler _harness.ContextReady, Sub(sender, args) RaiseEvent ContextReady(args)

            _history.Add(New HistoryMessage With {.role = "user", .content = userText})
            TrimHistory()
            Dim turn As New UserTurn With {
                .AppType = "Excel",
                .Text = userText,
                .Mode = If(String.Equals(executionMode, "execute", StringComparison.OrdinalIgnoreCase), "execute", "read_only"),
                .HostContextText = If(hostContextText, ""),
                .OfficeContext = officeContext,
                .ContextPack = contextPack
            }
            Dim result = Await _harness.RunAsync(turn, _cancellation.Token)
            _currentRunId = result.RunId
            If result.Status <> HarnessRunStatus.AwaitingApproval Then FinishRun(result)
            Return result
        End Function

        Public Async Function ApproveAsync(approved As Boolean) As Task(Of HarnessRunResult)
            If _harness Is Nothing OrElse String.IsNullOrWhiteSpace(_currentRunId) Then Return MissingRunResult()
            Dim result = Await _harness.ApproveAsync(_currentRunId, approved, CancellationToken.None)
            _currentRunId = result.RunId
            If result.Status <> HarnessRunStatus.AwaitingApproval Then FinishRun(result)
            Return result
        End Function

        Public Async Function CancelAsync() As Task(Of HarnessRunResult)
            If _harness Is Nothing OrElse String.IsNullOrWhiteSpace(_currentRunId) Then
                _cancellation?.Cancel()
                Return MissingRunResult()
            End If
            Dim result = Await _harness.CancelAsync(_currentRunId, CancellationToken.None)
            _cancellation?.Cancel()
            ' The original RunAsync owns terminal completion. It will observe the linked
            ' cancellation and emit exactly one Completed event after the loop unwinds.
            Return result
        End Function

        Private Sub HandlePhaseChanged(sender As Object, args As HarnessPhaseChangedEventArgs)
            If args IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(args.RunId) Then _currentRunId = args.RunId
            RaiseEvent PhaseChanged(args)
        End Sub

        Private Sub FinishRun(result As HarnessRunResult)
            If result IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(result.UserMessage) Then
                _history.Add(New HistoryMessage With {.role = "assistant", .content = result.UserMessage})
                TrimHistory()
            End If
            _cancellation?.Dispose()
            _cancellation = Nothing
            RaiseEvent Completed(result)
        End Sub

        Private Function MergeHistory(runHistory As List(Of HistoryMessage)) As List(Of HistoryMessage)
            If runHistory IsNot Nothing AndAlso runHistory.Count > 0 Then Return runHistory
            Return _history.Select(Function(item) New HistoryMessage With {.role = item.role, .content = item.content}).ToList()
        End Function

        Private Sub TrimHistory()
            While _history.Count > 20
                _history.RemoveAt(0)
            End While
        End Sub

        Private Shared Function MissingRunResult() As HarnessRunResult
            Return New HarnessRunResult With {
                .Status = HarnessRunStatus.Failed,
                .UserMessage = "没有可继续的任务",
                .ErrorCode = ExceptionClassifier.CodeApprovalUnavailable
            }
        End Function
    End Class
End Namespace
