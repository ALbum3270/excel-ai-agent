Imports System.Threading
Imports System.Threading.Tasks

Namespace Agent.Harness

    Public Class OfficeHarness
        Implements IOfficeHarness

        Private ReadOnly _kernel As AgentKernel
        Private ReadOnly _runTraceStore As IRunTraceStore
        Private _currentRunId As String = ""

        Public Event PhaseChanged As EventHandler(Of HarnessPhaseChangedEventArgs) Implements IOfficeHarness.PhaseChanged
        Public Event StepChanged As EventHandler(Of HarnessStepChangedEventArgs) Implements IOfficeHarness.StepChanged
        Public Event ContextReady As EventHandler(Of HarnessContextEventArgs) Implements IOfficeHarness.ContextReady

        Public Sub New(kernel As AgentKernel, Optional runTraceStore As IRunTraceStore = Nothing)
            If kernel Is Nothing Then Throw New ArgumentNullException(NameOf(kernel))
            _kernel = kernel
            _runTraceStore = If(runTraceStore, New NoopRunTraceStore())
            AddHandler _kernel.OnStatusChanged, AddressOf HandleKernelStatusChanged
            AddHandler _kernel.OnStepCompleted, AddressOf HandleKernelStepCompleted
            AddHandler _kernel.OnPlanGenerated, AddressOf HandleKernelPlanGenerated
            AddHandler _kernel.OnExecutionExplained, AddressOf HandleKernelExecutionExplained
        End Sub

        Public Async Function RunAsync(turn As UserTurn,
                                       cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.RunAsync
            Dim runId = Guid.NewGuid().ToString()
            _currentRunId = runId
            Dim startedAt = DateTime.Now

            If turn Is Nothing Then
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Failed,
                    .UserMessage = "用户请求为空",
                    .DebugMessage = "UserTurn is null",
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            End If

            Try
                cancellationToken.ThrowIfCancellationRequested()
                SafeStartRun(runId, turn, startedAt)
                RaisePhase(runId, "starting", "Harness run started")

                Dim officeContext = turn.OfficeContext
                If officeContext Is Nothing Then
                    officeContext = New Agent.Context.OfficeContext With {.AppType = turn.AppType}
                End If

                RaiseEvent ContextReady(Me, New HarnessContextEventArgs With {
                    .RunId = runId,
                    .AppType = turn.AppType,
                    .ContextText = If(turn.HostContextText, "")
                })

                RaisePhase(runId, "executing", "AgentKernel executing")
                Dim agentResult = Await _kernel.ExecuteAsync(If(turn.Text, ""),
                                                             If(turn.AppType, ""),
                                                             If(turn.HostContextText, ""),
                                                             officeContext)

                cancellationToken.ThrowIfCancellationRequested()

                Dim succeeded = agentResult IsNot Nothing AndAlso agentResult.Success
                Dim status = If(succeeded, HarnessRunStatus.Succeeded, HarnessRunStatus.Failed)
                Dim message = If(agentResult?.Message, If(succeeded, "执行完成", "执行失败"))

                RaisePhase(runId, If(succeeded, "completed", "failed"), message)
                SafeCompleteRun(runId, If(succeeded, "succeeded", "failed"), message, If(succeeded, "", ExceptionClassifier.CodeUnknown), DateTime.Now)
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = status,
                    .UserMessage = message,
                    .DebugMessage = message,
                    .AgentSessionId = If(agentResult?.SessionId, ""),
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            Catch ex As OperationCanceledException
                RaisePhase(runId, "cancelled", "Harness run cancelled")
                SafeCompleteRun(runId, "cancelled", "已取消", ExceptionClassifier.CodeCancelled, DateTime.Now)
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Cancelled,
                    .UserMessage = "已取消",
                    .DebugMessage = ex.Message,
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            Catch ex As Exception
                AppLogger.Error("OfficeHarness", "RunAsync exception", ex)
                Dim classified = ExceptionClassifier.Classify(ex)
                RaisePhase(runId, "failed", classified.UserMessage)
                SafeCompleteRun(runId, "failed", classified.UserMessage, classified.ErrorCode, DateTime.Now)
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Failed,
                    .UserMessage = classified.UserMessage,
                    .DebugMessage = classified.DebugDetail,
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            Finally
                _currentRunId = ""
            End Try
        End Function

        Private Sub HandleKernelStatusChanged(status As String)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            RaisePhase(_currentRunId, "status", If(status, ""))
        End Sub

        Private Sub HandleKernelStepCompleted(stepIndex As Integer, success As Boolean, message As String)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            RaiseEvent StepChanged(Me, New HarnessStepChangedEventArgs With {
                .RunId = _currentRunId,
                .StepIndex = stepIndex,
                .Status = If(success, "completed", "failed"),
                .Message = If(message, "")
            })
        End Sub

        Private Sub HandleKernelPlanGenerated(plan As ExecutionPlan)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            Dim count = If(plan?.Steps?.Count, 0)
            RaisePhase(_currentRunId, "planned", $"生成执行计划：{count} 步")
        End Sub

        Private Sub HandleKernelExecutionExplained(explanation As ExecutionExplanation)
            If String.IsNullOrWhiteSpace(_currentRunId) OrElse explanation Is Nothing Then Return
            Try
                _runTraceStore.AppendStep(_currentRunId,
                                          explanation.StepIndex,
                                          explanation.ToolId,
                                          If(explanation.Success, "succeeded", "failed"),
                                          If(explanation.Message, explanation.ExplanationText),
                                          If(explanation.FailureReason, ""),
                                          explanation,
                                          explanation.StartedAt,
                                          explanation.FinishedAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Append run trace step failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub SafeStartRun(runId As String, turn As UserTurn, startedAt As DateTime)
            Try
                _runTraceStore.StartRun(runId, turn, startedAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Start run trace failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub SafeCompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime)
            Try
                _runTraceStore.CompleteRun(runId, status, finalMessage, errorCode, finishedAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Complete run trace failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub RaisePhase(runId As String, phase As String, message As String)
            RaiseEvent PhaseChanged(Me, New HarnessPhaseChangedEventArgs With {
                .RunId = If(runId, ""),
                .Phase = If(phase, ""),
                .Message = If(message, "")
            })
        End Sub
    End Class

End Namespace
