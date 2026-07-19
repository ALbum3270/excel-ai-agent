Imports Newtonsoft.Json.Linq

Namespace Agent.Harness

    ''' <summary>
    ''' 宿主确定性 capability fast-path 的轻量安全与追踪适配器。
    ''' fast-path 不得绕过 Safety，也必须与 Agent Loop 写入同一 RunTrace 表。
    ''' </summary>
    Public Class HostCapabilityRunRecorder
        Private ReadOnly _store As IRunTraceStore
        Private ReadOnly _startedAt As DateTime
        Private _completed As Boolean

        Public ReadOnly Property RunId As String
        Public ReadOnly Property CapabilityId As String
        Public ReadOnly Property AppType As String
        Public ReadOnly Property RiskLevel As String

        Public Sub New(appType As String,
                       capabilityId As String,
                       userText As String,
                       riskLevel As String,
                       Optional sessionId As String = "",
                       Optional store As IRunTraceStore = Nothing)
            Me.RunId = Guid.NewGuid().ToString()
            Me.AppType = If(appType, "")
            Me.CapabilityId = If(capabilityId, "host.capability")
            Me.RiskLevel = NormalizeRisk(riskLevel)
            _store = If(store, New SqliteRunTraceStore())
            _startedAt = DateTime.Now

            _store.StartRun(RunId,
                            New UserTurn With {
                                .SessionId = If(sessionId, ""),
                                .AppType = Me.AppType,
                                .Text = If(userText, ""),
                                .Mode = "capability_fast_path"
                            },
                            _startedAt)
        End Sub

        Public Function EvaluateSafety(Optional params As JObject = Nothing) As Execution.SafetyDecision
            Dim descriptor As New ToolDescriptor With {
                .Id = CapabilityId,
                .Name = CapabilityId,
                .AppType = AppType,
                .RiskLevel = RiskLevel
            }
            Dim gate As New Execution.SafetyGate()
            Return gate.Evaluate(descriptor, If(params, New JObject()))
        End Function

        Public Sub Complete(result As ToolResult)
            If _completed Then Return
            _completed = True

            Dim finishedAt = DateTime.Now
            Dim success = result IsNot Nothing AndAlso result.Success
            Dim status = If(success, "succeeded", "failed")
            Dim message = If(result?.UserMessage, result?.Message)
            Dim errorCode = If(result?.ErrorCode, "")
            _store.AppendStep(RunId,
                              0,
                              CapabilityId,
                              status,
                              If(message, ""),
                              errorCode,
                              If(result?.Observation, result?.Data),
                              _startedAt,
                              finishedAt)
            _store.CompleteRun(RunId, status, If(message, ""), errorCode, finishedAt)
        End Sub

        Public Sub CompleteSafetyDecision(decision As Execution.SafetyDecision)
            If decision Is Nothing Then Return
            Dim observation As New JObject From {
                {"kind", "safety"},
                {"summary", If(decision.UserMessage, decision.Reason)},
                {"changed", False},
                {"targetRefs", New JArray($"{AppType}:ActiveDocument")},
                {"warnings", New JArray(If(decision.Reason, ""))},
                {"safetyAction", decision.Action.ToString()},
                {"riskLevel", decision.RiskLevel}
            }
            Complete(ToolResult.Failed(CapabilityId,
                                       If(decision.Reason, "Safety blocked"),
                                       errorCode:=If(decision.ErrorCode, ExceptionClassifier.CodeSafetyBlocked),
                                       userMessage:=If(decision.UserMessage, decision.Reason),
                                       recoverable:=False,
                                       observation:=observation))
        End Sub

        Private Shared Function NormalizeRisk(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "high", "risky"
                    Return "risky"
                Case "medium"
                    Return "medium"
                Case Else
                    Return "safe"
            End Select
        End Function
    End Class

End Namespace
