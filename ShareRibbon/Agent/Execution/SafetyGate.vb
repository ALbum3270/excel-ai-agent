Imports Newtonsoft.Json.Linq

Namespace Agent.Execution

    Public Enum SafetyAction
        Allow
        RequireApproval
        Deny
    End Enum

    Public Class SafetyDecision
        Public Property Action As SafetyAction = SafetyAction.Allow
        Public Property Reason As String = ""
        Public Property UserMessage As String = ""
        Public Property ErrorCode As String = ""
        Public Property RiskLevel As String = "safe"

        Public Shared Function Allow(Optional riskLevel As String = "safe") As SafetyDecision
            Return New SafetyDecision With {
                .Action = SafetyAction.Allow,
                .RiskLevel = If(riskLevel, "safe")
            }
        End Function

        Public Shared Function RequireApproval(reason As String,
                                               Optional userMessage As String = Nothing,
                                               Optional riskLevel As String = "risky") As SafetyDecision
            Return New SafetyDecision With {
                .Action = SafetyAction.RequireApproval,
                .Reason = If(reason, ""),
                .UserMessage = If(userMessage, reason),
                .ErrorCode = ExceptionClassifier.CodeSafetyNeedsApproval,
                .RiskLevel = If(riskLevel, "risky")
            }
        End Function

        Public Shared Function Deny(reason As String,
                                    errorCode As String,
                                    Optional userMessage As String = Nothing,
                                    Optional riskLevel As String = "risky") As SafetyDecision
            Return New SafetyDecision With {
                .Action = SafetyAction.Deny,
                .Reason = If(reason, ""),
                .UserMessage = If(userMessage, reason),
                .ErrorCode = If(errorCode, ExceptionClassifier.CodeSafetyBlocked),
                .RiskLevel = If(riskLevel, "risky")
            }
        End Function
    End Class

    Public Class SafetyGate
        Public Property VbaEnabled As Boolean = False
        Public Property RequireApprovalForRisky As Boolean = True
        Public Property RequireApprovalForDelete As Boolean = True

        Public Function Evaluate(tool As ToolDescriptor, params As JObject) As SafetyDecision
            If tool Is Nothing Then
                Return SafetyDecision.Deny("工具不存在，无法执行安全裁决",
                                           ExceptionClassifier.CodeNotFound,
                                           "工具不存在，无法执行")
            End If

            Dim toolId = If(tool.Id, "")
            Dim risk = If(String.IsNullOrWhiteSpace(tool.RiskLevel), "risky", tool.RiskLevel.Trim().ToLowerInvariant())

            If tool.IsVbaFallback OrElse String.Equals(toolId, "ExecuteVBA", StringComparison.OrdinalIgnoreCase) Then
                Return EvaluateVba(toolId, params, risk)
            End If

            If RequireApprovalForDelete AndAlso IsDestructiveTool(toolId, params) Then
                Return SafetyDecision.RequireApproval($"工具 {toolId} 可能删除或清空内容",
                                                      $"工具 {toolId} 需要确认后才能执行",
                                                      "risky")
            End If

            If RequireApprovalForRisky AndAlso String.Equals(risk, "risky", StringComparison.OrdinalIgnoreCase) Then
                Return SafetyDecision.RequireApproval($"高风险工具 {toolId} 需要用户确认",
                                                      $"高风险工具 {toolId} 需要确认后才能执行",
                                                      risk)
            End If

            Return SafetyDecision.Allow(risk)
        End Function

        Private Function EvaluateVba(toolId As String, params As JObject, risk As String) As SafetyDecision
            If Not VbaEnabled Then
                Return SafetyDecision.Deny("VBA 工具默认关闭",
                                           ExceptionClassifier.CodeVbaDisabled,
                                           "VBA 执行默认关闭，未进入宿主执行器",
                                           risk)
            End If

            Dim code = If(params?("code")?.ToString(), "")
            Dim subResult = SafetyChecker.Check(code)
            If subResult IsNot Nothing AndAlso Not subResult.IsSafe Then
                Return SafetyDecision.Deny(subResult.Reason,
                                           ExceptionClassifier.CodeSafetyBlocked,
                                           subResult.ToUserMessage(),
                                           risk)
            End If

            If subResult IsNot Nothing AndAlso subResult.NeedsConfirm Then
                Return SafetyDecision.RequireApproval(subResult.Reason,
                                                      subResult.ToUserMessage(),
                                                      risk)
            End If

            Return SafetyDecision.RequireApproval($"VBA 工具 {toolId} 需要用户确认",
                                                  $"VBA 工具 {toolId} 需要确认后才能执行",
                                                  risk)
        End Function

        Private Function IsDestructiveTool(toolId As String, params As JObject) As Boolean
            Dim id = If(toolId, "").ToLowerInvariant()
            If id.Contains("delete") OrElse id.Contains("clear") OrElse id.Contains("remove") Then Return True

            If String.Equals(id, "replacetext", StringComparison.OrdinalIgnoreCase) Then
                Dim rangeName = If(params?("range")?.ToString(), "all")
                Return IsWholeDocumentRange(rangeName)
            End If

            Dim scope = If(params?("scope")?.ToString(), "")
            Dim range = If(params?("range")?.ToString(), "")
            Return IsWholeDocumentRange(scope) OrElse IsWholeDocumentRange(range)
        End Function

        Private Function IsWholeDocumentRange(value As String) As Boolean
            Dim normalized = If(value, "").Trim().ToLowerInvariant()
            Return normalized = "all" OrElse
                   normalized = "document" OrElse
                   normalized = "workbook" OrElse
                   normalized = "presentation" OrElse
                   normalized = "全文"
        End Function
    End Class

End Namespace
