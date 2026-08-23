Imports System.Linq

Namespace Agent.Goals

    ''' <summary>
    ''' Fail-closed admission checks that must run after the exact user goal is frozen but
    ''' before planning or any Office mutation.  Unsupported verification is not converted
    ''' into a weaker goal or delegated back to the planner.
    ''' </summary>
    Friend NotInheritable Class GoalExecutionAdmission
        Private Sub New()
        End Sub

        Friend Shared Function Validate(spec As AgentTaskSpec) As String
            If spec?.GoalContract Is Nothing Then
                Return "执行准入失败：缺少冻结后的 GoalContract。"
            End If

            Dim isReadOnly = String.Equals(spec.MutationPolicy,
                                           "read_only",
                                           StringComparison.OrdinalIgnoreCase)
            If Not isReadOnly AndAlso
               spec.GoalCompilation IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(spec.GoalInterpretationFallbackReason) Then
                Return "执行准入失败：本轮只能保留用户原文，结构化目标解释未通过验证；在 SemanticVerifier 可用前，禁止据此执行 Office 写入。请重试本轮请求或补充必要信息。"
            End If

            Dim unsupportedConstraints = spec.GoalContract.Constraints.
                Where(Function(item) item IsNot Nothing AndAlso item.Required).
                Select(Function(item) item.Id).
                ToList()
            If unsupportedConstraints.Count > 0 Then
                Return $"执行准入失败：required Goal constraints 尚无受信任的 ConstraintVerifier：{String.Join(", ", unsupportedConstraints)}。约束已完整保留，但不能在无法验收时宣称完成。"
            End If

            Return ""
        End Function

    End Class

End Namespace
