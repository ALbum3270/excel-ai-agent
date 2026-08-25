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

            ' Exact-text fallback and required semantic constraints are not reasons to hide
            ' every write tool before the adaptive loop starts.  The raw request remains the
            ' authoritative Goal, SafetyGate still controls each mutation, required
            ' capabilities are checked from actual lineage, and GoalOutcomeProjection binds
            ' every required semantic obligation to verified host evidence before completion.

            Return ""
        End Function

    End Class

End Namespace
