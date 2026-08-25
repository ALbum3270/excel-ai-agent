Namespace Agent

    ''' <summary>
    ''' Immutable result of resolving whether the captured user turn authorizes Office mutation.
    ''' </summary>
    Friend NotInheritable Class TaskExecutionAuthority
        Public Sub New(allowsMutation As Boolean,
                       groundedByGoal As Boolean,
                       source As String)
            Me.AllowsMutation = allowsMutation
            Me.GroundedByGoal = groundedByGoal
            Me.Source = If(source, "none")
        End Sub

        Public ReadOnly Property AllowsMutation As Boolean
        Public ReadOnly Property GroundedByGoal As Boolean
        Public ReadOnly Property Source As String
    End Class

    ''' <summary>
    ''' Resolves task-level mutation authority independently from tool selection. There is one
    ''' authority source: an accepted execute mode bound to the validated Goal used by the loop.
    ''' </summary>
    Friend NotInheritable Class TaskExecutionAuthorityResolver
        Private Sub New()
        End Sub

        Public Shared Function Resolve(intent As IntentResult,
                                       compilation As Goals.GoalCompilationResult,
                                       hasStructuredAuthoritySource As Boolean) As TaskExecutionAuthority
            If Not hasStructuredAuthoritySource Then
                Return New TaskExecutionAuthority(False, False, "no_structured_goal")
            End If

            If IntentAcceptancePolicy.HasGroundedHostMutationAuthority(intent, compilation) Then
                Return New TaskExecutionAuthority(True, True, "validated_goal_execute")
            End If

            Return New TaskExecutionAuthority(False, False, "structured_goal_no_mutation")
        End Function
    End Class

End Namespace
