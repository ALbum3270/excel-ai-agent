Imports System.Collections.Generic
Imports System.Linq

''' <summary>
''' Closed interpretation of the model's routing mode.  Unspecified is reserved for the
''' deterministic compatibility result; Invalid represents an untrusted model value that must
''' never fail open into Office execution.
''' </summary>
Public Enum InteractionModeKind
    Unspecified
    Execute
    Answer
    Clarify
    Invalid
End Enum

''' <summary>
''' Single acceptance seam for model-produced intent routing facts.  A model candidate is either
''' committed as one coherent snapshot or rejected without changing the deterministic fallback.
''' Verification metadata is deliberately not treated as execution authority.
''' </summary>
Public NotInheritable Class IntentAcceptancePolicy
    Private Const MinimumAcceptedConfidence As Double = 0.3R

    Private Sub New()
    End Sub

    Public Shared Function ParseMode(value As String) As InteractionModeKind
        Dim normalized = If(value, "").Trim().ToLowerInvariant()
        Select Case normalized
            Case ""
                Return InteractionModeKind.Unspecified
            Case "execute"
                Return InteractionModeKind.Execute
            Case "answer"
                Return InteractionModeKind.Answer
            Case "clarify"
                Return InteractionModeKind.Clarify
            Case Else
                Return InteractionModeKind.Invalid
        End Select
    End Function

    Public Shared Function CanonicalMode(mode As InteractionModeKind) As String
        Select Case mode
            Case InteractionModeKind.Execute
                Return "execute"
            Case InteractionModeKind.Answer
                Return "answer"
            Case InteractionModeKind.Clarify
                Return "clarify"
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>
    ''' Atomically accepts the model-owned fields of a candidate.  Entity extraction and direct
    ''' command hints remain owned by the deterministic recognizer and are not overwritten here.
    ''' </summary>
    Friend Shared Function TryAcceptModelCandidate(target As IntentResult,
                                                   candidate As IntentResult) As Boolean
        If target Is Nothing OrElse candidate Is Nothing Then Return False
        If Double.IsNaN(candidate.Confidence) OrElse Double.IsInfinity(candidate.Confidence) OrElse
           candidate.Confidence <= MinimumAcceptedConfidence OrElse candidate.Confidence > 1.0R Then
            Return False
        End If

        Dim mode = ParseMode(candidate.ResponseMode)
        If mode = InteractionModeKind.Unspecified OrElse mode = InteractionModeKind.Invalid Then
            Return False
        End If

        Dim outputs = NormalizeOutputs(candidate.RequestedOutputs)
        If mode <> InteractionModeKind.Execute AndAlso outputs.Count > 0 Then
            Return False
        End If

        ' Nothing below can reject the candidate.  Commit only after the complete snapshot has
        ' passed validation so a low-confidence or internally inconsistent result cannot partly
        ' alter routing.
        target.OfficeIntent = candidate.OfficeIntent
        target.IntentType = candidate.IntentType
        target.Confidence = candidate.Confidence
        target.ResponseMode = CanonicalMode(mode)
        target.RequestedOutputs = outputs
        target.UserFriendlyDescription = If(candidate.UserFriendlyDescription, "")
        target.GoalInterpretation = candidate.GoalInterpretation
        Return True
    End Function

    Public Shared Function IsNonExecutionMode(mode As InteractionModeKind) As Boolean
        Return mode = InteractionModeKind.Answer OrElse
            mode = InteractionModeKind.Clarify OrElse
            mode = InteractionModeKind.Invalid
    End Function

    ''' <summary>
    ''' Only an actual tool requirement is evidence that an answer/clarification needs the Agent.
    ''' ExpectedOutputs and ExpectedSlideCount are completion assertions, not routing authority.
    ''' </summary>
    Public Shared Function HasExecutableToolRequirement(spec As Agent.AgentTaskSpec) As Boolean
        Return spec IsNot Nothing AndAlso
            spec.RequiredTools IsNot Nothing AndAlso
            spec.RequiredTools.Any(Function(toolId) Not String.IsNullOrWhiteSpace(toolId))
    End Function

    Private Shared Function NormalizeOutputs(values As IEnumerable(Of String)) As List(Of String)
        If values Is Nothing Then Return New List(Of String)()
        Return values.
            Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
            Select(Function(value) value.Trim().ToLowerInvariant()).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function
End Class
