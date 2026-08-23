Imports System.Collections.Generic

Namespace Agent.Goals

    ''' <summary>
    ''' Untrusted structured interpretation returned by the existing intent-model call.
    ''' RawUserRequest is never accepted from the model; the Adapter supplies the captured text.
    ''' </summary>
    Public Class GoalInterpretationPayload
        Public Property Candidate As CandidateGoalContract
        Public Property UnresolvedClauses As New List(Of String)()
        Public Property Assumptions As New List(Of String)()
        Public Property Diagnostics As New List(Of String)()
        Public Property RequiresClarification As Boolean
    End Class

    ''' <summary>
    ''' Seam between task intake and the Candidate -> Validate -> Freeze pipeline.
    ''' </summary>
    Friend Interface IGoalInterpretationAdapter
        Function Interpret(rawUserRequest As String) As GoalCompilationResult
    End Interface

    ''' <summary>
    ''' Normal production Adapter. It compiles the structured interpretation emitted by the
    ''' intent-model call while replacing any model-supplied raw text with the captured request.
    ''' </summary>
    Friend NotInheritable Class ModelGoalInterpretationAdapter
        Implements IGoalInterpretationAdapter

        Private ReadOnly _payload As GoalInterpretationPayload

        Public Sub New(payload As GoalInterpretationPayload)
            _payload = payload
        End Sub

        Public Function Interpret(rawUserRequest As String) As GoalCompilationResult Implements IGoalInterpretationAdapter.Interpret
            If _payload?.Candidate Is Nothing Then
                Throw New InvalidOperationException("Model goal interpretation has no candidate.")
            End If

            Dim diagnostics As New List(Of String)(If(_payload.Diagnostics, New List(Of String)())) From {
                "Compiled from the structured goal interpretation returned by the existing intent-model call."
            }
            Return GoalCompiler.CompileAuthoritative(
                _payload.Candidate,
                rawUserRequest,
                _payload.UnresolvedClauses,
                _payload.Assumptions,
                diagnostics,
                _payload.RequiresClarification)
        End Function
    End Class

    ''' <summary>
    ''' Failure Adapter. It preserves the complete request as opaque semantic meaning and never
    ''' imports legacy TaskSpec inference into the authoritative goal.
    ''' </summary>
    Friend NotInheritable Class RawPreservingGoalInterpretationAdapter
        Implements IGoalInterpretationAdapter

        Private ReadOnly _diagnostic As String

        Public Sub New(Optional diagnostic As String = Nothing)
            _diagnostic = If(diagnostic, "No structured model interpretation was available; preserved the exact request as opaque semantics.")
        End Sub

        Public Function Interpret(rawUserRequest As String) As GoalCompilationResult Implements IGoalInterpretationAdapter.Interpret
            Dim compilation = GoalCompiler.Compile(rawUserRequest)
            Dim diagnostics As New List(Of String)(compilation.Diagnostics) From {_diagnostic}
            Return New GoalCompilationResult(
                compilation.Candidate,
                compilation.CoverageMap,
                compilation.UnresolvedClauses,
                compilation.Assumptions,
                diagnostics,
                compilation.RequiresClarification)
        End Function
    End Class

End Namespace
