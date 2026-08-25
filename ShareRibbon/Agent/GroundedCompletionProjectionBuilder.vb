Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Turns the provisional Goal-to-requirement mapping into a completion candidate using
    ''' only satisfied runtime evidence. Exact target/value grounding remains deterministic.
    ''' </summary>
    Friend NotInheritable Class GroundedCompletionProjectionBuilder
        Private Sub New()
        End Sub

        Friend Shared Function Build(
            session As AgentSession,
            ByRef evidenceClaims As List(Of String),
            ByRef contract As OutcomeContract,
            groundProjection As Func(Of AgentSession, OutcomeContract, IList(Of String), String)) As String

            contract = Nothing
            Dim source = session?.Plan?.OutcomeContract
            If source Is Nothing OrElse source.Requirements Is Nothing OrElse source.Requirements.Count = 0 Then
                Return "No provisional Goal-bound verification mapping is available."
            End If

            Try
                contract = JObject.FromObject(source).ToObject(Of OutcomeContract)()
            Catch ex As Exception
                Return $"The provisional verification mapping could not be cloned: {ex.Message}"
            End Try

            If evidenceClaims Is Nothing Then evidenceClaims = New List(Of String)()
            For Each record In If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item IsNot Nothing).
                SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                Where(Function(item) item IsNot Nothing AndAlso
                    item.Satisfied AndAlso
                    Not String.IsNullOrWhiteSpace(item.EvidenceId) AndAlso
                    Not String.Equals(item.EffectType, "unclassified_mutation", StringComparison.OrdinalIgnoreCase))
                If Not evidenceClaims.Contains(record.EvidenceId, StringComparer.OrdinalIgnoreCase) Then
                    evidenceClaims.Add(record.EvidenceId)
                End If
            Next

            If groundProjection Is Nothing Then
                contract = Nothing
                Return "The completion projection grounder is unavailable."
            End If
            Dim groundingError = groundProjection(session, contract, evidenceClaims)
            If Not String.IsNullOrWhiteSpace(groundingError) Then
                contract = Nothing
                Return "The provisional Goal mapping could not be grounded in current host evidence: " & groundingError
            End If
            Return ""
        End Function
    End Class

End Namespace
