Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.Goals

    ''' <summary>
    ''' Compiles interpretation output into an untrusted candidate plus traceability data.
    ''' It never returns a GoalContract; validation and freezing are separate phases.
    ''' </summary>
    Public NotInheritable Class GoalCompiler
        Private Sub New()
        End Sub

        Public Shared Function Compile(candidate As CandidateGoalContract,
                                       Optional unresolvedClauses As IEnumerable(Of String) = Nothing,
                                       Optional assumptions As IEnumerable(Of String) = Nothing,
                                       Optional diagnostics As IEnumerable(Of String) = Nothing,
                                       Optional requiresClarification As Boolean = False) As GoalCompilationResult
            EnsureRawRequestPreserved(candidate)
            Dim coverage = BuildCoverageMap(candidate)
            Return New GoalCompilationResult(
                candidate,
                coverage,
                unresolvedClauses,
                assumptions,
                diagnostics,
                requiresClarification)
        End Function

        ''' <summary>
        ''' Transitional adapter for the current intake. It preserves only the exact user
        ''' request. Legacy TaskSpec projections are deliberately excluded because they may
        ''' contain model summaries, routing hints, or host policy rather than user semantics.
        ''' </summary>
        Public Shared Function Compile(rawUserRequest As String) As GoalCompilationResult
            Dim raw = If(rawUserRequest, "")
            Dim candidate As New CandidateGoalContract With {
                .RawUserRequest = raw
            }

            Return Compile(
                candidate,
                diagnostics:=New List(Of String) From {
                    "Compiled by the raw-preserving intake adapter; legacy TaskSpec projections were excluded from authoritative goal semantics."
                })
        End Function

        Private Shared Sub EnsureRawRequestPreserved(candidate As CandidateGoalContract)
            If candidate Is Nothing OrElse String.IsNullOrWhiteSpace(candidate.RawUserRequest) Then Return
            If candidate.SourceClauses Is Nothing Then candidate.SourceClauses = New List(Of CandidateGoalSourceClause)()
            If candidate.Criteria Is Nothing Then candidate.Criteria = New List(Of CandidateGoalCriterion)()

            Dim rawClause = candidate.SourceClauses.FirstOrDefault(
                Function(item) item IsNot Nothing AndAlso
                    item.IsExplicit AndAlso
                    String.Equals(If(item.Text, ""), candidate.RawUserRequest, StringComparison.Ordinal))
            If rawClause Is Nothing Then
                rawClause = New CandidateGoalSourceClause With {
                    .Id = NextAvailableId("clause-raw-user-request", candidate.SourceClauses.Select(Function(item) If(item?.Id, ""))),
                    .Text = candidate.RawUserRequest,
                    .IsExplicit = True
                }
                candidate.SourceClauses.Add(rawClause)
            End If

            Dim preserved = candidate.Criteria.Any(
                Function(item) item IsNot Nothing AndAlso
                    item.Required AndAlso
                    String.Equals(item.Kind, "semantic", StringComparison.OrdinalIgnoreCase) AndAlso
                    String.Equals(If(item.Statement, ""), candidate.RawUserRequest, StringComparison.Ordinal) AndAlso
                    item.SourceClauseIds IsNot Nothing AndAlso
                    item.SourceClauseIds.Any(Function(id) String.Equals(If(id, "").Trim(), If(rawClause.Id, "").Trim(), StringComparison.OrdinalIgnoreCase)))
            If preserved Then Return

            candidate.Criteria.Add(New CandidateGoalCriterion With {
                .Id = NextAvailableId("criterion-raw-user-request", candidate.Criteria.Select(Function(item) If(item?.Id, ""))),
                .Statement = candidate.RawUserRequest,
                .Kind = "semantic",
                .SourceClauseIds = New List(Of String) From {rawClause.Id},
                .Required = True,
                .VerificationCapability = "semantic"
            })
        End Sub

        Private Shared Function NextAvailableId(baseId As String, existingIds As IEnumerable(Of String)) As String
            Dim existing As New HashSet(Of String)(
                If(existingIds, Enumerable.Empty(Of String)()).Select(Function(value) If(value, "").Trim()),
                StringComparer.OrdinalIgnoreCase)
            If Not existing.Contains(baseId) Then Return baseId

            Dim suffix As Integer = 2
            While existing.Contains($"{baseId}-{suffix}")
                suffix += 1
            End While
            Return $"{baseId}-{suffix}"
        End Function

        Private Shared Function BuildCoverageMap(candidate As CandidateGoalContract) As List(Of GoalCoverageMapEntry)
            Dim result As New List(Of GoalCoverageMapEntry)()
            If candidate?.SourceClauses Is Nothing Then Return result

            For Each sourceClause In candidate.SourceClauses.Where(Function(item) item IsNot Nothing)
                Dim criterionIds As New List(Of String)()
                For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Where(Function(item) item IsNot Nothing)
                    If criterion.SourceClauseIds IsNot Nothing AndAlso
                       criterion.SourceClauseIds.Any(Function(id) String.Equals(id, sourceClause.Id, StringComparison.OrdinalIgnoreCase)) Then
                        criterionIds.Add(If(criterion.Id, ""))
                    End If
                Next
                result.Add(New GoalCoverageMapEntry(sourceClause.Id, criterionIds))
            Next
            Return result
        End Function

    End Class

End Namespace
