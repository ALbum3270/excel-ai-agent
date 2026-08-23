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
            Return CompileCore(
                CloneCandidate(candidate),
                unresolvedClauses,
                assumptions,
                diagnostics,
                requiresClarification)
        End Function

        Friend Shared Function CompileAuthoritative(candidate As CandidateGoalContract,
                                                     rawUserRequest As String,
                                                     unresolvedClauses As IEnumerable(Of String),
                                                     assumptions As IEnumerable(Of String),
                                                     diagnostics As IEnumerable(Of String),
                                                     requiresClarification As Boolean) As GoalCompilationResult
            Dim working = CloneCandidate(candidate)
            If working IsNot Nothing Then working.RawUserRequest = If(rawUserRequest, "")
            Return CompileCore(
                working,
                unresolvedClauses,
                assumptions,
                diagnostics,
                requiresClarification)
        End Function

        Private Shared Function CompileCore(candidate As CandidateGoalContract,
                                            unresolvedClauses As IEnumerable(Of String),
                                            assumptions As IEnumerable(Of String),
                                            diagnostics As IEnumerable(Of String),
                                            requiresClarification As Boolean) As GoalCompilationResult
            NormalizeInterpretationMetadata(candidate)
            ResolveSourceOrigins(candidate)
            EnsureRawRequestPreserved(candidate)
            Dim demotedHints = DemoteUnverifiableInterpretationHints(candidate)
            ReconcileCapabilityEvidence(candidate)
            Dim coverage = BuildCoverageMap(candidate)
            Dim compilationDiagnostics As New List(Of String)(If(diagnostics, Enumerable.Empty(Of String)()))
            If demotedHints > 0 Then
                compilationDiagnostics.Add($"Demoted {demotedHints} model-derived paraphrase(s) to non-authoritative planning hints.")
            End If
            Return New GoalCompilationResult(
                candidate,
                coverage,
                unresolvedClauses,
                assumptions,
                compilationDiagnostics,
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

            Return CompileCore(
                candidate,
                Nothing,
                Nothing,
                diagnostics:=New List(Of String) From {
                    "Compiled by the raw-preserving intake adapter; legacy TaskSpec projections were excluded from authoritative goal semantics."
                },
                requiresClarification:=False)
        End Function

        Private Shared Sub EnsureRawRequestPreserved(candidate As CandidateGoalContract)
            If candidate Is Nothing OrElse String.IsNullOrWhiteSpace(candidate.RawUserRequest) Then Return
            If candidate.SourceClauses Is Nothing Then candidate.SourceClauses = New List(Of CandidateGoalSourceClause)()
            If candidate.Criteria Is Nothing Then candidate.Criteria = New List(Of CandidateGoalCriterion)()

            ' A structured interpretation must stand on its complete, verified source spans;
            ' silently adding the whole raw request would mask omitted or polarity-stripped
            ' clauses.  Only the explicit raw-preserving adapter starts without source spans.
            If Not candidate.SourceClauses.Any(
                Function(item) item IsNot Nothing AndAlso item.IsExplicit AndAlso
                    Not String.IsNullOrWhiteSpace(item.Text)) Then
                candidate.SourceClauses.Add(New CandidateGoalSourceClause With {
                    .Id = NextAvailableId("clause-raw-user-request", candidate.SourceClauses.Select(Function(item) If(item?.Id, ""))),
                    .Text = candidate.RawUserRequest,
                    .IsExplicit = True,
                    .SourceStart = 0
                })
            End If

            Dim semanticIndex As Integer = 0
            For Each sourceClause In candidate.SourceClauses.Where(
                Function(item) item IsNot Nothing AndAlso item.IsExplicit AndAlso Not String.IsNullOrWhiteSpace(item.Text)).ToList()
                semanticIndex += 1
                Dim preserved = candidate.Criteria.Any(
                    Function(item) item IsNot Nothing AndAlso
                        item.Required AndAlso
                        String.Equals(If(item.Statement, ""), sourceClause.Text, StringComparison.Ordinal) AndAlso
                        item.SourceClauseIds IsNot Nothing AndAlso
                        item.SourceClauseIds.Any(Function(id) String.Equals(If(id, "").Trim(), If(sourceClause.Id, "").Trim(), StringComparison.OrdinalIgnoreCase)))
                If preserved Then Continue For

                Dim isRawClause = String.Equals(sourceClause.Text, candidate.RawUserRequest, StringComparison.Ordinal) AndAlso
                    sourceClause.SourceStart = 0
                Dim preferredId = If(isRawClause,
                                     "criterion-raw-user-request",
                                     $"criterion-semantic-{semanticIndex}")
                candidate.Criteria.Add(New CandidateGoalCriterion With {
                    .Id = NextAvailableId(preferredId, candidate.Criteria.Select(Function(item) If(item?.Id, ""))),
                    .Statement = sourceClause.Text,
                    .Kind = "semantic",
                    .SourceClauseIds = New List(Of String) From {sourceClause.Id},
                    .Required = True,
                    .VerificationCapability = "semantic"
                })
            Next
        End Sub

        ''' <summary>
        ''' Model metadata can suggest planning categories, but it cannot invent executable
        ''' verifier/capability identities or arbitrary criterion kinds.  Normalize that
        ''' metadata before it participates in hashes, prompts or freezing.
        ''' </summary>
        Private Shared Sub NormalizeInterpretationMetadata(candidate As CandidateGoalContract)
            If candidate Is Nothing Then Return
            Dim allowedKinds As New HashSet(Of String)(
                {"semantic", "compute", "state", "visual", "capability"},
                StringComparer.OrdinalIgnoreCase)
            For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Where(Function(item) item IsNot Nothing)
                Dim normalizedKind = If(criterion.Kind, "").Trim().ToLowerInvariant()
                If Not allowedKinds.Contains(normalizedKind) Then normalizedKind = "semantic"
                criterion.Kind = normalizedKind
                If String.Equals(normalizedKind, "capability", StringComparison.OrdinalIgnoreCase) Then
                    ' ReconcileCapabilityEvidence replaces all model-authored capability edges
                    ' with deterministic evidence derived from the exact request.
                    criterion.Required = False
                    criterion.VerificationCapability = ""
                    criterion.CapabilityId = ""
                Else
                    criterion.VerificationCapability = normalizedKind
                    criterion.CapabilityId = ""
                End If
            Next
            For Each constraint In If(candidate.Constraints, New List(Of CandidateGoalConstraint)()).Where(Function(item) item IsNot Nothing)
                constraint.Kind = "semantic"
            Next
        End Sub

        Friend Shared Function CloneCandidateSnapshot(candidate As CandidateGoalContract) As CandidateGoalContract
            If candidate Is Nothing Then Return Nothing

            Dim clone As New CandidateGoalContract With {
                .RawUserRequest = If(candidate.RawUserRequest, ""),
                .SourceClauses = New List(Of CandidateGoalSourceClause)(),
                .Criteria = New List(Of CandidateGoalCriterion)(),
                .Constraints = New List(Of CandidateGoalConstraint)(),
                .RequiredCapabilities = New List(Of String)(If(candidate.RequiredCapabilities, Enumerable.Empty(Of String)()))
            }
            For Each sourceClause In If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)())
                If sourceClause Is Nothing Then
                    clone.SourceClauses.Add(Nothing)
                Else
                    clone.SourceClauses.Add(New CandidateGoalSourceClause With {
                        .Id = If(sourceClause.Id, ""),
                        .Text = If(sourceClause.Text, ""),
                        .IsExplicit = sourceClause.IsExplicit,
                        .SourceStart = sourceClause.SourceStart
                    })
                End If
            Next
            For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)())
                If criterion Is Nothing Then
                    clone.Criteria.Add(Nothing)
                Else
                    clone.Criteria.Add(New CandidateGoalCriterion With {
                        .Id = If(criterion.Id, ""),
                        .Statement = If(criterion.Statement, ""),
                        .Kind = If(criterion.Kind, ""),
                        .SourceClauseIds = New List(Of String)(If(criterion.SourceClauseIds, Enumerable.Empty(Of String)())),
                        .Required = criterion.Required,
                        .VerificationCapability = If(criterion.VerificationCapability, ""),
                        .CapabilityId = If(criterion.CapabilityId, "")
                    })
                End If
            Next
            For Each constraint In If(candidate.Constraints, New List(Of CandidateGoalConstraint)())
                If constraint Is Nothing Then
                    clone.Constraints.Add(Nothing)
                Else
                    clone.Constraints.Add(New CandidateGoalConstraint With {
                        .Id = If(constraint.Id, ""),
                        .Statement = If(constraint.Statement, ""),
                        .Kind = If(constraint.Kind, ""),
                        .SourceClauseIds = New List(Of String)(If(constraint.SourceClauseIds, Enumerable.Empty(Of String)())),
                        .Required = constraint.Required
                    })
                End If
            Next
            Return clone
        End Function

        Private Shared Function CloneCandidate(candidate As CandidateGoalContract) As CandidateGoalContract
            Return CloneCandidateSnapshot(candidate)
        End Function

        ''' <summary>
        ''' Resolves explicit clauses to a verified occurrence in the captured raw request.
        ''' Unique occurrences need no model offset; repeated identical text must carry a valid
        ''' offset or validation fails closed instead of guessing which occurrence was meant.
        ''' </summary>
        Private Shared Sub ResolveSourceOrigins(candidate As CandidateGoalContract)
            If candidate?.SourceClauses Is Nothing Then Return
            Dim raw = If(candidate.RawUserRequest, "")
            For Each clause In candidate.SourceClauses.Where(Function(item) item IsNot Nothing AndAlso item.IsExplicit)
                Dim text = If(clause.Text, "")
                If text.Length = 0 OrElse raw.Length = 0 Then
                    clause.SourceStart = -1
                    Continue For
                End If
                If String.Equals(text, raw, StringComparison.Ordinal) Then
                    clause.SourceStart = 0
                    Continue For
                End If

                Dim occurrences As New List(Of Integer)()
                Dim searchStart As Integer = 0
                While searchStart <= raw.Length - text.Length
                    Dim found = raw.IndexOf(text, searchStart, StringComparison.Ordinal)
                    If found < 0 Then Exit While
                    occurrences.Add(found)
                    ' Advance one UTF-16 code unit so overlapping occurrences remain visible
                    ' (for example "aa" occurs twice in "aaa").
                    searchStart = found + 1
                End While
                If occurrences.Count = 1 Then
                    clause.SourceStart = occurrences(0)
                ElseIf occurrences.Count = 0 OrElse Not occurrences.Contains(clause.SourceStart) Then
                    clause.SourceStart = -1
                End If
            Next
        End Sub

        ''' <summary>
        ''' Rebuilds method constraints only from deterministic evidence in the exact request.
        ''' Model capability declarations are routing hints, never their own authorization.
        ''' </summary>
        Private Shared Sub ReconcileCapabilityEvidence(candidate As CandidateGoalContract)
            If candidate Is Nothing Then Return
            If candidate.RequiredCapabilities Is Nothing Then candidate.RequiredCapabilities = New List(Of String)()
            If candidate.SourceClauses Is Nothing Then candidate.SourceClauses = New List(Of CandidateGoalSourceClause)()
            If candidate.Criteria Is Nothing Then candidate.Criteria = New List(Of CandidateGoalCriterion)()

            Dim evidenceCapabilities = GoalCapabilityEvidenceResolver.Resolve(candidate.RawUserRequest)
            candidate.RequiredCapabilities.Clear()
            candidate.RequiredCapabilities.AddRange(evidenceCapabilities)
            For Each criterion In candidate.Criteria.Where(
                Function(item) item IsNot Nothing AndAlso
                    String.Equals(item.Kind, "capability", StringComparison.OrdinalIgnoreCase))
                criterion.Required = False
            Next

            For Each capability In evidenceCapabilities
                Dim sourceClause = candidate.SourceClauses.
                    Where(Function(item) item IsNot Nothing AndAlso item.IsExplicit AndAlso
                        GoalCapabilityEvidenceResolver.ClauseSupports(item.Text, capability)).
                    OrderBy(Function(item) If(item.Text, "").Length).
                    ThenBy(Function(item) item.SourceStart).
                    FirstOrDefault()
                If sourceClause Is Nothing Then
                    sourceClause = candidate.SourceClauses.FirstOrDefault(
                        Function(item) item IsNot Nothing AndAlso item.IsExplicit AndAlso
                            String.Equals(If(item.Text, ""), If(candidate.RawUserRequest, ""), StringComparison.Ordinal))
                End If
                If sourceClause Is Nothing Then Continue For

                candidate.Criteria.Add(New CandidateGoalCriterion With {
                    .Id = NextAvailableId("criterion-capability-" & capability.ToLowerInvariant(), candidate.Criteria.Select(Function(item) If(item?.Id, ""))),
                    .Statement = sourceClause.Text,
                    .Kind = "capability",
                    .SourceClauseIds = New List(Of String) From {sourceClause.Id},
                    .Required = True,
                    .VerificationCapability = capability,
                    .CapabilityId = capability
                })
            Next
        End Sub

        ''' <summary>
        ''' Model paraphrases can help planning but cannot become user authority.  Only a
        ''' statement copied verbatim from one of its explicit referenced clauses remains
        ''' Required; the validator independently enforces the same invariant at public seams.
        ''' </summary>
        Private Shared Function DemoteUnverifiableInterpretationHints(candidate As CandidateGoalContract) As Integer
            If candidate Is Nothing Then Return 0
            Dim clauseById = If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)()).
                Where(Function(item) item IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(item.Id)).
                GroupBy(Function(item) item.Id.Trim(), StringComparer.OrdinalIgnoreCase).
                ToDictionary(Function(group) group.Key, Function(group) group.First(), StringComparer.OrdinalIgnoreCase)
            Dim demoted As Integer = 0

            For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Where(
                Function(item) item IsNot Nothing AndAlso item.Required)
                If HasVerbatimSourceAuthority(criterion.Statement, criterion.SourceClauseIds, clauseById) Then Continue For
                criterion.Required = False
                demoted += 1
            Next
            For Each constraint In If(candidate.Constraints, New List(Of CandidateGoalConstraint)()).Where(
                Function(item) item IsNot Nothing AndAlso item.Required)
                If HasVerbatimSourceAuthority(constraint.Statement, constraint.SourceClauseIds, clauseById) Then Continue For
                constraint.Required = False
                demoted += 1
            Next
            Return demoted
        End Function

        Private Shared Function HasVerbatimSourceAuthority(statement As String,
                                                            sourceClauseIds As IEnumerable(Of String),
                                                            clauseById As Dictionary(Of String, CandidateGoalSourceClause)) As Boolean
            For Each sourceId In If(sourceClauseIds, Enumerable.Empty(Of String)())
                Dim clause As CandidateGoalSourceClause = Nothing
                If clauseById.TryGetValue(If(sourceId, "").Trim(), clause) AndAlso
                   clause IsNot Nothing AndAlso clause.IsExplicit AndAlso
                   String.Equals(If(statement, ""), If(clause.Text, ""), StringComparison.Ordinal) Then
                    Return True
                End If
            Next
            Return False
        End Function

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
