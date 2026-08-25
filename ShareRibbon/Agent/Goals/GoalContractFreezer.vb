Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.Goals

    ''' <summary>
    ''' The only construction path for the authoritative immutable GoalContract.
    ''' </summary>
    Public NotInheritable Class GoalContractFreezer
        Private Sub New()
        End Sub

        Public Shared Function Freeze(compilation As GoalCompilationResult,
                                      validation As GoalCoverageValidationResult) As GoalContract
            If compilation Is Nothing Then Throw New ArgumentNullException(NameOf(compilation))
            If validation Is Nothing Then Throw New ArgumentNullException(NameOf(validation))
            If Not validation.Succeeded Then
                Throw New InvalidOperationException("Goal coverage validation failed: " & String.Join("; ", validation.Errors))
            End If

            Dim candidateFingerprintBefore = GoalCoverageValidator.ComputeCandidateFingerprint(compilation.Candidate)
            If Not String.Equals(validation.CandidateFingerprint, candidateFingerprintBefore, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("Goal candidate does not match the candidate that passed coverage validation.")
            End If

            Dim candidate = GoalCompiler.CloneCandidateSnapshot(compilation.Candidate)
            Dim candidateFingerprintAfter = GoalCoverageValidator.ComputeCandidateFingerprint(compilation.Candidate)
            Dim snapshotFingerprint = GoalCoverageValidator.ComputeCandidateFingerprint(candidate)
            If Not String.Equals(candidateFingerprintBefore, candidateFingerprintAfter, StringComparison.Ordinal) OrElse
               Not String.Equals(candidateFingerprintBefore, snapshotFingerprint, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("Goal candidate changed while the freeze snapshot was being captured.")
            End If

            Dim snapshotCompilation As New GoalCompilationResult(
                candidate,
                compilation.CoverageMap,
                compilation.UnresolvedClauses,
                compilation.Assumptions,
                compilation.Diagnostics,
                compilation.RequiresClarification)

            ' Revalidate at the freeze seam so a stale validation result cannot authorize a
            ' candidate that was mutated after validation.
            Dim currentValidation = GoalCoverageValidator.Validate(snapshotCompilation)
            If Not currentValidation.Succeeded Then
                Throw New InvalidOperationException("Goal candidate changed after validation: " & String.Join("; ", currentValidation.Errors))
            End If

            Dim sourceClauses = candidate.SourceClauses.
                Where(Function(item) item IsNot Nothing).
                Select(Function(item) New GoalSourceClause(
                    NormalizeId(item.Id),
                    If(item.Text, ""),
                    item.IsExplicit,
                    item.SourceStart)).
                ToList()
            Dim criteria = candidate.Criteria.
                Where(Function(item) item IsNot Nothing AndAlso item.Required).
                Select(Function(item) New GoalCriterion(
                    NormalizeId(item.Id),
                    If(item.Statement, ""),
                    NormalizeId(item.Kind),
                    NormalizeIdList(item.SourceClauseIds),
                    item.Required,
                    NormalizeId(item.VerificationCapability),
                    NormalizeId(item.CapabilityId))).
                ToList()
            Dim constraints = candidate.Constraints.
                Where(Function(item) item IsNot Nothing AndAlso item.Required).
                Select(Function(item) New GoalConstraint(
                    NormalizeId(item.Id),
                    If(item.Statement, ""),
                    NormalizeId(item.Kind),
                    NormalizeIdList(item.SourceClauseIds),
                    item.Required)).
                ToList()
            Dim capabilities = NormalizeIdList(candidate.RequiredCapabilities)
            Dim contractHash = GoalHashing.ComputeContractHash(
                candidate.RawUserRequest,
                sourceClauses,
                criteria,
                constraints,
                capabilities)
            Dim semanticHash = GoalHashing.ComputeSemanticHash(
                candidate.RawUserRequest,
                sourceClauses,
                criteria,
                constraints,
                capabilities)
            Dim goalId = "goal-" & semanticHash.Substring(0, 16).ToLowerInvariant()

            Return New GoalContract(
                goalId,
                candidate.RawUserRequest,
                sourceClauses,
                criteria,
                constraints,
                capabilities,
                contractHash,
                semanticHash)
        End Function

        Private Shared Function NormalizeId(value As String) As String
            Return If(value, "").Trim()
        End Function

        Private Shared Function NormalizeIdList(values As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each value In If(values, Enumerable.Empty(Of String)())
                Dim normalized = NormalizeId(value)
                If normalized.Length > 0 AndAlso seen.Add(normalized) Then result.Add(normalized)
            Next
            Return result
        End Function

    End Class

End Namespace
