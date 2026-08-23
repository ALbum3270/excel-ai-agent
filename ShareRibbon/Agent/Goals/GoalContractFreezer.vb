Imports System.Collections.Generic
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text

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

            Dim candidateFingerprint = GoalCoverageValidator.ComputeCandidateFingerprint(compilation.Candidate)
            If Not String.Equals(validation.CandidateFingerprint, candidateFingerprint, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("Goal candidate does not match the candidate that passed coverage validation.")
            End If

            ' Revalidate at the freeze seam so a stale validation result cannot authorize a
            ' candidate that was mutated after validation.
            Dim currentValidation = GoalCoverageValidator.Validate(compilation)
            If Not currentValidation.Succeeded Then
                Throw New InvalidOperationException("Goal candidate changed after validation: " & String.Join("; ", currentValidation.Errors))
            End If

            Dim candidate = compilation.Candidate
            Dim sourceClauses = candidate.SourceClauses.
                Where(Function(item) item IsNot Nothing).
                Select(Function(item) New GoalSourceClause(
                    NormalizeId(item.Id),
                    If(item.Text, ""),
                    item.IsExplicit,
                    NormalizeId(item.RequiredCapability))).
                ToList()
            Dim criteria = candidate.Criteria.
                Where(Function(item) item IsNot Nothing).
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
                Where(Function(item) item IsNot Nothing).
                Select(Function(item) New GoalConstraint(
                    NormalizeId(item.Id),
                    If(item.Statement, ""),
                    NormalizeId(item.Kind),
                    NormalizeIdList(item.SourceClauseIds),
                    item.Required)).
                ToList()
            Dim capabilities = NormalizeIdList(candidate.RequiredCapabilities)
            Dim contractHash = ComputeContractHash(
                candidate.RawUserRequest,
                sourceClauses,
                criteria,
                constraints,
                capabilities)
            Dim goalId = "goal-" & contractHash.Substring(0, 16).ToLowerInvariant()

            Return New GoalContract(
                goalId,
                candidate.RawUserRequest,
                sourceClauses,
                criteria,
                constraints,
                capabilities,
                contractHash)
        End Function

        Private Shared Function ComputeContractHash(rawUserRequest As String,
                                                    sourceClauses As IEnumerable(Of GoalSourceClause),
                                                    criteria As IEnumerable(Of GoalCriterion),
                                                    constraints As IEnumerable(Of GoalConstraint),
                                                    capabilities As IEnumerable(Of String)) As String
            Dim canonical As New StringBuilder()
            Append(canonical, "raw", NormalizeLineEndings(rawUserRequest))
            For Each clause In sourceClauses.OrderBy(Function(item) item.Id, StringComparer.Ordinal)
                Append(canonical, "clause.id", clause.Id)
                Append(canonical, "clause.text", NormalizeLineEndings(clause.Text))
                Append(canonical, "clause.explicit", clause.IsExplicit.ToString())
                Append(canonical, "clause.capability", clause.RequiredCapability)
            Next
            For Each criterion In criteria.OrderBy(Function(item) item.Id, StringComparer.Ordinal)
                Append(canonical, "criterion.id", criterion.Id)
                Append(canonical, "criterion.statement", NormalizeLineEndings(criterion.Statement))
                Append(canonical, "criterion.kind", criterion.Kind)
                Append(canonical, "criterion.required", criterion.Required.ToString())
                Append(canonical, "criterion.verifier", criterion.VerificationCapability)
                Append(canonical, "criterion.capability", criterion.CapabilityId)
                For Each sourceId In criterion.SourceClauseIds.OrderBy(Function(value) value, StringComparer.Ordinal)
                    Append(canonical, "criterion.source", sourceId)
                Next
            Next
            For Each constraint In constraints.OrderBy(Function(item) item.Id, StringComparer.Ordinal)
                Append(canonical, "constraint.id", constraint.Id)
                Append(canonical, "constraint.statement", NormalizeLineEndings(constraint.Statement))
                Append(canonical, "constraint.kind", constraint.Kind)
                Append(canonical, "constraint.required", constraint.Required.ToString())
                For Each sourceId In constraint.SourceClauseIds.OrderBy(Function(value) value, StringComparer.Ordinal)
                    Append(canonical, "constraint.source", sourceId)
                Next
            Next
            For Each capability In capabilities.OrderBy(Function(value) value, StringComparer.Ordinal)
                Append(canonical, "requiredCapability", capability)
            Next

            Using sha = SHA256.Create()
                Dim bytes = Encoding.UTF8.GetBytes(canonical.ToString())
                Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "")
            End Using
        End Function

        Private Shared Sub Append(builder As StringBuilder, name As String, value As String)
            Dim safeValue = If(value, "")
            builder.Append(name.Length).Append(":"c).Append(name)
            builder.Append(safeValue.Length).Append(":"c).Append(safeValue).Append("|"c)
        End Sub

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

        Private Shared Function NormalizeLineEndings(value As String) As String
            Return If(value, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        End Function
    End Class

End Namespace
