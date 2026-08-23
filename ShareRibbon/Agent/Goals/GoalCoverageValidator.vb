Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text

Namespace Agent.Goals

    Public NotInheritable Class GoalCoverageValidationResult
        Private ReadOnly _errors As ReadOnlyCollection(Of String)
        Private ReadOnly _coverageMap As ReadOnlyCollection(Of GoalCoverageMapEntry)
        Private ReadOnly _candidateFingerprint As String

        Friend Sub New(errors As IEnumerable(Of String),
                       coverageMap As IEnumerable(Of GoalCoverageMapEntry),
                       candidateFingerprint As String)
            _errors = New List(Of String)(If(errors, New List(Of String)())).AsReadOnly()
            _coverageMap = New List(Of GoalCoverageMapEntry)(If(coverageMap, New List(Of GoalCoverageMapEntry)())).AsReadOnly()
            _candidateFingerprint = If(candidateFingerprint, "")
        End Sub

        Public ReadOnly Property Succeeded As Boolean
            Get
                Return _errors.Count = 0
            End Get
        End Property

        Public ReadOnly Property Errors As IReadOnlyList(Of String)
            Get
                Return _errors
            End Get
        End Property

        Public ReadOnly Property CoverageMap As IReadOnlyList(Of GoalCoverageMapEntry)
            Get
                Return _coverageMap
            End Get
        End Property

        Friend ReadOnly Property CandidateFingerprint As String
            Get
                Return _candidateFingerprint
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Deterministic structural validator.  It does not trust the model's claim that all
    ''' clauses were covered; it recomputes coverage from criterion source references.
    ''' </summary>
    Public NotInheritable Class GoalCoverageValidator
        Private Sub New()
        End Sub

        Public Shared Function Validate(compilation As GoalCompilationResult) As GoalCoverageValidationResult
            Dim errors As New List(Of String)()
            Dim coverage As New List(Of GoalCoverageMapEntry)()
            Dim candidate = compilation?.Candidate
            If candidate Is Nothing Then
                errors.Add("Goal candidate is missing.")
                Return New GoalCoverageValidationResult(errors, coverage, "")
            End If
            If String.IsNullOrWhiteSpace(candidate.RawUserRequest) Then
                errors.Add("RawUserRequest is missing; no authoritative user-goal source is available.")
            End If
            If compilation.RequiresClarification Then
                errors.Add("Goal interpretation requires clarification and cannot be frozen.")
            End If

            Dim clauseIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each clause In If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)())
                If clause Is Nothing OrElse String.IsNullOrWhiteSpace(clause.Id) Then
                    errors.Add("Every source clause must have a stable id.")
                    Continue For
                End If
                Dim clauseId = NormalizeId(clause.Id)
                If Not clauseIds.Add(clauseId) Then errors.Add($"Duplicate source clause id: {clauseId}")
                If clause.IsExplicit AndAlso String.IsNullOrWhiteSpace(clause.Text) Then
                    errors.Add($"Explicit source clause {clauseId} has no text.")
                End If
            Next
            If clauseIds.Count = 0 Then errors.Add("At least one source clause is required.")

            Dim criterionIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)())
                If criterion Is Nothing OrElse String.IsNullOrWhiteSpace(criterion.Id) Then
                    errors.Add("Every goal criterion must have a stable id.")
                    Continue For
                End If
                Dim criterionId = NormalizeId(criterion.Id)
                If Not criterionIds.Add(criterionId) Then errors.Add($"Duplicate criterion id: {criterionId}")
                If String.IsNullOrWhiteSpace(criterion.Statement) Then
                    errors.Add($"Criterion {criterionId} has no statement.")
                End If
                For Each sourceId In If(criterion.SourceClauseIds, New List(Of String)())
                    Dim normalizedSourceId = NormalizeId(sourceId)
                    If Not clauseIds.Contains(normalizedSourceId) Then
                        errors.Add($"Criterion {criterionId} references unknown source clause {normalizedSourceId}.")
                    End If
                Next
            Next

            Dim constraintIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each constraint In If(candidate.Constraints, New List(Of CandidateGoalConstraint)())
                If constraint Is Nothing OrElse String.IsNullOrWhiteSpace(constraint.Id) Then
                    errors.Add("Every goal constraint must have a stable id.")
                    Continue For
                End If
                Dim constraintId = NormalizeId(constraint.Id)
                If Not constraintIds.Add(constraintId) Then errors.Add($"Duplicate constraint id: {constraintId}")
                If String.IsNullOrWhiteSpace(constraint.Statement) Then
                    errors.Add($"Constraint {constraintId} has no statement.")
                End If
                For Each sourceId In If(constraint.SourceClauseIds, New List(Of String)())
                    Dim normalizedSourceId = NormalizeId(sourceId)
                    If Not clauseIds.Contains(normalizedSourceId) Then
                        errors.Add($"Constraint {constraintId} references unknown source clause {normalizedSourceId}.")
                    End If
                Next
            Next

            For Each clause In If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)()).Where(Function(item) item IsNot Nothing)
                Dim mapped = If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Where(
                    Function(criterion) criterion IsNot Nothing AndAlso
                        criterion.Required AndAlso
                        criterion.SourceClauseIds IsNot Nothing AndAlso
                        criterion.SourceClauseIds.Any(Function(id) String.Equals(NormalizeId(id), NormalizeId(clause.Id), StringComparison.OrdinalIgnoreCase))).ToList()
                coverage.Add(New GoalCoverageMapEntry(NormalizeId(clause.Id), mapped.Select(Function(item) NormalizeId(item.Id))))

                If clause.IsExplicit AndAlso mapped.Count = 0 Then
                    errors.Add($"Explicit source clause {clause.Id} is not covered by a required criterion.")
                End If

                Dim unresolved = compilation IsNot Nothing AndAlso
                    compilation.UnresolvedClauses.Any(Function(id) String.Equals(NormalizeId(id), NormalizeId(clause.Id), StringComparison.OrdinalIgnoreCase))
                If unresolved Then
                    Dim preservedVerbatim = mapped.Any(
                        Function(criterion) String.Equals(criterion.Kind, "semantic", StringComparison.OrdinalIgnoreCase) AndAlso
                            String.Equals(If(criterion.Statement, "").Trim(), If(clause.Text, "").Trim(), StringComparison.Ordinal))
                    If Not preservedVerbatim Then
                        errors.Add($"Unresolved clause {clause.Id} was not preserved verbatim as a semantic criterion.")
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(clause.RequiredCapability) Then
                    Dim normalizedClauseCapability = NormalizeId(clause.RequiredCapability)
                    Dim capabilityDeclared = If(candidate.RequiredCapabilities, New List(Of String)()).Any(
                        Function(value) String.Equals(NormalizeId(value), normalizedClauseCapability, StringComparison.OrdinalIgnoreCase))
                    If Not capabilityDeclared Then
                        errors.Add($"Source clause {NormalizeId(clause.Id)} requires capability {normalizedClauseCapability}, but the candidate omitted it.")
                    Else
                        Dim tracedCapabilityCriterion = If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Any(
                            Function(criterion) criterion IsNot Nothing AndAlso
                                criterion.Required AndAlso
                                String.Equals(criterion.Kind, "capability", StringComparison.OrdinalIgnoreCase) AndAlso
                                String.Equals(NormalizeId(criterion.CapabilityId), normalizedClauseCapability, StringComparison.OrdinalIgnoreCase) AndAlso
                                criterion.SourceClauseIds IsNot Nothing AndAlso
                                criterion.SourceClauseIds.Any(Function(id) String.Equals(NormalizeId(id), NormalizeId(clause.Id), StringComparison.OrdinalIgnoreCase)))
                        If Not tracedCapabilityCriterion Then
                            errors.Add($"Required capability {normalizedClauseCapability} is not traced to source clause {NormalizeId(clause.Id)}.")
                        End If
                    End If
                End If
            Next

            For Each unresolvedClauseId In compilation.UnresolvedClauses
                Dim normalizedUnresolvedId = NormalizeId(unresolvedClauseId)
                If Not clauseIds.Contains(normalizedUnresolvedId) Then
                    errors.Add($"Unresolved clause references unknown source clause {normalizedUnresolvedId}.")
                End If
            Next

            Dim requiredCapabilities As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each capability In If(candidate.RequiredCapabilities, New List(Of String)())
                If String.IsNullOrWhiteSpace(capability) Then
                    errors.Add("Required capability ids cannot be empty.")
                    Continue For
                End If
                Dim normalizedCapability = NormalizeId(capability)
                If Not requiredCapabilities.Add(normalizedCapability) Then
                    errors.Add($"Duplicate required capability: {normalizedCapability}")
                End If
                Dim capabilityCriterion = If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Any(
                    Function(criterion) criterion IsNot Nothing AndAlso
                        criterion.Required AndAlso
                        String.Equals(criterion.Kind, "capability", StringComparison.OrdinalIgnoreCase) AndAlso
                        String.Equals(NormalizeId(criterion.CapabilityId), normalizedCapability, StringComparison.OrdinalIgnoreCase))
                If Not capabilityCriterion Then
                    errors.Add($"Required capability {normalizedCapability} has no required capability criterion.")
                End If
            Next

            Return New GoalCoverageValidationResult(errors, coverage, ComputeCandidateFingerprint(candidate))
        End Function

        Friend Shared Function ComputeCandidateFingerprint(candidate As CandidateGoalContract) As String
            If candidate Is Nothing Then Return ""

            Dim canonical As New StringBuilder()
            AppendFingerprintValue(canonical, "raw", candidate.RawUserRequest)
            For Each clause In If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)()).
                Where(Function(item) item IsNot Nothing).
                OrderBy(Function(item) If(item.Id, ""), StringComparer.Ordinal).
                ThenBy(Function(item) If(item.Text, ""), StringComparer.Ordinal)
                AppendFingerprintValue(canonical, "clause.id", clause.Id)
                AppendFingerprintValue(canonical, "clause.text", clause.Text)
                AppendFingerprintValue(canonical, "clause.explicit", clause.IsExplicit.ToString())
                AppendFingerprintValue(canonical, "clause.capability", clause.RequiredCapability)
            Next
            For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).
                Where(Function(item) item IsNot Nothing).
                OrderBy(Function(item) If(item.Id, ""), StringComparer.Ordinal).
                ThenBy(Function(item) If(item.Statement, ""), StringComparer.Ordinal)
                AppendFingerprintValue(canonical, "criterion.id", criterion.Id)
                AppendFingerprintValue(canonical, "criterion.statement", criterion.Statement)
                AppendFingerprintValue(canonical, "criterion.kind", criterion.Kind)
                AppendFingerprintValue(canonical, "criterion.required", criterion.Required.ToString())
                AppendFingerprintValue(canonical, "criterion.verifier", criterion.VerificationCapability)
                AppendFingerprintValue(canonical, "criterion.capability", criterion.CapabilityId)
                For Each sourceId In If(criterion.SourceClauseIds, New List(Of String)()).OrderBy(Function(value) value, StringComparer.Ordinal)
                    AppendFingerprintValue(canonical, "criterion.source", sourceId)
                Next
            Next
            For Each constraint In If(candidate.Constraints, New List(Of CandidateGoalConstraint)()).
                Where(Function(item) item IsNot Nothing).
                OrderBy(Function(item) If(item.Id, ""), StringComparer.Ordinal).
                ThenBy(Function(item) If(item.Statement, ""), StringComparer.Ordinal)
                AppendFingerprintValue(canonical, "constraint.id", constraint.Id)
                AppendFingerprintValue(canonical, "constraint.statement", constraint.Statement)
                AppendFingerprintValue(canonical, "constraint.kind", constraint.Kind)
                AppendFingerprintValue(canonical, "constraint.required", constraint.Required.ToString())
                For Each sourceId In If(constraint.SourceClauseIds, New List(Of String)()).OrderBy(Function(value) value, StringComparer.Ordinal)
                    AppendFingerprintValue(canonical, "constraint.source", sourceId)
                Next
            Next
            For Each capability In If(candidate.RequiredCapabilities, New List(Of String)()).OrderBy(Function(value) value, StringComparer.Ordinal)
                AppendFingerprintValue(canonical, "requiredCapability", capability)
            Next

            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", "")
            End Using
        End Function

        Private Shared Sub AppendFingerprintValue(builder As StringBuilder, name As String, value As String)
            Dim safeName = If(name, "")
            Dim safeValue = If(value, "")
            builder.Append(safeName.Length).Append(":"c).Append(safeName)
            builder.Append(safeValue.Length).Append(":"c).Append(safeValue).Append("|"c)
        End Sub

        Private Shared Function NormalizeId(value As String) As String
            Return If(value, "").Trim()
        End Function
    End Class

End Namespace
