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
            Dim rawUserRequest = If(candidate.RawUserRequest, "")

            Dim clauseIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim explicitClauseIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim clauseById As New Dictionary(Of String, CandidateGoalSourceClause)(StringComparer.OrdinalIgnoreCase)
            Dim explicitOrigins As New HashSet(Of String)(StringComparer.Ordinal)
            Dim assumptionTexts As New HashSet(Of String)(
                If(compilation.Assumptions, New List(Of String)()).
                    Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                    Select(Function(value) value.Trim()),
                StringComparer.Ordinal)
            For Each clause In If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)())
                If clause Is Nothing OrElse String.IsNullOrWhiteSpace(clause.Id) Then
                    errors.Add("Every source clause must have a stable id.")
                    Continue For
                End If
                Dim clauseId = NormalizeId(clause.Id)
                If Not clauseIds.Add(clauseId) Then errors.Add($"Duplicate source clause id: {clauseId}")
                If Not clauseById.ContainsKey(clauseId) Then clauseById.Add(clauseId, clause)
                If Not clause.IsExplicit Then
                    errors.Add($"Source clause {clauseId} is non-explicit; inferred content belongs in Assumptions, not GoalContract.")
                End If
                If clause.IsExplicit AndAlso String.IsNullOrWhiteSpace(clause.Text) Then
                    errors.Add($"Explicit source clause {clauseId} has no text.")
                ElseIf clause.IsExplicit Then
                    explicitClauseIds.Add(clauseId)
                    If clause.SourceStart < 0 OrElse clause.SourceStart > rawUserRequest.Length - clause.Text.Length OrElse
                       Not String.Equals(rawUserRequest.Substring(clause.SourceStart, clause.Text.Length), clause.Text, StringComparison.Ordinal) Then
                        errors.Add($"Explicit source clause {clauseId} has no verified exact occurrence in RawUserRequest.")
                    ElseIf Not GoalSourceAuthority.IsCompleteSemanticSpan(
                        rawUserRequest,
                        clause.SourceStart,
                        clause.Text) Then
                        errors.Add($"Explicit source clause {clauseId} is an incomplete semantic span; surrounding polarity or modality was omitted.")
                    Else
                        Dim origin = clause.SourceStart.ToString(Globalization.CultureInfo.InvariantCulture) & ":" & clause.Text.Length.ToString(Globalization.CultureInfo.InvariantCulture)
                        If Not explicitOrigins.Add(origin) Then
                            errors.Add($"Duplicate explicit source occurrence at {origin}.")
                        End If
                    End If
                End If
            Next
            If clauseIds.Count = 0 Then errors.Add("At least one source clause is required.")
            If explicitClauseIds.Count > 0 AndAlso Not GoalSourceAuthority.CoversAuthoritativeText(
                rawUserRequest,
                clauseById.Values.Where(Function(item) item IsNot Nothing AndAlso item.IsExplicit)) Then
                errors.Add("Explicit source clauses do not cover all authoritative user text; a requirement or governing modifier was omitted.")
            End If

            Dim criterionIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each criterion In If(candidate.Criteria, New List(Of CandidateGoalCriterion)()).Where(
                Function(item) item IsNot Nothing AndAlso item.Required)
                If String.IsNullOrWhiteSpace(criterion.Id) Then
                    errors.Add("Every goal criterion must have a stable id.")
                    Continue For
                End If
                Dim criterionId = NormalizeId(criterion.Id)
                If Not criterionIds.Add(criterionId) Then errors.Add($"Duplicate criterion id: {criterionId}")
                If String.IsNullOrWhiteSpace(criterion.Statement) Then
                    errors.Add($"Criterion {criterionId} has no statement.")
                End If
                If criterion.Required AndAlso assumptionTexts.Contains(If(criterion.Statement, "").Trim()) Then
                    errors.Add($"Required criterion {criterionId} duplicates a non-authoritative assumption.")
                End If
                Dim criterionSources = If(criterion.SourceClauseIds, New List(Of String)())
                If criterion.Required AndAlso criterionSources.Count = 0 Then
                    errors.Add($"Required criterion {criterionId} must trace to an explicit source clause.")
                End If
                Dim seenCriterionSources As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each sourceId In criterionSources
                    Dim normalizedSourceId = NormalizeId(sourceId)
                    If Not seenCriterionSources.Add(normalizedSourceId) Then
                        errors.Add($"Criterion {criterionId} contains duplicate source clause reference {normalizedSourceId}.")
                    End If
                    If Not clauseIds.Contains(normalizedSourceId) Then
                        errors.Add($"Criterion {criterionId} references unknown source clause {normalizedSourceId}.")
                    ElseIf criterion.Required AndAlso Not explicitClauseIds.Contains(normalizedSourceId) Then
                        errors.Add($"Required criterion {criterionId} references non-explicit source clause {normalizedSourceId}.")
                    End If
                Next
                If criterion.Required AndAlso criterionSources.Count > 0 Then
                    Dim verbatimAuthorized = criterionSources.Any(
                        Function(sourceId)
                            Dim source As CandidateGoalSourceClause = Nothing
                            Return clauseById.TryGetValue(NormalizeId(sourceId), source) AndAlso
                                source IsNot Nothing AndAlso source.IsExplicit AndAlso
                                String.Equals(If(criterion.Statement, ""), If(source.Text, ""), StringComparison.Ordinal)
                        End Function)
                    If Not verbatimAuthorized Then
                        errors.Add($"Required criterion {criterionId} is a model paraphrase; required semantics must equal a referenced explicit clause verbatim.")
                    End If
                End If
                If criterion.Required AndAlso String.Equals(criterion.Kind, "capability", StringComparison.OrdinalIgnoreCase) Then
                    Dim capabilityId = NormalizeId(criterion.CapabilityId)
                    If String.IsNullOrWhiteSpace(capabilityId) OrElse
                       Not If(candidate.RequiredCapabilities, New List(Of String)()).Any(
                           Function(value) String.Equals(NormalizeId(value), capabilityId, StringComparison.OrdinalIgnoreCase)) Then
                        errors.Add($"Required capability criterion {criterionId} lacks independently authorized capability evidence.")
                    ElseIf Not criterionSources.Any(
                        Function(sourceId)
                            Dim source As CandidateGoalSourceClause = Nothing
                            Return clauseById.TryGetValue(NormalizeId(sourceId), source) AndAlso
                                source IsNot Nothing AndAlso
                                GoalCapabilityEvidenceResolver.ClauseSupports(source.Text, capabilityId)
                        End Function) Then
                        errors.Add($"Required capability criterion {criterionId} is not attached to exact capability evidence.")
                    End If
                End If
            Next

            Dim constraintIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each constraint In If(candidate.Constraints, New List(Of CandidateGoalConstraint)()).Where(
                Function(item) item IsNot Nothing AndAlso item.Required)
                If String.IsNullOrWhiteSpace(constraint.Id) Then
                    errors.Add("Every goal constraint must have a stable id.")
                    Continue For
                End If
                Dim constraintId = NormalizeId(constraint.Id)
                If Not constraintIds.Add(constraintId) Then errors.Add($"Duplicate constraint id: {constraintId}")
                If String.IsNullOrWhiteSpace(constraint.Statement) Then
                    errors.Add($"Constraint {constraintId} has no statement.")
                End If
                If constraint.Required AndAlso assumptionTexts.Contains(If(constraint.Statement, "").Trim()) Then
                    errors.Add($"Required constraint {constraintId} duplicates a non-authoritative assumption.")
                End If
                Dim constraintSources = If(constraint.SourceClauseIds, New List(Of String)())
                If constraint.Required AndAlso constraintSources.Count = 0 Then
                    errors.Add($"Required constraint {constraintId} must trace to an explicit source clause.")
                End If
                Dim seenConstraintSources As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each sourceId In constraintSources
                    Dim normalizedSourceId = NormalizeId(sourceId)
                    If Not seenConstraintSources.Add(normalizedSourceId) Then
                        errors.Add($"Constraint {constraintId} contains duplicate source clause reference {normalizedSourceId}.")
                    End If
                    If Not clauseIds.Contains(normalizedSourceId) Then
                        errors.Add($"Constraint {constraintId} references unknown source clause {normalizedSourceId}.")
                    ElseIf constraint.Required AndAlso Not explicitClauseIds.Contains(normalizedSourceId) Then
                        errors.Add($"Required constraint {constraintId} references non-explicit source clause {normalizedSourceId}.")
                    End If
                Next
                If constraint.Required AndAlso constraintSources.Count > 0 Then
                    Dim verbatimAuthorized = constraintSources.Any(
                        Function(sourceId)
                            Dim source As CandidateGoalSourceClause = Nothing
                            Return clauseById.TryGetValue(NormalizeId(sourceId), source) AndAlso
                                source IsNot Nothing AndAlso source.IsExplicit AndAlso
                                String.Equals(If(constraint.Statement, ""), If(source.Text, ""), StringComparison.Ordinal)
                        End Function)
                    If Not verbatimAuthorized Then
                        errors.Add($"Required constraint {constraintId} is a model paraphrase; required semantics must equal a referenced explicit clause verbatim.")
                    End If
                End If
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

                If clause.IsExplicit Then
                    Dim preservedVerbatim = mapped.Any(
                        Function(criterion) String.Equals(If(criterion.Statement, ""), If(clause.Text, ""), StringComparison.Ordinal))
                    If Not preservedVerbatim Then
                        errors.Add($"Explicit source clause {clause.Id} was not preserved verbatim as a required criterion.")
                    End If
                End If

                Dim unresolved = compilation IsNot Nothing AndAlso
                    compilation.UnresolvedClauses.Any(Function(id) String.Equals(NormalizeId(id), NormalizeId(clause.Id), StringComparison.OrdinalIgnoreCase))
                If unresolved Then
                    Dim preservedUnresolvedVerbatim = mapped.Any(
                        Function(criterion) String.Equals(If(criterion.Statement, ""), If(clause.Text, ""), StringComparison.Ordinal))
                    If Not preservedUnresolvedVerbatim Then
                        errors.Add($"Unresolved clause {clause.Id} was not preserved verbatim as a required criterion.")
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

            Dim evidencedCapabilities As New HashSet(Of String)(
                GoalCapabilityEvidenceResolver.Resolve(rawUserRequest),
                StringComparer.OrdinalIgnoreCase)
            If Not requiredCapabilities.SetEquals(evidencedCapabilities) Then
                Dim missingEvidence = requiredCapabilities.Where(Function(value) Not evidencedCapabilities.Contains(value)).ToList()
                Dim omittedEvidence = evidencedCapabilities.Where(Function(value) Not requiredCapabilities.Contains(value)).ToList()
                If missingEvidence.Count > 0 Then
                    errors.Add("Required capabilities lack exact user-language evidence: " & String.Join(", ", missingEvidence))
                End If
                If omittedEvidence.Count > 0 Then
                    errors.Add("Explicit method capabilities were omitted: " & String.Join(", ", omittedEvidence))
                End If
            End If

            Return New GoalCoverageValidationResult(errors, coverage, ComputeCandidateFingerprint(candidate))
        End Function

        Friend Shared Function ComputeCandidateFingerprint(candidate As CandidateGoalContract) As String
            If candidate Is Nothing Then Return ""

            Dim canonical As New StringBuilder()
            AppendFingerprintValue(canonical, "raw", candidate.RawUserRequest)
            Dim clauses = If(candidate.SourceClauses, New List(Of CandidateGoalSourceClause)())
            AppendFingerprintValue(canonical, "clauses.count", clauses.Count.ToString(Globalization.CultureInfo.InvariantCulture))
            For clauseIndex = 0 To clauses.Count - 1
                Dim clause = clauses(clauseIndex)
                AppendFingerprintValue(canonical, "clause.index", clauseIndex.ToString(Globalization.CultureInfo.InvariantCulture))
                If clause Is Nothing Then
                    AppendFingerprintValue(canonical, "clause.null", "true")
                    Continue For
                End If
                AppendFingerprintValue(canonical, "clause.id", clause.Id)
                AppendFingerprintValue(canonical, "clause.text", clause.Text)
                AppendFingerprintValue(canonical, "clause.explicit", clause.IsExplicit.ToString())
                AppendFingerprintValue(canonical, "clause.sourceStart", clause.SourceStart.ToString(Globalization.CultureInfo.InvariantCulture))
            Next
            Dim criteria = If(candidate.Criteria, New List(Of CandidateGoalCriterion)())
            AppendFingerprintValue(canonical, "criteria.count", criteria.Count.ToString(Globalization.CultureInfo.InvariantCulture))
            For criterionIndex = 0 To criteria.Count - 1
                Dim criterion = criteria(criterionIndex)
                AppendFingerprintValue(canonical, "criterion.index", criterionIndex.ToString(Globalization.CultureInfo.InvariantCulture))
                If criterion Is Nothing Then
                    AppendFingerprintValue(canonical, "criterion.null", "true")
                    Continue For
                End If
                AppendFingerprintValue(canonical, "criterion.id", criterion.Id)
                AppendFingerprintValue(canonical, "criterion.statement", criterion.Statement)
                AppendFingerprintValue(canonical, "criterion.kind", criterion.Kind)
                AppendFingerprintValue(canonical, "criterion.required", criterion.Required.ToString())
                AppendFingerprintValue(canonical, "criterion.verifier", criterion.VerificationCapability)
                AppendFingerprintValue(canonical, "criterion.capability", criterion.CapabilityId)
                Dim criterionSources = If(criterion.SourceClauseIds, New List(Of String)())
                AppendFingerprintValue(canonical, "criterion.sources.count", criterionSources.Count.ToString(Globalization.CultureInfo.InvariantCulture))
                For sourceIndex = 0 To criterionSources.Count - 1
                    AppendFingerprintValue(canonical, "criterion.source.index", sourceIndex.ToString(Globalization.CultureInfo.InvariantCulture))
                    AppendFingerprintValue(canonical, "criterion.source", criterionSources(sourceIndex))
                Next
            Next
            Dim constraints = If(candidate.Constraints, New List(Of CandidateGoalConstraint)())
            AppendFingerprintValue(canonical, "constraints.count", constraints.Count.ToString(Globalization.CultureInfo.InvariantCulture))
            For constraintIndex = 0 To constraints.Count - 1
                Dim constraint = constraints(constraintIndex)
                AppendFingerprintValue(canonical, "constraint.index", constraintIndex.ToString(Globalization.CultureInfo.InvariantCulture))
                If constraint Is Nothing Then
                    AppendFingerprintValue(canonical, "constraint.null", "true")
                    Continue For
                End If
                AppendFingerprintValue(canonical, "constraint.id", constraint.Id)
                AppendFingerprintValue(canonical, "constraint.statement", constraint.Statement)
                AppendFingerprintValue(canonical, "constraint.kind", constraint.Kind)
                AppendFingerprintValue(canonical, "constraint.required", constraint.Required.ToString())
                Dim constraintSources = If(constraint.SourceClauseIds, New List(Of String)())
                AppendFingerprintValue(canonical, "constraint.sources.count", constraintSources.Count.ToString(Globalization.CultureInfo.InvariantCulture))
                For sourceIndex = 0 To constraintSources.Count - 1
                    AppendFingerprintValue(canonical, "constraint.source.index", sourceIndex.ToString(Globalization.CultureInfo.InvariantCulture))
                    AppendFingerprintValue(canonical, "constraint.source", constraintSources(sourceIndex))
                Next
            Next
            Dim capabilities = If(candidate.RequiredCapabilities, New List(Of String)())
            AppendFingerprintValue(canonical, "capabilities.count", capabilities.Count.ToString(Globalization.CultureInfo.InvariantCulture))
            For capabilityIndex = 0 To capabilities.Count - 1
                AppendFingerprintValue(canonical, "capability.index", capabilityIndex.ToString(Globalization.CultureInfo.InvariantCulture))
                AppendFingerprintValue(canonical, "requiredCapability", capabilities(capabilityIndex))
            Next

            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", "")
            End Using
        End Function

        Friend Shared Function ComputeCompilationFingerprint(compilation As GoalCompilationResult) As String
            If compilation Is Nothing Then Return ""
            Dim canonical As New StringBuilder()
            AppendFingerprintValue(canonical, "candidate", ComputeCandidateFingerprint(compilation.Candidate))
            AppendFingerprintValue(canonical, "requiresClarification", compilation.RequiresClarification.ToString())
            AppendFingerprintList(canonical, "unresolved", compilation.UnresolvedClauses)
            AppendFingerprintList(canonical, "assumptions", compilation.Assumptions)
            AppendFingerprintList(canonical, "diagnostics", compilation.Diagnostics)
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()))).Replace("-", "")
            End Using
        End Function

        Private Shared Sub AppendFingerprintList(builder As StringBuilder,
                                                 name As String,
                                                 values As IEnumerable(Of String))
            Dim items = If(values, Enumerable.Empty(Of String)()).ToList()
            AppendFingerprintValue(builder, name & ".count", items.Count.ToString(Globalization.CultureInfo.InvariantCulture))
            For index = 0 To items.Count - 1
                AppendFingerprintValue(builder, name & ".index", index.ToString(Globalization.CultureInfo.InvariantCulture))
                AppendFingerprintValue(builder, name & ".value", items(index))
            Next
        End Sub

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
