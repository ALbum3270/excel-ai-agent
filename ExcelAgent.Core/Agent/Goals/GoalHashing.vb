Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text

Namespace Agent.Goals

    ''' <summary>
    ''' Local hashing Module with three deliberately separate meanings:
    ''' candidate fingerprint is an exact pre-freeze snapshot, ContractHash protects the exact
    ''' frozen payload, and SemanticHash canonicalizes verified source occurrences while
    ''' ignoring model-generated ids and collection order. Canonicalization is polynomial in
    ''' request length and graph size; it never performs permutation search on the request path.
    ''' </summary>
    Friend Module GoalHashing

        Friend Function ComputeContractHash(rawUserRequest As String,
                                            sourceClauses As IEnumerable(Of GoalSourceClause),
                                            criteria As IEnumerable(Of GoalCriterion),
                                            constraints As IEnumerable(Of GoalConstraint),
                                            capabilities As IEnumerable(Of String)) As String
            Dim clauses = If(sourceClauses, Enumerable.Empty(Of GoalSourceClause)()).ToList()
            Dim criterionList = If(criteria, Enumerable.Empty(Of GoalCriterion)()).ToList()
            Dim constraintList = If(constraints, Enumerable.Empty(Of GoalConstraint)()).ToList()
            Dim capabilityList = If(capabilities, Enumerable.Empty(Of String)()).ToList()
            Dim serialized As New StringBuilder()
            AppendToken(serialized, "domain", "office-ai.goal.contract.v2")
            AppendToken(serialized, "raw", If(rawUserRequest, ""))

            AppendCount(serialized, "clauses", clauses.Count)
            For Each clause In clauses
                AppendToken(serialized, "clause.id", clause.Id)
                AppendToken(serialized, "clause.text", clause.Text)
                AppendToken(serialized, "clause.explicit", If(clause.IsExplicit, "1", "0"))
                AppendToken(serialized, "clause.sourceStart", clause.SourceStart.ToString(CultureInfo.InvariantCulture))
            Next

            AppendCount(serialized, "criteria", criterionList.Count)
            For Each criterion In criterionList
                AppendToken(serialized, "criterion.id", criterion.Id)
                AppendToken(serialized, "criterion.statement", criterion.Statement)
                AppendToken(serialized, "criterion.kind", criterion.Kind)
                AppendToken(serialized, "criterion.required", If(criterion.Required, "1", "0"))
                AppendToken(serialized, "criterion.verifier", criterion.VerificationCapability)
                AppendToken(serialized, "criterion.capability", criterion.CapabilityId)
                AppendCount(serialized, "criterion.sources", criterion.SourceClauseIds.Count)
                For Each sourceId In criterion.SourceClauseIds
                    AppendToken(serialized, "criterion.source", sourceId)
                Next
            Next

            AppendCount(serialized, "constraints", constraintList.Count)
            For Each constraint In constraintList
                AppendToken(serialized, "constraint.id", constraint.Id)
                AppendToken(serialized, "constraint.statement", constraint.Statement)
                AppendToken(serialized, "constraint.kind", constraint.Kind)
                AppendToken(serialized, "constraint.required", If(constraint.Required, "1", "0"))
                AppendCount(serialized, "constraint.sources", constraint.SourceClauseIds.Count)
                For Each sourceId In constraint.SourceClauseIds
                    AppendToken(serialized, "constraint.source", sourceId)
                Next
            Next

            AppendCount(serialized, "capabilities", capabilityList.Count)
            For Each capability In capabilityList
                AppendToken(serialized, "requiredCapability", capability)
            Next
            Return HashText(serialized.ToString())
        End Function

        Friend Function ComputeSemanticHash(rawUserRequest As String,
                                            sourceClauses As IEnumerable(Of GoalSourceClause),
                                            criteria As IEnumerable(Of GoalCriterion),
                                            constraints As IEnumerable(Of GoalConstraint),
                                            capabilities As IEnumerable(Of String)) As String
            Dim raw = If(rawUserRequest, "")
            Dim clauses = If(sourceClauses, Enumerable.Empty(Of GoalSourceClause)()).ToList()
            Dim criterionList = If(criteria, Enumerable.Empty(Of GoalCriterion)()).ToList()
            Dim constraintList = If(constraints, Enumerable.Empty(Of GoalConstraint)()).ToList()
            Dim idToSourceIdentity As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim sourceIdentities As New HashSet(Of String)(StringComparer.Ordinal)
            Dim clauseRecords As New List(Of String)()

            For Each clause In clauses
                If idToSourceIdentity.ContainsKey(clause.Id) Then
                    Throw New InvalidOperationException("Semantic hashing requires unique source clause ids.")
                End If
                Dim sourceIdentity = VerifiedSourceIdentity(raw, clause)
                If Not sourceIdentities.Add(sourceIdentity) Then
                    Throw New InvalidOperationException("Semantic hashing requires unique verified source occurrences.")
                End If
                idToSourceIdentity.Add(clause.Id, sourceIdentity)
                clauseRecords.Add(SemanticClauseRecord(clause, sourceIdentity))
            Next

            Dim builder As New StringBuilder()
            AppendToken(builder, "domain", "office-ai.goal.semantic.v2")
            AppendToken(builder, "raw", NormalizeNaturalText(raw))

            clauseRecords.Sort(StringComparer.Ordinal)
            AppendCount(builder, "clauses", clauseRecords.Count)
            For Each record In clauseRecords
                AppendToken(builder, "clause", record)
            Next

            Dim criterionRecords = criterionList.
                Select(Function(item) SemanticCriterionRecord(item, idToSourceIdentity)).
                OrderBy(Function(value) value, StringComparer.Ordinal).
                ToList()
            AppendCount(builder, "criteria", criterionRecords.Count)
            For Each record In criterionRecords
                AppendToken(builder, "criterion", record)
            Next

            Dim constraintRecords = constraintList.
                Select(Function(item) SemanticConstraintRecord(item, idToSourceIdentity)).
                OrderBy(Function(value) value, StringComparer.Ordinal).
                ToList()
            AppendCount(builder, "constraints", constraintRecords.Count)
            For Each record In constraintRecords
                AppendToken(builder, "constraint", record)
            Next

            Dim normalizedCapabilities = If(capabilities, Enumerable.Empty(Of String)()).
                Select(Function(value) NormalizeIdentifier(value)).
                OrderBy(Function(value) value, StringComparer.Ordinal).
                ToList()
            AppendCount(builder, "capabilities", normalizedCapabilities.Count)
            For Each capability In normalizedCapabilities
                AppendToken(builder, "requiredCapability", capability)
            Next
            Return HashText(builder.ToString())
        End Function

        Private Function VerifiedSourceIdentity(rawUserRequest As String,
                                                clause As GoalSourceClause) As String
            If clause Is Nothing OrElse Not clause.IsExplicit Then
                Throw New InvalidOperationException("Semantic hashing accepts only explicit source clauses.")
            End If
            Dim text = If(clause.Text, "")
            If clause.SourceStart < 0 OrElse clause.SourceStart > rawUserRequest.Length - text.Length OrElse
               Not String.Equals(rawUserRequest.Substring(clause.SourceStart, text.Length), text, StringComparison.Ordinal) Then
                Throw New InvalidOperationException("Semantic hashing requires a verified source occurrence.")
            End If

            Dim builder As New StringBuilder()
            AppendToken(builder, "prefix", NormalizeNaturalText(rawUserRequest.Substring(0, clause.SourceStart)))
            AppendToken(builder, "text", NormalizeNaturalText(text))
            Return builder.ToString()
        End Function

        Private Function SemanticClauseRecord(clause As GoalSourceClause,
                                              sourceIdentity As String) As String
            Dim builder As New StringBuilder()
            AppendToken(builder, "origin", sourceIdentity)
            AppendToken(builder, "explicit", If(clause.IsExplicit, "1", "0"))
            Return builder.ToString()
        End Function

        Private Function SemanticCriterionRecord(criterion As GoalCriterion,
                                                 idToSourceIdentity As Dictionary(Of String, String)) As String
            Dim builder As New StringBuilder()
            AppendToken(builder, "statement", NormalizeNaturalText(criterion.Statement))
            AppendToken(builder, "kind", NormalizeIdentifier(criterion.Kind))
            AppendToken(builder, "required", If(criterion.Required, "1", "0"))
            AppendToken(builder, "verifier", NormalizeIdentifier(criterion.VerificationCapability))
            AppendToken(builder, "capability", NormalizeIdentifier(criterion.CapabilityId))
            AppendSourceIdentities(builder, criterion.SourceClauseIds, idToSourceIdentity)
            Return builder.ToString()
        End Function

        Private Function SemanticConstraintRecord(constraint As GoalConstraint,
                                                  idToSourceIdentity As Dictionary(Of String, String)) As String
            Dim builder As New StringBuilder()
            AppendToken(builder, "statement", NormalizeNaturalText(constraint.Statement))
            AppendToken(builder, "kind", NormalizeIdentifier(constraint.Kind))
            AppendToken(builder, "required", If(constraint.Required, "1", "0"))
            AppendSourceIdentities(builder, constraint.SourceClauseIds, idToSourceIdentity)
            Return builder.ToString()
        End Function

        Private Sub AppendSourceIdentities(builder As StringBuilder,
                                           sourceIds As IEnumerable(Of String),
                                           idToSourceIdentity As Dictionary(Of String, String))
            Dim identities As New List(Of String)()
            For Each sourceId In If(sourceIds, Enumerable.Empty(Of String)())
                Dim identity As String = Nothing
                If Not idToSourceIdentity.TryGetValue(If(sourceId, "").Trim(), identity) Then
                    Throw New InvalidOperationException("Semantic hashing found an unknown source clause reference.")
                End If
                identities.Add(identity)
            Next
            identities.Sort(StringComparer.Ordinal)
            AppendCount(builder, "sources", identities.Count)
            For Each identity In identities
                AppendToken(builder, "source", identity)
            Next
        End Sub

        Private Function NormalizeNaturalText(value As String) As String
            Return If(value, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Normalize(NormalizationForm.FormC)
        End Function

        Private Function NormalizeIdentifier(value As String) As String
            Return NormalizeNaturalText(If(value, "").Trim()).ToLowerInvariant()
        End Function

        Private Sub AppendCount(builder As StringBuilder, name As String, count As Integer)
            AppendToken(builder, name & ".count", count.ToString(CultureInfo.InvariantCulture))
        End Sub

        Private Sub AppendToken(builder As StringBuilder, name As String, value As String)
            Dim safeName = If(name, "")
            Dim safeValue = If(value, "")
            Dim valueBytes = Encoding.UTF8.GetBytes(safeValue)
            builder.Append(safeName.Length.ToString(CultureInfo.InvariantCulture)).Append(":"c).Append(safeName)
            builder.Append(valueBytes.Length.ToString(CultureInfo.InvariantCulture)).Append(":"c)
            builder.Append(Convert.ToBase64String(valueBytes)).Append("|"c)
        End Sub

        Private Function HashText(value As String) As String
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(If(value, "")))).Replace("-", "")
            End Using
        End Function
    End Module

End Namespace
