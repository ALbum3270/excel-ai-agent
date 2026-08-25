Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        ''' <summary>
        ''' Converts the model's semantic completion verdict (Goal criterion -> observed
        ''' evidence) into the exact low-level projection consumed by AgentGoalVerifier.  The
        ''' model decides whether the observed data satisfies the user request; the harness
        ''' owns targets, effects, values and dependency lineage and therefore never asks the
        ''' model to transcribe an OutcomeContract.
        ''' </summary>
        Private Function BuildObservedCompletionProjection(session As AgentSession,
                                                            ByRef evidenceClaims As List(Of String),
                                                            criterionEvidence As IDictionary(Of String, List(Of String)),
                                                            ByRef contract As OutcomeContract) As String
            contract = Nothing
            If session?.Spec Is Nothing Then Return "Task specification is missing."

            Dim requiredCriterionIds = Goals.GoalOutcomeProjection.RequiredHostCriterionIds(session.Spec)
            If requiredCriterionIds.Count = 0 Then
                Return "The frozen Goal has no host-verifiable criterion."
            End If

            Dim allEvidence = If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item IsNot Nothing).
                SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                Where(Function(item) item IsNot Nothing).
                ToList()
            Dim evidenceById = allEvidence.
                Where(Function(item) Not String.IsNullOrWhiteSpace(item.EvidenceId)).
                GroupBy(Function(item) item.EvidenceId, StringComparer.OrdinalIgnoreCase).
                ToDictionary(Function(group) group.Key, Function(group) group.First(), StringComparer.OrdinalIgnoreCase)

            Dim mapping As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
            If criterionEvidence IsNot Nothing AndAlso criterionEvidence.Count > 0 Then
                For Each pair In criterionEvidence
                    Dim criterionId = If(pair.Key, "").Trim()
                    If Not requiredCriterionIds.Contains(criterionId, StringComparer.OrdinalIgnoreCase) Then
                        Return $"criterionEvidence references unknown or non-host criterion {criterionId}."
                    End If
                    mapping(criterionId) = New List(Of String)(If(pair.Value, New List(Of String)()))
                Next
            Else
                Dim sharedClaims = New List(Of String)(If(evidenceClaims, New List(Of String)()))
                If sharedClaims.All(Function(value) String.IsNullOrWhiteSpace(value)) Then
                    sharedClaims = allEvidence.
                        Where(Function(record) record.Satisfied AndAlso
                            Not String.Equals(record.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase) AndAlso
                            Not String.Equals(record.EffectType, "compute_artifact", StringComparison.OrdinalIgnoreCase) AndAlso
                            Not String.Equals(record.EffectType, "unclassified_mutation", StringComparison.OrdinalIgnoreCase)).
                        OrderByDescending(Function(record) record.WorldRevision).
                        GroupBy(Function(record) record.EvidenceId, StringComparer.OrdinalIgnoreCase).
                        Select(Function(group) group.First().EvidenceId).
                        ToList()
                End If
                If sharedClaims.Count = 0 Then
                    Return "The completion decision cited no satisfied observable final-state evidence."
                End If
                If requiredCriterionIds.Count <> 1 Then
                    Return "The Goal has multiple independent host criteria, so one shared evidence set cannot be copied onto every criterion. Use the provisional Goal mapping or provide criterionEvidence."
                End If
                mapping(requiredCriterionIds(0)) = sharedClaims
            End If

            For Each criterionId In requiredCriterionIds
                If Not mapping.ContainsKey(criterionId) OrElse
                   mapping(criterionId) Is Nothing OrElse
                   mapping(criterionId).All(Function(value) String.IsNullOrWhiteSpace(value)) Then
                    Return $"No inspected evidence was mapped to required Goal criterion {criterionId}."
                End If
            Next

            Dim normalizedClaims As New List(Of String)()
            Dim projection As New OutcomeContract With {.SchemaVersion = "1.0"}
            Dim ordinal As Integer = 0
            Dim isReadOnly = String.Equals(session.Spec.MutationPolicy, "read_only", StringComparison.OrdinalIgnoreCase)

            For Each criterionId In requiredCriterionIds
                Dim records As New List(Of OutcomeEvidenceRecord)()
                For Each claim In mapping(criterionId)
                    Dim matched = ResolveCompletionEvidenceClaim(session, evidenceById, claim)
                    If matched.Count = 0 Then Return $"Completion cites unknown evidence {claim}."
                    For Each record In matched
                        If Not record.Satisfied OrElse
                           String.Equals(record.EffectType, "unclassified_mutation", StringComparison.OrdinalIgnoreCase) Then
                            Return $"Completion evidence {record.EvidenceId} is not a satisfied host assertion."
                        End If
                        If Not records.Any(Function(existing) String.Equals(existing.EvidenceId, record.EvidenceId, StringComparison.OrdinalIgnoreCase)) Then
                            records.Add(record)
                        End If
                    Next
                Next

                Dim goalRecords = records.Where(
                    Function(record)
                        If isReadOnly Then Return True
                        Return Not String.Equals(record.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase) AndAlso
                            Not String.Equals(record.EffectType, "compute_artifact", StringComparison.OrdinalIgnoreCase)
                    End Function).ToList()
                If goalRecords.Count = 0 Then
                    Return $"Evidence mapped to {criterionId} contains no observable final state; inspect the destination and cite that observation."
                End If

                For Each record In goalRecords
                    ordinal += 1
                    Dim requirement = CreateObservedRequirement(session, record, criterionId, ordinal)
                    If requirement Is Nothing Then
                        Return $"Evidence {record.EvidenceId} cannot be converted into a verifiable host requirement."
                    End If
                    projection.Requirements.Add(requirement)
                    AddUniqueClaim(normalizedClaims, record.EvidenceId)
                Next
            Next

            ' A final write derived from an explicitly required compute capability must retain
            ' the exact source-read lineage. Add only those read assertions that are ancestors
            ' of the compute result; unrelated reads never become decorative proof.
            For Each outputRequirement In projection.Requirements.Where(
                Function(item) item IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(item.DerivedFromCapability)).ToList()
                Dim outputRecord = allEvidence.FirstOrDefault(
                    Function(item) String.Equals(item.TargetRef, outputRequirement.TargetRef, StringComparison.OrdinalIgnoreCase) AndAlso
                                   String.Equals(item.EffectType, outputRequirement.EffectType, StringComparison.OrdinalIgnoreCase))
                If outputRecord Is Nothing Then Continue For
                Dim readRecords = FindComputeSourceReadEvidence(session, outputRecord, outputRequirement.DerivedFromCapability)
                If readRecords.Count = 0 Then
                    Return $"Output evidence {outputRecord.EvidenceId} has no verified source read lineage for {outputRequirement.DerivedFromCapability}."
                End If
                For Each readRecord In readRecords
                    ordinal += 1
                    projection.Requirements.Add(New OutcomeRequirement With {
                        .Id = $"observed-support-{ordinal}",
                        .AppType = session.AppType,
                        .TargetRef = readRecord.TargetRef,
                        .EffectType = "read_coverage",
                        .Operator = "covers",
                        .Required = True,
                        .Description = "Verified source data used by " & outputRequirement.DerivedFromCapability
                    })
                    AddUniqueClaim(normalizedClaims, readRecord.EvidenceId)
                Next
            Next

            evidenceClaims = normalizedClaims
            contract = projection
            Return ""
        End Function

        ''' <summary>
        ''' Reuses the planning model's semantic Goal-to-requirement mapping after execution,
        ''' but replaces every guessed target/value with exact typed host evidence. This is the
        ''' normal completion path when a provider returns a final answer without redundantly
        ''' transcribing criterionEvidence. It is not an auto-complete path: the model has
        ''' already observed the last tool result and explicitly chosen decision=complete.
        ''' </summary>
        Private Function BuildGroundedProvisionalCompletionProjection(session As AgentSession,
                                                                       ByRef evidenceClaims As List(Of String),
                                                                       ByRef contract As OutcomeContract) As String
            Return GroundedCompletionProjectionBuilder.Build(
                session,
                evidenceClaims,
                contract,
                AddressOf GroundCompletionProjectionInEvidence)
        End Function

        Private Function CreateObservedRequirement(session As AgentSession,
                                                   record As OutcomeEvidenceRecord,
                                                   criterionId As String,
                                                   ordinal As Integer) As OutcomeRequirement
            If record Is Nothing OrElse String.IsNullOrWhiteSpace(record.TargetRef) OrElse
               String.IsNullOrWhiteSpace(record.EffectType) Then Return Nothing

            Dim requirement As New OutcomeRequirement With {
                .Id = $"observed-goal-{ordinal}",
                .AppType = session.AppType,
                .TargetRef = record.TargetRef,
                .EffectType = record.EffectType,
                .PropertyName = If(record.PropertyName, "").Trim(),
                .CriterionIds = New List(Of String) From {criterionId},
                .Required = True,
                .Description = ResolveGoalCriterionStatement(session, criterionId)
            }
            Dim effect = requirement.EffectType.Trim().ToLowerInvariant()
            Select Case effect
                Case "object_exists", "object_absent", "artifact"
                    requirement.PropertyName = ""
                    requirement.Operator = "exists"
                    requirement.ExpectedValue = Nothing
                Case "read_coverage"
                    requirement.Operator = "covers"
                    requirement.ExpectedValue = Nothing
                Case Else
                    requirement.Operator = "equals"
                    requirement.ExpectedValue = OutcomeProjectionValue.ClonePropertyValue(
                        record.Actual,
                        requirement.PropertyName)
                    If requirement.ExpectedValue Is Nothing AndAlso record.RequestVerified Then
                        requirement.ExpectedValue = OutcomeProjectionValue.ClonePropertyValue(
                            record.VerifiedRequest,
                            requirement.PropertyName)
                    End If
            End Select

            Dim computeProducer = ResolveRequiredComputeProducer(session, record)
            If Not String.IsNullOrWhiteSpace(computeProducer) Then
                requirement.DerivedFromCapability = computeProducer
                If effect = "data_state" Then requirement.ExpectedValue = Nothing
            End If
            Return requirement
        End Function

        Private Function ResolveRequiredComputeProducer(session As AgentSession,
                                                        record As OutcomeEvidenceRecord) As String
            Dim ancestors = GetObservedDependencyClosure(session, record?.DerivedFromEvidenceIds)
            For Each capability In AgentExecutionContract.ResolveRequiredCapabilities(session?.Spec)
                Dim descriptor = _toolRegistry.GetTool(capability)
                If descriptor Is Nothing OrElse
                   Not String.Equals(descriptor.AccessMode, "compute", StringComparison.OrdinalIgnoreCase) Then Continue For
                If session.Iterations.Any(
                    Function(item) item IsNot Nothing AndAlso ancestors.Contains(item.EvidenceId) AndAlso
                                   item.Action IsNot Nothing AndAlso
                                   String.Equals(item.Action.ToolId, descriptor.Id, StringComparison.OrdinalIgnoreCase) AndAlso
                                   item.Explanation IsNot Nothing AndAlso item.Explanation.Success) Then
                    Return descriptor.Id
                End If
            Next
            Return ""
        End Function

        Private Shared Function ResolveCompletionEvidenceClaim(session As AgentSession,
                                                               evidenceById As IDictionary(Of String, OutcomeEvidenceRecord),
                                                               claim As String) As List(Of OutcomeEvidenceRecord)
            Dim normalized = If(claim, "").Trim()
            Dim exact As OutcomeEvidenceRecord = Nothing
            If normalized.Length > 0 AndAlso evidenceById.TryGetValue(normalized, exact) Then
                Return New List(Of OutcomeEvidenceRecord) From {exact}
            End If
            Dim iteration = If(session?.Iterations, New List(Of ReActIteration)()).FirstOrDefault(
                Function(item) item IsNot Nothing AndAlso
                               String.Equals(item.EvidenceId, normalized, StringComparison.OrdinalIgnoreCase))
            Return If(iteration?.ContractEvidence, New List(Of OutcomeEvidenceRecord)()).
                Where(Function(item) item IsNot Nothing AndAlso item.Satisfied).
                ToList()
        End Function

        Private Shared Function GetObservedDependencyClosure(session As AgentSession,
                                                             directDependencies As IEnumerable(Of String)) As HashSet(Of String)
            Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim pending As New Queue(Of String)(If(directDependencies, Enumerable.Empty(Of String)()).
                                                Where(Function(value) Not String.IsNullOrWhiteSpace(value)))
            While pending.Count > 0
                Dim dependency = pending.Dequeue()
                If Not visited.Add(dependency) Then Continue While
                Dim iteration = If(session?.Iterations, New List(Of ReActIteration)()).FirstOrDefault(
                    Function(item) item IsNot Nothing AndAlso
                                   String.Equals(item.EvidenceId, dependency, StringComparison.OrdinalIgnoreCase))
                If iteration Is Nothing Then Continue While
                For Each parent In If(iteration.DependsOnEvidenceIds, New List(Of String)())
                    If Not String.IsNullOrWhiteSpace(parent) Then pending.Enqueue(parent)
                Next
            End While
            Return visited
        End Function

        Private Shared Function FindComputeSourceReadEvidence(session As AgentSession,
                                                              outputRecord As OutcomeEvidenceRecord,
                                                              capability As String) As List(Of OutcomeEvidenceRecord)
            Dim outputAncestors = GetObservedDependencyClosure(session, outputRecord?.DerivedFromEvidenceIds)
            Dim computeIteration = If(session?.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item IsNot Nothing AndAlso outputAncestors.Contains(item.EvidenceId) AndAlso
                                     item.Action IsNot Nothing AndAlso
                                     String.Equals(item.Action.ToolId, capability, StringComparison.OrdinalIgnoreCase) AndAlso
                                     item.Explanation IsNot Nothing AndAlso item.Explanation.Success).
                OrderByDescending(Function(item) item.Index).
                FirstOrDefault()
            If computeIteration Is Nothing Then Return New List(Of OutcomeEvidenceRecord)()
            Dim computeAncestors = GetObservedDependencyClosure(session, computeIteration.DependsOnEvidenceIds)
            Return If(session?.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item IsNot Nothing AndAlso computeAncestors.Contains(item.EvidenceId)).
                SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                Where(Function(item) item IsNot Nothing AndAlso item.Satisfied AndAlso
                                     String.Equals(item.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase)).
                GroupBy(Function(item) item.EvidenceId, StringComparer.OrdinalIgnoreCase).
                Select(Function(group) group.First()).
                ToList()
        End Function

        Private Shared Function ResolveGoalCriterionStatement(session As AgentSession,
                                                              criterionId As String) As String
            Dim criterion = session?.Spec?.GoalContract?.Criteria?.FirstOrDefault(
                Function(item) item IsNot Nothing AndAlso
                               String.Equals(item.Id, criterionId, StringComparison.OrdinalIgnoreCase))
            If criterion IsNot Nothing Then Return If(criterion.Statement, criterionId)
            Dim constraint = session?.Spec?.GoalContract?.Constraints?.FirstOrDefault(
                Function(item) item IsNot Nothing AndAlso
                               String.Equals(item.Id, criterionId, StringComparison.OrdinalIgnoreCase))
            Return If(constraint?.Statement, criterionId)
        End Function

        Private Shared Sub AddUniqueClaim(claims As IList(Of String),
                                          evidenceId As String)
            If claims Is Nothing OrElse String.IsNullOrWhiteSpace(evidenceId) Then Return
            If Not claims.Contains(evidenceId, StringComparer.OrdinalIgnoreCase) Then claims.Add(evidenceId)
        End Sub

        ''' <summary>
        ''' Grounds a model-selected verification projection in the exact typed host evidence
        ''' cited by the completion decision. The model owns the semantic mapping from a Goal
        ''' criterion to evidence; it does not get to re-describe the observed target or value.
        ''' This prevents worksheet-root guesses, invented output shapes, and provider-specific
        ''' contract prose from turning an already successful task into a completion retry loop.
        ''' </summary>
        Private Function GroundCompletionProjectionInEvidence(session As AgentSession,
                                                               contract As OutcomeContract,
                                                               evidenceClaims As IList(Of String)) As String
            If session Is Nothing OrElse contract Is Nothing OrElse contract.Requirements Is Nothing Then
                Return "Completion projection is missing."
            End If

            Dim claims = New HashSet(Of String)(
                If(evidenceClaims, Enumerable.Empty(Of String)()).
                    Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                    Select(Function(value) value.Trim()),
                StringComparer.OrdinalIgnoreCase)
            Dim restrictToModelClaims = claims.Count > 0
            Dim activeWorkbookName = AgentGoalVerifier.ResolveContextWorkbookName(
                TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack))

            Dim availableEvidence = If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item IsNot Nothing).
                SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                Where(Function(record) record IsNot Nothing AndAlso record.Satisfied).
                ToList()

            For Each requirement In contract.Requirements.
                Where(Function(item) item IsNot Nothing AndAlso item.Required)
                If Not restrictToModelClaims AndAlso
                   requirement.CriterionIds IsNot Nothing AndAlso requirement.CriterionIds.Count > 0 AndAlso
                   String.IsNullOrWhiteSpace(requirement.DerivedFromCapability) Then
                    Return $"Requirement {requirement.Id} has no trusted producer binding; " &
                        "the next adaptive decision must select evidence semantically."
                End If
                Dim requiredTarget = AgentGoalVerifier.BindActiveWorkbookReference(
                    requirement.TargetRef,
                    activeWorkbookName)
                Dim candidates = availableEvidence.Where(
                    Function(record)
                        If restrictToModelClaims AndAlso Not claims.Contains(record.EvidenceId) Then Return False
                        If Not String.Equals(record.EffectType,
                                             requirement.EffectType,
                                             StringComparison.OrdinalIgnoreCase) Then Return False
                        If Not AgentGoalVerifier.TargetsOverlap(requiredTarget, record.TargetRef) Then Return False
                        Dim producer = If(requirement.DerivedFromCapability, "").Trim()
                        Return producer.Length = 0 OrElse
                            String.Equals(record.SourceToolId, producer, StringComparison.OrdinalIgnoreCase) OrElse
                            String.Equals(requirement.EffectType, "data_state", StringComparison.OrdinalIgnoreCase)
                    End Function).
                    ToList()

                candidates = NarrowCompletionEvidenceCandidates(requirement, candidates)

                If candidates.Count = 0 Then
                    Return $"No current host evidence matches requirement {requirement.Id} " &
                        $"(effect={requirement.EffectType}, target={requiredTarget})."
                End If
                If candidates.Count > 1 Then
                    Dim latestRevision = candidates.Max(Function(item) item.WorldRevision)
                    candidates = candidates.Where(Function(item) item.WorldRevision = latestRevision).ToList()
                    If candidates.Count > 1 Then
                        Return $"Host evidence is ambiguous for requirement {requirement.Id}: " &
                            String.Join(", ", candidates.Select(Function(item) item.EvidenceId))
                    End If
                End If

                Dim evidence = candidates(0)
                If evidenceClaims IsNot Nothing AndAlso Not claims.Contains(evidence.EvidenceId) Then
                    evidenceClaims.Add(evidence.EvidenceId)
                    claims.Add(evidence.EvidenceId)
                End If
                requirement.TargetRef = evidence.TargetRef
                requirement.PropertyName = If(evidence.PropertyName, "").Trim()

                Dim effect = If(requirement.EffectType, "").Trim().ToLowerInvariant()
                Select Case effect
                    Case "object_exists", "object_absent"
                        ' Host adapters may verify creation/deletion through a concrete member
                        ' such as Worksheet.Name. That member is evidence metadata, not part of
                        ' the lifecycle assertion: existence and property state are deliberately
                        ' separate contract slots.
                        requirement.PropertyName = ""
                        requirement.Operator = "exists"
                        requirement.ExpectedValue = Nothing
                    Case "artifact"
                        requirement.Operator = "exists"
                        requirement.ExpectedValue = Nothing
                    Case "read_coverage"
                        requirement.Operator = "covers"
                    Case "compute_artifact"
                        ' Compute artifacts are verified through their typed dependency lineage.
                    Case Else
                        Dim observedValue = OutcomeProjectionValue.ClonePropertyValue(
                            evidence.Actual,
                            requirement.PropertyName)
                        If observedValue Is Nothing AndAlso evidence.RequestVerified Then
                            observedValue = OutcomeProjectionValue.ClonePropertyValue(
                                evidence.VerifiedRequest,
                                requirement.PropertyName)
                        End If
                        If observedValue Is Nothing Then Continue For
                        requirement.Operator = "equals"
                        requirement.ExpectedValue = observedValue
                End Select
            Next

            Return ""
        End Function

        ''' <summary>
        ''' Selects the host assertion that corresponds to one provisional requirement without
        ''' relying on task wording.  A single tool call can emit several assertions against the
        ''' same object (for example chart type, title and legend position); target/effect alone
        ''' is therefore not a sufficient identity.  Prefer the exact host property, then a
        ''' verified public-request alias, then an exact expected value.  Revision remains only
        ''' a recency tie-breaker.
        ''' </summary>
        Private Shared Function NarrowCompletionEvidenceCandidates(
            requirement As OutcomeRequirement,
            candidates As List(Of OutcomeEvidenceRecord)) As List(Of OutcomeEvidenceRecord)

            If requirement Is Nothing OrElse candidates Is Nothing OrElse candidates.Count <= 1 Then
                Return If(candidates, New List(Of OutcomeEvidenceRecord)())
            End If

            Dim scored = candidates.
                Select(Function(record) New With {
                    .Record = record,
                    .Score = CompletionEvidenceMatchScore(requirement, record)
                }).
                ToList()
            Dim bestScore = scored.Max(Function(item) item.Score)
            If bestScore <= 0 Then Return candidates
            Return scored.
                Where(Function(item) item.Score = bestScore).
                Select(Function(item) item.Record).
                ToList()
        End Function

        Private Shared Function CompletionEvidenceMatchScore(requirement As OutcomeRequirement,
                                                               record As OutcomeEvidenceRecord) As Integer
            If requirement Is Nothing OrElse record Is Nothing Then Return 0
            Dim score As Integer = 0
            Dim propertyName = If(requirement.PropertyName, "").Trim()
            If Not String.IsNullOrWhiteSpace(propertyName) Then
                If String.Equals(record.PropertyName, propertyName, StringComparison.OrdinalIgnoreCase) Then
                    score += 100
                ElseIf ObjectContainsProperty(record.VerifiedRequest, propertyName) Then
                    score += 90
                ElseIf ObjectContainsProperty(record.Actual, propertyName) Then
                    score += 80
                End If
            End If

            If requirement.ExpectedValue IsNot Nothing AndAlso
               requirement.ExpectedValue.Type <> JTokenType.Null Then
                If CompletionEvidenceValueMatches(record.VerifiedRequest,
                                                  propertyName,
                                                  requirement.ExpectedValue) Then score += 40
                If CompletionEvidenceValueMatches(record.Actual,
                                                  If(String.IsNullOrWhiteSpace(propertyName), record.PropertyName, propertyName),
                                                  requirement.ExpectedValue) Then score += 30
            End If
            Return score
        End Function

        Private Shared Function ObjectContainsProperty(value As JToken,
                                                       propertyName As String) As Boolean
            Dim obj = TryCast(value, JObject)
            Return obj IsNot Nothing AndAlso
                obj.GetValue(propertyName, StringComparison.OrdinalIgnoreCase) IsNot Nothing
        End Function

        Private Shared Function CompletionEvidenceValueMatches(value As JToken,
                                                               propertyName As String,
                                                               expected As JToken) As Boolean
            If value Is Nothing OrElse expected Is Nothing Then Return False
            Dim candidate = value
            Dim obj = TryCast(value, JObject)
            If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(propertyName) Then
                candidate = obj.GetValue(propertyName, StringComparison.OrdinalIgnoreCase)
                If candidate Is Nothing Then Return False
            End If
            Return JToken.DeepEquals(candidate, expected)
        End Function

        Private Function SealVerificationProjection(session As AgentSession,
                                                     plan As ExecutionPlan,
                                                     executionAppType As String) As String
            If session?.Spec Is Nothing Then Return "任务规格缺失，无法密封验收投影。"
            Dim existingContract = session.Spec.OutcomeContract
            If existingContract IsNot Nothing AndAlso existingContract.Frozen Then
                Return Goals.GoalOutcomeProjection.ValidateIntegrity(session.Spec, existingContract)
            End If

            ' Validate the planner proposal as a local candidate. A rejected projection must
            ' not become session state and poison the next bounded repair attempt.
            Dim contract As OutcomeContract = Nothing
            If plan IsNot Nothing AndAlso plan.OutcomeContract IsNot Nothing Then
                contract = plan.OutcomeContract
            Else
                contract = existingContract
            End If
            If contract Is Nothing OrElse contract.Requirements Is Nothing OrElse contract.Requirements.Count = 0 Then
                If String.Equals(executionAppType, "Excel", StringComparison.OrdinalIgnoreCase) Then
                    Return "规划失败：模型未生成结构化结果合同；为避免局部成功被误报为任务完成，未执行任何 Office 写入。"
                End If
                Return ""
            End If
            contract.ValidatedComputeCapabilities.Clear()
            contract.ValidatedProducerCapabilities.Clear()

            Dim activeWorkbookName = AgentGoalVerifier.ResolveContextWorkbookName(
                TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack))
            If String.Equals(executionAppType, "Excel", StringComparison.OrdinalIgnoreCase) Then
                contract.BoundWorkbook = activeWorkbookName
                For Each requirement In contract.Requirements.Where(Function(item) item IsNot Nothing)
                    requirement.TargetRef = AgentGoalVerifier.BindActiveWorkbookReference(
                        requirement.TargetRef,
                        activeWorkbookName)
                Next
            End If

            Dim allowedEffects As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each descriptor In _toolRegistry.GetAvailableTools(executionAppType)
                For Each effect In OutcomeEffectCatalog.GetEffects(descriptor)
                    allowedEffects.Add(effect)
                Next
            Next
            Dim ids As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim allowedOperators As New HashSet(Of String)(
                {"equals", "contains", "covers", "exists"},
                StringComparer.OrdinalIgnoreCase)
            For Each requirement In contract.Requirements.Where(Function(item) item IsNot Nothing AndAlso item.Required)
                If String.IsNullOrWhiteSpace(requirement.Id) OrElse Not ids.Add(requirement.Id) Then
                    Return "规划失败：结果合同 requirement.id 缺失或重复。"
                End If
                If String.IsNullOrWhiteSpace(requirement.EffectType) OrElse
                   Not allowedEffects.Contains(requirement.EffectType) Then
                    Return $"规划失败：结果合同 {requirement.Id} 使用了未注册的 effectType {requirement.EffectType}。"
                End If
                If String.Equals(requirement.EffectType, "unclassified_mutation", StringComparison.OrdinalIgnoreCase) Then
                    Return $"规划失败：结果合同 {requirement.Id} 使用了仅供失效传播的 unclassified_mutation；它不能作为任务完成证明。"
                End If
                If String.IsNullOrWhiteSpace(requirement.TargetRef) AndAlso
                   Not String.Equals(requirement.EffectType, "compute_artifact", StringComparison.OrdinalIgnoreCase) Then
                    Return $"规划失败：结果合同 {requirement.Id} 缺少目标对象引用。"
                End If
                Dim targetReferenceError = AgentGoalVerifier.ValidateContractTargetReference(
                    executionAppType,
                    requirement.TargetRef)
                If Not String.IsNullOrWhiteSpace(targetReferenceError) Then
                    Return $"规划失败：结果合同 {requirement.Id} 的目标无效：{targetReferenceError}"
                End If
                If Not String.IsNullOrWhiteSpace(requirement.AppType) AndAlso
                   Not String.Equals(requirement.AppType, executionAppType, StringComparison.OrdinalIgnoreCase) Then
                    Return $"规划失败：结果合同 {requirement.Id} 的应用类型与当前宿主不一致。"
                End If
                If String.IsNullOrWhiteSpace(requirement.Operator) OrElse
                   Not allowedOperators.Contains(requirement.Operator) Then
                    Return $"规划失败：结果合同 {requirement.Id} 使用了不支持的 operator {requirement.Operator}。"
                End If
                If String.Equals(requirement.Operator, "covers", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(requirement.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase) Then
                    Return $"规划失败：结果合同 {requirement.Id} 只有 read_coverage 才能使用 covers。"
                End If
                If String.Equals(requirement.Operator, "exists", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(requirement.EffectType, "object_exists", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(requirement.EffectType, "object_absent", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(requirement.EffectType, "artifact", StringComparison.OrdinalIgnoreCase) Then
                    Return $"规划失败：结果合同 {requirement.Id} 用 exists 弱化了具体状态断言；exists 只允许对象生命周期或 artifact。"
                End If
                If String.Equals(requirement.Operator, "contains", StringComparison.OrdinalIgnoreCase) AndAlso
                   IsVacuousContractExpected(requirement.ExpectedValue) Then
                    Return $"规划失败：结果合同 {requirement.Id} 的 contains expectedValue 为空，无法证明任何具体结果。"
                End If
                If String.Equals(requirement.EffectType, "property_state", StringComparison.OrdinalIgnoreCase) AndAlso
                   String.IsNullOrWhiteSpace(requirement.PropertyName) Then
                    Return $"规划失败：结果合同 {requirement.Id} 的 property_state 缺少 property。"
                End If
                If (String.Equals(requirement.EffectType, "object_exists", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(requirement.EffectType, "object_absent", StringComparison.OrdinalIgnoreCase)) AndAlso
                   (Not String.IsNullOrWhiteSpace(requirement.PropertyName) OrElse
                    (requirement.ExpectedValue IsNot Nothing AndAlso requirement.ExpectedValue.Type <> JTokenType.Null)) Then
                    Return $"规划失败：结果合同 {requirement.Id} 把对象存在性和对象属性混为一条断言；存在/不存在 requirement 不能携带 property 或 expectedValue。"
                End If
                If Not String.IsNullOrWhiteSpace(requirement.DerivedFromCapability) Then
                    Dim producerDescriptor = _toolRegistry.GetTool(requirement.DerivedFromCapability)
                    Dim producerEffects = OutcomeEffectCatalog.GetEffects(producerDescriptor)
                    Dim isComputeProducer = producerDescriptor IsNot Nothing AndAlso
                        String.Equals(producerDescriptor.AccessMode, "compute", StringComparison.OrdinalIgnoreCase)
                    Dim isDirectEffectProducer = producerDescriptor IsNot Nothing AndAlso
                        producerEffects.Contains(requirement.EffectType, StringComparer.OrdinalIgnoreCase)
                    If Not isComputeProducer AndAlso Not isDirectEffectProducer Then
                        Return $"验收投影无效：{requirement.Id} 的 derivedFromCapability={requirement.DerivedFromCapability} 未注册，或不能直接产生 effectType={requirement.EffectType}。"
                    End If
                    If Not contract.ValidatedProducerCapabilities.Contains(
                        producerDescriptor.Id,
                        StringComparer.OrdinalIgnoreCase) Then
                        contract.ValidatedProducerCapabilities.Add(producerDescriptor.Id)
                    End If
                    If isComputeProducer AndAlso Not contract.ValidatedComputeCapabilities.Contains(
                        producerDescriptor.Id,
                        StringComparer.OrdinalIgnoreCase) Then
                        contract.ValidatedComputeCapabilities.Add(producerDescriptor.Id)
                    End If
                End If
                Dim requiresExpectedValue =
                    String.Equals(requirement.EffectType, "property_state", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(requirement.EffectType, "formula_state", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(requirement.EffectType, "order_state", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(requirement.EffectType, "filter_state", StringComparison.OrdinalIgnoreCase) OrElse
                    (String.Equals(requirement.EffectType, "data_state", StringComparison.OrdinalIgnoreCase) AndAlso
                     String.IsNullOrWhiteSpace(requirement.DerivedFromCapability))
                If requiresExpectedValue AndAlso
                   (requirement.ExpectedValue Is Nothing OrElse requirement.ExpectedValue.Type = JTokenType.Null) Then
                    Return $"规划失败：结果合同 {requirement.Id} 缺少可验证的 expectedValue。"
                End If
            Next

            Dim bindingError = Goals.GoalOutcomeProjection.ValidateBinding(
                session.Spec,
                contract,
                requireFrozenGoalBinding:=False)
            If Not String.IsNullOrWhiteSpace(bindingError) Then
                Return "规划失败：" & bindingError
            End If

            Dim expectedCriterionIds = New HashSet(Of String)(
                Goals.GoalOutcomeProjection.RequiredHostCriterionIds(session.Spec),
                StringComparer.OrdinalIgnoreCase)
            If expectedCriterionIds.Count > 0 Then
                Dim assertionSlotOwners As New Dictionary(Of String, String)(StringComparer.Ordinal)
                For Each requirement In contract.Requirements.Where(Function(item) item IsNot Nothing AndAlso item.Required)
                    If requirement.CriterionIds IsNot Nothing AndAlso requirement.CriterionIds.Count = 1 Then
                        Dim assertionSignature = BuildOutcomeAssertionSlot(requirement, executionAppType)
                        Dim existingOwner As String = Nothing
                        If assertionSlotOwners.TryGetValue(assertionSignature, existingOwner) Then
                            Return $"规划失败：结果合同 {requirement.Id} 与另一条成功标准复用了同一宿主断言槽位（{existingOwner}）。请拆分目标范围或属性，使独立成功标准由不同的可观察事实证明。"
                        End If
                        assertionSlotOwners(assertionSignature) = requirement.CriterionIds(0)
                    End If
                Next
            End If

            Dim isReadOnly = String.Equals(session.Spec.MutationPolicy, "read_only", StringComparison.OrdinalIgnoreCase)
            If isReadOnly AndAlso Not contract.Requirements.Any(
                Function(item) item.Required AndAlso String.Equals(item.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase)) Then
                Return "规划失败：只读任务的结果合同缺少 read_coverage。"
            End If
            If Not isReadOnly AndAlso Not contract.Requirements.Any(
                Function(item) item.Required AndAlso
                    Not String.Equals(item.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase) AndAlso
                    Not String.Equals(item.EffectType, "compute_artifact", StringComparison.OrdinalIgnoreCase)) Then
                Return "规划失败：修改任务的结果合同没有声明最终可观察状态。"
            End If

            For Each capability In AgentExecutionContract.ResolveRequiredCapabilities(session.Spec)
                Dim descriptor = _toolRegistry.GetTool(capability)
                If descriptor Is Nothing OrElse
                   Not String.Equals(descriptor.AccessMode, "compute", StringComparison.OrdinalIgnoreCase) OrElse isReadOnly Then Continue For
                If Not contract.Requirements.Any(
                    Function(item) item.Required AndAlso
                        String.Equals(item.DerivedFromCapability, capability, StringComparison.OrdinalIgnoreCase) AndAlso
                        Not String.Equals(item.EffectType, "compute_artifact", StringComparison.OrdinalIgnoreCase)) Then
                    Return $"规划失败：用户要求使用 {capability}，但结果合同没有声明最终产物来源；不能只执行计算或创建空对象后宣称完成。"
                End If
                If Not contract.Requirements.Any(
                    Function(item) item.Required AndAlso
                        String.Equals(item.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase)) Then
                    Return $"规划失败：结果合同要求最终产物来自 {capability}，但没有声明其输入数据的 read_coverage；无法建立完整数据血缘。"
                End If
            Next

            Goals.GoalOutcomeProjection.Seal(session.Spec, contract)
            session.Spec.OutcomeContract = contract
            If plan IsNot Nothing Then plan.OutcomeContract = contract
            Return ""
        End Function

        ''' <summary>
        ''' Planner lineage is a hint about how an observed result was produced. A registered
        ''' compute tool may supply transitive lineage through a later write; a registered host
        ''' tool may name itself only when it directly emits the requirement's effect type.
        ''' Everything else is removed before deterministic validation.
        ''' </summary>
        Private Sub NormalizePlannerLineageHints(contract As OutcomeContract,
                                                 executionAppType As String)
            If contract?.Requirements Is Nothing Then Return

            For Each requirement In contract.Requirements.Where(Function(item) item IsNot Nothing)
                Dim proposedCapability = If(requirement.DerivedFromCapability, "").Trim()
                If String.IsNullOrWhiteSpace(proposedCapability) Then Continue For

                Dim descriptor = _toolRegistry.GetTool(proposedCapability)
                Dim isCompute = descriptor IsNot Nothing AndAlso
                    String.Equals(descriptor.AccessMode, "compute", StringComparison.OrdinalIgnoreCase)
                Dim directlyProducesEffect = descriptor IsNot Nothing AndAlso
                    OutcomeEffectCatalog.GetEffects(descriptor).
                        Contains(requirement.EffectType, StringComparer.OrdinalIgnoreCase)
                If Not isCompute AndAlso Not directlyProducesEffect Then
                    requirement.DerivedFromCapability = ""
                    AppLogger.Warn(
                        "LoopEngine",
                        $"Removed untrusted planner lineage hint before validation; requirement={AppLogger.Redact(requirement.Id)}; capability={AppLogger.Redact(proposedCapability)}; app={AppLogger.Redact(executionAppType)}")
                    Continue For
                End If

                requirement.DerivedFromCapability = descriptor.Id
            Next
        End Sub

        Private Shared Function IsVacuousContractExpected(value As JToken) As Boolean
            If value Is Nothing OrElse value.Type = JTokenType.Null OrElse value.Type = JTokenType.Undefined Then Return True
            If value.Type = JTokenType.Object Then Return Not DirectCast(value, JObject).HasValues
            If value.Type = JTokenType.Array Then Return DirectCast(value, JArray).Count = 0
            If value.Type = JTokenType.String Then Return String.IsNullOrEmpty(value.Value(Of String)())
            Return False
        End Function

        Private Shared Function BuildOutcomeAssertionSlot(requirement As OutcomeRequirement,
                                                          executionAppType As String) As String
            If requirement Is Nothing Then Return ""
            Dim effectiveAppType = If(requirement.AppType, "").Trim()
            If String.IsNullOrWhiteSpace(effectiveAppType) Then effectiveAppType = If(executionAppType, "").Trim()
            Dim objectLifecycle = String.Equals(requirement.EffectType, "object_exists", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(requirement.EffectType, "object_absent", StringComparison.OrdinalIgnoreCase)
            Return String.Join("|", {
                effectiveAppType.ToLowerInvariant(),
                AgentGoalVerifier.CanonicalContractTargetReference(requirement.TargetRef),
                If(requirement.EffectType, "").Trim().ToLowerInvariant(),
                If(objectLifecycle, "", If(requirement.PropertyName, "").Trim().ToLowerInvariant()),
                If(requirement.DerivedFromCapability, "").Trim().ToLowerInvariant()
            })
        End Function

    End Class

End Namespace
