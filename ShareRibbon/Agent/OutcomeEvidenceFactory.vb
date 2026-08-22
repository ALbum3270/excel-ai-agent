Imports System.Collections.Generic
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Converts a runtime action and its host observation into the canonical evidence
    ''' consumed by AgentGoalVerifier. Failed/partial writes may emit invalidation-only
    ''' records: they can revoke stale proof but can never prove task completion.
    ''' </summary>
    Public NotInheritable Class OutcomeEvidenceFactory
        Private Sub New()
        End Sub

        Public Shared Function Create(tool As ToolDescriptor,
                                      toolCall As ToolCall,
                                      result As ToolResult,
                                      iterationEvidenceId As String,
                                      dependencies As IEnumerable(Of String),
                                      Optional worldRevision As Long = 0,
                                      Optional activeWorkbookName As String = "") As List(Of OutcomeEvidenceRecord)
            Dim records As New List(Of OutcomeEvidenceRecord)()
            If tool Is Nothing OrElse toolCall Is Nothing OrElse result Is Nothing Then Return records

            Dim observation = ToObject(result.Observation)
            BindObservationReferences(observation, activeWorkbookName)
            Dim accessMode = If(tool.AccessMode, "write").Trim().ToLowerInvariant()
            Dim writeExpectedToken = observation?("writeExpected")
            If String.Equals(tool.Id, "OfficeObjectOperation", StringComparison.OrdinalIgnoreCase) AndAlso
               writeExpectedToken IsNot Nothing AndAlso
               writeExpectedToken.Type = JTokenType.Boolean AndAlso
               Not writeExpectedToken.Value(Of Boolean)() Then accessMode = "read"
            Dim mutationObserved = accessMode <> "read" AndAlso accessMode <> "compute" AndAlso
                ObservationConfirmsChange(observation)
            Dim uncertainMutation = FailureMayHaveMutated(tool, result)
            Dim expected = If(toolCall.Parameters?.DeepClone(), New JObject())
            Dim dataHash = ComputeHash(result.Data)
            Dim dependencyList = If(dependencies, Enumerable.Empty(Of String)()).
                Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            ' A verification array is a set of atomic host assertions.  Preserve each
            ' (effect, target, property, actual) tuple instead of multiplying batch-level
            ' effects by batch-level targets.  The latter can turn "delete Chart1 and read
            ' Chart2.Name" into false evidence that Chart2 was deleted.
            Dim verificationItems = TryCast(observation?("verification"), JArray)
            If verificationItems IsNot Nothing AndAlso verificationItems.Count > 0 Then
                AppendAtomicVerificationEvidence(
                    records,
                    tool,
                    result,
                    observation,
                    verificationItems,
                    iterationEvidenceId,
                    expected,
                    dataHash,
                    dependencyList,
                    worldRevision,
                    accessMode,
                    mutationObserved,
                    activeWorkbookName)
                AppendArtifactAnchorEvidence(records, tool, result, observation, iterationEvidenceId,
                                             expected, dataHash, dependencyList, worldRevision,
                                             activeWorkbookName)
                AppendIdentityTransitionTombstones(records, observation, iterationEvidenceId,
                                                   tool.Id, worldRevision, activeWorkbookName)
                Dim preciseInvalidations = AppendExplicitInvalidationTombstones(
                    records, observation, iterationEvidenceId, tool.Id, worldRevision, activeWorkbookName)
                If uncertainMutation AndAlso preciseInvalidations = 0 Then
                    AddInvalidationRecord(records, iterationEvidenceId, tool.Id, worldRevision, "*")
                End If
                Return records
            End If

            Dim effects = ResolveVerifiedEffects(tool, observation)
            If effects.Count = 0 Then
                Dim preciseInvalidations = AppendExplicitInvalidationTombstones(
                    records, observation, iterationEvidenceId, tool.Id, worldRevision, activeWorkbookName)
                If (mutationObserved OrElse uncertainMutation) AndAlso preciseInvalidations = 0 Then
                    AddInvalidationRecord(records, iterationEvidenceId, tool.Id, worldRevision, "*")
                End If
                Return records
            End If
            Dim targets = ExtractTargetRefs(effects(0), observation, result.Data)
            targets = targets.Select(
                Function(value) AgentGoalVerifier.BindActiveWorkbookReference(value, activeWorkbookName)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
            If targets.Count = 0 Then
                If result.Success Then
                    targets.Add($"artifact:{iterationEvidenceId}")
                ElseIf mutationObserved OrElse uncertainMutation Then
                    ' A host-confirmed mutation with unknown scope must conservatively revoke
                    ' every older state assertion. It is never positive completion evidence.
                    targets.Add("*")
                Else
                    Return records
                End If
            End If
            Dim ordinal As Integer = 0
            For Each effect In effects
                For Each target In targets
                    Dim actual = ExtractVerifiedActual(accessMode, observation, result.Data, target, targets.Count)
                    Dim invalidationOnly = String.Equals(effect, "unclassified_mutation", StringComparison.OrdinalIgnoreCase)
                    Dim satisfied = Not invalidationOnly AndAlso result.Success AndAlso
                        IsHostSatisfied(accessMode, observation, target, targets.Count)
                    Dim verifiedRequest = ExtractVerifiedRequest(accessMode, observation, target, targets.Count)
                    Dim requestVerified = result.Success AndAlso verifiedRequest IsNot Nothing AndAlso
                        verifiedRequest.Type <> JTokenType.Null
                    Dim invalidatesPrior = accessMode <> "read" AndAlso accessMode <> "compute" AndAlso
                        (invalidationOnly OrElse mutationObserved OrElse
                         (Not satisfied AndAlso HasTargetObservation(observation, target, targets.Count)))
                    If Not satisfied AndAlso Not invalidatesPrior Then Continue For

                    ordinal += 1
                    records.Add(New OutcomeEvidenceRecord With {
                        .EvidenceId = $"{iterationEvidenceId}/e{ordinal}",
                        .IterationEvidenceId = iterationEvidenceId,
                        .TargetRef = target,
                        .EffectType = effect,
                        .PropertyName = If(satisfied,
                                           ExtractVerifiedPropertyName(observation, target, targets.Count),
                                           ""),
                        .Expected = expected.DeepClone(),
                        .Actual = actual?.DeepClone(),
                        .Satisfied = satisfied,
                        .InvalidatesPrior = invalidatesPrior,
                        .RequestVerified = requestVerified,
                        .VerifiedRequest = verifiedRequest?.DeepClone(),
                        .SourceToolId = tool.Id,
                        .DataHash = dataHash,
                        .WorldRevision = worldRevision,
                        .DerivedFromEvidenceIds = New List(Of String)(dependencyList)
                    })
                Next
            Next

            AppendIdentityTransitionTombstones(
                records,
                observation,
                iterationEvidenceId,
                tool.Id,
                worldRevision,
                activeWorkbookName)
            AppendArtifactAnchorEvidence(records, tool, result, observation, iterationEvidenceId,
                                         expected, dataHash, dependencyList, worldRevision,
                                         activeWorkbookName)
            Dim explicitInvalidationCount = AppendExplicitInvalidationTombstones(
                records, observation, iterationEvidenceId, tool.Id, worldRevision, activeWorkbookName)
            If uncertainMutation AndAlso explicitInvalidationCount = 0 Then
                AddInvalidationRecord(records, iterationEvidenceId, tool.Id, worldRevision, "*")
            End If

            Return records
        End Function

        ''' <summary>
        ''' True when a mutating tool returned a host state observation that must advance the
        ''' evidence timeline. Synthetic validation/approval failures do not advance it.
        ''' </summary>
        Public Shared Function ObservationAdvancesWorld(tool As ToolDescriptor,
                                                        result As ToolResult) As Boolean
            If tool Is Nothing OrElse result Is Nothing Then Return False
            Dim accessMode = If(tool.AccessMode, "write").Trim().ToLowerInvariant()
            If accessMode = "read" OrElse accessMode = "compute" Then Return False
            If FailureMayHaveMutated(tool, result) Then Return True
            Dim observation = ToObject(result.Observation)
            If observation Is Nothing Then Return False
            Dim writeExpectedToken = observation("writeExpected")
            If String.Equals(tool.Id, "OfficeObjectOperation", StringComparison.OrdinalIgnoreCase) AndAlso
               writeExpectedToken IsNot Nothing AndAlso
               writeExpectedToken.Type = JTokenType.Boolean AndAlso
               Not writeExpectedToken.Value(Of Boolean)() Then Return False
            If ObservationConfirmsChange(observation) Then Return True
            If observation("verification") IsNot Nothing OrElse observation("after") IsNot Nothing Then Return True
            If HasObservationItems(observation, "invalidationRefs") OrElse
               HasObservationItems(observation, "invalidatedTargetRefs") OrElse
               HasObservationItems(observation, "artifactAnchors") Then Return True
            Dim satisfied = observation("satisfied")
            Return result.Success AndAlso satisfied IsNot Nothing AndAlso satisfied.Type = JTokenType.Boolean
        End Function

        Private Shared Function HasObservationItems(observation As JObject,
                                                     propertyName As String) As Boolean
            If observation Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return False
            Dim token = observation(propertyName)
            If token Is Nothing OrElse token.Type = JTokenType.Null OrElse token.Type = JTokenType.Undefined Then Return False
            If token.Type = JTokenType.Array Then Return DirectCast(token, JArray).Count > 0
            If token.Type = JTokenType.Object Then Return DirectCast(token, JObject).HasValues
            Return Not String.IsNullOrWhiteSpace(token.ToString())
        End Function

        Private Shared Function FailureMayHaveMutated(tool As ToolDescriptor,
                                                      result As ToolResult) As Boolean
            If tool Is Nothing OrElse result Is Nothing OrElse result.Success Then Return False
            Dim accessMode = If(tool.AccessMode, "write").Trim().ToLowerInvariant()
            If accessMode = "read" OrElse accessMode = "compute" Then Return False

            ' Composite tools may have committed an earlier sub-batch before a later batch
            ' failed.  Their adapter has the only complete call-level view, so an explicit
            ' mayHaveMutated observation takes precedence over the final sub-batch error code.
            Dim observation = ToObject(result.Observation)
            Dim mayHaveMutated = observation?("mayHaveMutated")
            If mayHaveMutated IsNot Nothing AndAlso
               mayHaveMutated.Type = JTokenType.Boolean AndAlso
               mayHaveMutated.Value(Of Boolean)() Then Return True

            ' These failures are produced before dispatch reaches a mutating host member. All
            ' other write failures are treated as possibly applied and invalidate stale proof.
            Select Case If(result.ErrorCode, "").Trim().ToUpperInvariant()
                Case ExceptionClassifier.CodeToolNotAllowed,
                     ExceptionClassifier.CodeSafetyBlocked,
                     ExceptionClassifier.CodeSafetyNeedsApproval,
                     ExceptionClassifier.CodeApprovalUnavailable,
                     ExceptionClassifier.CodeOperationSchemaInvalid,
                     ExceptionClassifier.CodeObjectRefInvalid,
                     ExceptionClassifier.CodeObjectNotFound,
                     ExceptionClassifier.CodeCapabilityNotFound,
                     ExceptionClassifier.CodeMemberNotExecutable,
                     ExceptionClassifier.CodeArgument,
                     ExceptionClassifier.CodeNotFound,
                     ExceptionClassifier.CodeHostUnsupported,
                     ExceptionClassifier.CodeDocMissing,
                     ExceptionClassifier.CodeVbaDisabled
                    Return False
            End Select
            Return True
        End Function

        Private Shared Sub AddInvalidationRecord(records As List(Of OutcomeEvidenceRecord),
                                                 iterationEvidenceId As String,
                                                 toolId As String,
                                                 worldRevision As Long,
                                                 targetRef As String)
            If records Is Nothing Then Return
            If records.Any(Function(item) item IsNot Nothing AndAlso item.InvalidatesPrior AndAlso
                                          String.Equals(item.EffectType, "unclassified_mutation", StringComparison.OrdinalIgnoreCase) AndAlso
                                          String.Equals(item.TargetRef, targetRef, StringComparison.OrdinalIgnoreCase)) Then Return
            records.Add(New OutcomeEvidenceRecord With {
                .EvidenceId = $"{iterationEvidenceId}/e{records.Count + 1}",
                .IterationEvidenceId = iterationEvidenceId,
                .TargetRef = If(String.IsNullOrWhiteSpace(targetRef), "*", targetRef.Trim()),
                .EffectType = "unclassified_mutation",
                .PropertyName = "",
                .Satisfied = False,
                .InvalidatesPrior = True,
                .RequestVerified = False,
                .SourceToolId = If(toolId, ""),
                .WorldRevision = worldRevision
            })
        End Sub

        Private Shared Sub AppendIdentityTransitionTombstones(records As List(Of OutcomeEvidenceRecord),
                                                              observation As JObject,
                                                              iterationEvidenceId As String,
                                                              toolId As String,
                                                              worldRevision As Long,
                                                              Optional activeWorkbookName As String = "")
            Dim operations = TryCast(observation?("operations"), JArray)
            If operations Is Nothing Then Return
            For Each operation In operations.OfType(Of JObject)()
                If Not String.Equals(operation("status")?.ToString(), "succeeded", StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim memberId = If(operation("memberId")?.ToString(), "")
                If memberId.IndexOf(".property.Name", StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
                Dim oldRef = If(operation("targetRef")?.ToString(), "").Trim()
                Dim newRef = If(operation("resultRef")?.ToString(), "").Trim()
                If String.IsNullOrWhiteSpace(oldRef) OrElse String.IsNullOrWhiteSpace(newRef) OrElse
                   String.Equals(NormalizeRef(oldRef), NormalizeRef(newRef), StringComparison.OrdinalIgnoreCase) Then Continue For
                AddInvalidationRecord(
                    records,
                    iterationEvidenceId,
                    toolId,
                    worldRevision,
                    AgentGoalVerifier.BindActiveWorkbookReference(oldRef, activeWorkbookName))
            Next
        End Sub

        Private Shared Sub AppendAtomicVerificationEvidence(records As List(Of OutcomeEvidenceRecord),
                                                             tool As ToolDescriptor,
                                                             result As ToolResult,
                                                             observation As JObject,
                                                             verification As JArray,
                                                             iterationEvidenceId As String,
                                                             expected As JToken,
                                                             dataHash As String,
                                                             dependencies As List(Of String),
                                                             worldRevision As Long,
                                                             accessMode As String,
                                                             mutationObserved As Boolean,
                                                             activeWorkbookName As String)
            If records Is Nothing OrElse tool Is Nothing OrElse verification Is Nothing Then Return
            Dim declaredEffects = OutcomeEffectCatalog.GetEffects(tool)
            Dim fallbackEffect = If(declaredEffects.Count = 1,
                                    declaredEffects(0),
                                    If(observation?("effectType")?.ToString(), ""))
            Dim fallbackTargets = ExtractTargetRefs(fallbackEffect, observation, result?.Data).
                Select(Function(value) AgentGoalVerifier.BindActiveWorkbookReference(value, activeWorkbookName)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
            Dim ambiguousUnscopedObservation As Boolean = False

            For Each item In verification.OfType(Of JObject)()
                Dim targetRef = If(item("targetRef")?.ToString(), "").Trim()
                If String.IsNullOrWhiteSpace(targetRef) Then
                    If fallbackTargets.Count <> 1 Then
                        ambiguousUnscopedObservation = True
                        Continue For
                    End If
                    targetRef = fallbackTargets(0)
                Else
                    targetRef = AgentGoalVerifier.BindActiveWorkbookReference(targetRef, activeWorkbookName)
                End If

                Dim effect = ResolveAtomicEffect(declaredEffects, item, observation)
                If String.Equals(accessMode, "read", StringComparison.OrdinalIgnoreCase) Then
                    effect = "read_coverage"
                End If
                If String.IsNullOrWhiteSpace(effect) Then effect = "unclassified_mutation"
                Dim invalidationOnly = String.Equals(effect, "unclassified_mutation", StringComparison.OrdinalIgnoreCase)
                Dim passed = String.Equals(item("status")?.ToString(), "passed", StringComparison.OrdinalIgnoreCase)
                Dim satisfied = Not invalidationOnly AndAlso result.Success AndAlso passed
                Dim hasObservation = item("actual") IsNot Nothing OrElse
                    Not String.IsNullOrWhiteSpace(item("status")?.ToString())
                Dim invalidatesPrior = accessMode <> "read" AndAlso accessMode <> "compute" AndAlso
                    (invalidationOnly OrElse mutationObserved OrElse (Not satisfied AndAlso hasObservation))
                If Not satisfied AndAlso Not invalidatesPrior Then Continue For

                Dim propertyName = If(item("property")?.ToString(), "").Trim()
                Dim actual = BuildAtomicActual(propertyName, item("actual"))
                Dim verifiedRequest As JToken = If(satisfied, BuildAtomicVerifiedRequest(item), Nothing)
                records.Add(New OutcomeEvidenceRecord With {
                    .EvidenceId = $"{iterationEvidenceId}/e{records.Count + 1}",
                    .IterationEvidenceId = iterationEvidenceId,
                    .TargetRef = targetRef,
                    .EffectType = effect,
                    .PropertyName = If(satisfied, propertyName, ""),
                    .Expected = expected?.DeepClone(),
                    .Actual = actual,
                    .Satisfied = satisfied,
                    .InvalidatesPrior = invalidatesPrior,
                    .RequestVerified = verifiedRequest IsNot Nothing,
                    .VerifiedRequest = verifiedRequest,
                    .SourceToolId = tool.Id,
                    .DataHash = dataHash,
                    .WorldRevision = worldRevision,
                    .DerivedFromEvidenceIds = New List(Of String)(dependencies)
                })
            Next

            If ambiguousUnscopedObservation AndAlso mutationObserved Then
                AddInvalidationRecord(records, iterationEvidenceId, tool.Id, worldRevision, "*")
            End If
        End Sub

        Private Shared Function ResolveAtomicEffect(declaredEffects As List(Of String),
                                                     item As JObject,
                                                     observation As JObject) As String
            If declaredEffects Is Nothing OrElse declaredEffects.Count = 0 Then Return ""
            ' Dedicated tools have one declared semantic outcome.  Their internal COM
            ' properties are implementation details (for example CreateSheet verifies Name
            ' while proving object_exists).  The generic bridge instead relies on each
            ' operation's host-inferred effect.
            If declaredEffects.Count = 1 Then Return declaredEffects(0)
            Dim itemEffect = If(item?("effectType")?.ToString(), "").Trim()
            If declaredEffects.Contains(itemEffect, StringComparer.OrdinalIgnoreCase) Then Return itemEffect
            Dim observedEffect = If(observation?("effectType")?.ToString(), "").Trim()
            If declaredEffects.Contains(observedEffect, StringComparer.OrdinalIgnoreCase) Then Return observedEffect
            Return ""
        End Function

        Private Shared Function BuildAtomicActual(propertyName As String, actual As JToken) As JToken
            Dim value = If(actual?.DeepClone(), JValue.CreateNull())
            If String.IsNullOrWhiteSpace(propertyName) Then Return value
            Dim state As New JObject()
            state(propertyName) = value
            Return state
        End Function

        Private Shared Function BuildAtomicVerifiedRequest(item As JObject) As JToken
            If item Is Nothing OrElse item("requestExpected") Is Nothing Then Return Nothing
            Dim requestProperty = If(item("requestProperty")?.ToString(), "").Trim()
            If String.IsNullOrWhiteSpace(requestProperty) Then Return item("requestExpected").DeepClone()
            Dim request As New JObject()
            request(requestProperty) = item("requestExpected").DeepClone()
            Return request
        End Function

        Private Shared Sub AppendArtifactAnchorEvidence(records As List(Of OutcomeEvidenceRecord),
                                                        tool As ToolDescriptor,
                                                        result As ToolResult,
                                                        observation As JObject,
                                                        iterationEvidenceId As String,
                                                        expected As JToken,
                                                        dataHash As String,
                                                        dependencies As List(Of String),
                                                        worldRevision As Long,
                                                        activeWorkbookName As String)
            If records Is Nothing OrElse tool Is Nothing OrElse result Is Nothing OrElse Not result.Success Then Return
            If HasFailedRequiredVerification(observation) Then Return
            If Not OutcomeEffectCatalog.GetEffects(tool).Contains("artifact", StringComparer.OrdinalIgnoreCase) Then Return
            Dim anchors = TryCast(observation?("artifactAnchors"), JArray)
            If anchors Is Nothing Then Return
            For Each anchor In anchors.OfType(Of JObject)()
                If Not String.Equals(anchor("status")?.ToString(), "passed", StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim targetRef = AgentGoalVerifier.BindActiveWorkbookReference(
                    If(anchor("targetRef")?.ToString(), ""), activeWorkbookName)
                Dim artifactRef = AgentGoalVerifier.BindActiveWorkbookReference(
                    If(anchor("artifactRef")?.ToString(), ""), activeWorkbookName)
                If String.IsNullOrWhiteSpace(targetRef) Then Continue For
                records.Add(New OutcomeEvidenceRecord With {
                    .EvidenceId = $"{iterationEvidenceId}/e{records.Count + 1}",
                    .IterationEvidenceId = iterationEvidenceId,
                    .TargetRef = targetRef,
                    .EffectType = "artifact",
                    .Expected = expected?.DeepClone(),
                    .Actual = If(String.IsNullOrWhiteSpace(artifactRef),
                                 anchor.DeepClone(),
                                 New JValue(artifactRef)),
                    .Satisfied = True,
                    .InvalidatesPrior = True,
                    .SourceToolId = tool.Id,
                    .DataHash = dataHash,
                    .WorldRevision = worldRevision,
                    .RelatedTargetRefs = If(String.IsNullOrWhiteSpace(artifactRef),
                                            New List(Of String)(),
                                            New List(Of String) From {artifactRef}),
                    .DerivedFromEvidenceIds = New List(Of String)(dependencies)
                })
            Next
        End Sub

        Private Shared Function HasFailedRequiredVerification(observation As JObject) As Boolean
            Dim verification = TryCast(observation?("verification"), JArray)
            If verification Is Nothing Then Return False
            Return verification.OfType(Of JObject)().Any(
                Function(item)
                    ' Match the executor's fail-closed contract: verification is required
                    ' unless the producer explicitly marks it optional.
                    Dim required = If(item("required")?.Value(Of Boolean?)(), True)
                    If Not required Then Return False
                    Return Not String.Equals(item("status")?.ToString(), "passed", StringComparison.OrdinalIgnoreCase)
                End Function)
        End Function

        Private Shared Function AppendExplicitInvalidationTombstones(records As List(Of OutcomeEvidenceRecord),
                                                                     observation As JObject,
                                                                     iterationEvidenceId As String,
                                                                     toolId As String,
                                                                     worldRevision As Long,
                                                                     activeWorkbookName As String) As Integer
            If observation Is Nothing Then Return 0
            Dim refs As New List(Of String)()
            AddRefs(refs, observation("invalidationRefs"))
            AddRefs(refs, observation("invalidatedTargetRefs"))
            Dim normalized = refs.
                Select(Function(value) AgentGoalVerifier.BindActiveWorkbookReference(value, activeWorkbookName)).
                Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
            For Each targetRef In normalized
                AddInvalidationRecord(records, iterationEvidenceId, toolId, worldRevision, targetRef)
            Next
            Return normalized.Count
        End Function

        Private Shared Sub BindObservationReferences(observation As JObject,
                                                     activeWorkbookName As String)
            If observation Is Nothing OrElse String.IsNullOrWhiteSpace(activeWorkbookName) Then Return
            BindReferenceArray(TryCast(observation("targetRefs"), JArray), activeWorkbookName)
            BindReferenceArray(TryCast(observation("invalidationRefs"), JArray), activeWorkbookName)
            BindReferenceArray(TryCast(observation("invalidatedTargetRefs"), JArray), activeWorkbookName)

            Dim verificationArray = TryCast(observation("verification"), JArray)
            If verificationArray IsNot Nothing Then
                For Each item In verificationArray.OfType(Of JObject)()
                    BindReferenceProperty(item, "targetRef", activeWorkbookName)
                Next
            End If
            Dim verificationObject = TryCast(observation("verification"), JObject)
            BindReferenceProperty(verificationObject, "targetRef", activeWorkbookName)

            Dim operations = TryCast(observation("operations"), JArray)
            If operations IsNot Nothing Then
                For Each operation In operations.OfType(Of JObject)()
                    BindReferenceProperty(operation, "targetRef", activeWorkbookName)
                    BindReferenceProperty(operation, "resultRef", activeWorkbookName)
                Next
            End If

            Dim anchors = TryCast(observation("artifactAnchors"), JArray)
            If anchors IsNot Nothing Then
                For Each anchor In anchors.OfType(Of JObject)()
                    BindReferenceProperty(anchor, "targetRef", activeWorkbookName)
                    BindReferenceProperty(anchor, "artifactRef", activeWorkbookName)
                Next
            End If
        End Sub

        Private Shared Sub BindReferenceArray(values As JArray,
                                              activeWorkbookName As String)
            If values Is Nothing Then Return
            For index = 0 To values.Count - 1
                If values(index)?.Type <> JTokenType.String Then Continue For
                values(index) = AgentGoalVerifier.BindActiveWorkbookReference(values(index).ToString(), activeWorkbookName)
            Next
        End Sub

        Private Shared Sub BindReferenceProperty(container As JObject,
                                                 propertyName As String,
                                                 activeWorkbookName As String)
            If container Is Nothing OrElse container(propertyName)?.Type <> JTokenType.String Then Return
            container(propertyName) = AgentGoalVerifier.BindActiveWorkbookReference(
                container(propertyName).ToString(),
                activeWorkbookName)
        End Sub

        Private Shared Function ResolveVerifiedEffects(tool As ToolDescriptor,
                                                       observation As JObject) As List(Of String)
            Dim declared = OutcomeEffectCatalog.GetEffects(tool)
            Dim observedEffect = If(observation?("effectType")?.ToString(), "").Trim()
            If Not String.IsNullOrWhiteSpace(observedEffect) AndAlso
               declared.Contains(observedEffect, StringComparer.OrdinalIgnoreCase) Then
                Return New List(Of String) From {observedEffect}
            End If

            ' Without a typed host effect, only an unambiguous capability declaration can
            ' become evidence. A multi-effect declaration would otherwise create a false
            ' Cartesian product of effects and targets.
            If declared.Count = 1 Then Return declared
            Return New List(Of String)()
        End Function

        Private Shared Function IsHostSatisfied(accessMode As String,
                                                observation As JObject,
                                                targetRef As String,
                                                resolvedTargetCount As Integer) As Boolean
            If accessMode = "read" OrElse accessMode = "compute" Then Return True
            If observation Is Nothing Then Return False

            Dim verification = TryCast(observation("verification"), JArray)
            If verification IsNot Nothing Then
                Dim relevant = GetRelevantVerificationItems(verification, targetRef, resolvedTargetCount)
                If relevant.Count = 0 Then Return False
                Return relevant.All(
                    Function(item) Not If(item("required")?.Value(Of Boolean)(), True) OrElse
                                   String.Equals(item("status")?.ToString(), "passed", StringComparison.OrdinalIgnoreCase))
            End If

            Dim verificationObject = TryCast(observation("verification"), JObject)
            If VerificationObjectApplies(verificationObject, targetRef, resolvedTargetCount) Then
                Dim verified = verificationObject("satisfied")
                Return verified IsNot Nothing AndAlso verified.Type = JTokenType.Boolean AndAlso verified.Value(Of Boolean)()
            End If

            Dim visualPassed = observation.SelectToken("visualVerification.passed")
            If resolvedTargetCount = 1 AndAlso visualPassed IsNot Nothing AndAlso visualPassed.Type = JTokenType.Boolean Then
                Return visualPassed.Value(Of Boolean)()
            End If

            Dim explicitSatisfied = observation("satisfied")
            If resolvedTargetCount = 1 AndAlso explicitSatisfied IsNot Nothing AndAlso explicitSatisfied.Type = JTokenType.Boolean Then
                Return explicitSatisfied.Value(Of Boolean)()
            End If

            ' changed is audit information only.  It cannot prove that the requested state
            ' was reached, and an idempotent satisfied operation may legitimately not change.
            Return False
        End Function

        Private Shared Function ExtractVerifiedRequest(accessMode As String,
                                                       observation As JObject,
                                                       targetRef As String,
                                                       resolvedTargetCount As Integer) As JToken
            If accessMode = "read" OrElse accessMode = "compute" OrElse observation Is Nothing Then Return Nothing

            Dim verificationArray = TryCast(observation("verification"), JArray)
            If verificationArray IsNot Nothing Then
                Dim claims As New JObject()
                For Each item In GetRelevantVerificationItems(verificationArray, targetRef, resolvedTargetCount)
                    If Not String.Equals(item("status")?.ToString(), "passed", StringComparison.OrdinalIgnoreCase) Then Continue For
                    Dim requestProperty = If(item("requestProperty")?.ToString(), "").Trim()
                    If String.IsNullOrWhiteSpace(requestProperty) OrElse item("requestExpected") Is Nothing Then Continue For
                    claims(requestProperty) = item("requestExpected").DeepClone()
                Next
                If claims.HasValues Then Return claims
                Return Nothing
            End If

            Dim verificationObject = TryCast(observation("verification"), JObject)
            If VerificationObjectApplies(verificationObject, targetRef, resolvedTargetCount) AndAlso
               verificationObject("satisfied")?.Value(Of Boolean)() AndAlso
               verificationObject("requestExpected") IsNot Nothing Then
                Return verificationObject("requestExpected").DeepClone()
            End If
            Return Nothing
        End Function

        Private Shared Function GetRelevantVerificationItems(verification As JArray,
                                                               targetRef As String,
                                                               resolvedTargetCount As Integer) As List(Of JObject)
            If verification Is Nothing Then Return New List(Of JObject)()
            Dim normalizedTarget = NormalizeRef(targetRef)
            Return verification.OfType(Of JObject)().Where(
                Function(item)
                    Dim itemTarget = NormalizeRef(item("targetRef")?.ToString())
                    Return (String.IsNullOrWhiteSpace(itemTarget) AndAlso resolvedTargetCount = 1) OrElse
                        String.Equals(itemTarget, normalizedTarget, StringComparison.OrdinalIgnoreCase)
                End Function).ToList()
        End Function

        Private Shared Function VerificationObjectApplies(verification As JObject,
                                                           targetRef As String,
                                                           resolvedTargetCount As Integer) As Boolean
            If verification Is Nothing Then Return False
            Dim verificationTarget = NormalizeRef(verification("targetRef")?.ToString())
            If String.IsNullOrWhiteSpace(verificationTarget) Then Return resolvedTargetCount = 1
            Return String.Equals(
                verificationTarget,
                NormalizeRef(targetRef),
                StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function HasTargetObservation(observation As JObject,
                                                     targetRef As String,
                                                     resolvedTargetCount As Integer) As Boolean
            If observation Is Nothing Then Return False
            Dim verification = TryCast(observation("verification"), JArray)
            If verification IsNot Nothing Then
                Return GetRelevantVerificationItems(verification, targetRef, resolvedTargetCount).Any(
                    Function(item) item("actual") IsNot Nothing OrElse
                                   Not String.IsNullOrWhiteSpace(item("status")?.ToString()))
            End If
            Dim verificationObject = TryCast(observation("verification"), JObject)
            If VerificationObjectApplies(verificationObject, targetRef, resolvedTargetCount) Then
                Return verificationObject("actual") IsNot Nothing OrElse
                    verificationObject("satisfied") IsNot Nothing
            End If
            Return resolvedTargetCount = 1 AndAlso observation("after") IsNot Nothing
        End Function

        Private Shared Function ObservationConfirmsChange(observation As JObject) As Boolean
            Dim changed = observation?("changed")
            Return changed IsNot Nothing AndAlso changed.Type = JTokenType.Boolean AndAlso changed.Value(Of Boolean)()
        End Function

        Private Shared Function ExtractTargetRefs(effectType As String,
                                                  observation As JObject,
                                                  data As Object) As List(Of String)
            Dim result As New List(Of String)()
            Dim verification = TryCast(observation?("verification"), JArray)
            If verification IsNot Nothing Then
                For Each item In verification.OfType(Of JObject)()
                    AddRef(result, item("targetRef")?.ToString())
                Next
            End If

            Dim verificationObject = TryCast(observation?("verification"), JObject)
            If verificationObject IsNot Nothing Then AddRef(result, verificationObject("targetRef")?.ToString())

            Dim operations = TryCast(observation?("operations"), JArray)
            If operations IsNot Nothing Then
                For Each operation In operations.OfType(Of JObject)()
                    If String.Equals(effectType, "object_exists", StringComparison.OrdinalIgnoreCase) Then
                        AddRef(result, operation("resultRef")?.ToString())
                    ElseIf String.Equals(effectType, "object_absent", StringComparison.OrdinalIgnoreCase) Then
                        AddRef(result, operation("targetRef")?.ToString())
                    End If
                Next
            End If

            If result.Count = 0 Then AddRefs(result, observation?("targetRefs"))
            If result.Count = 0 Then
                Dim dataToken = ToToken(data)
                Dim dataObject = TryCast(dataToken, JObject)
                If dataObject IsNot Nothing Then
                    AddRefs(result, dataObject("targetRefs"))
                    Dim sheet = dataObject("sheet")?.ToString()
                    Dim address = dataObject("address")?.ToString()
                    If Not String.IsNullOrWhiteSpace(address) Then
                        result.Add(If(String.IsNullOrWhiteSpace(sheet), address, $"Excel:{sheet}!{address}"))
                    End If
                End If
            End If
            Return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        End Function

        Private Shared Sub AddRef(result As List(Of String), value As String)
            If result Is Nothing OrElse String.IsNullOrWhiteSpace(value) Then Return
            result.Add(value.Trim())
        End Sub

        Private Shared Function ExtractVerifiedActual(accessMode As String,
                                                      observation As JObject,
                                                      data As Object,
                                                      targetRef As String,
                                                      resolvedTargetCount As Integer) As JToken
            If accessMode = "read" OrElse accessMode = "compute" Then
                Return ToToken(data)?.DeepClone()
            End If
            If observation Is Nothing Then Return Nothing

            Dim verificationArray = TryCast(observation("verification"), JArray)
            If verificationArray IsNot Nothing Then
                Dim state As New JObject()
                Dim scalar As JToken = Nothing
                Dim scalarCount As Integer = 0
                For Each item In GetRelevantVerificationItems(verificationArray, targetRef, resolvedTargetCount)
                    Dim actual = item("actual")?.DeepClone()
                    Dim propertyName = If(item("property")?.ToString(), "").Trim()
                    If String.IsNullOrWhiteSpace(propertyName) Then
                        scalar = actual
                        scalarCount += 1
                    Else
                        state(propertyName) = If(actual, JValue.CreateNull())
                    End If
                Next
                If state.HasValues Then Return state
                If scalarCount = 1 Then Return scalar
                ' An explicit verification array is authoritative. Never borrow a global
                ' `after` value for a target that failed or was absent from that array.
                Return Nothing
            End If

            Dim verificationObject = TryCast(observation("verification"), JObject)
            If VerificationObjectApplies(verificationObject, targetRef, resolvedTargetCount) Then
                If verificationObject("actual") IsNot Nothing Then Return verificationObject("actual").DeepClone()
                ' Some high-level adapters return the normalized state that was matched in
                ' `expected` plus satisfied=true. Once the host has performed that comparison,
                ' the matched state is valid evidence; the raw tool request alone is not.
                If IsHostSatisfied(accessMode, observation, targetRef, resolvedTargetCount) AndAlso verificationObject("expected") IsNot Nothing Then
                    Return verificationObject("expected").DeepClone()
                End If
            End If

            If resolvedTargetCount = 1 AndAlso observation("after") IsNot Nothing Then Return observation("after").DeepClone()
            Return Nothing
        End Function

        Private Shared Function ExtractVerifiedPropertyName(observation As JObject,
                                                            targetRef As String,
                                                            resolvedTargetCount As Integer) As String
            Dim verification = TryCast(observation?("verification"), JArray)
            If verification Is Nothing Then Return ""
            Dim properties = GetRelevantVerificationItems(verification, targetRef, resolvedTargetCount).
                Select(Function(item) If(item("property")?.ToString(), "").Trim()).
                Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
            If properties.Count = 1 Then Return properties(0)
            Return ""
        End Function

        Private Shared Function NormalizeRef(value As String) As String
            Return Uri.UnescapeDataString(If(value, "")).Trim().TrimEnd("/"c).ToLowerInvariant()
        End Function

        Private Shared Sub AddRefs(result As List(Of String), token As JToken)
            If result Is Nothing OrElse token Is Nothing Then Return
            If token.Type = JTokenType.Array Then
                For Each item In token
                    Dim value = item?.ToString()
                    If Not String.IsNullOrWhiteSpace(value) Then result.Add(value.Trim())
                Next
            Else
                Dim value = token.ToString()
                If Not String.IsNullOrWhiteSpace(value) Then result.Add(value.Trim())
            End If
        End Sub

        Private Shared Function ToObject(value As Object) As JObject
            Return TryCast(ToToken(value), JObject)
        End Function

        Private Shared Function ToToken(value As Object) As JToken
            If value Is Nothing Then Return Nothing
            Try
                Return If(TryCast(value, JToken), JToken.FromObject(value))
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ComputeHash(value As Object) As String
            Dim token = ToToken(value)
            If token Is Nothing Then Return ""
            Dim text = token.ToString(Formatting.None)
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function
    End Class

End Namespace
