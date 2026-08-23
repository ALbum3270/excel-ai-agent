Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Completion boundary for the adaptive loop. A local tool success is never equivalent
    ''' to goal success: every frozen outcome requirement must match canonical host evidence.
    ''' </summary>
    Public NotInheritable Partial Class AgentGoalVerifier
        Private Sub New()
        End Sub

        Public Shared Function Validate(session As AgentSession,
                                        contextPack As Context.ContextPack,
                                        evidenceClaims As IEnumerable(Of String),
                                        readOnlyEvidenceCount As Integer) As String
            If session Is Nothing OrElse session.Spec Is Nothing Then
                Return "任务未完成：缺少可验收的任务规格。"
            End If

            Dim contract = session.Spec.OutcomeContract
            If contract Is Nothing OrElse contract.Requirements Is Nothing OrElse
               Not contract.Requirements.Any(Function(item) item IsNot Nothing AndAlso item.Required) Then
                If String.Equals(session.AppType, "Excel", StringComparison.OrdinalIgnoreCase) Then
                    Return "任务未完成：首次规划没有生成冻结的结构化结果合同；为防止把任意局部写入误报为完成，已失败关闭。"
                End If
                Return ValidateLegacyNonExcelOutcome(session)
            End If
            Dim integrityError = Goals.GoalOutcomeProjection.ValidateIntegrity(session.Spec, contract)
            If Not String.IsNullOrWhiteSpace(integrityError) Then Return integrityError

            Dim capabilityError = AgentExecutionContract.ValidateOutcome(session)
            If Not String.IsNullOrWhiteSpace(capabilityError) Then Return capabilityError

            ' ExpectedOutputs is a mutable legacy projection. Once GoalContract exists its
            ' semantics are already represented by Goal criteria and cannot be reintroduced
            ' as an independent completion authority.
            If session.Spec.GoalContract Is Nothing Then
                Dim outputError = ValidateExpectedOutputs(session)
                If Not String.IsNullOrWhiteSpace(outputError) Then Return outputError
            End If

            If String.Equals(session.AppType, "Excel", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(contract.BoundWorkbook) Then
                Dim currentWorkbook = ResolveContextWorkbookName(contextPack)
                If Not String.IsNullOrWhiteSpace(currentWorkbook) AndAlso
                   Not String.Equals(NormalizeWorkbookName(currentWorkbook),
                                     NormalizeWorkbookName(contract.BoundWorkbook),
                                     StringComparison.OrdinalIgnoreCase) Then
                    Return $"任务未完成：活动工作簿已从 {contract.BoundWorkbook} 切换为 {currentWorkbook}；旧工作簿中的证据不能证明当前工作簿状态。"
                End If
            End If

            Dim claims = New HashSet(Of String)(
                If(evidenceClaims, Enumerable.Empty(Of String)()).
                    Where(Function(value) Not String.IsNullOrWhiteSpace(value)).
                    Select(Function(value) value.Trim()),
                StringComparer.OrdinalIgnoreCase)
            ' Keep invalidation-only records from failed/partial writes in the timeline. They
            ' can revoke older proof, while EvidenceMatchesRequirementShape still prevents
            ' them from proving a positive requirement.
            Dim evidence = If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item IsNot Nothing).
                SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                Where(Function(item) item IsNot Nothing).
                ToList()

            For Each requirement In contract.Requirements.Where(Function(item) item IsNot Nothing AndAlso item.Required)
                Dim candidates = evidence.Where(
                    Function(record) EvidenceMatchesRequirementShape(session, record, requirement) AndAlso
                                     Not IsSuperseded(record, evidence, requirement) AndAlso
                                     RequirementValueMatches(record, requirement)).ToList()
                Dim availableMatches = SelectCoveringEvidence(candidates, requirement.TargetRef, requirement.EffectType)
                If availableMatches.Count = 0 Then
                    Return $"任务未完成：结果合同 {requirement.Id} 尚无匹配的宿主证据（effect={requirement.EffectType}, target={requirement.TargetRef}）。"
                End If

                Dim claimedMatches = SelectCoveringEvidence(
                    candidates.Where(Function(record) claims.Contains(record.EvidenceId)),
                    requirement.TargetRef,
                    requirement.EffectType)
                If claims.Count = 0 OrElse claimedMatches.Count = 0 Then
                    Return $"任务未完成：请在 complete.evidence 中引用结果合同 {requirement.Id} 的真实 evidenceId：{String.Join(", ", availableMatches.Select(Function(item) item.EvidenceId))}。"
                End If
            Next

            For Each claim In claims
                If Not evidence.Any(Function(record) String.Equals(record.EvidenceId, claim, StringComparison.OrdinalIgnoreCase)) Then
                    Return $"任务未完成：模型引用了不存在或未成功的证据 {claim}。"
                End If
            Next

            Return ""
        End Function

        Private Shared Function EvidenceMatchesRequirementExceptTarget(session As AgentSession,
                                                                       record As OutcomeEvidenceRecord,
                                                                       requirement As OutcomeRequirement) As Boolean
            If Not EvidenceMatchesRequirementShape(session, record, requirement) Then Return False
            ' The requested tool parameters are audit context, not proof. Completion is
            ' matched only against the state returned by the host verifier.
            If Not RequirementValueMatches(record, requirement) Then Return False
            Return True
        End Function

        Private Shared Function RequirementValueMatches(record As OutcomeEvidenceRecord,
                                                         requirement As OutcomeRequirement) As Boolean
            If ExpectedMatches(record.Actual,
                               requirement.ExpectedValue,
                               requirement.PropertyName,
                               requirement.Operator) Then Return True
            Return record.RequestVerified AndAlso
                ExpectedMatches(record.VerifiedRequest,
                                 requirement.ExpectedValue,
                                requirement.PropertyName,
                                requirement.Operator)
        End Function

        Private Shared Function EvidenceMatchesRequirementShape(session As AgentSession,
                                                                 record As OutcomeEvidenceRecord,
                                                                 requirement As OutcomeRequirement) As Boolean
            If record Is Nothing OrElse requirement Is Nothing OrElse Not record.Satisfied Then Return False
            If String.Equals(record.EffectType, "unclassified_mutation", StringComparison.OrdinalIgnoreCase) Then Return False
            If Not String.Equals(record.EffectType, requirement.EffectType, StringComparison.OrdinalIgnoreCase) Then Return False
            If (String.Equals(requirement.EffectType, "object_exists", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(requirement.EffectType, "object_absent", StringComparison.OrdinalIgnoreCase)) AndAlso
               (Not String.IsNullOrWhiteSpace(requirement.PropertyName) OrElse
                (requirement.ExpectedValue IsNot Nothing AndAlso requirement.ExpectedValue.Type <> JTokenType.Null)) Then Return False
            If Not PropertyMatches(record, requirement) Then Return False
            If Not String.IsNullOrWhiteSpace(requirement.DerivedFromCapability) AndAlso
               Not IsDerivedFromCapability(session, record, requirement) Then Return False
            Return True
        End Function

        Private Shared Function IsSuperseded(record As OutcomeEvidenceRecord,
                                             allEvidence As IEnumerable(Of OutcomeEvidenceRecord),
                                             requirement As OutcomeRequirement) As Boolean
            If record Is Nothing OrElse allEvidence Is Nothing Then Return False
            Return allEvidence.Any(
                Function(later)
                     If later Is Nothing OrElse (Not later.Satisfied AndAlso Not later.InvalidatesPrior) OrElse
                        later.WorldRevision <= record.WorldRevision Then Return False
                    If Not EffectsCanSupersede(record, later, requirement) Then Return False
                    Return EvidenceTargetsOverlap(record, later)
                End Function)
        End Function

        Private Shared Function EvidenceTargetsOverlap(left As OutcomeEvidenceRecord,
                                                       right As OutcomeEvidenceRecord) As Boolean
            If left Is Nothing OrElse right Is Nothing Then Return False
            Dim leftRefs = New List(Of String) From {left.TargetRef}
            leftRefs.AddRange(If(left.RelatedTargetRefs, New List(Of String)()))
            Dim rightRefs = New List(Of String) From {right.TargetRef}
            rightRefs.AddRange(If(right.RelatedTargetRefs, New List(Of String)()))
            Return leftRefs.Where(Function(value) Not String.IsNullOrWhiteSpace(value)).Any(
                Function(leftRef) rightRefs.Where(Function(value) Not String.IsNullOrWhiteSpace(value)).Any(
                    Function(rightRef) TargetsOverlap(leftRef, rightRef)))
        End Function

        Private Shared Function EffectsCanSupersede(earlier As OutcomeEvidenceRecord,
                                                     later As OutcomeEvidenceRecord,
                                                     requirement As OutcomeRequirement) As Boolean
            Dim earlierEffect = If(earlier?.EffectType, "").Trim().ToLowerInvariant()
            Dim laterEffect = If(later?.EffectType, "").Trim().ToLowerInvariant()
            If later.InvalidatesPrior AndAlso String.Equals(later.TargetRef, "*", StringComparison.Ordinal) Then Return True
            If later.InvalidatesPrior AndAlso laterEffect = "artifact" Then Return earlierEffect = "artifact"
            If later.InvalidatesPrior AndAlso laterEffect = "unclassified_mutation" Then Return True
            If laterEffect = "read_coverage" OrElse laterEffect = "compute_artifact" OrElse
               laterEffect = "artifact" Then Return False

            ' A stable artifact anchor (for example "chart created at D1") is related to the
            ' generated Office object. Any later mutation of that object may move, replace, or
            ' otherwise detach it from the frozen anchor. Require a fresh anchor observation
            ' instead of preserving stale creation proof merely because the object still exists.
            If earlierEffect = "artifact" Then Return True

            ' Deleting or recreating a worksheet invalidates every earlier assertion scoped to
            ' that worksheet. Reusing the same name never resurrects evidence from the old
            ' object instance.
            If laterEffect = "object_absent" OrElse laterEffect = "object_exists" Then Return True

            If earlierEffect = "object_exists" OrElse earlierEffect = "object_absent" Then
                ' A successful range/object mutation proves that an earlier absence assertion
                ' is stale, but ordinary child mutations do not invalidate existence.
                If earlierEffect = "object_absent" Then Return True
                Return False
            End If

            If earlierEffect = "property_state" Then
                If laterEffect <> "property_state" Then Return False
                Dim propertyName = If(requirement?.PropertyName, "").Trim()
                If String.IsNullOrWhiteSpace(propertyName) Then Return True
                Return RecordMentionsProperty(later, propertyName)
            End If

            Dim structuralEffects As New HashSet(Of String)(
                {"data_state", "formula_state", "order_state", "filter_state"},
                StringComparer.OrdinalIgnoreCase)
            If earlierEffect = "read_coverage" Then Return structuralEffects.Contains(laterEffect)
            Return structuralEffects.Contains(earlierEffect) AndAlso structuralEffects.Contains(laterEffect)
        End Function

        Private Shared Function RecordMentionsProperty(record As OutcomeEvidenceRecord,
                                                        propertyName As String) As Boolean
            If record Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return True
            If record.InvalidatesPrior AndAlso String.IsNullOrWhiteSpace(record.PropertyName) Then Return True
            If String.Equals(record.PropertyName, propertyName, StringComparison.OrdinalIgnoreCase) Then Return True
            Dim actualObject = TryCast(record.Actual, JObject)
            If actualObject IsNot Nothing AndAlso
               actualObject.GetValue(propertyName, StringComparison.OrdinalIgnoreCase) IsNot Nothing Then Return True
            Dim verifiedRequest = If(record.RequestVerified, TryCast(record.VerifiedRequest, JObject), Nothing)
            Return verifiedRequest IsNot Nothing AndAlso
                verifiedRequest.GetValue(propertyName, StringComparison.OrdinalIgnoreCase) IsNot Nothing
        End Function

        Private Shared Function TargetsOverlap(left As String, right As String) As Boolean
            If String.Equals(If(left, "").Trim(), "*", StringComparison.Ordinal) OrElse
               String.Equals(If(right, "").Trim(), "*", StringComparison.Ordinal) Then Return True
            Dim leftRange As ExcelRangeRef = Nothing
            Dim rightRange As ExcelRangeRef = Nothing
            Dim leftIsRange = TryParseExcelRange(left, leftRange)
            Dim rightIsRange = TryParseExcelRange(right, rightRange)
            If leftIsRange OrElse rightIsRange Then
                If leftIsRange AndAlso rightIsRange Then
                    Return RangeTargetsOverlap(leftRange, rightRange) AndAlso
                        RangesIntersect(leftRange, rightRange)
                End If

                Dim rangeRef = If(leftIsRange, leftRange, rightRange)
                Dim objectRef = If(leftIsRange, right, left)
                Return IsWorksheetRootReference(objectRef) AndAlso
                    String.Equals(
                        CanonicalObjectIdentity(objectRef),
                        BuildWorksheetIdentity(rangeRef.Workbook, rangeRef.Sheet),
                        StringComparison.OrdinalIgnoreCase)
            End If
            Dim leftIdentity = CanonicalObjectIdentity(left)
            Dim rightIdentity = CanonicalObjectIdentity(right)
            If String.Equals(leftIdentity, rightIdentity, StringComparison.OrdinalIgnoreCase) Then Return True

            ' Canonical object refs form a containment tree. A lifecycle/invalidation record
            ' for ChartObject1, a ListObject or a PivotTable must invalidate evidence for its
            ' descendants, while a segment boundary keeps siblings (Chart1/Chart10) separate.
            Return IsCanonicalAncestor(leftIdentity, rightIdentity) OrElse
                IsCanonicalAncestor(rightIdentity, leftIdentity)
        End Function

        Private Shared Function IsCanonicalAncestor(parentIdentity As String,
                                                     childIdentity As String) As Boolean
            If String.IsNullOrWhiteSpace(parentIdentity) OrElse String.IsNullOrWhiteSpace(childIdentity) Then Return False
            Return childIdentity.StartsWith(parentIdentity.TrimEnd("/"c) & "/",
                                            StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsWorksheetRootReference(value As String) As Boolean
            Dim identity = CanonicalObjectIdentity(value)
            Return Regex.IsMatch(identity, "^workbook:[^/]+/worksheet:[^/]+$", RegexOptions.IgnoreCase)
        End Function

        Private Shared Function PropertyMatches(record As OutcomeEvidenceRecord,
                                                requirement As OutcomeRequirement) As Boolean
            If String.IsNullOrWhiteSpace(requirement.PropertyName) Then Return True
            If String.Equals(record.PropertyName, requirement.PropertyName, StringComparison.OrdinalIgnoreCase) Then Return True
            Dim actualObject = TryCast(record.Actual, JObject)
            If actualObject IsNot Nothing AndAlso actualObject.GetValue(
                requirement.PropertyName,
                StringComparison.OrdinalIgnoreCase) IsNot Nothing Then Return True
            Dim verifiedRequest = If(record.RequestVerified, TryCast(record.VerifiedRequest, JObject), Nothing)
            Return verifiedRequest IsNot Nothing AndAlso verifiedRequest.GetValue(
                       requirement.PropertyName,
                       StringComparison.OrdinalIgnoreCase) IsNot Nothing
        End Function

        Private Shared Function ExpectedMatches(actualRequest As JToken,
                                                expected As JToken,
                                                propertyName As String,
                                                comparisonOperator As String) As Boolean
            If expected Is Nothing OrElse expected.Type = JTokenType.Null Then Return True
            If actualRequest Is Nothing Then Return False

            Dim actual = actualRequest
            If Not String.IsNullOrWhiteSpace(propertyName) AndAlso actualRequest.Type = JTokenType.Object Then
                actual = DirectCast(actualRequest, JObject).GetValue(propertyName, StringComparison.OrdinalIgnoreCase)
                If actual Is Nothing Then Return False
            End If

            Select Case If(comparisonOperator, "equals").Trim().ToLowerInvariant()
                Case "contains"
                    Return TokenContains(actual, expected, propertyName)
                Case "exists"
                    Return actual IsNot Nothing AndAlso actual.Type <> JTokenType.Null
                Case "covers"
                    Return TokenContains(actual, expected, propertyName)
                Case Else
                    Return TokensEqual(actual, expected, propertyName)
            End Select
        End Function

        Private Shared Function TokenContains(actual As JToken,
                                              expected As JToken,
                                              Optional propertyName As String = "") As Boolean
            If expected Is Nothing Then Return True
            If actual Is Nothing Then Return False
            If expected.Type = JTokenType.Object Then
                If actual.Type <> JTokenType.Object Then Return False
                Dim actualObject = DirectCast(actual, JObject)
                For Each prop In DirectCast(expected, JObject).Properties()
                    Dim actualValue = actualObject.GetValue(prop.Name, StringComparison.OrdinalIgnoreCase)
                    If Not TokenContains(actualValue, prop.Value, prop.Name) Then Return False
                Next
                Return True
            End If
            If expected.Type = JTokenType.Array Then
                If actual.Type <> JTokenType.Array Then Return False
                Dim actualArray = DirectCast(actual, JArray)
                Dim expectedArray = DirectCast(expected, JArray)
                If expectedArray.Count = 0 Then Return True
                Dim actualIndex As Integer = 0
                For Each expectedItem In expectedArray
                    Dim found As Boolean = False
                    While actualIndex < actualArray.Count
                        If TokenContains(actualArray(actualIndex), expectedItem, propertyName) Then
                            found = True
                            actualIndex += 1
                            Exit While
                        End If
                        actualIndex += 1
                    End While
                    If Not found Then Return False
                Next
                Return True
            End If
            If IsNumericToken(actual) OrElse IsNumericToken(expected) Then Return TokensEqual(actual, expected, propertyName)
            If actual.Type = JTokenType.String AndAlso expected.Type = JTokenType.String Then
                Dim comparison = If(IsCaseInsensitiveSemanticProperty(propertyName),
                                    StringComparison.OrdinalIgnoreCase,
                                    StringComparison.Ordinal)
                Return NormalizeSemanticString(actual.Value(Of String)(), propertyName).
                    IndexOf(NormalizeSemanticString(expected.Value(Of String)(), propertyName), comparison) >= 0
            End If
            Return TokensEqual(actual, expected, propertyName)
        End Function

        Private Shared Function TokensEqual(actual As JToken,
                                            expected As JToken,
                                            Optional propertyName As String = "") As Boolean
            If actual Is Nothing OrElse expected Is Nothing Then Return actual Is expected
            If IsNumericToken(actual) AndAlso IsNumericToken(expected) Then
                Dim actualNumber As Decimal
                Dim expectedNumber As Decimal
                Return Decimal.TryParse(actual.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, actualNumber) AndAlso
                       Decimal.TryParse(expected.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, expectedNumber) AndAlso
                       actualNumber = expectedNumber
            End If
            If actual.Type <> expected.Type Then Return False
            If actual.Type = JTokenType.String Then
                Dim comparison = If(IsCaseInsensitiveSemanticProperty(propertyName),
                                    StringComparison.OrdinalIgnoreCase,
                                    StringComparison.Ordinal)
                Return String.Equals(
                    NormalizeSemanticString(actual.Value(Of String)(), propertyName),
                    NormalizeSemanticString(expected.Value(Of String)(), propertyName),
                    comparison)
            End If
            Return JToken.DeepEquals(actual, expected)
        End Function

        Private Shared Function IsCaseInsensitiveSemanticProperty(propertyName As String) As Boolean
            Dim normalized = If(propertyName, "").Trim().ToLowerInvariant()
            Return New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "style", "type", "numberformat", "numberformatlocal", "fontcolor",
                "fillcolor", "backgroundcolor", "color", "horizontalalignment",
                "verticalalignment", "legendposition", "charttype", "format",
                "operator", "rule"
            }.Contains(normalized)
        End Function

        Private Shared Function NormalizeSemanticString(value As String,
                                                        propertyName As String) As String
            Dim text = If(value, "")
            If String.Equals(propertyName, "NumberFormat", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(propertyName, "NumberFormatLocal", StringComparison.OrdinalIgnoreCase) Then
                Return Regex.Replace(text.Trim(), "\s+", "")
            End If
            If IsCaseInsensitiveSemanticProperty(propertyName) Then Return text.Trim()
            Return text
        End Function

        Private Shared Function IsNumericToken(token As JToken) As Boolean
            Return token IsNot Nothing AndAlso
                (token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float)
        End Function

        Private Shared Function IsDerivedFromCapability(session As AgentSession,
                                                        record As OutcomeEvidenceRecord,
                                                        requirement As OutcomeRequirement) As Boolean
            If session?.Iterations Is Nothing Then Return False
            Dim capability = requirement.DerivedFromCapability
            Dim validatedCapabilities = session.Spec?.OutcomeContract?.ValidatedComputeCapabilities
            If validatedCapabilities Is Nothing OrElse
               Not validatedCapabilities.Contains(capability, StringComparer.OrdinalIgnoreCase) Then Return False
            Dim producers = session.Iterations.Where(
                Function(item) item IsNot Nothing AndAlso
                               item.Explanation IsNot Nothing AndAlso item.Explanation.Success AndAlso
                               item.Action IsNot Nothing AndAlso
                               String.Equals(item.Action.ToolId, capability, StringComparison.OrdinalIgnoreCase)).
                Where(Function(item) Not String.IsNullOrWhiteSpace(item.EvidenceId)).
                ToList()
            If producers.Count = 0 Then Return False

            Dim outputAncestors = GetDependencyClosure(session, record.DerivedFromEvidenceIds)
            For Each producer In producers
                If Not outputAncestors.Contains(producer.EvidenceId) Then Continue For

                ' A persisted compute result is only grounded when the exact computation that
                ' fed the write also descends from the read evidence required by the contract.
                ' This prevents an embedded sample from being computed and written while an
                ' unrelated full-range read is used as decorative evidence.
                Dim readRequirements = session.Spec.OutcomeContract.Requirements.Where(
                    Function(item) item IsNot Nothing AndAlso item.Required AndAlso
                                   String.Equals(item.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase)).ToList()
                If readRequirements.Count = 0 Then Return False

                Dim producerAncestors = GetDependencyClosure(session, producer.DependsOnEvidenceIds)
                Dim allTimelineEvidence = session.Iterations.Where(
                    Function(item) item IsNot Nothing).
                    SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                    Where(Function(item) item IsNot Nothing).
                    ToList()
                Dim ancestorReadEvidence = session.Iterations.Where(
                    Function(item) item IsNot Nothing AndAlso producerAncestors.Contains(item.EvidenceId)).
                    SelectMany(Function(item) If(item.ContractEvidence, New List(Of OutcomeEvidenceRecord)())).
                    Where(Function(item) item IsNot Nothing AndAlso item.Satisfied AndAlso
                                         String.Equals(item.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase)).
                    ToList()
                Dim allSourcesCovered = readRequirements.All(
                    Function(sourceRequirement)
                        Dim sourceCandidates = ancestorReadEvidence.Where(
                            Function(sourceRecord) EvidenceMatchesRequirementExceptTarget(
                                session,
                                sourceRecord,
                                sourceRequirement) AndAlso
                                Not IsSuperseded(sourceRecord, allTimelineEvidence, sourceRequirement)).ToList()
                        Return SelectCoveringEvidence(
                            sourceCandidates,
                            sourceRequirement.TargetRef,
                            sourceRequirement.EffectType).Count > 0
                    End Function)
                If allSourcesCovered Then Return True
            Next
            Return False
        End Function

        Private Shared Function GetDependencyClosure(session As AgentSession,
                                                     directDependencies As IEnumerable(Of String)) As HashSet(Of String)
            Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim pending As New Queue(Of String)(If(directDependencies, Enumerable.Empty(Of String)()).
                                                Where(Function(value) Not String.IsNullOrWhiteSpace(value)))
            While pending.Count > 0
                Dim dependency = pending.Dequeue()
                If Not visited.Add(dependency) Then Continue While
                Dim dependencyIteration = session.Iterations.FirstOrDefault(
                    Function(item) item IsNot Nothing AndAlso
                                   String.Equals(item.EvidenceId, dependency, StringComparison.OrdinalIgnoreCase))
                If dependencyIteration Is Nothing Then Continue While
                For Each parent In If(dependencyIteration.DependsOnEvidenceIds, New List(Of String)())
                    If Not String.IsNullOrWhiteSpace(parent) Then pending.Enqueue(parent)
                Next
            End While
            Return visited
        End Function

        Private Shared Function SelectCoveringEvidence(candidates As IEnumerable(Of OutcomeEvidenceRecord),
                                                       requiredTarget As String,
                                                       effectType As String) As List(Of OutcomeEvidenceRecord)
            Dim available = If(candidates, Enumerable.Empty(Of OutcomeEvidenceRecord)()).
                Where(Function(item) item IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(item.TargetRef)).
                ToList()
            If String.IsNullOrWhiteSpace(requiredTarget) Then Return available.Take(1).ToList()

            Dim requiredRange As ExcelRangeRef = Nothing
            If TryParseExcelRange(requiredTarget, requiredRange) Then
                Dim rectangles As New List(Of KeyValuePair(Of OutcomeEvidenceRecord, ExcelRangeRef))()
                For Each record In available
                    Dim actualRange As ExcelRangeRef = Nothing
                    If Not TryParseExcelRange(record.TargetRef, actualRange) Then Continue For
                    If Not RangeScopeMatches(actualRange, requiredRange) Then Continue For
                    If Not RangesIntersect(actualRange, requiredRange) Then Continue For
                    rectangles.Add(New KeyValuePair(Of OutcomeEvidenceRecord, ExcelRangeRef)(record, actualRange))
                Next
                If rectangles.Count = 0 OrElse Not RectanglesCover(requiredRange, rectangles.Select(Function(item) item.Value)) Then
                    Return New List(Of OutcomeEvidenceRecord)()
                End If

                ' Return only evidence that contributes to the covered target. Completion must
                ' cite every returned ID; unrelated overlapping evidence is not accepted.
                Return rectangles.Select(Function(item) item.Key).
                    GroupBy(Function(item) item.EvidenceId, StringComparer.OrdinalIgnoreCase).
                    Select(Function(group) group.First()).
                    ToList()
            End If

            Dim requiredIdentity = CanonicalObjectIdentity(requiredTarget)
            Return available.Where(
                Function(record) String.Equals(
                    CanonicalObjectIdentity(record.TargetRef),
                    requiredIdentity,
                    StringComparison.OrdinalIgnoreCase)).Take(1).ToList()
        End Function

        Private Shared Function SheetMatches(actual As String, required As String) As Boolean
            If String.IsNullOrWhiteSpace(required) Then Return String.IsNullOrWhiteSpace(actual)
            Return String.Equals(NormalizeWorksheetName(actual),
                                 NormalizeWorksheetName(required),
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function RangeScopeMatches(actual As ExcelRangeRef,
                                                  required As ExcelRangeRef) As Boolean
            If actual Is Nothing OrElse required Is Nothing Then Return False
            If Not String.Equals(NormalizeWorkbookName(actual.Workbook),
                                 NormalizeWorkbookName(required.Workbook),
                                 StringComparison.OrdinalIgnoreCase) OrElse
               Not SheetMatches(actual.Sheet, required.Sheet) Then Return False
            ' A base-range requirement may be proven by a verified child state (for example
            ' its Interior), but a child requirement is never interchangeable with a sibling.
            Return String.IsNullOrWhiteSpace(required.ChildPath) OrElse
                String.Equals(NormalizeChildPath(actual.ChildPath),
                              NormalizeChildPath(required.ChildPath),
                              StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function RangeTargetsOverlap(left As ExcelRangeRef,
                                                    right As ExcelRangeRef) As Boolean
            If left Is Nothing OrElse right Is Nothing Then Return False
            If Not String.Equals(NormalizeWorkbookName(left.Workbook),
                                 NormalizeWorkbookName(right.Workbook),
                                 StringComparison.OrdinalIgnoreCase) OrElse
               Not SheetMatches(left.Sheet, right.Sheet) Then Return False
            Dim leftChild = NormalizeChildPath(left.ChildPath)
            Dim rightChild = NormalizeChildPath(right.ChildPath)
            Return String.IsNullOrWhiteSpace(leftChild) OrElse
                String.IsNullOrWhiteSpace(rightChild) OrElse
                String.Equals(leftChild, rightChild, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function NormalizeWorksheetName(value As String) As String
            Return Uri.UnescapeDataString(If(value, "")).Trim().Trim("'"c).ToLowerInvariant()
        End Function

        Private Shared Function RangesIntersect(left As ExcelRangeRef, right As ExcelRangeRef) As Boolean
            Return left.StartRow <= right.EndRow AndAlso left.EndRow >= right.StartRow AndAlso
                   left.StartColumn <= right.EndColumn AndAlso left.EndColumn >= right.StartColumn
        End Function

        Private Shared Function RectanglesCover(required As ExcelRangeRef,
                                                ranges As IEnumerable(Of ExcelRangeRef)) As Boolean
            Dim intersections = If(ranges, Enumerable.Empty(Of ExcelRangeRef)()).
                Where(Function(item) item IsNot Nothing AndAlso RangesIntersect(item, required)).
                Select(Function(item) New ExcelRangeRef With {
                    .Workbook = required.Workbook,
                    .Sheet = required.Sheet,
                    .StartColumn = Math.Max(item.StartColumn, required.StartColumn),
                    .EndColumn = Math.Min(item.EndColumn, required.EndColumn),
                    .StartRow = Math.Max(item.StartRow, required.StartRow),
                    .EndRow = Math.Min(item.EndRow, required.EndRow)
                }).ToList()
            If intersections.Count = 0 Then Return False

            ' Sweep only rectangle boundaries, not worksheet cells. With at most one rectangle
            ' per action this remains small even for whole-column or million-row ranges.
            Dim xBoundaries As New SortedSet(Of Long) From {
                required.StartColumn,
                CLng(required.EndColumn) + 1
            }
            Dim yBoundaries As New SortedSet(Of Long) From {
                required.StartRow,
                CLng(required.EndRow) + 1
            }
            For Each item In intersections
                xBoundaries.Add(item.StartColumn)
                xBoundaries.Add(CLng(item.EndColumn) + 1)
                yBoundaries.Add(item.StartRow)
                yBoundaries.Add(CLng(item.EndRow) + 1)
            Next

            Dim xs = xBoundaries.ToList()
            Dim ys = yBoundaries.ToList()
            For xIndex = 0 To xs.Count - 2
                For yIndex = 0 To ys.Count - 2
                    Dim sampleColumn = xs(xIndex)
                    Dim sampleRow = ys(yIndex)
                    If sampleColumn < required.StartColumn OrElse sampleColumn > required.EndColumn OrElse
                       sampleRow < required.StartRow OrElse sampleRow > required.EndRow Then Continue For
                    If Not intersections.Any(
                        Function(item) sampleColumn >= item.StartColumn AndAlso sampleColumn <= item.EndColumn AndAlso
                                       sampleRow >= item.StartRow AndAlso sampleRow <= item.EndRow) Then Return False
                Next
            Next
            Return True
        End Function

        Friend Shared Function ValidateContractTargetReference(appType As String,
                                                               targetRef As String) As String
            If Not String.Equals(appType, "Excel", StringComparison.OrdinalIgnoreCase) Then Return ""
            Dim parsed As ExcelRangeRef = Nothing
            If TryParseExcelRange(targetRef, parsed) Then
                If String.IsNullOrWhiteSpace(parsed.Sheet) Then
                    Return $"Excel 范围目标 {targetRef} 未限定工作表，不能作为稳定结果合同引用。"
                End If
                Return ""
            End If

            Dim decoded = Uri.UnescapeDataString(If(targetRef, "")).Trim()
            If decoded.IndexOf("!"c) >= 0 OrElse
               decoded.IndexOf("/ranges/", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               Regex.IsMatch(decoded, "^[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?$", RegexOptions.IgnoreCase) Then
                Return $"Excel 目标引用 {targetRef} 不是可解析的稳定范围或对象引用。"
            End If
            Return ""
        End Function

        Private Shared Function TryParseExcelRange(value As String, ByRef result As ExcelRangeRef) As Boolean
            result = Nothing
            If String.IsNullOrWhiteSpace(value) Then Return False
            Dim decoded As String
            Try
                decoded = Uri.UnescapeDataString(value).Replace("$", "").Trim()
            Catch
                Return False
            End Try

            Dim match = Regex.Match(
                decoded,
                "^(?:Excel:)?(?<sheet>[^!/:]+)!(?<start>[A-Za-z]{1,3}\d+)(?::(?<end>[A-Za-z]{1,3}\d+))?$",
                RegexOptions.IgnoreCase)
            Dim workbook = "active"
            If Not match.Success Then
                match = Regex.Match(
                    decoded.Replace("\", "/"),
                    "^(?:Excel:)?workbooks/(?<workbook>[^/]+)/worksheets/(?<sheet>[^/]+)/ranges/(?<start>[A-Za-z]{1,3}\d+)(?::(?<end>[A-Za-z]{1,3}\d+))?(?:/(?<child>font|interior|borders(?:/\d+)?|rows|columns))?$",
                    RegexOptions.IgnoreCase)
                If match.Success Then workbook = match.Groups("workbook").Value
            End If
            If Not match.Success Then
                match = Regex.Match(
                    decoded.Replace("\", "/"),
                    "^(?:Excel:)?(?:[^/]+/)*worksheets/(?<sheet>[^/]+)/ranges/(?<start>[A-Za-z]{1,3}\d+)(?::(?<end>[A-Za-z]{1,3}\d+))?(?:/(?<child>font|interior|borders(?:/\d+)?|rows|columns))?$",
                    RegexOptions.IgnoreCase)
            End If
            If Not match.Success Then
                match = Regex.Match(
                    decoded,
                    "^(?<start>[A-Za-z]{1,3}\d+)(?::(?<end>[A-Za-z]{1,3}\d+))?$",
                    RegexOptions.IgnoreCase)
            End If
            If Not match.Success Then Return False
            Dim startColumn As Integer = 0
            Dim startRow As Integer = 0
            Dim endColumn As Integer = 0
            Dim endRow As Integer = 0
            If Not ParseCell(match.Groups("start").Value, startColumn, startRow) Then Return False
            Dim endValue = If(match.Groups("end").Success, match.Groups("end").Value, match.Groups("start").Value)
            If Not ParseCell(endValue, endColumn, endRow) Then Return False
            result = New ExcelRangeRef With {
                .Workbook = workbook,
                .Sheet = match.Groups("sheet").Value.Trim("'"c),
                .ChildPath = If(match.Groups("child").Success, match.Groups("child").Value, ""),
                .StartColumn = Math.Min(startColumn, endColumn),
                .EndColumn = Math.Max(startColumn, endColumn),
                .StartRow = Math.Min(startRow, endRow),
                .EndRow = Math.Max(startRow, endRow)
            }
            Return True
        End Function

        Private Shared Function ParseCell(value As String, ByRef column As Integer, ByRef row As Integer) As Boolean
            Dim match = Regex.Match(If(value, ""), "^(?<column>[A-Za-z]{1,3})(?<row>\d+)$")
            If Not match.Success OrElse Not Integer.TryParse(match.Groups("row").Value, row) OrElse row <= 0 Then Return False
            column = 0
            For Each ch In match.Groups("column").Value.ToUpperInvariant()
                column = column * 26 + (AscW(ch) - AscW("A"c) + 1)
            Next
            Return column > 0
        End Function

        Private Shared Function NormalizeTarget(value As String) As String
            Dim decoded = Uri.UnescapeDataString(If(value, "")).Trim().ToLowerInvariant()
            Return Regex.Replace(decoded, "[\s'""“”‘’]", "")
        End Function

        Private Shared Function NormalizeWorkbookName(value As String) As String
            Dim normalized = Uri.UnescapeDataString(If(value, "active")).Trim().Trim("'"c).ToLowerInvariant()
            Return If(String.IsNullOrWhiteSpace(normalized), "active", normalized)
        End Function

        Private Shared Function NormalizeChildPath(value As String) As String
            Return Uri.UnescapeDataString(If(value, "")).Trim().Trim("/"c).ToLowerInvariant()
        End Function

        Private Shared Function Normalize(value As String) As String
            Return Regex.Replace(If(value, "").Trim().ToLowerInvariant(), "\s+", "")
        End Function

        Private Shared Function ValidateExpectedOutputs(session As AgentSession) As String
            If session?.Spec Is Nothing Then Return ""
            Dim successfulIterations = If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item?.Explanation IsNot Nothing AndAlso item.Explanation.Success).
                ToList()

            If ContainsExpectedOutput(session.Spec, "images") Then
                Dim hasImageArtifact = successfulIterations.Any(
                    Function(item) item.OutcomeArtifacts IsNot Nothing AndAlso
                                   item.OutcomeArtifacts.SelectTokens("$..*").Any(
                                       Function(token) token.Type = JTokenType.String AndAlso
                                                       Regex.IsMatch(token.ToString(), "\.(png|jpe?g|gif|webp|svg)$", RegexOptions.IgnoreCase)))
                If Not hasImageArtifact Then Return "任务未完成：要求真实图片，但宿主没有返回可访问的图片产物。"
            End If

            If session.Spec.ExpectedSlideCount > 0 Then
                Dim created = successfulIterations.SelectMany(
                    Function(item) If(item.OutcomeEvidence?.SelectTokens("$..targetRefs[*]"), Enumerable.Empty(Of JToken)())).
                    Select(Function(token) token.ToString()).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    Count(Function(target) target.IndexOf("slide", StringComparison.OrdinalIgnoreCase) >= 0)
                If created < session.Spec.ExpectedSlideCount Then
                    Return $"任务未完成：要求至少 {session.Spec.ExpectedSlideCount} 张幻灯片，宿主证据仅确认 {created} 张。"
                End If
            End If
            Return ""
        End Function

        Private Shared Function ContainsExpectedOutput(spec As AgentTaskSpec, value As String) As Boolean
            Return spec?.ExpectedOutputs IsNot Nothing AndAlso
                spec.ExpectedOutputs.Any(Function(item) String.Equals(item, value, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Shared Function ValidateLegacyNonExcelOutcome(session As AgentSession) As String
            Dim verifiedMutation = If(session.Iterations, New List(Of ReActIteration)()).Any(
                Function(item) item?.Explanation IsNot Nothing AndAlso item.Explanation.Success AndAlso
                               item.OutcomeEvidence IsNot Nothing AndAlso
                               (BooleanValue(item.OutcomeEvidence("satisfied")) OrElse
                                BooleanValue(item.OutcomeEvidence.SelectToken("visualVerification.passed"))))
            If verifiedMutation Then Return ""
            Return "任务未完成：当前宿主尚未迁移结构化结果合同，且没有明确的宿主验收证据。"
        End Function

        Private Shared Function BooleanValue(token As JToken) As Boolean
            Return token IsNot Nothing AndAlso token.Type = JTokenType.Boolean AndAlso token.Value(Of Boolean)()
        End Function

        Private NotInheritable Class ExcelRangeRef
            Public Property Workbook As String = "active"
            Public Property Sheet As String
            Public Property ChildPath As String = ""
            Public Property StartColumn As Integer
            Public Property EndColumn As Integer
            Public Property StartRow As Integer
            Public Property EndRow As Integer
        End Class
    End Class

End Namespace
