Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        Private Shared Function CloneToken(value As Object) As JToken
            If value Is Nothing Then Return Nothing
            Try
                Dim token = TryCast(value, JToken)
                If token Is Nothing Then token = JToken.FromObject(value)
                Return token.DeepClone()
            Catch
                Return Nothing
            End Try
        End Function

        Private Function FreezeInitialOutcomeContract(session As AgentSession,
                                                      plan As ExecutionPlan,
                                                      executionAppType As String) As String
            If session?.Spec Is Nothing Then Return "任务规格缺失，无法冻结结果合同。"
            If session.Spec.GoalContract IsNot Nothing Then
                ' A mutable legacy projection must never outrank the proposal generated for
                ' the current frozen goal.  Preserve only an already sealed resume contract.
                If session.Spec.OutcomeContract Is Nothing OrElse Not session.Spec.OutcomeContract.Frozen Then
                    session.Spec.OutcomeContract = plan?.OutcomeContract
                End If
            ElseIf session.Spec.OutcomeContract Is Nothing AndAlso plan?.OutcomeContract IsNot Nothing Then
                session.Spec.OutcomeContract = plan.OutcomeContract
            End If
            Dim contract = session.Spec.OutcomeContract
            If contract Is Nothing OrElse contract.Requirements Is Nothing OrElse contract.Requirements.Count = 0 Then
                If String.Equals(executionAppType, "Excel", StringComparison.OrdinalIgnoreCase) Then
                    Return "规划失败：模型未生成结构化结果合同；为避免局部成功被误报为任务完成，未执行任何 Office 写入。"
                End If
                Return ""
            End If
            If contract.Frozen Then
                Return Goals.GoalOutcomeProjection.ValidateIntegrity(session.Spec, contract)
            End If
            contract.ValidatedComputeCapabilities.Clear()

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
                    If producerDescriptor Is Nothing OrElse
                       Not String.Equals(producerDescriptor.AccessMode, "compute", StringComparison.OrdinalIgnoreCase) Then
                        Return $"规划失败：结果合同 {requirement.Id} 的 derivedFromCapability 必须引用已注册的 compute 工具；{requirement.DerivedFromCapability} 不是可验证的计算来源。"
                    End If
                    If Not contract.ValidatedComputeCapabilities.Contains(
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
            Return ""
        End Function

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
