Imports System.Collections.Generic
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json

Namespace Agent.Goals

    ''' <summary>
    ''' Transitional one-way adapter from the immutable user goal to the legacy host-evidence
    ''' contract.  OutcomeContract may describe how to verify a goal, but it is never a source
    ''' from which goal semantics can be reconstructed.
    ''' </summary>
    Friend NotInheritable Class GoalOutcomeProjection
        Private Sub New()
        End Sub

        Friend Shared Function RequiredHostCriterionIds(spec As AgentTaskSpec) As List(Of String)
            If spec?.GoalContract IsNot Nothing Then
                Return spec.GoalContract.Criteria.
                    Where(Function(item) item IsNot Nothing AndAlso
                        item.Required AndAlso
                        Not String.Equals(item.Kind, "capability", StringComparison.OrdinalIgnoreCase)).
                    Select(Function(item) If(item.Id, "").Trim()).
                    Where(Function(id) id.Length > 0).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()
            End If

            ' Compatibility only for persisted sessions created before GoalContract existed.
            Dim result As New List(Of String)()
            For index = 0 To If(spec?.SuccessCriteria, New List(Of String)()).Count - 1
                result.Add($"criterion-{index + 1}")
            Next
            Return result
        End Function

        Friend Shared Function ValidateBinding(spec As AgentTaskSpec,
                                               contract As OutcomeContract,
                                               requireFrozenGoalBinding As Boolean) As String
            If spec Is Nothing Then Return "任务规格缺失，无法验证结果合同与用户目标的绑定。"
            If contract Is Nothing Then Return "结果合同缺失，无法验证其是否覆盖用户目标。"

            Dim goal = spec.GoalContract
            If goal IsNot Nothing AndAlso goal.Constraints.Any(
                Function(item) item IsNot Nothing AndAlso item.Required) Then
                Return "冻结目标包含 required constraints，但当前没有受信任的 GoalConstraintVerifier；不能把普通 OutcomeRequirement 当作约束证明。"
            End If
            If goal IsNot Nothing AndAlso requireFrozenGoalBinding Then
                If String.IsNullOrWhiteSpace(contract.BoundGoalContractHash) Then
                    Return "任务未完成：结果合同没有绑定冻结后的 GoalContract。"
                End If
                If Not String.Equals(contract.BoundGoalContractHash,
                                     goal.ContractHash,
                                     StringComparison.Ordinal) Then
                    Return "[GOAL_BINDING_MISMATCH] 任务未完成：结果合同绑定的 GoalContract 已变化，旧验证映射不能证明当前目标。"
                End If
            End If

            Dim expected = New HashSet(Of String)(
                RequiredHostCriterionIds(spec),
                StringComparer.OrdinalIgnoreCase)
            Dim mapped As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim unboundSupportRequirements As New List(Of OutcomeRequirement)()

            For Each requirement In If(contract.Requirements, New List(Of OutcomeRequirement)()).
                Where(Function(item) item IsNot Nothing AndAlso item.Required)
                Dim criterionIds = If(requirement.CriterionIds, New List(Of String)()).
                    Where(Function(id) Not String.IsNullOrWhiteSpace(id)).
                    Select(Function(id) id.Trim()).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()
                If criterionIds.Count > 1 Then
                    Return $"结果合同 {requirement.Id} 把多个独立目标折叠成同一条宿主断言；每条 requirement 最多映射一个 Goal criterion。"
                End If
                If goal IsNot Nothing AndAlso criterionIds.Count = 0 Then
                    Dim isSupport = String.Equals(requirement.EffectType, "read_coverage", StringComparison.OrdinalIgnoreCase) OrElse
                        String.Equals(requirement.EffectType, "compute_artifact", StringComparison.OrdinalIgnoreCase)
                    If Not isSupport Then
                        Return $"结果合同 {requirement.Id} 是未绑定 Goal criterion 的额外目标；验证计划不能扩大冻结后的用户目标。"
                    End If
                    unboundSupportRequirements.Add(requirement)
                End If
                For Each criterionId In criterionIds
                    If Not expected.Contains(criterionId) Then
                        Return $"结果合同 {requirement.Id} 引用了不属于当前冻结目标的 criterion {criterionId}。"
                    End If
                    mapped.Add(criterionId)
                Next
            Next

            If expected.Count = 0 Then Return ""
            Dim omitted = expected.Where(Function(id) Not mapped.Contains(id)).ToList()
            If omitted.Count > 0 Then
                Return $"结果合同未覆盖冻结目标的全部 required criteria：{String.Join(", ", omitted)}。"
            End If
            If goal IsNot Nothing AndAlso unboundSupportRequirements.Count > 0 Then
                Dim hasMappedComputeOutput = contract.Requirements.Any(
                    Function(item) item IsNot Nothing AndAlso item.Required AndAlso
                        item.CriterionIds IsNot Nothing AndAlso item.CriterionIds.Count > 0 AndAlso
                        Not String.IsNullOrWhiteSpace(item.DerivedFromCapability))
                If Not hasMappedComputeOutput Then
                    Return "结果合同包含未绑定目标的读取/计算支持断言，但没有任何已绑定的计算产物需要这些支持证据。"
                End If
            End If
            Return ""
        End Function

        Friend Shared Sub Seal(spec As AgentTaskSpec,
                               contract As OutcomeContract)
            If spec Is Nothing Then Throw New ArgumentNullException(NameOf(spec))
            If contract Is Nothing Then Throw New ArgumentNullException(NameOf(contract))
            Dim hasGoal = spec.GoalContract IsNot Nothing
            contract.BindToGoal(
                If(hasGoal, spec.GoalContract.ContractHash, ""),
                If(hasGoal, "goal-v1", "legacy-v1"))
            contract.Frozen = True
            contract.SealIntegrity(ComputeIntegrityHash(contract))
        End Sub

        Friend Shared Function ValidateIntegrity(spec As AgentTaskSpec,
                                                 contract As OutcomeContract) As String
            If contract Is Nothing OrElse Not contract.Frozen Then
                Return "任务未完成：结果合同尚未冻结。"
            End If
            Dim expectedMode = If(spec?.GoalContract IsNot Nothing, "goal-v1", "legacy-v1")
            If Not String.Equals(contract.BindingMode, expectedMode, StringComparison.Ordinal) Then
                Return "任务未完成：结果合同使用了错误的目标绑定模式。"
            End If
            Dim bindingError = ValidateBinding(spec, contract, requireFrozenGoalBinding:=True)
            If Not String.IsNullOrWhiteSpace(bindingError) Then Return bindingError
            If String.IsNullOrWhiteSpace(contract.FrozenOutcomeContractHash) OrElse
               Not String.Equals(contract.FrozenOutcomeContractHash,
                                 ComputeIntegrityHash(contract),
                                 StringComparison.Ordinal) Then
                Return "[OUTCOME_CONTRACT_MUTATED] 任务未完成：冻结后的结果合同已被修改，旧证据不能用于当前验收。"
            End If
            Return ""
        End Function

        Private Shared Function ComputeIntegrityHash(contract As OutcomeContract) As String
            Dim canonical As New StringBuilder()
            AppendField(canonical, contract.SchemaVersion)
            AppendField(canonical, contract.BoundGoalContractHash)
            AppendField(canonical, contract.BindingMode)
            AppendField(canonical, contract.BoundWorkbook)
            For Each capability In If(contract.ValidatedComputeCapabilities, New List(Of String)()).
                Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                Select(Function(item) item.Trim()).
                OrderBy(Function(item) item, StringComparer.OrdinalIgnoreCase)
                AppendField(canonical, capability)
            Next
            For Each requirement In If(contract.Requirements, New List(Of OutcomeRequirement)()).
                OrderBy(Function(item) If(item?.Id, ""), StringComparer.OrdinalIgnoreCase)
                If requirement Is Nothing Then
                    AppendField(canonical, "<null-requirement>")
                    Continue For
                End If
                AppendField(canonical, requirement.Id)
                AppendField(canonical, requirement.AppType)
                AppendField(canonical, requirement.TargetRef)
                AppendField(canonical, requirement.EffectType)
                AppendField(canonical, requirement.PropertyName)
                AppendField(canonical, requirement.Operator)
                AppendField(canonical, If(requirement.ExpectedValue?.ToString(Formatting.None), "<null>"))
                AppendField(canonical, requirement.DerivedFromCapability)
                For Each criterionId In If(requirement.CriterionIds, New List(Of String)()).
                    Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                    Select(Function(item) item.Trim()).
                    OrderBy(Function(item) item, StringComparer.OrdinalIgnoreCase)
                    AppendField(canonical, criterionId)
                Next
                AppendField(canonical, requirement.Required.ToString(Globalization.CultureInfo.InvariantCulture))
            Next

            Using sha = SHA256.Create()
                Dim bytes = Encoding.UTF8.GetBytes(canonical.ToString())
                Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "")
            End Using
        End Function

        Private Shared Sub AppendField(target As StringBuilder,
                                       value As String)
            Dim text = If(value, "")
            target.Append(text.Length.ToString(Globalization.CultureInfo.InvariantCulture)).
                Append(":"c).
                Append(text).
                Append("|"c)
        End Sub

    End Class

End Namespace
