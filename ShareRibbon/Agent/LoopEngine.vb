Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' One adaptive decision loop: Think -> choose one action -> SafetyGate -> Act ->
    ''' Observe -> verify -> update state. A plan is explanatory UI guidance only.
    ''' </summary>
    Public Partial Class LoopEngine
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _undoManager As Core.UndoManager

        Private Const MaxIterations As Integer = 15
        Private Const MaxReplanAttempts As Integer = 2
        Private Const MaxIdenticalRetryAttempts As Integer = 2

        Public Property OnStatusChanged As Action(Of String)
        Public Property OnIterationUpdate As Action(Of ReActIteration)
        Public Property OnStepCompleted As Action(Of Integer, Boolean, String)
        Public Property OnExecutionExplained As Action(Of ExecutionExplanation)
        Public Property OnRequestApproval As Func(Of String, Task(Of Boolean))
        Public Property OnPlanGenerated As Action(Of ExecutionPlan)
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, Task(Of String))
        Public Property CaptureContextPack As Func(Of Context.ContextPack)

        Public Sub New(toolRegistry As ToolRegistry,
                       memory As AgentMemory,
                       promptManager As PromptManager)
            _toolRegistry = toolRegistry
            _memory = memory
            _promptManager = promptManager
            _undoManager = New Core.UndoManager()
        End Sub

        Public Async Function RunAsync(session As AgentSession,
                                        systemPrompt As String,
                                        Optional skill As AgentSkill = Nothing) As Task(Of AgentResult)
            Dim decisionCount As Integer = 0
            Dim replanAttempts As Integer = 0
            Dim modelDeclaredComplete As Boolean = False
            Dim completionMessage As String = ""
            Dim readOnlyEvidence As New JArray()
            Dim toolDataflow As New AgentToolDataflow()
            Dim blockedActionSignatures As New Dictionary(Of String, BlockedActionState)(StringComparer.Ordinal)
            Dim retryableFailureCounts As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim worldRevision As Long = 0
            Dim frozenGoalHash As String = ""

            Try
                OnStatusChanged?.Invoke("正在分析任务...")
                If session.Spec Is Nothing Then
                    session.Spec = Await GenerateSpecAsync(session)
                Else
                    AppLogger.Info("LoopEngine", $"Use precomputed TaskSpec goal={AppLogger.Redact(session.Spec.Goal)}")
                End If

                OnStatusChanged?.Invoke("正在制定高层任务骨架...")
                Dim goalContractError = EstablishFrozenGoalContract(session)
                If Not String.IsNullOrWhiteSpace(goalContractError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(goalContractError)
                    Return AgentResult.Failed(session.Id, goalContractError)
                End If
                frozenGoalHash = session.Spec.GoalContract.ContractHash

                session.Plan = Await GeneratePlanAsync(session, systemPrompt, skill)
                If session.Plan Is Nothing Then
                    session.Plan = CreateFallbackPlan(session)
                ElseIf session.Plan.Steps Is Nothing OrElse session.Plan.Steps.Count = 0 Then
                    session.Plan.Steps = CreateFallbackPlan(session).Steps
                End If
                Dim outcomeContractError = FreezeInitialOutcomeContract(session, session.Plan, executionAppType:=session.AppType)
                If Not String.IsNullOrWhiteSpace(outcomeContractError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(outcomeContractError)
                    Return AgentResult.Failed(session.Id, outcomeContractError)
                End If
                OnPlanGenerated?.Invoke(session.Plan)

                session.Status = AgentStatus.Executing
                OnStatusChanged?.Invoke($"任务骨架已生成（{session.Plan.Steps.Count} 个提示），进入自适应执行...")
                Dim executionContext = ToolExecutionContext.FromSession(session, skill)
                Dim iterationLimit = Math.Min(MaxIterations, Math.Max(1, session.MaxIterations))

                While decisionCount < iterationLimit AndAlso Not modelDeclaredComplete
                    Dim goalInvariantError = ValidateFrozenGoalInvariant(session, frozenGoalHash)
                    If Not String.IsNullOrWhiteSpace(goalInvariantError) Then
                        session.Status = AgentStatus.Failed
                        OnStatusChanged?.Invoke(goalInvariantError)
                        Return AgentResult.Failed(session.Id, goalInvariantError)
                    End If
                    Dim planStep = GetCurrentPlanHint(session.Plan)
                    decisionCount += 1

                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"决策 {decisionCount}/{iterationLimit}: 正在根据最新观察选择下一动作...")
                    Dim rawDecision = Await ThinkAsync(session, planStep, systemPrompt)
                    Dim decision = ParseReactDecision(rawDecision)

                    goalInvariantError = ValidateFrozenGoalInvariant(session, frozenGoalHash)
                    If Not String.IsNullOrWhiteSpace(goalInvariantError) Then
                        session.Status = AgentStatus.Failed
                        OnStatusChanged?.Invoke(goalInvariantError)
                        Return AgentResult.Failed(session.Id, goalInvariantError)
                    End If

                    If decision Is Nothing Then
                        _memory.SetWorking(
                            "lastObservation",
                            "模型响应不是可解析的 act/complete/replan/fail 决策；请基于当前目标、最新观察和已注册工具重新决定。")
                        Continue While
                    End If

                    Select Case decision.Kind
                        Case "complete"
                            Dim currentContext = TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack)
                            Dim completionError = AgentGoalVerifier.Validate(
                                session,
                                currentContext,
                                decision.Evidence,
                                readOnlyEvidence.Count)
                            If String.IsNullOrWhiteSpace(completionError) Then
                                modelDeclaredComplete = True
                                completionMessage = decision.Message
                                MarkRemainingPlanHintsSkipped(session.Plan)
                                _memory.SetWorking(
                                    "lastObservation",
                                    "The model declared completion and deterministic goal verification passed.")
                                Exit While
                            End If

                            _memory.SetWorking(
                                "lastObservation",
                                $"Completion was rejected by deterministic goal verification: {completionError} Choose exactly one next action that closes this evidence gap.")
                            AppLogger.Warn("LoopEngine", $"Rejected premature completion: {completionError}")
                            Continue While

                        Case "replan"
                            If replanAttempts >= MaxReplanAttempts Then
                                _memory.SetWorking(
                                    "lastObservation",
                                    "Replan limit reached. Keep the current soft skeleton and choose one executable action, complete with evidence, or fail explicitly.")
                                Continue While
                            End If

                            replanAttempts += 1
                            session.Status = AgentStatus.Reflecting
                            OnStatusChanged?.Invoke("正在根据最新观察更新高层任务骨架...")
                            Dim replanned = Await GeneratePlanAsync(session, systemPrompt, skill)
                            If session.Spec.GoalContract Is Nothing OrElse
                               Not String.Equals(session.Spec.GoalContract.ContractHash, frozenGoalHash, StringComparison.Ordinal) Then
                                Dim invariantError = "GoalContract changed during replanning; execution stopped before further Office actions."
                                session.Status = AgentStatus.Failed
                                OnStatusChanged?.Invoke(invariantError)
                                Return AgentResult.Failed(session.Id, invariantError)
                            End If
                            If replanned Is Nothing OrElse replanned.Steps Is Nothing OrElse replanned.Steps.Count = 0 Then
                                _memory.SetWorking(
                                    "lastObservation",
                                    "Replan did not produce a useful high-level skeleton. The existing skeleton remains advisory; choose the next action directly.")
                            Else
                                session.Plan = replanned
                                OnPlanGenerated?.Invoke(session.Plan)
                            End If
                            Continue While

                        Case "fail"
                            Dim explicitFailure = If(
                                String.IsNullOrWhiteSpace(decision.Message),
                                "模型根据当前事实和工具确认任务无法安全继续。",
                                decision.Message)
                            session.Status = AgentStatus.Failed
                            OnStatusChanged?.Invoke(explicitFailure)
                            Return AgentResult.Failed(session.Id, explicitFailure)
                    End Select

                    Dim toolCall = decision.Action
                    If toolCall Is Nothing Then
                        _memory.SetWorking("lastObservation", "act 决策缺少工具调用；请选择一个已注册工具。")
                        Continue While
                    End If

                    Dim stepStartedAt = DateTime.Now
                    Dim undoPoint As Core.UndoManager.UndoPoint = Nothing
                    Dim toolResult As ToolResult = Nothing
                    Dim tool As ToolDescriptor = Nothing
                    Dim protectFromDuplicateSideEffects As Boolean = False
                    Dim blockedByGuard As Boolean = False
                    Dim evidenceId = $"obs-{session.Iterations.Count + 1}"

                    ' Runtime dependencies are bound before schema normalization. The model never
                    ' has to copy a previous tool's full payload through its token stream.
                    Dim inputDependencies = toolDataflow.BindInputsWithDependencies(toolCall)
                    Dim actionSignature = BuildActionSignature(toolCall)
                    Dim normalizeMessage As String = ""
                    If Not _toolRegistry.TryNormalizeToolCall(
                        session.AppType,
                        toolCall,
                        executionContext,
                        normalizeMessage) Then
                        toolResult = ToolResult.Failed(
                            toolCall.ToolId,
                            normalizeMessage,
                            New With {.availableTools = BuildAvailableToolHint(session.AppType, executionContext)},
                            ExceptionClassifier.CodeNotFound,
                            normalizeMessage,
                            normalizeMessage,
                            recoverable:=False,
                            observation:=New JObject From {
                                {"kind", "action_validation"},
                                {"summary", normalizeMessage},
                                {"changed", False}
                            })
                    Else
                        If Not String.IsNullOrWhiteSpace(normalizeMessage) Then
                            AppLogger.Info("LoopEngine", normalizeMessage)
                        End If
                        actionSignature = BuildActionSignature(toolCall)
                        tool = _toolRegistry.GetTool(toolCall.ToolId)
                        protectFromDuplicateSideEffects = ToolMayMutate(tool)
                        If IsActionBlocked(blockedActionSignatures, actionSignature, worldRevision) Then
                            blockedByGuard = True
                            toolResult = ToolResult.Failed(
                                toolCall.ToolId,
                                "相同工具与参数此前已返回不可原样重试的失败；未再次执行宿主操作",
                                errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                recoverable:=False,
                                observation:=New JObject From {
                                    {"kind", "duplicate_action_guard"},
                                    {"summary", "阻止重复执行不可原样重试的相同调用"},
                                    {"changed", False}
                                })
                        Else
                            If tool IsNot Nothing AndAlso
                               String.Equals(tool.RiskLevel, "risky", StringComparison.OrdinalIgnoreCase) Then
                                OnStatusChanged?.Invoke($"将执行高风险工具 {toolCall.ToolId}；安全策略可能要求确认。")
                            End If

                            session.Status = AgentStatus.Executing
                            If _undoManager IsNot Nothing Then
                                undoPoint = _undoManager.CreateUndoPoint(
                                    If(session.AppType, "Unknown"),
                                    $"决策 {decisionCount}: {toolCall.ToolId}",
                                    planStep.Description)
                            End If

                            toolResult = ValidateObservedOutcome(
                                Await _toolRegistry.ExecuteToolAsync(
                                    executionContext,
                                    toolCall.ToolId,
                                    toolCall.Parameters))

                            If toolResult IsNot Nothing AndAlso
                               Not toolResult.Success AndAlso
                               String.Equals(
                                   toolResult.ErrorCode,
                                   ExceptionClassifier.CodeSafetyNeedsApproval,
                                   StringComparison.OrdinalIgnoreCase) Then
                                If OnRequestApproval Is Nothing Then
                                    toolResult = ToolResult.Failed(
                                        toolCall.ToolId,
                                        "工具需要审批，但当前运行时没有审批处理器",
                                        errorCode:=ExceptionClassifier.CodeApprovalUnavailable,
                                        userMessage:="当前运行时无法请求审批，任务已安全停止",
                                        recoverable:=False,
                                        observation:=New JObject From {
                                            {"kind", "approval"},
                                            {"summary", "没有可用的审批处理器"},
                                            {"changed", False},
                                            {"warnings", New JArray("approval_handler_unavailable")}
                                        },
                                        taskFatal:=True)
                                Else
                                    session.Status = AgentStatus.WaitingApproval
                                    OnStatusChanged?.Invoke($"工具 {toolCall.ToolId} 正在等待用户确认...")
                                    Dim approved As Boolean = False
                                    Dim approvalResolved As Boolean = False
                                    Try
                                        approved = Await OnRequestApproval(If(toolResult.UserMessage, toolResult.Message))
                                        approvalResolved = True
                                    Catch ex As Exception
                                        toolResult = ToolResult.Failed(
                                            toolCall.ToolId,
                                            "审批请求处理失败",
                                            errorCode:=ExceptionClassifier.CodeApprovalUnavailable,
                                            userMessage:="无法完成审批交互，任务已安全停止",
                                            debugDetail:=ex.Message,
                                            recoverable:=False,
                                            observation:=New JObject From {
                                                {"kind", "approval"},
                                                {"summary", "审批处理器不可用"},
                                                {"changed", False},
                                                {"warnings", New JArray("approval_handler_failed")}
                                            },
                                            taskFatal:=True)
                                    End Try

                                    session.Status = AgentStatus.Executing
                                    If approvalResolved AndAlso approved Then
                                        executionContext.ApproveTool(toolCall.ToolId, toolCall.Parameters)
                                        toolResult = ValidateObservedOutcome(
                                            Await _toolRegistry.ExecuteToolAsync(
                                                executionContext,
                                                toolCall.ToolId,
                                                toolCall.Parameters))
                                    ElseIf approvalResolved Then
                                        toolResult = ToolResult.Failed(
                                            toolCall.ToolId,
                                            "用户拒绝高风险操作",
                                            errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                            userMessage:="已取消该高风险操作",
                                            recoverable:=False,
                                            observation:=New JObject From {
                                                {"kind", "approval"},
                                                {"summary", "用户拒绝高风险操作"},
                                                {"changed", False},
                                                {"warnings", New JArray("approval_rejected")}
                                            },
                                            taskFatal:=True)
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If toolResult Is Nothing Then
                        toolResult = ToolResult.Failed(
                            toolCall.ToolId,
                            "工具没有返回执行结果",
                            errorCode:=ExceptionClassifier.CodeUnknown,
                            recoverable:=False)
                    End If

                    If toolResult.Success AndAlso IsReadOnlyAnswerSpec(session.Spec) Then
                        Dim evidenceError = AppendReadOnlyEvidence(readOnlyEvidence, toolCall, toolResult)
                        If Not String.IsNullOrWhiteSpace(evidenceError) Then
                            toolResult = ToolResult.Failed(
                                toolCall.ToolId,
                                evidenceError,
                                errorCode:=ExceptionClassifier.CodeVerifyFailed,
                                userMessage:=evidenceError,
                                recoverable:=False,
                                observation:=New JObject From {
                                    {"kind", "read_evidence"},
                                    {"summary", evidenceError},
                                    {"changed", False}
                                })
                        End If
                    End If
                    toolResult = ApplyRetryPolicy(tool, toolResult)
                    If Not toolResult.Success AndAlso toolResult.Retryable Then
                        Dim retryFailureCount As Integer = 0
                        retryableFailureCounts.TryGetValue(actionSignature, retryFailureCount)
                        retryFailureCount += 1
                        retryableFailureCounts(actionSignature) = retryFailureCount
                        If retryFailureCount >= MaxIdenticalRetryAttempts Then
                            ' Bound identical-call retries even for read/compute tools. A subsequent
                            ' ReAct decision must change the action or wait for a new world revision.
                            toolResult.Retryable = False
                        End If
                    ElseIf toolResult.Success Then
                        retryableFailureCounts.Remove(actionSignature)
                    End If

                    If OutcomeEvidenceFactory.ObservationAdvancesWorld(tool, toolResult) Then worldRevision += 1
                    Dim evidenceContext = TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack)
                    Dim evidenceWorkbook = AgentGoalVerifier.ResolveContextWorkbookName(evidenceContext)
                    Dim contractEvidence = OutcomeEvidenceFactory.Create(
                        tool,
                        toolCall,
                        toolResult,
                        evidenceId,
                        inputDependencies,
                        worldRevision,
                        evidenceWorkbook)
                    If toolResult.Success Then toolDataflow.RecordSuccess(toolResult, evidenceId)
                    If Not toolResult.Success AndAlso
                       Not toolResult.Retryable AndAlso
                       Not blockedByGuard Then
                        BlockAction(
                            blockedActionSignatures,
                            actionSignature,
                            worldRevision,
                            protectFromDuplicateSideEffects AndAlso
                                Not ObservationConfirmsNoChange(toolResult))
                    End If

                    session.Status = AgentStatus.Observing
                    Dim observation = FormatObservation(toolResult)
                    _memory.SetWorking("lastObservation", observation)
                    Dim stepFinishedAt = DateTime.Now
                    Dim stepIndex As Integer = 0
                    If session.Plan IsNot Nothing AndAlso session.Plan.Steps IsNot Nothing Then
                        stepIndex = Math.Max(0, session.Plan.Steps.IndexOf(planStep))
                    End If
                    Dim explanation = BuildExecutionExplanation(
                        stepIndex,
                        planStep,
                        toolCall,
                        toolResult,
                        undoPoint,
                        observation,
                        stepStartedAt,
                        stepFinishedAt)
                    planStep.LastExplanation = explanation

                    Dim iteration = New ReActIteration With {
                        .Index = session.CurrentIteration,
                        .EvidenceId = evidenceId,
                        .Thought = rawDecision,
                        .Action = toolCall,
                        .AccessMode = If(tool?.AccessMode, ""),
                        .DependsOnEvidenceIds = New List(Of String)(inputDependencies),
                        .Observation = observation,
                        .OutcomeEvidence = CloneObservation(toolResult.Observation),
                        .OutcomeArtifacts = CloneToken(toolResult.Artifacts),
                        .ContractEvidence = contractEvidence,
                        .Explanation = explanation
                    }
                    session.Iterations.Add(iteration)
                    session.CurrentIteration += 1
                    OnExecutionExplained?.Invoke(explanation)
                    OnIterationUpdate?.Invoke(iteration)

                    UpdatePlanHintForObservation(session.Plan, planStep, toolCall, toolResult)

                    If toolResult.SessionFatal OrElse toolResult.TaskFatal Then
                        Dim fatalMessage = $"任务因终止性宿主错误停止：{toolResult.ToObserveSummary()}"
                        session.Status = AgentStatus.Failed
                        OnStatusChanged?.Invoke(fatalMessage)
                        Return AgentResult.Failed(
                            session.Id,
                            fatalMessage,
                            taskFatal:=toolResult.TaskFatal,
                            sessionFatal:=toolResult.SessionFatal,
                            errorCode:=toolResult.ErrorCode)
                    End If

                    ' Every ordinary failure is now just another observation. The next model
                    ' decision may change parameters, choose another tool, replan, or fail.
                End While

                If Not modelDeclaredComplete Then
                    session.Status = AgentStatus.Failed
                    Dim currentContext = TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack)
                    Dim acceptanceError = AgentGoalVerifier.Validate(
                        session,
                        currentContext,
                        Enumerable.Empty(Of String)(),
                        readOnlyEvidence.Count)
                    Dim failMsg = If(
                        String.IsNullOrWhiteSpace(acceptanceError),
                        $"任务尚未形成模型完成决策，已达到自适应决策上限 {iterationLimit}。",
                        $"任务尚未通过验收：{acceptanceError}")
                    OnStatusChanged?.Invoke(failMsg)
                    Return AgentResult.Failed(session.Id, failMsg)
                End If

                Dim finalOutput As String = ""
                If IsReadOnlyAnswerSpec(session.Spec) Then
                    OnStatusChanged?.Invoke("正在根据已读取的完整数据生成答案...")
                    finalOutput = Await GenerateReadOnlyAnswerAsync(session, readOnlyEvidence)
                    If String.IsNullOrWhiteSpace(finalOutput) Then
                        session.Status = AgentStatus.Failed
                        Dim answerFailure = "已读取工作簿数据，但未能生成可验证的最终答案；未对数据作任何修改"
                        OnStatusChanged?.Invoke(answerFailure)
                        Return AgentResult.Failed(session.Id, answerFailure)
                    End If
                End If

                session.Status = AgentStatus.Completed
                Dim finalMsg = If(
                    String.IsNullOrWhiteSpace(completionMessage),
                    $"任务完成，共执行 {session.CurrentIteration} 个工具迭代",
                    completionMessage)
                OnStatusChanged?.Invoke(finalMsg)
                Dim userMessage = If(String.IsNullOrWhiteSpace(finalOutput), finalMsg, finalOutput)
                Return AgentResult.SuccessResult(session.Id, userMessage, finalOutput)

            Catch ex As Exception
                session.Status = AgentStatus.Failed
                Dim classified = ExceptionClassifier.Classify(ex)
                OnStatusChanged?.Invoke($"执行出错: {classified.UserMessage}")
                AppLogger.Error("LoopEngine", "RunAsync unhandled exception", ex)
                Return AgentResult.Failed(
                    session.Id,
                    $"执行异常: [{classified.ErrorCode}] {classified.UserMessage}",
                    taskFatal:=classified.TaskFatal,
                    sessionFatal:=classified.SessionFatal,
                    errorCode:=classified.ErrorCode)
            End Try
        End Function

        Private Shared Function CreateFallbackPlan(session As AgentSession) As ExecutionPlan
            Dim goal = If(session?.Spec?.Goal, session?.UserRequest)
            Dim plan As New ExecutionPlan With {
                .Understanding = goal,
                .Summary = "运行时根据最新观察逐步选择动作"
            }
            plan.Steps.Add(New PlanStep With {
                .StepNumber = 1,
                .Description = If(String.IsNullOrWhiteSpace(goal), "完成用户目标", goal),
                .ToolHint = ""
            })
            Return plan
        End Function

        Private Shared Function GetCurrentPlanHint(plan As ExecutionPlan) As PlanStep
            If plan?.Steps IsNot Nothing Then
                Dim pending = plan.Steps.FirstOrDefault(
                    Function(item) item.Status <> StepStatus.Completed AndAlso item.Status <> StepStatus.Skipped)
                If pending IsNot Nothing Then
                    pending.Status = StepStatus.Running
                    Return pending
                End If
            End If

            Dim nextNumber As Integer = 1
            If plan IsNot Nothing AndAlso plan.Steps IsNot Nothing Then nextNumber = plan.Steps.Count + 1
            Return New PlanStep With {
                .StepNumber = nextNumber,
                .Description = "检查整体目标是否已由当前状态和真实观察满足",
                .ToolHint = "",
                .Status = StepStatus.Running
            }
        End Function

        Private Sub UpdatePlanHintForObservation(plan As ExecutionPlan,
                                                 currentHint As PlanStep,
                                                 toolCall As ToolCall,
                                                 result As ToolResult)
            If currentHint Is Nothing OrElse toolCall Is Nothing OrElse result Is Nothing Then Return

            Dim index As Integer = -1
            If plan IsNot Nothing AndAlso plan.Steps IsNot Nothing Then index = plan.Steps.IndexOf(currentHint)
            If Not result.Success Then
                currentHint.Status = StepStatus.Failed
                currentHint.ErrorMessage = result.ToObserveSummary()
                If index >= 0 Then OnStepCompleted?.Invoke(index, False, If(result.UserMessage, result.Message))
                Return
            End If

            currentHint.ErrorMessage = ""
            If ActionMatchesPlanHint(currentHint, toolCall) Then
                currentHint.Status = StepStatus.Completed
                If index >= 0 Then OnStepCompleted?.Invoke(index, True, result.Message)
            Else
                currentHint.Status = StepStatus.Running
                OnStatusChanged?.Invoke($"动作 {toolCall.ToolId} 已成功；任务骨架仅作提示，下一轮将按最新状态继续决策。")
            End If
        End Sub

        Private Function ActionMatchesPlanHint(planStep As PlanStep,
                                               toolCall As ToolCall) As Boolean
            If planStep Is Nothing OrElse toolCall Is Nothing Then Return False
            Dim expectedTool = If(planStep.ToolHint, "").Trim()
            If String.IsNullOrWhiteSpace(expectedTool) AndAlso Not String.IsNullOrWhiteSpace(planStep.Code) Then
                expectedTool = If(ParsePlannedToolCall(planStep.Code)?.ToolId, "").Trim()
            End If
            If String.IsNullOrWhiteSpace(expectedTool) Then Return False
            Return String.Equals(expectedTool, toolCall.ToolId, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Sub MarkRemainingPlanHintsSkipped(plan As ExecutionPlan)
            If plan?.Steps Is Nothing Then Return
            For Each item In plan.Steps
                If item.Status = StepStatus.Completed Then Continue For
                item.Status = StepStatus.Skipped
                If String.IsNullOrWhiteSpace(item.ErrorMessage) Then
                    item.ErrorMessage = "整体目标已由当前 World Snapshot / Observation 验证满足"
                End If
            Next
        End Sub

        Private Shared Function BuildActionSignature(toolCall As ToolCall) As String
            If toolCall Is Nothing Then Return ""
            Dim parameters = If(toolCall.Parameters Is Nothing,
                                 "{}",
                                 CanonicalizeJsonToken(toolCall.Parameters).
                                     ToString(Newtonsoft.Json.Formatting.None))
            Return If(toolCall.ToolId, "").Trim().ToLowerInvariant() & "|" & parameters
        End Function

        Private Shared Function CanonicalizeJsonToken(token As JToken) As JToken
            If token Is Nothing Then Return JValue.CreateNull()
            If token.Type = JTokenType.Object Then
                Dim canonicalObject As New JObject()
                For Each prop In DirectCast(token, JObject).
                    Properties().
                    OrderBy(Function(item) item.Name, StringComparer.Ordinal)
                    canonicalObject.Add(prop.Name, CanonicalizeJsonToken(prop.Value))
                Next
                Return canonicalObject
            End If
            If token.Type = JTokenType.Array Then
                Dim canonicalArray As New JArray()
                For Each item In DirectCast(token, JArray)
                    canonicalArray.Add(CanonicalizeJsonToken(item))
                Next
                Return canonicalArray
            End If
            Return token.DeepClone()
        End Function

        Private Shared Function IsActionBlocked(blocked As Dictionary(Of String, BlockedActionState),
                                                signature As String,
                                                worldRevision As Long) As Boolean
            If blocked Is Nothing OrElse String.IsNullOrWhiteSpace(signature) Then Return False
            Dim state As BlockedActionState = Nothing
            If Not blocked.TryGetValue(signature, state) OrElse state Is Nothing Then Return False
            If state.Permanent OrElse state.WorldRevision = worldRevision Then Return True

            ' A deterministic no-change failure belongs to the world snapshot in which it
            ' occurred. Once another verified action changes the host, the same call may be valid.
            blocked.Remove(signature)
            Return False
        End Function

        Private Shared Sub BlockAction(blocked As Dictionary(Of String, BlockedActionState),
                                       signature As String,
                                       worldRevision As Long,
                                       permanent As Boolean)
            If blocked Is Nothing OrElse String.IsNullOrWhiteSpace(signature) Then Return
            Dim existing As BlockedActionState = Nothing
            If blocked.TryGetValue(signature, existing) AndAlso existing IsNot Nothing Then
                If existing.Permanent Then Return
                existing.Permanent = permanent
                existing.WorldRevision = worldRevision
                Return
            End If
            blocked(signature) = New BlockedActionState With {
                .WorldRevision = worldRevision,
                .Permanent = permanent
            }
        End Sub

        Private Shared Function ApplyRetryPolicy(tool As ToolDescriptor,
                                                 result As ToolResult) As ToolResult
            If result Is Nothing OrElse result.Success Then Return result
            If result.TaskFatal OrElse result.SessionFatal Then
                result.Retryable = False
                Return result
            End If

            ' Identical calls are automatically retryable only when the tool contract proves
            ' that no Office state can be mutated. A mutating executor may still recover by
            ' choosing different parameters/tooling, but never by blindly repeating the write.
            If ToolMayMutate(tool) Then
                result.Retryable = False
                Return result
            End If

            ' Retryability is derived from the stable error taxonomy, never trusted from an
            ' adapter flag. Otherwise a syntax/validation failure could opt itself into an
            ' identical retry and waste another model/tool cycle.
            result.Retryable = False
            Select Case If(result.ErrorCode, "").Trim().ToUpperInvariant()
                Case ExceptionClassifier.CodeNetwork,
                     ExceptionClassifier.CodeTimeout,
                     ExceptionClassifier.CodeCom,
                     ExceptionClassifier.CodeIo
                    result.Retryable = True
            End Select
            Return result
        End Function

        Private Shared Function ObservationConfirmsChange(result As ToolResult) As Boolean
            Return ReadObservationBoolean(result, "changed", True)
        End Function

        Private Shared Function ObservationConfirmsNoChange(result As ToolResult) As Boolean
            Return ReadObservationBoolean(result, "changed", False)
        End Function

        Private Shared Function ReadObservationBoolean(result As ToolResult,
                                                       propertyName As String,
                                                       expectedValue As Boolean) As Boolean
            If result Is Nothing OrElse result.Observation Is Nothing Then Return False
            Try
                Dim token = TryCast(result.Observation, JToken)
                If token Is Nothing Then token = JToken.FromObject(result.Observation)
                Dim value = token?(propertyName)
                Return value IsNot Nothing AndAlso
                       value.Type = JTokenType.Boolean AndAlso
                       value.Value(Of Boolean)() = expectedValue
            Catch
                Return False
            End Try
        End Function

        Private Shared Function EstablishFrozenGoalContract(session As AgentSession) As String
            If session?.Spec Is Nothing Then Return "Task specification is missing; the user goal cannot be frozen."
            Try
                If String.IsNullOrWhiteSpace(session.Spec.RawUserRequest) Then
                    Dim rawRequest = If(session.UserRequest, "")
                    If String.IsNullOrWhiteSpace(rawRequest) Then rawRequest = session.Spec.Goal
                    session.Spec.CaptureRawUserRequest(rawRequest)
                End If

                If session.Spec.GoalContract IsNot Nothing Then Return ""

                ' Freeze only semantics derived from the exact user request. TaskSpec fields
                ' are mutable legacy projections and must never flow back into GoalContract.
                Dim compilation = Goals.GoalCompiler.Compile(session.Spec.RawUserRequest)
                Dim validation = Goals.GoalCoverageValidator.Validate(compilation)
                If Not validation.Succeeded Then
                    Return "Goal compilation failed: " & String.Join("; ", validation.Errors)
                End If

                Dim frozenGoal = Goals.GoalContractFreezer.Freeze(compilation, validation)
                session.Spec.SetGoalContractOnce(frozenGoal)
                Return ""
            Catch ex As Exception
                Return "Goal compilation failed: " & ex.Message
            End Try
        End Function

        Private Shared Function ValidateFrozenGoalInvariant(session As AgentSession, expectedHash As String) As String
            If session?.Spec?.GoalContract Is Nothing Then
                Return "Frozen GoalContract is missing; execution stopped before further Office actions."
            End If
            If Not String.Equals(session.Spec.GoalContract.ContractHash, expectedHash, StringComparison.Ordinal) Then
                Return "GoalContract changed during execution; execution stopped before further Office actions."
            End If
            Return ""
        End Function

        Private Shared Function ToolMayMutate(tool As ToolDescriptor) As Boolean
            If tool Is Nothing Then Return False
            Dim mode = If(tool.AccessMode, "").Trim().ToLowerInvariant()
            Return mode <> "read" AndAlso mode <> "compute"
        End Function

        Private Class BlockedActionState
            Public Property WorldRevision As Long
            Public Property Permanent As Boolean
        End Class

        Private Shared Function CloneObservation(value As Object) As JObject
            If value Is Nothing Then Return Nothing
            Try
                Dim token = TryCast(value, JToken)
                If token Is Nothing Then token = JToken.FromObject(value)
                Dim obj = TryCast(token, JObject)
                If obj Is Nothing Then Return Nothing
                Return DirectCast(obj.DeepClone(), JObject)
            Catch
                Return Nothing
            End Try
        End Function

        Public ReadOnly Property UndoManager As Core.UndoManager
            Get
                Return _undoManager
            End Get
        End Property
    End Class

End Namespace
