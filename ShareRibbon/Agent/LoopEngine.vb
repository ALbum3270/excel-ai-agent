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
                                        Optional skill As AgentSkill = Nothing,
                                        Optional cancellationToken As Threading.CancellationToken = Nothing) As Task(Of AgentResult)
            Dim decisionCount As Integer = 0
            Dim replanAttempts As Integer = 0
            Dim modelDeclaredComplete As Boolean = False
            Dim completionMessage As String = ""
            Dim readOnlyEvidence As New JArray()
            Dim toolDataflow As New AgentToolDataflow()
            Dim blockedActionSignatures As New Dictionary(Of String, BlockedActionState)(StringComparer.Ordinal)
            Dim successfulMutationRevisions As New Dictionary(Of String, Long)(StringComparer.Ordinal)
            Dim retryableFailureCounts As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim worldRevision As Long = 0
            Dim frozenGoalHash As String = ""

            Try
                cancellationToken.ThrowIfCancellationRequested()
                OnStatusChanged?.Invoke("正在分析任务...")
                If session.Spec Is Nothing Then
                    session.Spec = Await AwaitWithCancellation(GenerateSpecAsync(session), cancellationToken)
                Else
                    AppLogger.Info("LoopEngine", $"Use precomputed TaskSpec goal={AppLogger.Redact(session.Spec.Goal)}")
                End If

                OnStatusChanged?.Invoke("正在确认用户目标...")
                Dim goalContractError = EstablishFrozenGoalContract(session)
                If Not String.IsNullOrWhiteSpace(goalContractError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(goalContractError)
                    Return AgentResult.Failed(session.Id, goalContractError)
                End If
                frozenGoalHash = session.Spec.GoalContract.ContractHash

                Dim admissionError = Goals.GoalExecutionAdmission.Validate(session.Spec)
                If Not String.IsNullOrWhiteSpace(admissionError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(admissionError)
                    Return AgentResult.Failed(session.Id, admissionError)
                End If

                _memory.BeginTaskContext(TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack))

                cancellationToken.ThrowIfCancellationRequested()
                Dim outcomeContractError = Await AwaitWithCancellation(
                    PrepareInitialPlanAsync(session, systemPrompt, skill),
                    cancellationToken)
                If Not String.IsNullOrWhiteSpace(outcomeContractError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(outcomeContractError)
                    Return AgentResult.Failed(session.Id, outcomeContractError)
                End If
                OnPlanGenerated?.Invoke(session.Plan)

                session.Status = AgentStatus.Executing
                OnStatusChanged?.Invoke("用户目标已确认；将根据最新观察逐步决定下一动作...")
                Dim executionContext = ToolExecutionContext.FromSession(session, skill)
                Dim iterationLimit = Math.Min(MaxIterations, Math.Max(1, session.MaxIterations))

                While decisionCount < iterationLimit AndAlso Not modelDeclaredComplete
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim goalInvariantError = ValidateFrozenGoalInvariant(session, frozenGoalHash)
                    If Not String.IsNullOrWhiteSpace(goalInvariantError) Then
                        session.Status = AgentStatus.Failed
                        OnStatusChanged?.Invoke(goalInvariantError)
                        Return AgentResult.Failed(session.Id, goalInvariantError)
                    End If
                    decisionCount += 1
                    Dim planStep = CreateCurrentDecisionStep(decisionCount)

                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"正在根据最新观察选择下一动作（自适应决策第 {decisionCount} 轮）...")
                    Dim rawDecision = Await AwaitWithCancellation(
                        ThinkAsync(session, planStep, systemPrompt),
                        cancellationToken)
                    cancellationToken.ThrowIfCancellationRequested()
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

                    If HasConflictingCompletionPayload(decision) Then
                        _memory.SetWorking(
                            "lastObservation",
                            "The previous model response was rejected before execution because decision=act also carried a final message or completion evidence/outcomeContract. No tool was called. Choose exactly one control state: return decision=complete with message and evidence if the Goal is satisfied, otherwise return decision=act with exactly one action and no completion payload.")
                        Continue While
                    End If

                    Select Case decision.Kind
                        Case "complete"
                            If TryAcceptCompletionDecision(
                                session,
                                decision,
                                readOnlyEvidence.Count,
                                completionMessage) Then
                                modelDeclaredComplete = True
                                Exit While
                            End If
                            Continue While

                        Case "replan"
                            If replanAttempts >= MaxReplanAttempts Then
                                _memory.SetWorking(
                                    "lastObservation",
                                    "Strategy reset limit reached. Choose one executable next action from the latest observations, complete with evidence, or fail explicitly.")
                                Continue While
                            End If

                            replanAttempts += 1
                            session.Status = AgentStatus.Reflecting
                            OnStatusChanged?.Invoke("正在根据最新观察重新评估下一动作...")
                            Dim observationBeforeStrategyReset = _memory.GetWorkingString("lastObservation")
                            _memory.SetWorking(
                                "lastObservation",
                                $"Latest host observation: {observationBeforeStrategyReset}{vbCrLf}The previous decision requested a strategy reset: {If(decision.Message, decision.Thought)}. No future-step plan was generated. Re-evaluate the frozen goal and choose exactly one next action from current observations.")
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
                    Dim normalizeErrorCode As String = ExceptionClassifier.CodeNotFound
                    If Not _toolRegistry.TryNormalizeToolCall(
                        session.AppType,
                        toolCall,
                        executionContext,
                        normalizeMessage,
                        normalizeErrorCode) Then
                        toolResult = ToolResult.Failed(
                            toolCall.ToolId,
                            normalizeMessage,
                            New With {.availableTools = BuildAvailableToolHint(session.AppType, executionContext)},
                            normalizeErrorCode,
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
                        If protectFromDuplicateSideEffects AndAlso
                           WasSuccessfulMutationAlreadyApplied(
                               successfulMutationRevisions,
                               actionSignature,
                               worldRevision) Then
                            blockedByGuard = True
                            toolResult = ToolResult.Failed(
                                toolCall.ToolId,
                                "The identical mutating action already succeeded in the current Office state; the host call was not repeated.",
                                errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                userMessage:="相同修改已经成功执行，系统已阻止重复操作",
                                recoverable:=True,
                                observation:=New JObject From {
                                    {"kind", "duplicate_success_guard"},
                                    {"summary", "Identical successful mutation was not dispatched again"},
                                    {"changed", False},
                                    {"actionSignature", actionSignature},
                                    {"worldRevision", worldRevision}
                                })
                        ElseIf IsActionBlocked(blockedActionSignatures, actionSignature, worldRevision) Then
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

                            cancellationToken.ThrowIfCancellationRequested()
                            toolResult = ValidateObservedOutcome(
                                Await _toolRegistry.ExecuteToolAsync(
                                    executionContext,
                                    toolCall.ToolId,
                                    toolCall.Parameters,
                                    cancellationToken))
                            ' Office COM calls are not interrupted halfway through.  Cancellation
                            ' is observed immediately after the host returns, before another model
                            ' or tool action can start.
                            cancellationToken.ThrowIfCancellationRequested()

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
                                        approved = Await AwaitWithCancellation(
                                            OnRequestApproval(If(toolResult.UserMessage, toolResult.Message)),
                                            cancellationToken)
                                        approvalResolved = True
                                    Catch ex As OperationCanceledException
                                        Throw
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
                                        cancellationToken.ThrowIfCancellationRequested()
                                        toolResult = ValidateObservedOutcome(
                                            Await _toolRegistry.ExecuteToolAsync(
                                                executionContext,
                                                toolCall.ToolId,
                                                toolCall.Parameters,
                                                cancellationToken))
                                        cancellationToken.ThrowIfCancellationRequested()
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
                    If toolResult.Success AndAlso
                       protectFromDuplicateSideEffects AndAlso
                       HasRequestBoundVerification(toolResult) Then
                        successfulMutationRevisions(actionSignature) = worldRevision
                    End If
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
                    _memory.SetWorking("lastObservation", FormatModelObservation(toolResult, observation))
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
                        .OutcomeArtifacts = OutcomeProjectionValue.CloneToken(toolResult.Artifacts),
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
                    Dim failMsg =
                        $"任务未完成：已达到自适应决策上限 {iterationLimit}，但模型未返回明确的 decision=complete 或 decision=fail；最后一次观察已保留，未执行终局验收或额外 Office 写入。"
                    OnStatusChanged?.Invoke(failMsg)
                    Return AgentResult.Failed(session.Id, failMsg)
                End If

                ' The same model turn that observes the final tool result must produce the
                ' user-facing response.  Tool execution never terminates the loop by itself,
                ' and no second answer-only model pipeline is needed after completion.
                Dim finalOutput As String = completionMessage

                session.Status = AgentStatus.Completed
                Dim finalMsg = If(
                    String.IsNullOrWhiteSpace(completionMessage),
                    $"任务完成，共执行 {session.CurrentIteration} 个工具迭代",
                    completionMessage)
                OnStatusChanged?.Invoke(finalMsg)
                Return AgentResult.SuccessResult(session.Id, finalMsg, finalOutput)

            Catch ex As OperationCanceledException
                session.Status = AgentStatus.Failed
                OnStatusChanged?.Invoke("已终止")
                Return AgentResult.Failed(
                    session.Id,
                    "已终止",
                    errorCode:=ExceptionClassifier.CodeCancelled)
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

        ''' <summary>
        ''' Stop waiting for non-mutating asynchronous work (primarily model calls) as
        ''' soon as the request token is cancelled. The model transport receives the same
        ''' token and shuts down cooperatively. Host mutation tools deliberately use boundary
        ''' checks instead, so a COM call is never abandoned mid-write.
        ''' </summary>
        Private Shared Async Function AwaitWithCancellation(Of T)(
            operation As Task(Of T),
            cancellationToken As Threading.CancellationToken) As Task(Of T)

            If operation Is Nothing Then Throw New ArgumentNullException(NameOf(operation))
            If Not cancellationToken.CanBeCanceled Then Return Await operation

            Dim cancellationSignal = New TaskCompletionSource(Of Boolean)(
                TaskCreationOptions.RunContinuationsAsynchronously)
            Using cancellationRegistration = cancellationToken.Register(
                Sub() cancellationSignal.TrySetResult(True))
                Dim completed = Await Task.WhenAny(operation, cancellationSignal.Task)
                If completed IsNot operation Then ObserveAbandonedTask(operation)
                cancellationToken.ThrowIfCancellationRequested()
                Return Await operation
            End Using
        End Function

        Private Shared Sub ObserveAbandonedTask(Of T)(operation As Task(Of T))
            operation.ContinueWith(
                Sub(faultedTask)
                    Dim ignored = faultedTask.Exception
                End Sub,
                Threading.CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted Or TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default)
        End Sub

        Private Shared Function CreateFallbackPlan(session As AgentSession) As ExecutionPlan
            Dim goal = If(session?.Spec?.Goal, session?.UserRequest)
            Return New ExecutionPlan With {
                .Understanding = goal,
                .Summary = "运行时根据最新观察逐步选择动作，不预生成未来步骤"
            }
        End Function

        Private Shared Function CreateCurrentDecisionStep(decisionNumber As Integer) As PlanStep
            Return New PlanStep With {
                .StepNumber = Math.Max(1, decisionNumber),
                .Description = "根据最新 ContextPack、Observation 与证据选择当前唯一下一动作",
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

                ' Freeze only semantics derived from the exact captured request. TaskSpec fields
                ' are mutable legacy projections and must never flow back into GoalContract.
                Dim compilation = session.Spec.GoalCompilation
                If compilation Is Nothing Then
                    session.Spec.RecordGoalInterpretationFallback(
                        "No intake GoalCompilation was attached; exact captured request was preserved.")
                    compilation = New Goals.RawPreservingGoalInterpretationAdapter(
                        "No intake GoalCompilation was attached; LoopEngine preserved the exact captured request.").
                        Interpret(session.Spec.RawUserRequest)
                End If
                Dim validation = Goals.GoalCoverageValidator.Validate(compilation)
                If Not validation.Succeeded Then
                    If compilation.RequiresClarification Then
                        Return "Goal compilation requires clarification: " & String.Join("; ", validation.Errors)
                    End If

                    ' A malformed model interpretation may not block the task or import guessed
                    ' meaning. Degrade to one opaque exact-text semantic criterion.
                    session.Spec.RecordGoalInterpretationFallback(
                        "Structured interpretation rejected by deterministic validation: " & String.Join("; ", validation.Errors))
                    compilation = New Goals.RawPreservingGoalInterpretationAdapter(
                        "Structured goal interpretation failed deterministic validation: " & String.Join("; ", validation.Errors)).
                        Interpret(session.Spec.RawUserRequest)
                    validation = Goals.GoalCoverageValidator.Validate(compilation)
                    If Not validation.Succeeded Then
                        Return "Goal compilation failed: " & String.Join("; ", validation.Errors)
                    End If
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
