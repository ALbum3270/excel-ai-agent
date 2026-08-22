Imports System.Text
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' ReAct 循环引擎 - 核心执行逻辑
    ''' Think -> Plan -> Act -> Observe -> Reflect
    ''' </summary>
    Public Partial Class LoopEngine
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _undoManager As Core.UndoManager
        Private _multimodalRepairDisabledForRun As Boolean

        ' 循环限制
        Private Const MaxIterations As Integer = 15
        Private Const MaxNoProgress As Integer = 3
        Private Const MaxReplanAttempts As Integer = 2

        ' 回调
        Public Property OnStatusChanged As Action(Of String)
        Public Property OnIterationUpdate As Action(Of ReActIteration)
        Public Property OnStepCompleted As Action(Of Integer, Boolean, String)
        Public Property OnExecutionExplained As Action(Of ExecutionExplanation)
        Public Property OnRequestApproval As Func(Of String, Task(Of Boolean))
        Public Property OnPlanGenerated As Action(Of ExecutionPlan)
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, Task(Of String))
        Public Property CaptureContextPack As Func(Of Context.ContextPack)

        Public Sub New(toolRegistry As ToolRegistry, memory As AgentMemory, promptManager As PromptManager)
            _toolRegistry = toolRegistry
            _memory = memory
            _promptManager = promptManager
            _undoManager = New Core.UndoManager()
        End Sub

        ''' <summary>
        ''' 执行 ReAct 循环
        ''' </summary>
        Public Async Function RunAsync(session As AgentSession,
                                        systemPrompt As String,
                                        Optional skill As AgentSkill = Nothing) As Task(Of AgentResult)
            Dim noProgressCount As Integer = 0
            Dim replanAttempts As Integer = 0
            Dim decisionCount As Integer = 0
            Dim modelDeclaredComplete As Boolean = False
            Dim readOnlyEvidence As New JArray()
            Dim toolDataflow As New AgentToolDataflow()
            _multimodalRepairDisabledForRun = False

            Try
                ' Phase 1: 生成 Spec
                OnStatusChanged?.Invoke("正在分析任务...")
                If session.Spec Is Nothing Then
                    session.Spec = Await GenerateSpecAsync(session)
                Else
                    AppLogger.Info("LoopEngine", $"Use precomputed TaskSpec goal={AppLogger.Redact(session.Spec.Goal)}")
                End If

                ' Phase 2: 生成计划
                OnStatusChanged?.Invoke("正在制定执行计划...")
                session.Plan = Await GeneratePlanAsync(session, systemPrompt, skill)
                If session.Plan Is Nothing OrElse session.Plan.Steps.Count = 0 Then
                    Dim capabilityGap = If(session.Plan?.CapabilityGap, "")
                    Dim planFailure = If(String.IsNullOrWhiteSpace(capabilityGap),
                                         "规划失败：模型未生成可执行计划",
                                         "当前能力无法完整执行：" & capabilityGap)
                    OnStatusChanged?.Invoke(planFailure)
                    Return AgentResult.Failed(session.Id, planFailure)
                End If

                Dim coverageError = ValidatePlanCoverage(session.Spec, session.Plan)
                If Not String.IsNullOrWhiteSpace(coverageError) Then
                    OnStatusChanged?.Invoke(coverageError)
                    Return AgentResult.Failed(session.Id, coverageError)
                End If

                ' 通知计划已生成
                OnPlanGenerated?.Invoke(session.Plan)

                session.Status = AgentStatus.Executing
                OnStatusChanged?.Invoke($"规划完成（共 {session.Plan.Steps.Count} 步），进入自治执行...")
                Dim executionContext As ToolExecutionContext = ToolExecutionContext.FromSession(session, skill)

                ' Phase 3: ReAct Loop
                Dim stepIndex As Integer = 0
                While decisionCount < MaxIterations AndAlso Not modelDeclaredComplete
                    Dim isCompletionCheck = stepIndex >= session.Plan.Steps.Count
                    Dim planStep As PlanStep
                    If isCompletionCheck Then
                        planStep = New PlanStep With {
                            .StepNumber = session.Plan.Steps.Count + 1,
                            .Description = "检查整体目标是否已由真实观察满足；如未满足，选择当前唯一下一动作"
                        }
                    Else
                        planStep = session.Plan.Steps(stepIndex)
                    End If
                    planStep.Status = StepStatus.Running
                    decisionCount += 1

                    ' --- THINK ---
                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"步骤 {stepIndex + 1}/{session.Plan.Steps.Count}: {planStep.Description}")
                    ' The plan is only a high-level task skeleton. Tool parameters that depend
                    ' on future observations cannot be authoritative at planning time, so every
                    ' step asks the model for the current action after the previous observation
                    ' has been recorded. This is the adaptive ReAct seam.
                    OnStatusChanged?.Invoke($"步骤 {stepIndex + 1}/{session.Plan.Steps.Count}: 正在结合最新观察决定下一动作...")
                    Dim thought = Await ThinkAsync(session, planStep, systemPrompt)
                    Dim decision = ParseReactDecision(thought)

                    If decision IsNot Nothing AndAlso decision.Kind = "complete" Then
                        Dim unfinishedMilestones = session.Plan.Steps.
                            Where(Function(item) item.Status <> StepStatus.Completed).
                            Count()
                        Dim completionError = If(unfinishedMilestones > 0,
                                                 $"高层任务骨架仍有 {unfinishedMilestones} 个步骤未由匹配的成功工具观察完成。",
                                                 ValidateExecutionOutcome(session))
                        If String.IsNullOrWhiteSpace(completionError) Then
                            modelDeclaredComplete = True
                            noProgressCount = 0
                            _memory.SetWorking("lastObservation", "The model declared completion and deterministic acceptance passed.")
                            Exit While
                        End If

                        noProgressCount += 1
                        _memory.SetWorking("lastObservation",
                                           $"Completion was rejected by deterministic acceptance: {completionError} Choose the next action that closes this gap.")
                        AppLogger.Warn("LoopEngine", $"Rejected premature completion: {completionError}")
                        If noProgressCount >= MaxNoProgress Then Exit While
                        Continue While
                    End If

                    If decision IsNot Nothing AndAlso decision.Kind = "replan" Then
                        If replanAttempts >= MaxReplanAttempts Then
                            _memory.SetWorking("lastObservation", "Replan limit reached. Choose an executable action or fail explicitly.")
                            noProgressCount += 1
                            If noProgressCount >= MaxNoProgress Then Exit While
                            Continue While
                        End If

                        replanAttempts += 1
                        session.Status = AgentStatus.Reflecting
                        OnStatusChanged?.Invoke("正在根据最新观察重新规划...")
                        Dim replanned = Await GeneratePlanAsync(session, systemPrompt, skill)
                        If replanned Is Nothing OrElse replanned.Steps.Count = 0 Then
                            noProgressCount += 1
                            _memory.SetWorking("lastObservation", "Replan did not produce an executable high-level skeleton.")
                            If noProgressCount >= MaxNoProgress Then Exit While
                            Continue While
                        End If

                        Dim replanCoverageError = ValidatePlanCoverage(session.Spec, replanned)
                        If Not String.IsNullOrWhiteSpace(replanCoverageError) Then
                            noProgressCount += 1
                            _memory.SetWorking("lastObservation", $"Replan was rejected: {replanCoverageError}")
                            If noProgressCount >= MaxNoProgress Then Exit While
                            Continue While
                        End If

                        session.Plan = replanned
                        stepIndex = 0
                        noProgressCount = 0
                        OnPlanGenerated?.Invoke(session.Plan)
                        Continue While
                    End If

                    If decision IsNot Nothing AndAlso decision.Kind = "fail" Then
                        Dim explicitFailure = If(String.IsNullOrWhiteSpace(decision.Message),
                                                 "The model determined that the task cannot be completed safely with the current facts and tools.",
                                                 decision.Message)
                        session.Status = AgentStatus.Failed
                        OnStatusChanged?.Invoke(explicitFailure)
                        Return AgentResult.Failed(session.Id, explicitFailure)
                    End If

                    Dim toolCall = If(decision?.Action, Nothing)

                    ' --- PARSE ACTION ---
                    If toolCall Is Nothing Then
                        noProgressCount += 1
                        planStep.ErrorMessage = "无法解析当前 ReAct 决策"

                        _memory.SetWorking("lastObservation",
                                           $"步骤 {stepIndex + 1} 的模型响应不是可解析的 act/complete/replan/fail 决策；请基于当前目标和已注册工具重新决定。")
                        If noProgressCount >= MaxNoProgress Then Exit While
                        Continue While
                    End If

                    If isCompletionCheck Then
                        planStep.Description = $"完成目标所需的附加动作：{toolCall.ToolId}"
                        planStep.ToolHint = toolCall.ToolId
                        session.Plan.Steps.Add(planStep)
                    End If

                    Dim normalizeMessage As String = ""
                    If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, executionContext, normalizeMessage) Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = normalizeMessage
                        OnStepCompleted?.Invoke(stepIndex, False, normalizeMessage)

                        _memory.SetWorking("lastObservation",
                                           $"工具调用未通过规范化：{normalizeMessage}。请使用已注册工具的原始 ID 和有效参数重新决定。")
                        If noProgressCount >= MaxNoProgress Then Exit While
                        Continue While
                    End If
                    If Not String.IsNullOrWhiteSpace(normalizeMessage) Then
                        Debug.WriteLine($"[LoopEngine] {normalizeMessage}")
                    End If
                    Dim contractedToolId = If(AgentExecutionContract.IsMandatoryTool(session.Spec, toolCall.ToolId),
                                              toolCall.ToolId,
                                              "")

                    ' --- RISK NOTICE (autonomous mode) ---
                    Dim tool = _toolRegistry.GetTool(toolCall.ToolId)
                    If tool IsNot Nothing AndAlso tool.RiskLevel = "risky" Then
                        Debug.WriteLine($"[LoopEngine] 自治模式执行高风险工具: {toolCall.ToolId}")
                        OnStatusChanged?.Invoke($"步骤 {stepIndex + 1} 使用高风险工具 {toolCall.ToolId}，已记录风险并继续执行")
                    End If

                    ' --- ACT (增强版 - 多轮自修复 + 撤销点) ---
                    session.Status = AgentStatus.Executing

                    ' 创建撤销点（执行前保存状态）
                    Dim undoPoint As Core.UndoManager.UndoPoint = Nothing
                    If _undoManager IsNot Nothing Then
                        undoPoint = _undoManager.CreateUndoPoint(
                            If(session.AppType, "Unknown"),
                            $"步骤 {stepIndex + 1}: {toolCall.ToolId}",
                            planStep.Description)
                        If undoPoint IsNot Nothing Then
                            Debug.WriteLine($"[LoopEngine] 创建撤销点: {undoPoint.Name}")
                        End If
                    End If

                    Const MaxFixAttempts As Integer = 3
                    Dim fixAttempt As Integer = 0
                    Dim toolResult As ToolResult = Nothing
                    Dim stepStartedAt = DateTime.Now

                    ' 多轮修复循环
                    While fixAttempt < MaxFixAttempts
                        ' A plan is produced before read/compute outputs exist. Bind the actual
                        ' successful runtime result before schema validation and execution.
                        toolDataflow.BindInputs(toolCall)
                        Dim normalized As String = ""
                        If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, executionContext, normalized) Then
                            toolResult = ToolResult.Failed(toolCall.ToolId,
                                                           normalized,
                                                           New With {.availableTools = BuildAvailableToolHint(session.AppType, executionContext)},
                                                           ExceptionClassifier.CodeNotFound,
                                                           normalized,
                                                           normalized,
                                                           recoverable:=True)
                        Else
                            If Not String.IsNullOrWhiteSpace(normalized) Then Debug.WriteLine($"[LoopEngine] {normalized}")
                            ' 执行工具
                            toolResult = ValidateObservedOutcome(
                                Await _toolRegistry.ExecuteToolAsync(executionContext, toolCall.ToolId, toolCall.Parameters))
                        End If

                        If Not toolResult.Success AndAlso
                           String.Equals(toolResult.ErrorCode, ExceptionClassifier.CodeSafetyNeedsApproval, StringComparison.OrdinalIgnoreCase) AndAlso
                           OnRequestApproval IsNot Nothing Then
                            OnStatusChanged?.Invoke($"工具 {toolCall.ToolId} 正在等待用户确认...")
                            Dim approved = Await OnRequestApproval(If(toolResult.UserMessage, toolResult.Message))
                            If approved Then
                                executionContext.ApproveTool(toolCall.ToolId, toolCall.Parameters)
                                OnStatusChanged?.Invoke($"用户已批准工具 {toolCall.ToolId}，继续执行...")
                                toolResult = ValidateObservedOutcome(
                                    Await _toolRegistry.ExecuteToolAsync(executionContext, toolCall.ToolId, toolCall.Parameters))
                            Else
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
                                    })
                            End If
                        End If

                        If toolResult.Success Then
                            toolDataflow.RecordSuccess(toolResult)
                            If IsReadOnlyAnswerSpec(session.Spec) Then
                                Dim evidenceError = AppendReadOnlyEvidence(readOnlyEvidence, toolCall, toolResult)
                                If Not String.IsNullOrWhiteSpace(evidenceError) Then
                                    session.Status = AgentStatus.Failed
                                    OnStatusChanged?.Invoke(evidenceError)
                                    Return AgentResult.Failed(session.Id, evidenceError)
                                End If
                            End If
                            ' 成功，跳出循环
                            Exit While
                        End If

                        ' 失败：尝试自动修复
                        fixAttempt += 1
                        If fixAttempt < MaxFixAttempts Then
                            ' Non-recoverable tool failures skip AI repair and go straight to observe/reflect.
                            If Not toolResult.Recoverable Then
                                AppLogger.Warn("LoopEngine", $"Skip repair for non-recoverable tool failure: {toolResult.ToObserveSummary()}")
                                Exit While
                            End If

                            OnStatusChanged?.Invoke($"代码执行失败，AI 正在修复（尝试 {fixAttempt}/{MaxFixAttempts}）...")
                            AppLogger.Info("LoopEngine", $"Repair attempt {fixAttempt}/{MaxFixAttempts}: {toolResult.ToObserveSummary()}")

                            ' 构建修复提示词（含结构化错误契约）
                            Dim fixPrompt = $"上一次执行失败：

错误码: {If(toolResult.ErrorCode, ExceptionClassifier.CodeUnknown)}
用户可见说明: {If(toolResult.UserMessage, toolResult.Message)}
调试细节: {If(toolResult.DebugDetail, toolResult.Message)}
可自动修复: {toolResult.Recoverable}

原工具调用: {toolCall.ToolId}
原参数: {Newtonsoft.Json.JsonConvert.SerializeObject(toolCall.Parameters)}

当前 {If(session.AppType, "Office")} 可用工具（必须使用原样工具 ID，不要自创 snake_case 或未注册命令）:
{BuildAvailableToolHint(session.AppType, executionContext)}

请分析错误原因并返回修正后的工具调用。只返回 JSON，格式：
```json
{{
  ""toolId"": ""..."",
  ""parameters"": {{...}}
}}
```"

                            Try
                                ' 请求 AI 修复
                                Dim fixedResponse = Await SendRepairRequestAsync(fixPrompt, systemPrompt, toolResult)
                                Dim fixedJson = ExtractJson(fixedResponse)

                                If Not String.IsNullOrEmpty(fixedJson) Then
                                    Dim fixedObj = JObject.Parse(fixedJson)
                                    Dim fixedToolCall = ParseFixedToolCall(fixedObj, toolCall)

                                    If Not String.IsNullOrWhiteSpace(contractedToolId) AndAlso
                                       Not String.Equals(fixedToolCall.ToolId, contractedToolId, StringComparison.OrdinalIgnoreCase) Then
                                        Dim contractMessage = $"自动修复不能把任务合同要求的工具 {contractedToolId} 替换为 {fixedToolCall.ToolId}；已停止执行。"
                                        toolResult = ToolResult.Failed(
                                            contractedToolId,
                                            contractMessage,
                                            errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                            userMessage:=contractMessage,
                                            debugDetail:="Mandatory tool substitution rejected by execution contract.",
                                            recoverable:=False,
                                            observation:=New JObject From {
                                                {"kind", "execution_contract"},
                                                {"summary", contractMessage},
                                                {"changed", False},
                                                {"warnings", New JArray("mandatory_tool_substitution_rejected")}
                                            })
                                        AppLogger.Warn("LoopEngine", contractMessage)
                                        Exit While
                                    End If

                                    ' 使用修复后的工具调用
                                    toolCall = fixedToolCall
                                    AppLogger.Info("LoopEngine", "AI generated repair plan")
                                Else
                                    AppLogger.Warn("LoopEngine", "Unable to parse repair response; stop repair")
                                    Exit While
                                End If
                            Catch ex As Exception
                                AppLogger.Error("LoopEngine", "Repair loop exception", ex)
                                Exit While
                            End Try
                        End If
                    End While

                    ' --- OBSERVE ---
                    session.Status = AgentStatus.Observing
                    Dim observation = FormatObservation(toolResult)

                    ' 如果失败且已达最大修复次数，追加提示
                    If Not toolResult.Success AndAlso fixAttempt >= MaxFixAttempts Then
                        observation &= $" (AI 已尝试自动修复 {fixAttempt} 次，仍然失败)"
                    ElseIf fixAttempt > 0 AndAlso toolResult.Success Then
                        observation &= $" (AI 第 {fixAttempt} 次修复成功)"
                    End If

                    _memory.SetWorking("lastObservation", observation)
                    Dim stepFinishedAt = DateTime.Now
                    Dim explanation = BuildExecutionExplanation(stepIndex, planStep, toolCall, toolResult, fixAttempt, undoPoint, observation, stepStartedAt, stepFinishedAt)
                    planStep.LastExplanation = explanation

                    ' 记录迭代
                    Dim iteration = New ReActIteration With {
                        .Index = session.CurrentIteration,
                        .Thought = thought,
                        .Action = toolCall,
                        .Observation = observation,
                        .Explanation = explanation
                    }
                    session.Iterations.Add(iteration)
                    session.CurrentIteration += 1
                    OnExecutionExplained?.Invoke(explanation)
                    OnIterationUpdate?.Invoke(iteration)

                    ' 更新步骤状态
                    If toolResult.Success Then
                        noProgressCount = 0
                        If ActionCompletesPlanStep(planStep, toolCall) Then
                            planStep.Status = StepStatus.Completed
                            planStep.ErrorMessage = ""
                            OnStepCompleted?.Invoke(stepIndex, True, toolResult.Message)
                        Else
                            planStep.Status = StepStatus.Running
                            OnStatusChanged?.Invoke($"支持动作 {toolCall.ToolId} 已成功；继续完成当前里程碑 {planStep.ToolHint}。")
                        End If
                    Else
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = toolResult.ToObserveSummary()
                        noProgressCount += 1
                        OnStepCompleted?.Invoke(stepIndex, False, If(toolResult.UserMessage, toolResult.Message))
                        AppLogger.Warn("LoopEngine", $"Step failed: {toolResult.ToObserveSummary()}")

                        If Not toolResult.Recoverable Then
                            Dim terminalFailure = $"任务因不可恢复错误停止: {toolResult.ToObserveSummary()}"
                            session.Status = AgentStatus.Failed
                            OnStatusChanged?.Invoke(terminalFailure)
                            AppLogger.Warn("LoopEngine", terminalFailure)
                            Return AgentResult.Failed(session.Id, terminalFailure)
                        End If

                        ' 失败且多轮修复失败，提示可以撤销
                        If fixAttempt >= MaxFixAttempts AndAlso undoPoint IsNot Nothing Then
                            Dim undoHint = _undoManager.GetUndoHint(If(session.AppType, "Unknown"))
                            AppLogger.Info("LoopEngine", $"Execution failed; {undoHint}")
                        End If

                        ' --- REFLECT (连续失败) ---
                        If noProgressCount >= MaxNoProgress Then
                            If replanAttempts >= MaxReplanAttempts Then
                                Dim failMsg = $"步骤多次失败，已达最大重规划次数: {toolResult.ToObserveSummary()}"
                                AppLogger.Error("LoopEngine", failMsg)
                                Return AgentResult.Failed(session.Id, failMsg)
                            End If

                            session.Status = AgentStatus.Reflecting
                            OnStatusChanged?.Invoke("正在分析失败原因并重新规划...")
                            replanAttempts += 1
                            AppLogger.Info("LoopEngine", $"Reflect/replan attempt {replanAttempts}: {toolResult.ToObserveSummary()}")

                            Dim newPlan = Await ReflectAndReplanAsync(session, toolResult.ToObserveSummary(), systemPrompt)
                            If newPlan IsNot Nothing AndAlso newPlan.Steps.Count > 0 Then
                                session.Plan = newPlan
                                stepIndex = 0
                                noProgressCount = 0
                                Continue While
                            Else
                                Dim replanFail = $"重新规划失败: {toolResult.ToObserveSummary()}"
                                AppLogger.Error("LoopEngine", replanFail)
                                Return AgentResult.Failed(session.Id, replanFail)
                            End If
                        End If
                    End If

                    If toolResult.Success AndAlso planStep.Status = StepStatus.Completed Then
                        stepIndex += 1
                    ElseIf Not toolResult.Success Then
                        ' A recoverable failure is a new fact, not a reason to blindly advance
                        ' through the old plan. Keep the current high-level step and let the next
                        ' ReAct iteration choose a different action from the observation.
                        Continue While
                    End If
                End While

                If Not modelDeclaredComplete Then
                    session.Status = AgentStatus.Failed
                    Dim acceptanceError = ValidateExecutionOutcome(session)
                    Dim failMsg = If(String.IsNullOrWhiteSpace(acceptanceError),
                                     $"任务未形成可验收的完成决策，已达自适应决策上限 {MaxIterations}。",
                                     $"任务未通过验收：{acceptanceError}")
                    OnStatusChanged?.Invoke(failMsg)
                    AppLogger.Warn("LoopEngine", failMsg)
                    Return AgentResult.Failed(session.Id, failMsg)
                End If

                Dim incompleteSteps = session.Plan.Steps.
                    Where(Function(s) s.Status <> StepStatus.Completed).
                    ToList()
                If incompleteSteps.Count > 0 Then
                    session.Status = AgentStatus.Failed
                    Dim failMsg = $"任务未完成，失败/未执行步骤 {incompleteSteps.Count} 个: {String.Join("; ", incompleteSteps.Select(Function(s) s.ErrorMessage).Where(Function(m) Not String.IsNullOrWhiteSpace(m)).Take(3))}"
                    OnStatusChanged?.Invoke(failMsg)
                    AppLogger.Warn("LoopEngine", failMsg)
                    Return AgentResult.Failed(session.Id, failMsg)
                End If

                Dim outcomeError = ValidateExecutionOutcome(session)
                If Not String.IsNullOrWhiteSpace(outcomeError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(outcomeError)
                    AppLogger.Warn("LoopEngine", outcomeError)
                    Return AgentResult.Failed(session.Id, outcomeError)
                End If

                Dim finalOutput As String = ""
                If IsReadOnlyAnswerSpec(session.Spec) Then
                    OnStatusChanged?.Invoke("正在根据已读取的完整数据生成答案...")
                    finalOutput = Await GenerateReadOnlyAnswerAsync(session, readOnlyEvidence)
                    If String.IsNullOrWhiteSpace(finalOutput) Then
                        Dim answerFailure = "已读取工作簿数据，但未能生成可验证的最终答案；未对数据作任何修改"
                        session.Status = AgentStatus.Failed
                        OnStatusChanged?.Invoke(answerFailure)
                        Return AgentResult.Failed(session.Id, answerFailure)
                    End If
                End If

                ' 完成
                session.Status = AgentStatus.Completed
                Dim finalMsg = $"任务完成，共执行 {session.CurrentIteration} 个迭代"
                OnStatusChanged?.Invoke(finalMsg)
                AppLogger.Info("LoopEngine", finalMsg)
                Dim userMessage = If(String.IsNullOrWhiteSpace(finalOutput), finalMsg, finalOutput)
                Return AgentResult.SuccessResult(session.Id, userMessage, finalOutput)

            Catch ex As Exception
                session.Status = AgentStatus.Failed
                Dim classified = ExceptionClassifier.Classify(ex)
                OnStatusChanged?.Invoke($"执行出错: {classified.UserMessage}")
                AppLogger.Error("LoopEngine", "RunAsync unhandled exception", ex)
                Return AgentResult.Failed(session.Id, $"执行异常: [{classified.ErrorCode}] {classified.UserMessage}")
            End Try
        End Function

        Private Function ActionCompletesPlanStep(planStep As PlanStep,
                                                 toolCall As ToolCall) As Boolean
            If planStep Is Nothing OrElse toolCall Is Nothing Then Return False

            Dim expectedTool = If(planStep.ToolHint, "").Trim()
            If String.IsNullOrWhiteSpace(expectedTool) AndAlso Not String.IsNullOrWhiteSpace(planStep.Code) Then
                expectedTool = If(ParsePlannedToolCall(planStep.Code)?.ToolId, "").Trim()
            End If
            If String.IsNullOrWhiteSpace(expectedTool) Then Return True

            Return String.Equals(expectedTool, toolCall.ToolId, StringComparison.OrdinalIgnoreCase)
        End Function


        ''' <summary>
        ''' 获取撤销管理器（供外部访问）
        ''' </summary>
        Public ReadOnly Property UndoManager As Core.UndoManager
            Get
                Return _undoManager
            End Get
        End Property

    End Class

End Namespace
