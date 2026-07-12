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
    Public Class LoopEngine
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _undoManager As Core.UndoManager

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

        Public Sub New(toolRegistry As ToolRegistry, memory As AgentMemory, promptManager As PromptManager)
            _toolRegistry = toolRegistry
            _memory = memory
            _promptManager = promptManager
            _undoManager = New Core.UndoManager(10) ' 最多保存 10 个撤销点
        End Sub

        ''' <summary>
        ''' 执行 ReAct 循环
        ''' </summary>
        Public Async Function RunAsync(session As AgentSession,
                                        systemPrompt As String,
                                        Optional skill As AgentSkill = Nothing) As Task(Of AgentResult)
            Dim noProgressCount As Integer = 0
            Dim replanAttempts As Integer = 0

            Try
                ' Phase 1: 生成 Spec
                OnStatusChanged?.Invoke("正在分析任务...")
                session.Spec = Await GenerateSpecAsync(session)

                ' Phase 2: 生成计划
                OnStatusChanged?.Invoke("正在制定执行计划...")
                session.Plan = Await GeneratePlanAsync(session, systemPrompt, skill)
                If session.Plan Is Nothing OrElse session.Plan.Steps.Count = 0 Then
                    Return AgentResult.Failed(session.Id, "规划失败：无法生成执行计划")
                End If

                ' 通知计划已生成
                OnPlanGenerated?.Invoke(session.Plan)

                session.Status = AgentStatus.Executing
                OnStatusChanged?.Invoke($"规划完成（共 {session.Plan.Steps.Count} 步），进入自治执行...")

                ' Phase 3: ReAct Loop
                Dim stepIndex As Integer = 0
                While stepIndex < session.Plan.Steps.Count AndAlso session.CurrentIteration < MaxIterations
                    Dim planStep = session.Plan.Steps(stepIndex)
                    planStep.Status = StepStatus.Running

                    ' --- THINK ---
                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"步骤 {stepIndex + 1}/{session.Plan.Steps.Count}: {planStep.Description}")
                    Dim thought = Await ThinkAsync(session, planStep, systemPrompt)

                    ' --- PARSE ACTION ---
                    Dim toolCall = ParseToolCall(thought)
                    If toolCall Is Nothing Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = "无法解析工具调用"
                        OnStepCompleted?.Invoke(stepIndex, False, "解析失败")

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If

                    Dim normalizeMessage As String = ""
                    If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, normalizeMessage) Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = normalizeMessage
                        OnStepCompleted?.Invoke(stepIndex, False, normalizeMessage)

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If
                    If Not String.IsNullOrWhiteSpace(normalizeMessage) Then
                        Debug.WriteLine($"[LoopEngine] {normalizeMessage}")
                    End If

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
                    Dim originalToolCall = toolCall
                    Dim stepStartedAt = DateTime.Now

                    ' 多轮修复循环
                    While fixAttempt < MaxFixAttempts
                        Dim normalized As String = ""
                        If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, normalized) Then
                            toolResult = ToolResult.Failed(toolCall.ToolId,
                                                           normalized,
                                                           New With {.availableTools = BuildAvailableToolHint(session.AppType)},
                                                           ExceptionClassifier.CodeNotFound,
                                                           normalized,
                                                           normalized,
                                                           recoverable:=True)
                        Else
                            If Not String.IsNullOrWhiteSpace(normalized) Then Debug.WriteLine($"[LoopEngine] {normalized}")
                            ' 执行工具
                            toolResult = Await _toolRegistry.ExecuteToolAsync(toolCall.ToolId, toolCall.Parameters)
                        End If

                        If toolResult.Success Then
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
{BuildAvailableToolHint(session.AppType)}

请分析错误原因并返回修正后的工具调用。只返回 JSON，格式：
```json
{{
  ""toolId"": ""..."",
  ""parameters"": {{...}}
}}
```"

                            Try
                                ' 请求 AI 修复
                                Dim fixedResponse = Await SendAIRequest(fixPrompt, systemPrompt, Nothing)
                                Dim fixedJson = ExtractJson(fixedResponse)

                                If Not String.IsNullOrEmpty(fixedJson) Then
                                    Dim fixedObj = JObject.Parse(fixedJson)
                                    Dim fixedToolCall = ParseFixedToolCall(fixedObj, toolCall)

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
                        planStep.Status = StepStatus.Completed
                        noProgressCount = 0
                        OnStepCompleted?.Invoke(stepIndex, True, toolResult.Message)
                    Else
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = toolResult.ToObserveSummary()
                        noProgressCount += 1
                        OnStepCompleted?.Invoke(stepIndex, False, If(toolResult.UserMessage, toolResult.Message))
                        AppLogger.Warn("LoopEngine", $"Step failed: {toolResult.ToObserveSummary()}")

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

                    stepIndex += 1
                End While

                If session.CurrentIteration = 0 Then
                    session.Status = AgentStatus.Failed
                    Dim failMsg = "任务未执行任何工具调用。可能是计划步骤没有生成可解析的 action，或当前宿主工具未加载。"
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

                ' 完成
                session.Status = AgentStatus.Completed
                Dim finalMsg = $"任务完成，共执行 {session.CurrentIteration} 个迭代"
                OnStatusChanged?.Invoke(finalMsg)
                AppLogger.Info("LoopEngine", finalMsg)
                Return AgentResult.SuccessResult(session.Id, finalMsg)

            Catch ex As Exception
                session.Status = AgentStatus.Failed
                Dim classified = ExceptionClassifier.Classify(ex)
                OnStatusChanged?.Invoke($"执行出错: {classified.UserMessage}")
                AppLogger.Error("LoopEngine", "RunAsync unhandled exception", ex)
                Return AgentResult.Failed(session.Id, $"执行异常: [{classified.ErrorCode}] {classified.UserMessage}")
            End Try
        End Function

#Region "Private Methods"

        Private Function BuildExecutionExplanation(stepIndex As Integer,
                                                   planStep As PlanStep,
                                                   toolCall As ToolCall,
                                                   toolResult As ToolResult,
                                                   fixAttempts As Integer,
                                                   undoPoint As Core.UndoManager.UndoPoint,
                                                   observation As String,
                                                   startedAt As DateTime,
                                                   finishedAt As DateTime) As ExecutionExplanation
            Dim tool = If(toolCall Is Nothing, Nothing, _toolRegistry.GetTool(toolCall.ToolId))
            Dim toolId = If(toolCall?.ToolId, "")
            Dim toolName = If(tool?.Name, toolId)
            Dim category = If(tool?.Category, "")
            Dim risk = If(tool?.RiskLevel, "unknown")
            Dim paramsJson = If(toolCall?.Parameters Is Nothing, "{}", toolCall.Parameters.ToString(Formatting.None))
            Dim success = toolResult IsNot Nothing AndAlso toolResult.Success
            Dim message = If(toolResult?.Message, "")
            Dim skillName = ExtractResultDataValue(toolResult, "skillName")
            Dim scriptFileName = ExtractResultDataValue(toolResult, "scriptFileName")
            Dim mcpToolName = ExtractResultDataValue(toolResult, "mcpToolName")
            Dim mcpStatus = ExtractResultDataValue(toolResult, "mcpStatus")
            Dim failureReason = If(success, "", ExtractResultDataValue(toolResult, "failureReason"))
            If String.IsNullOrWhiteSpace(failureReason) AndAlso Not success Then failureReason = message
            Dim verb = If(success, "已完成", "未完成")
            Dim fixText = If(fixAttempts > 0, $"，期间自动修复 {fixAttempts} 次", "")
            Dim skillText = If(String.IsNullOrWhiteSpace(skillName), "", $"，Skill: {skillName}")
            Dim scriptText = If(String.IsNullOrWhiteSpace(scriptFileName), "", $"，脚本: {scriptFileName}")
            Dim mcpText = If(String.IsNullOrWhiteSpace(mcpToolName), "", $"，MCP: {mcpToolName}")
            Dim elapsedMs = CLng(Math.Max(0, (finishedAt - startedAt).TotalMilliseconds))
            If toolResult IsNot Nothing AndAlso toolResult.ElapsedMs > 0 Then elapsedMs = toolResult.ElapsedMs
            Dim undoHint = If(undoPoint Is Nothing, "", _undoManager.GetUndoHint(If(undoPoint.AppType, "")))
            Dim undoPointName = If(undoPoint?.Name, "")
            Dim canUndo = undoPoint IsNot Nothing AndAlso undoPoint.CanUndo
            Dim beforeSummary = $"准备执行步骤 {stepIndex + 1}: {If(planStep?.Description, "")}；工具: {toolId}；参数: {paramsJson}"
            Dim afterSummary = If(observation, message)
            Dim repairSummary = If(fixAttempts > 0,
                                   If(success, $"AI 自动修复 {fixAttempts} 次后成功", $"AI 自动修复 {fixAttempts} 次后仍失败"),
                                   "")
            Dim text = $"步骤 {stepIndex + 1} {verb}：调用 {toolName}（{toolId}）{skillText}{scriptText}{mcpText}{fixText}。{message}"

            Return New ExecutionExplanation With {
                .StepIndex = stepIndex,
                .StepDescription = If(planStep?.Description, ""),
                .ToolId = toolId,
                .ToolName = toolName,
                .ToolCategory = category,
                .RiskLevel = risk,
                .ParametersJson = paramsJson,
                .StartedAt = startedAt,
                .FinishedAt = finishedAt,
                .ElapsedMs = elapsedMs,
                .BeforeSummary = beforeSummary,
                .AfterSummary = afterSummary,
                .Success = success,
                .Message = message,
                .SkillName = skillName,
                .ScriptFileName = scriptFileName,
                .McpToolName = mcpToolName,
                .McpStatus = mcpStatus,
                .FailureReason = failureReason,
                .UndoPointName = undoPointName,
                .UndoHint = undoHint,
                .CanUndo = canUndo,
                .AutoRepairSummary = repairSummary,
                .FixAttempts = fixAttempts,
                .ExplanationText = text
            }
        End Function

        Private Function ExtractResultDataValue(toolResult As ToolResult, key As String) As String
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""

            Try
                Dim obj = JObject.FromObject(toolResult.Data)
                Dim token = obj.SelectToken(key)
                If token Is Nothing Then Return ""
                Return token.ToString()
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", $"ExtractResultDataValue key={key} failed", ex)
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 生成任务 Spec
        ''' </summary>
        Private Async Function GenerateSpecAsync(session As AgentSession) As Task(Of AgentTaskSpec)
            Dim spec As New AgentTaskSpec()
            Try
                Dim prompt = $"分析以下需求，提取结构化任务规格：

需求: {session.UserRequest}

返回 JSON：
```json
{{
  ""goal"": ""一句话描述核心目标"",
  ""constraints"": [""约束1""],
  ""success_criteria"": [""成功标准1""],
  ""complexity"": ""simple|medium|complex""
}}
```

complexity 规则：
- simple：单一操作，步骤数 <= 2，无需用户确认
- medium：2-5个步骤，建议用户确认
- complex：步骤多或逻辑复杂，必须用户确认"

                Dim response = Await SendAIRequest(prompt,
                    "你是一个任务分析专家。只返回JSON，不要解释。", Nothing)

                Dim jsonStr = ExtractJson(response)
                If Not String.IsNullOrEmpty(jsonStr) Then
                    Dim obj = JObject.Parse(jsonStr)
                    spec.Goal = If(obj("goal")?.ToString(), session.UserRequest)
                    spec.Complexity = If(obj("complexity")?.ToString(), "medium")

                    Dim constraints = TryCast(obj("constraints"), JArray)
                    If constraints IsNot Nothing Then
                        For Each c In constraints
                            spec.Constraints.Add(c.ToString())
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] Spec生成失败: {ex.Message}")
            End Try
            Return spec
        End Function

        ''' <summary>
        ''' 生成执行计划
        ''' </summary>
        Private Async Function GeneratePlanAsync(session As AgentSession,
                                                  systemPrompt As String,
                                                  skill As AgentSkill) As Task(Of ExecutionPlan)
            Dim plan As New ExecutionPlan()
            Try
                Dim prompt = _promptManager.BuildPlanningPrompt(session, systemPrompt, skill)
                Dim response = Await SendAIRequest(prompt, systemPrompt, _memory.GetRecentMessages(5))

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim obj = JObject.Parse(jsonStr)
                plan.Understanding = obj("understanding")?.ToString()
                plan.Summary = obj("summary")?.ToString()

                Dim stepsArray = TryCast(obj("steps"), JArray)
                If stepsArray IsNot Nothing Then
                    Dim stepNum = 1
                    For Each item In stepsArray
                        plan.Steps.Add(New PlanStep With {
                            .StepNumber = stepNum,
                            .Description = item("description")?.ToString(),
                            .Code = item("code")?.ToString(),
                            .Language = If(item("language")?.ToString(), "json")
                        })
                        stepNum += 1
                    Next
                End If
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 规划失败: {ex.Message}")
                Return Nothing
            End Try
            Return plan
        End Function

        ''' <summary>
        ''' Think：调用LLM生成思考+行动
        ''' </summary>
        Private Async Function ThinkAsync(session As AgentSession,
                                           planStep As PlanStep,
                                           systemPrompt As String) As Task(Of String)
            Dim lastObservation = _memory.GetWorkingString("lastObservation")
            Dim prompt = _promptManager.BuildReactPrompt(planStep, _memory, lastObservation)
            Dim history = _memory.GetRecentMessages(10)
            Return Await SendAIRequest(prompt, systemPrompt, history)
        End Function

        ''' <summary>
        ''' 反思并重新规划
        ''' </summary>
        Private Async Function ReflectAndReplanAsync(session As AgentSession,
                                                      observation As String,
                                                      systemPrompt As String) As Task(Of ExecutionPlan)
            Try
                Dim prompt = _promptManager.BuildReflectionPrompt(session, observation)
                Dim response = Await SendAIRequest(prompt, systemPrompt, Nothing)

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim decision = JObject.Parse(jsonStr)
                Dim strategy = decision("strategy")?.ToString()?.ToLower()

                Select Case strategy
                    Case "retry"
                        ' 重试当前计划
                        Return session.Plan
                    Case "skip"
                        ' 跳过当前步骤继续
                        Return session.Plan
                    Case "replan"
                        ' 重新生成计划
                        Return Await GeneratePlanAsync(session, systemPrompt, session.Skill)
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 反思失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 解析工具调用
        ''' </summary>
        Private Function ParseToolCall(response As String) As ToolCall
            Try
                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim obj = JObject.Parse(jsonStr)
                Dim action = obj("action")
                If action Is Nothing Then Return Nothing

                Dim toolId = action("tool")?.ToString()
                Dim params = TryCast(action("params"), JObject)
                If String.IsNullOrEmpty(toolId) Then Return Nothing
                If params Is Nothing Then params = New JObject()

                Return New ToolCall With {
                    .ToolId = toolId,
                    .Parameters = params
                }
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 解析工具调用失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Private Function ParseFixedToolCall(fixedObj As JObject, fallback As ToolCall) As ToolCall
            If fixedObj Is Nothing Then Return fallback

            Dim toolId = fixedObj("toolId")?.ToString()
            Dim params = TryCast(fixedObj("parameters"), JObject)

            If String.IsNullOrWhiteSpace(toolId) Then
                Dim action = TryCast(fixedObj("action"), JObject)
                If action IsNot Nothing Then
                    toolId = action("tool")?.ToString()
                    If params Is Nothing Then params = TryCast(action("params"), JObject)
                End If
            End If

            If String.IsNullOrWhiteSpace(toolId) Then toolId = fallback.ToolId
            If params Is Nothing Then params = If(fallback.Parameters, New JObject())

            Return New ToolCall With {
                .ToolId = toolId,
                .Parameters = params
            }
        End Function

        Private Function BuildAvailableToolHint(appType As String) As String
            Dim tools = _toolRegistry.GetAvailableTools(appType).
                OrderBy(Function(t) t.Category).
                ThenBy(Function(t) t.Id).
                Take(40).
                Select(Function(t) $"- {t.Id}: {t.Name}")
            Return String.Join(vbCrLf, tools)
        End Function

        ''' <summary>
        ''' 格式化观察结果
        ''' </summary>
        Private Function FormatObservation(result As ToolResult) As String
            If result Is Nothing Then
                Return "❌ [unknown] 无工具结果"
            End If
            If result.Success Then
                Return $"✅ [{result.ToolId}] 执行成功: {result.Message}"
            End If
            ' Structured observe payload for repair/reflect (P0-4).
            Return $"❌ [{result.ToolId}] {result.ToObserveSummary()}"
        End Function

        ''' <summary>
        ''' 从响应中提取JSON
        ''' </summary>
        Private Function ExtractJson(response As String) As String
            If String.IsNullOrWhiteSpace(response) Then Return Nothing

            ' 查找 ```json 代码块
            Dim start = response.IndexOf("```json")
            If start >= 0 Then
                start = response.IndexOf("{"c, start)
                If start >= 0 Then
                    Dim endIdx = response.LastIndexOf("}"c)
                    If endIdx > start Then
                        Return response.Substring(start, endIdx - start + 1)
                    End If
                End If
            End If

            ' 查找纯 JSON
            start = response.IndexOf("{"c)
            If start >= 0 Then
                Dim endIdx = response.LastIndexOf("}"c)
                If endIdx > start Then
                    Return response.Substring(start, endIdx - start + 1)
                End If
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 等待用户确认
        ''' </summary>
        Private Async Function WaitForApprovalAsync(message As String) As Task(Of Boolean)
            If OnRequestApproval IsNot Nothing Then
                Return Await OnRequestApproval(message)
            End If
            ' 无回调时默认批准
            Return True
        End Function

        ''' <summary>
        ''' 获取撤销管理器（供外部访问）
        ''' </summary>
        Public ReadOnly Property UndoManager As Core.UndoManager
            Get
                Return _undoManager
            End Get
        End Property

#End Region

    End Class

End Namespace
