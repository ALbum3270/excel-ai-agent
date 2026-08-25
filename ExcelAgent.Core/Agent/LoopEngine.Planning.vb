Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Namespace Agent
    Public Partial Class LoopEngine
        #Region "Planning, Observation, And Explanation Helpers"
        Private Const MaxReadOnlyEvidenceChars As Integer = 250000

        Private Shared Function IsReadOnlyAnswerSpec(spec As AgentTaskSpec) As Boolean
            Return spec IsNot Nothing AndAlso
                String.Equals(If(spec.MutationPolicy, ""), "read_only", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsConversationOnlySpec(spec As AgentTaskSpec) As Boolean
            Return IsReadOnlyAnswerSpec(spec) AndAlso Not spec.RequiresHostExecution
        End Function

        Private Shared Function AppendReadOnlyEvidence(evidence As JArray,
                                                       toolCall As ToolCall,
                                                       toolResult As ToolResult) As String
            If evidence Is Nothing OrElse toolResult Is Nothing OrElse Not toolResult.Success Then Return ""
            If toolResult.Data Is Nothing Then
                Return "只读工具没有返回可供验证的数据；已停止作答，未对工作簿作任何修改"
            End If

            Try
                Dim dataToken = TryCast(toolResult.Data, JToken)
                If dataToken Is Nothing Then dataToken = JToken.FromObject(toolResult.Data)
                Dim item As New JObject From {
                    {"toolId", If(toolCall?.ToolId, "")},
                    {"data", dataToken.DeepClone()}
                }
                Dim projectedSize = evidence.ToString(Formatting.None).Length + item.ToString(Formatting.None).Length
                If projectedSize > MaxReadOnlyEvidenceChars Then
                    Return $"精确读取结果超过当前回答证据上限（{MaxReadOnlyEvidenceChars} 字符）；已停止作答，未截断数据或进行估算"
                End If
                evidence.Add(item)
                Return ""
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Read-only evidence capture failed: {AppLogger.Redact(ex.Message)}")
                Return "无法保存只读工具返回的结构化证据；已停止作答，未对工作簿作任何修改"
            End Try
        End Function

        Private Function BuildExecutionExplanation(stepIndex As Integer,
                                                   planStep As PlanStep,
                                                   toolCall As ToolCall,
                                                   toolResult As ToolResult,
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
            Dim failureReason = If(success, "", ExtractResultDataValue(toolResult, "failureReason"))
            If String.IsNullOrWhiteSpace(failureReason) AndAlso Not success Then failureReason = message
            Dim verb = If(success, "已完成", "未完成")
            Dim elapsedMs = CLng(Math.Max(0, (finishedAt - startedAt).TotalMilliseconds))
            If toolResult IsNot Nothing AndAlso toolResult.ElapsedMs > 0 Then elapsedMs = toolResult.ElapsedMs
            Dim undoHint = If(undoPoint Is Nothing, "", _undoManager.GetUndoHint(If(undoPoint.AppType, "")))
            Dim undoPointName = If(undoPoint?.Name, "")
            Dim canUndo = undoPoint IsNot Nothing AndAlso undoPoint.CanUndo
            Dim beforeSummary = $"准备执行步骤 {stepIndex + 1}: {If(planStep?.Description, "")}；工具: {toolId}；参数: {paramsJson}"
            Dim afterSummary = If(observation, message)
            Dim observationJson = SerializeCompactJson(If(toolResult?.Observation, Nothing), 8192)
            Dim dataSummaryJson = SerializeCompactJson(BuildDataSummary(toolResult), 4096)
            Dim text = $"步骤 {stepIndex + 1} {verb}：调用 {toolName}（{toolId}）。{message}"

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
                .ObservationJson = observationJson,
                .DataSummaryJson = dataSummaryJson,
                .Success = success,
                .Message = message,
                .FailureReason = failureReason,
                .UndoPointName = undoPointName,
                .UndoHint = undoHint,
                .CanUndo = canUndo,
                .AutoRepairSummary = "",
                .FixAttempts = 0,
                .ExplanationText = text
            }
        End Function

        Private Function BuildDataSummary(toolResult As ToolResult) As Object
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing Then Return Nothing

            Try
                Dim token = JToken.FromObject(toolResult.Data)
                If token.Type = JTokenType.Array Then
                    Dim arr = CType(token, JArray)
                    Return New JObject From {
                        {"type", "array"},
                        {"count", arr.Count},
                        {"preview", CloneFirstItems(arr, 5)}
                    }
                End If

                If token.Type = JTokenType.Object Then
                    Dim obj = CType(token, JObject)
                    Dim summary As New JObject From {
                        {"type", "object"},
                        {"keys", BuildStringArray(obj.Properties().Select(Function(p) p.Name).Take(20))}
                    }

                    If obj("total") IsNot Nothing Then summary("total") = obj("total")
                    If obj("returned") IsNot Nothing Then summary("returned") = obj("returned")
                    If obj("truncated") IsNot Nothing Then summary("truncated") = obj("truncated")
                    If obj("items") IsNot Nothing AndAlso obj("items").Type = JTokenType.Array Then
                        summary("itemsPreview") = CloneFirstItems(CType(obj("items"), JArray), 5)
                    End If

                    Return summary
                End If

                Return token
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", "BuildDataSummary failed", ex)
                Return Nothing
            End Try
        End Function

        Private Function CloneFirstItems(items As JArray, maxItems As Integer) As JArray
            Dim result As New JArray()
            If items Is Nothing Then Return result

            For Each item In items.Take(Math.Max(0, maxItems))
                result.Add(item.DeepClone())
            Next

            Return result
        End Function

        Private Function BuildStringArray(items As IEnumerable(Of String)) As JArray
            Dim result As New JArray()
            If items Is Nothing Then Return result

            For Each item In items
                result.Add(If(item, ""))
            Next

            Return result
        End Function

        Private Function SerializeCompactJson(value As Object, maxLength As Integer) As String
            If value Is Nothing Then Return ""

            Try
                Dim json = JsonConvert.SerializeObject(value, Formatting.None)
                If maxLength > 0 AndAlso json.Length > maxLength Then
                    Return json.Substring(0, maxLength) & "...(truncated)"
                End If
                Return json
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", "SerializeCompactJson failed", ex)
                Return ""
            End Try
        End Function

        Private Function ExtractResultDataValue(toolResult As ToolResult, key As String) As String
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""

            Try
                Dim tokenValue = TryCast(toolResult.Data, JToken)
                If tokenValue Is Nothing Then tokenValue = JToken.FromObject(toolResult.Data)
                Dim obj = TryCast(tokenValue, JObject)
                If obj Is Nothing Then Return ""
                Dim token = obj.SelectToken(key)
                If token Is Nothing Then Return ""
                Return token.ToString()
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", $"ExtractResultDataValue key={key} failed", ex)
                Return ""
            End Try
        End Function

        Private Async Function GenerateSpecAsync(session As AgentSession) As Task(Of AgentTaskSpec)
            Dim spec As New AgentTaskSpec()
            Try
                Dim prompt = $"分析以下需求，提取结构化任务规格：

需求: {session.UserRequest}

返回 JSON：
```json
{{
  ""goal"": ""一句话描述核心目标"",
  ""target_object"": ""需要读取或改变的 Office 对象；未知时为空"",
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
                    spec.TargetObject = If(obj("target_object")?.ToString(), "")
                    spec.Complexity = If(obj("complexity")?.ToString(), "medium")
                    Dim constraints = TryCast(obj("constraints"), JArray)
                    If constraints IsNot Nothing Then
                        For Each c In constraints
                            spec.Constraints.Add(c.ToString())
                        Next
                    End If
                    Dim successCriteria = TryCast(obj("success_criteria"), JArray)
                    If successCriteria IsNot Nothing Then
                        For Each criterion In successCriteria
                            Dim criterionText = If(criterion?.ToString(), "").Trim()
                            If Not String.IsNullOrWhiteSpace(criterionText) Then spec.SuccessCriteria.Add(criterionText)
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] Spec生成失败: {ex.Message}")
            End Try
            Return spec
        End Function

        Private Async Function GeneratePlanAsync(session As AgentSession,
                                                  systemPrompt As String,
                                                  skill As AgentSkill,
                                                  Optional planningValidationFeedback As String = "") As Task(Of ExecutionPlan)
            Dim plan As New ExecutionPlan()
            Dim failureReason As String = ""
            Try
                Dim prompt = _promptManager.BuildPlanningPrompt(
                    session,
                    systemPrompt,
                    skill,
                    _memory,
                    planningValidationFeedback)
                Dim response = Await SendAIRequest(prompt, systemPrompt, _memory.GetRecentMessages(5))

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then
                    failureReason = "响应中没有 JSON"
                Else
                    Dim obj = JObject.Parse(jsonStr)
                    plan.Understanding = obj("understanding")?.ToString()
                    plan.Summary = obj("summary")?.ToString()
                    plan.CapabilityGap = obj("capabilityGap")?.ToString()
                    plan.OutcomeContract = ParseOutcomeContract(TryCast(obj("outcomeContract"), JObject), session.AppType)

                    ' Deliberately ignore legacy `steps`.  They are predictions about future
                    ' world state and must never steer adaptive execution.  The initial turn
                    ' only supplies the frozen acceptance projection and user-facing summary.
                End If
            Catch ex As Exception
                failureReason = ex.Message
            End Try

            If Not String.IsNullOrWhiteSpace(failureReason) Then
                AppLogger.Warn("LoopEngine", $"初始验收投影生成不完整: {AppLogger.Redact(failureReason)}")
            End If
            Return plan
        End Function

        Private Shared Function ParseOutcomeContract(value As JObject,
                                                       appType As String) As OutcomeContract
            If value Is Nothing Then Return Nothing
            Dim requirements = TryCast(value("requirements"), JArray)
            If requirements Is Nothing OrElse requirements.Count = 0 Then Return Nothing

            Dim contract As New OutcomeContract With {
                .SchemaVersion = If(value("schemaVersion")?.ToString(), "1.0")
            }
            Dim ids As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For index = 0 To requirements.Count - 1
                Dim item = TryCast(requirements(index), JObject)
                If item Is Nothing Then Continue For
                Dim effectType = If(item("effectType")?.ToString(), "").Trim().ToLowerInvariant()
                If String.IsNullOrWhiteSpace(effectType) Then Continue For
                Dim requirementId = If(item("id")?.ToString(), $"goal-{index + 1}").Trim()
                If String.IsNullOrWhiteSpace(requirementId) OrElse ids.Contains(requirementId) Then Continue For
                ids.Add(requirementId)
                contract.Requirements.Add(New OutcomeRequirement With {
                    .Id = requirementId,
                    .AppType = If(item("appType")?.ToString(), appType),
                    .TargetRef = If(item("targetRef")?.ToString(), "").Trim(),
                    .EffectType = effectType,
                    .PropertyName = If(item("property")?.ToString(), "").Trim(),
                    .Operator = If(item("operator")?.ToString(), "equals").Trim().ToLowerInvariant(),
                    .ExpectedValue = item("expectedValue")?.DeepClone(),
                    .DerivedFromCapability = If(item("derivedFromCapability")?.ToString(), "").Trim(),
                    .CriterionIds = ParseStringList(TryCast(item("criterionIds"), JArray)),
                    .Required = If(item("required")?.Value(Of Boolean)(), True),
                    .Description = If(item("description")?.ToString(), "").Trim()
                })
            Next
            If contract.Requirements.Count = 0 Then Return Nothing
            Return contract
        End Function

        Private Shared Function ParseStringList(values As JArray) As List(Of String)
            Dim result As New List(Of String)()
            If values Is Nothing Then Return result
            For Each value In values
                Dim text = If(value?.ToString(), "").Trim()
                If Not String.IsNullOrWhiteSpace(text) AndAlso
                   Not result.Contains(text, StringComparer.OrdinalIgnoreCase) Then result.Add(text)
            Next
            Return result
        End Function

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

        Private Function BuildAvailableToolHint(appType As String,
                                                Optional executionContext As ToolExecutionContext = Nothing) As String
            Dim tools = _toolRegistry.GetVisibleTools(appType, executionContext).
                OrderBy(Function(t) t.Category).
                ThenBy(Function(t) t.Id).
                Take(40).
                Select(Function(t) $"- {t.Id}: {t.Name}")
            Return String.Join(vbCrLf, tools)
        End Function

        Private Function FormatObservation(result As ToolResult) As String
            If result Is Nothing Then
                Return "❌ [unknown] 无工具结果"
            End If
            If result.Success Then
                Dim summary = result.ToObserveSummary()
                Dim dataSummary = FormatResultData(result.Data)
                If Not String.IsNullOrWhiteSpace(dataSummary) Then
                    Return $"✅ [{result.ToolId}] 执行成功: {summary}{vbCrLf}data={dataSummary}"
                End If
                Return $"✅ [{result.ToolId}] 执行成功: {summary}"
            End If
            ' Structured observe payload for repair/reflect (P0-4).
            Return $"❌ [{result.ToolId}] {result.ToObserveSummary()}"
        End Function

        ''' <summary>
        ''' Adds redacted execution diagnostics to the model-only observation channel. The
        ''' public iteration text remains concise, while the next adaptive decision receives
        ''' the concrete parser/runtime error needed to repair code instead of guessing.
        ''' </summary>
        Private Shared Function FormatModelObservation(result As ToolResult,
                                                       publicObservation As String) As String
            If result Is Nothing OrElse result.Success Then
                Return If(publicObservation, "")
            End If

            Dim diagnostic As New JObject From {
                {"success", False},
                {"toolId", If(result.ToolId, "")},
                {"errorCode", If(result.ErrorCode, ExceptionClassifier.CodeUnknown)},
                {"message", If(result.Message, "")},
                {"userMessage", If(result.UserMessage, "")},
                {"debugDetail", If(result.DebugDetail, "")},
                {"recoverable", result.Recoverable},
                {"retryable", result.Retryable},
                {"taskFatal", result.TaskFatal},
                {"sessionFatal", result.SessionFatal},
                {"data", CloneModelDiagnosticToken(result.Data)},
                {"observation", CloneModelDiagnosticToken(result.Observation)}
            }
            Dim payload = AppLogger.Redact(diagnostic.ToString(Formatting.None))
            Return $"{If(publicObservation, "")}{vbCrLf}Full structured tool failure for the next model decision (redacted):{vbCrLf}{payload}"
        End Function

        Private Shared Function CloneModelDiagnosticToken(value As Object) As JToken
            If value Is Nothing Then Return JValue.CreateNull()
            Try
                Dim token = TryCast(value, JToken)
                If token Is Nothing Then token = JToken.FromObject(value)
                Return token.DeepClone()
            Catch ex As Exception
                Return New JObject From {
                    {"serializationError", AppLogger.Redact(ex.Message)},
                    {"fallback", AppLogger.Redact(Convert.ToString(value, Globalization.CultureInfo.InvariantCulture))}
                }
            End Try
        End Function

        ''' <summary>
        ''' A host call returning Success=True is not sufficient for a write operation.
        ''' Structured write observations must confirm an actual change, and operation
        ''' batches must not contain unhandled failed/partial steps.
        ''' </summary>
        Private Function ValidateObservedOutcome(result As ToolResult) As ToolResult
            If result Is Nothing OrElse Not result.Success OrElse result.Observation Is Nothing Then Return result

            Try
                Dim observation = TryCast(result.Observation, JToken)
                If observation Is Nothing Then observation = JToken.FromObject(result.Observation)
                If observation Is Nothing OrElse observation.Type <> JTokenType.Object Then Return result

                Dim kind = If(observation("kind")?.ToString(), "").Trim().ToLowerInvariant()
                Dim writeExpectedToken = observation("writeExpected")
                Dim writeExpected = writeExpectedToken IsNot Nothing AndAlso
                                    writeExpectedToken.Type = JTokenType.Boolean AndAlso
                                    writeExpectedToken.Value(Of Boolean)()
                If kind <> "write" AndAlso (kind <> "office_operation_batch" OrElse Not writeExpected) Then
                    Return result
                End If

                Dim operations = TryCast(observation("operations"), JArray)
                If operations IsNot Nothing Then
                    Dim failedCount = operations.OfType(Of JObject)().Count(
                        Function(item) Not String.Equals(item("status")?.ToString(), "succeeded", StringComparison.OrdinalIgnoreCase))
                    If failedCount > 0 Then
                        Return ToolResult.Failed(
                            result.ToolId,
                            $"操作批次存在 {failedCount} 个未成功步骤",
                            data:=result.Data,
                            errorCode:=ExceptionClassifier.CodePartialApply,
                            userMessage:="部分 Office 操作未完成，正在尝试修复",
                            recoverable:=True,
                            observation:=result.Observation,
                            artifacts:=result.Artifacts)
                    End If
                End If

                Dim satisfiedToken = observation("satisfied")
                Dim hasSemanticVerification = satisfiedToken IsNot Nothing AndAlso satisfiedToken.Type = JTokenType.Boolean
                If hasSemanticVerification AndAlso Not satisfiedToken.Value(Of Boolean)() Then
                    Return ToolResult.Failed(
                        result.ToolId,
                        "宿主产生了变化，但观察结果不符合工具请求的预期状态",
                        data:=result.Data,
                        errorCode:=ExceptionClassifier.CodeVerifyFailed,
                        userMessage:="Office 已产生变化，但实际结果与请求不一致，正在尝试修复",
                        recoverable:=True,
                        observation:=result.Observation,
                        artifacts:=result.Artifacts)
                End If

                Dim changedToken = observation("changed")
                If changedToken IsNot Nothing AndAlso
                   changedToken.Type = JTokenType.Boolean AndAlso
                   Not changedToken.Value(Of Boolean)() AndAlso
                   Not (hasSemanticVerification AndAlso satisfiedToken.Value(Of Boolean)()) Then
                    Return ToolResult.Failed(
                        result.ToolId,
                        "宿主返回成功，但观察结果未检测到实际变化",
                        data:=result.Data,
                        errorCode:=ExceptionClassifier.CodeVerifyFailed,
                        userMessage:="Office 未产生预期变化，正在重新验证或修复",
                        recoverable:=True,
                        observation:=result.Observation,
                        artifacts:=result.Artifacts)
                End If
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Validate observation failed closed: {AppLogger.Redact(ex.Message)}")
                Return ToolResult.Failed(
                    result.ToolId,
                    "无法解析宿主返回的结构化观察，不能确认操作结果",
                    data:=result.Data,
                    errorCode:=ExceptionClassifier.CodeObservationFailed,
                    userMessage:="插件无法安全验证 Office 操作结果；请选择其他验证或执行路径",
                    debugDetail:=ex.Message,
                    recoverable:=False,
                    observation:=result.Observation,
                    artifacts:=result.Artifacts)
            End Try
            Return result
        End Function

        Private Function FormatResultData(data As Object) As String
            If data Is Nothing Then Return ""

            Try
                Dim text As String
                If TypeOf data Is JToken Then
                    text = DirectCast(data, JToken).ToString(Formatting.None)
                Else
                    text = JsonConvert.SerializeObject(data, Formatting.None)
                End If

                Const maxLen As Integer = 12000
                If text.Length > maxLen Then
                    Return text.Substring(0, maxLen) & "...(truncated)"
                End If
                Return text
            Catch ex As Exception
                Return ""
            End Try
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

        #End Region

    End Class

End Namespace
