Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        Private Async Function ThinkAsync(session As AgentSession,
                                           planStep As PlanStep,
                                           systemPrompt As String) As Task(Of String)
            If CaptureContextPack IsNot Nothing Then
                Try
                    Dim refreshedContext = CaptureContextPack.Invoke()
                    If refreshedContext IsNot Nothing Then
                        _memory.SetWorking("lastContextPack", refreshedContext)
                    End If
                Catch ex As Exception
                    AppLogger.Warn("LoopEngine", $"ContextPack refresh failed; preserving previous context: {ex.Message}")
                End Try
            End If
            Dim lastObservation = _memory.GetWorkingString("lastObservation")
            Dim prompt = _promptManager.BuildReactPrompt(session, planStep, _memory, lastObservation)
            Dim history = _memory.GetRecentMessages(10)
            Return Await SendAIRequest(prompt, systemPrompt, history)
        End Function

        Private NotInheritable Class ReactDecision
            Public Property Kind As String = ""
            Public Property Thought As String = ""
            Public Property Action As ToolCall
            Public Property Message As String = ""
        End Class

        Private Function ParseReactDecision(response As String) As ReactDecision
            Try
                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrWhiteSpace(jsonStr) Then Return Nothing

                Dim obj = JObject.Parse(jsonStr)
                Dim action = ParsePlannedToolCall(jsonStr)
                Dim kind = If(obj("decision")?.ToString(), "").Trim().ToLowerInvariant()
                If String.IsNullOrWhiteSpace(kind) AndAlso action IsNot Nothing Then kind = "act"

                Select Case kind
                    Case "action", "continue", "next"
                        kind = "act"
                    Case "done", "finished", "success"
                        kind = "complete"
                    Case "re-plan", "revise"
                        kind = "replan"
                    Case "error", "blocked"
                        kind = "fail"
                End Select

                If kind <> "act" AndAlso kind <> "complete" AndAlso kind <> "replan" AndAlso kind <> "fail" Then
                    Return Nothing
                End If
                If kind = "act" AndAlso action Is Nothing Then Return Nothing

                Return New ReactDecision With {
                    .Kind = kind,
                    .Thought = If(obj("thought")?.ToString(), If(obj("analysis")?.ToString(), "")),
                    .Action = action,
                    .Message = If(obj("message")?.ToString(), If(obj("reason")?.ToString(), If(obj("finalAnswer")?.ToString(), "")))
                }
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Unable to parse ReAct decision: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Private Function ValidateMilestoneTools(plan As ExecutionPlan) As String
            For Each stepItem In plan.Steps
                Dim milestoneTool = If(stepItem.ToolHint, "").Trim()
                If String.IsNullOrWhiteSpace(milestoneTool) Then
                    milestoneTool = If(ParsePlannedToolCall(stepItem.Code)?.ToolId, "").Trim()
                End If
                If String.IsNullOrWhiteSpace(milestoneTool) Then
                    Return $"规划步骤 {stepItem.StepNumber} 缺少可验证的 toolHint；高层骨架中每个里程碑必须对应一个已注册工具。"
                End If
                If _toolRegistry.GetTool(milestoneTool) Is Nothing Then
                    Return $"规划步骤 {stepItem.StepNumber} 引用了未注册的 toolHint：{milestoneTool}。"
                End If
            Next
            Return ""
        End Function

    End Class

End Namespace
