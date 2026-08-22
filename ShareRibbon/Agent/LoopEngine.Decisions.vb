Imports System.Collections.Generic
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
            Public Property Evidence As New List(Of String)()
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

                Dim evidence As New List(Of String)()
                Dim evidenceArray = TryCast(obj("evidence"), JArray)
                If evidenceArray IsNot Nothing Then
                    For Each item In evidenceArray
                        Dim value = If(item?.ToString(), "").Trim()
                        If Not String.IsNullOrWhiteSpace(value) Then evidence.Add(value)
                    Next
                Else
                    Dim value = If(obj("evidence")?.ToString(), "").Trim()
                    If Not String.IsNullOrWhiteSpace(value) Then evidence.Add(value)
                End If

                Return New ReactDecision With {
                    .Kind = kind,
                    .Thought = If(obj("thought")?.ToString(), If(obj("analysis")?.ToString(), "")),
                    .Action = action,
                    .Message = If(obj("message")?.ToString(), If(obj("reason")?.ToString(), If(obj("finalAnswer")?.ToString(), ""))),
                    .Evidence = evidence
                }
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Unable to parse ReAct decision: {ex.Message}")
                Return Nothing
            End Try
        End Function

    End Class

End Namespace
