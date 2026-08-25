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
                        _memory.ObserveTaskContext(refreshedContext)
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
            Public Property CriterionEvidence As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
            Public Property OutcomeContract As OutcomeContract
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

                Dim criterionEvidence As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
                Dim criterionEvidenceObject = TryCast(obj("criterionEvidence"), JObject)
                If criterionEvidenceObject IsNot Nothing Then
                    For Each prop In criterionEvidenceObject.Properties()
                        Dim criterionId = If(prop.Name, "").Trim()
                        If String.IsNullOrWhiteSpace(criterionId) Then Continue For
                        Dim values As New List(Of String)()
                        Dim array = TryCast(prop.Value, JArray)
                        If array IsNot Nothing Then
                            For Each item In array
                                Dim value = If(item?.ToString(), "").Trim()
                                If value.Length > 0 AndAlso Not values.Contains(value, StringComparer.OrdinalIgnoreCase) Then values.Add(value)
                            Next
                        Else
                            Dim value = If(prop.Value?.ToString(), "").Trim()
                            If value.Length > 0 Then values.Add(value)
                        End If
                        If values.Count > 0 Then criterionEvidence(criterionId) = values
                    Next
                End If

                Return New ReactDecision With {
                    .Kind = kind,
                    .Thought = If(obj("thought")?.ToString(), If(obj("analysis")?.ToString(), "")),
                    .Action = action,
                    .Message = If(obj("message")?.ToString(), If(obj("reason")?.ToString(), If(obj("finalAnswer")?.ToString(), ""))),
                    .Evidence = evidence,
                    .CriterionEvidence = criterionEvidence,
                    .OutcomeContract = ParseOutcomeContract(TryCast(obj("outcomeContract"), JObject), "")
                }
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Unable to parse ReAct decision: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' An act decision and a final response/completion projection are mutually exclusive
        ''' control states. Executing an action from a response that simultaneously claims
        ''' completion can repeat an already-proven mutation. Reject the contradictory turn
        ''' and let the model choose one state on the next observation cycle.
        ''' </summary>
        Private Shared Function HasConflictingCompletionPayload(decision As ReactDecision) As Boolean
            If decision Is Nothing OrElse Not String.Equals(decision.Kind, "act", StringComparison.OrdinalIgnoreCase) Then Return False
            Return Not String.IsNullOrWhiteSpace(decision.Message) OrElse
                (decision.Evidence IsNot Nothing AndAlso decision.Evidence.Count > 0) OrElse
                (decision.CriterionEvidence IsNot Nothing AndAlso decision.CriterionEvidence.Count > 0) OrElse
                decision.OutcomeContract IsNot Nothing
        End Function

    End Class

End Namespace
