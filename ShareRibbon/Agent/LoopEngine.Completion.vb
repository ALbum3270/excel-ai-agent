Imports System.Collections.Generic

Namespace Agent

    Public Partial Class LoopEngine

        ''' <summary>
        ''' Accepts a model semantic completion verdict only after projecting its cited host
        ''' evidence into the deterministic goal verifier.  The model owns the semantic
        ''' judgment; the harness owns exact targets, effects, values and lineage.
        ''' </summary>
        Private Function TryAcceptCompletionDecision(session As AgentSession,
                                                     decision As ReactDecision,
                                                     readOnlyEvidenceCount As Integer,
                                                     ByRef completionMessage As String) As Boolean
            ' Complete is both a control decision and the final assistant turn.  A tool result
            ' can never close the loop, and an empty control-only completion would leave the
            ' user with no answer even when the underlying work succeeded.
            If String.IsNullOrWhiteSpace(decision?.Message) Then
                _memory.SetWorking(
                    "lastObservation",
                    "Completion was rejected because it contained no final user-facing response. Based on the latest tool observation, return decision=complete with a self-contained message that directly answers the user's question or states the observed result. Do not call another tool merely to produce wording.")
                Return False
            End If

            ' A read-only answer is delivered in chat; it is not an Office host-state effect.
            ' Requiring an OutcomeContract here makes an exact ReadRange/PythonCompute result
            ' impossible to complete because there is intentionally no workbook mutation to
            ' project.  The deterministic acceptance boundary is instead: structured evidence
            ' was captured and every capability explicitly required by the user succeeded.
            If IsReadOnlyAnswerSpec(session?.Spec) Then
                If session.Spec.RequiresHostExecution AndAlso readOnlyEvidenceCount <= 0 Then
                    _memory.SetWorking(
                        "lastObservation",
                        "Completion was rejected because no structured read/compute evidence has been captured. Read the smallest complete workbook range needed for the answer, then decide again.")
                    Return False
                End If

                Dim capabilityError = AgentExecutionContract.ValidateOutcome(session)
                If Not String.IsNullOrWhiteSpace(capabilityError) Then
                    _memory.SetWorking(
                        "lastObservation",
                        $"Completion was rejected because a user-required capability has not succeeded: {capabilityError}")
                    Return False
                End If

                completionMessage = decision.Message.Trim()
                MarkRemainingPlanHintsSkipped(session.Plan)
                _memory.SetWorking(
                    "lastObservation",
                    If(session.Spec.RequiresHostExecution,
                       "The model declared the read-only analysis complete; structured evidence and required capabilities were verified. The final answer will now be delivered without modifying Office.",
                       "The model completed a conversational turn inside the unified adaptive Agent; no Office host evidence was required and no mutation was permitted."))
                Return True
            End If

            Dim persistedResumeContract = session.Spec.OutcomeContract IsNot Nothing AndAlso
                session.Spec.OutcomeContract.Frozen
            Dim completionEvidence As New List(Of String)(If(decision.Evidence, New List(Of String)()))

            If Not persistedResumeContract Then
                Dim completionContract = decision.OutcomeContract
                If completionContract Is Nothing Then
                    Dim observedProjectionError As String = ""
                    If (decision.CriterionEvidence Is Nothing OrElse decision.CriterionEvidence.Count = 0) AndAlso
                       session.Plan?.OutcomeContract IsNot Nothing Then
                        observedProjectionError = BuildGroundedProvisionalCompletionProjection(
                            session,
                            completionEvidence,
                            completionContract)
                    End If
                    If completionContract Is Nothing Then
                        Dim evidenceProjectionError = BuildObservedCompletionProjection(
                            session,
                            completionEvidence,
                            decision.CriterionEvidence,
                            completionContract)
                        If Not String.IsNullOrWhiteSpace(observedProjectionError) AndAlso
                           Not String.IsNullOrWhiteSpace(evidenceProjectionError) Then
                            observedProjectionError &= " " & evidenceProjectionError
                        Else
                            observedProjectionError = evidenceProjectionError
                        End If
                    End If
                    If Not String.IsNullOrWhiteSpace(observedProjectionError) Then
                        _memory.SetWorking(
                            "lastObservation",
                            $"Completion evidence was insufficient: {observedProjectionError} Inspect the real destination with a read/discovery tool if needed, then return complete with evidence (and criterionEvidence when the Goal has multiple criteria). Do not mutate Office merely to manufacture verification metadata.")
                        Return False
                    End If
                Else
                    Dim groundingError = GroundCompletionProjectionInEvidence(session, completionContract, completionEvidence)
                    If Not String.IsNullOrWhiteSpace(groundingError) Then
                        _memory.SetWorking(
                            "lastObservation",
                            $"Completion verification projection was rejected: {groundingError} Cite the exact evidence IDs for each Goal criterion. Do not execute another Office mutation when the requested state already exists.")
                        Return False
                    End If
                End If

                NormalizePlannerLineageHints(completionContract, session.AppType)
                Dim completionPlan As New ExecutionPlan With {.OutcomeContract = completionContract}
                Dim projectionError = SealVerificationProjection(
                    session,
                    completionPlan,
                    executionAppType:=session.AppType)
                If Not String.IsNullOrWhiteSpace(projectionError) Then
                    ResetTransientCompletionContract(session)
                    _memory.SetWorking(
                        "lastObservation",
                        $"Completion verification projection was rejected: {projectionError} Revise only the verification projection from current host evidence; do not mutate Office to make a guessed contract true.")
                    Return False
                End If
            End If

            Dim currentContext = TryCast(_memory.GetWorking("lastContextPack"), Context.ContextPack)
            Dim completionError = AgentGoalVerifier.Validate(
                session,
                currentContext,
                completionEvidence,
                readOnlyEvidenceCount)
            If String.IsNullOrWhiteSpace(completionError) Then
                completionMessage = decision.Message.Trim()
                MarkRemainingPlanHintsSkipped(session.Plan)
                _memory.SetWorking(
                    "lastObservation",
                    "The model declared completion and deterministic goal verification passed.")
                Return True
            End If

            _memory.SetWorking(
                "lastObservation",
                $"Completion was rejected by deterministic goal verification: {completionError} If Office already contains the requested result, revise the verification projection from the cited host evidence instead of changing Office. Otherwise choose exactly one next action that closes the real evidence gap.")
            AppLogger.Warn("LoopEngine", $"Rejected premature completion: {completionError}")
            If Not persistedResumeContract Then ResetTransientCompletionContract(session)
            Return False
        End Function

        Private Shared Sub ResetTransientCompletionContract(session As AgentSession)
            session.Spec.OutcomeContract = Nothing
        End Sub

    End Class

End Namespace
