Imports System.Threading.Tasks

Namespace Agent

    Public Partial Class LoopEngine

        ''' <summary>
        ''' Captures a user-facing understanding and an optional provisional verification
        ''' projection.  New runs do not seal model-authored outcome assertions before Office
        ''' has produced observations; only a persisted, already-sealed resume contract is
        ''' authoritative here.
        ''' </summary>
        Private Async Function PrepareInitialPlanAsync(session As AgentSession,
                                                       systemPrompt As String,
                                                       skill As AgentSkill) As Task(Of String)
            Dim existingContract = session.Spec.OutcomeContract

            ' Resume state is authoritative.  Verify its seal before any model request and
            ' preserve the persisted projection; a model must never repair or replace a frozen
            ' verification contract.  If only the contract survived, use a metadata-only plan
            ' so adaptive ReAct can continue without inventing future actions.
            If existingContract IsNot Nothing AndAlso existingContract.Frozen Then
                Dim outcomeContractError = SealVerificationProjection(
                    session,
                    session.Plan,
                    executionAppType:=session.AppType)
                If Not String.IsNullOrWhiteSpace(outcomeContractError) Then
                    Return outcomeContractError
                End If

                If session.Plan Is Nothing Then
                    session.Plan = CreateFallbackPlan(session)
                End If
                session.Plan.OutcomeContract = existingContract
                Return ""
            End If

            ' A conversational turn still uses the adaptive Agent and its complete/fail
            ' protocol, but it has no future host state to project. Avoid a redundant model
            ' planning call and let the first ReAct decision answer directly.
            If IsConversationOnlySpec(session.Spec) Then
                session.Plan = CreateFallbackPlan(session)
                Return ""
            End If

            Dim candidatePlan = Await GeneratePlanAsync(session, systemPrompt, skill)
            If candidatePlan Is Nothing Then candidatePlan = CreateFallbackPlan(session)

            ' Legacy providers may still emit a full list of future steps.  Discard it.  The
            ' only executable decision is produced later by ThinkAsync from the latest world.
            candidatePlan.Steps.Clear()

            ' Keep a provisional projection only as compatibility input for a later completion
            ' attempt.  It is deliberately absent from AgentTaskSpec and is neither frozen nor
            ' allowed to gate actions.  Completion may replace it using post-observation facts.
            session.Spec.OutcomeContract = Nothing
            session.Plan = candidatePlan
            Return ""
        End Function

    End Class

End Namespace
