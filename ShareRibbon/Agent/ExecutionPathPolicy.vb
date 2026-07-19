' ShareRibbon\Agent\ExecutionPathPolicy.vb
' Single source of truth for product execution path selection.

Namespace Agent

    ''' <summary>
    ''' Product path policy for Office AI Agent.
    ''' Primary path: Analyze (AiNativeRuntime) → AgentKernel/LoopEngine → Tool/Capability Executor.
    '''
    ''' Legacy paths (plain chat with intent and Ralph startLoop protocol) are compatibility only
    ''' and must not become new feature entry points.
    ''' </summary>
    Public NotInheritable Class ExecutionPathPolicy

        Private Sub New()
        End Sub

        ''' <summary>
        ''' When AgentKernel fails to start, allow falling back to plain chat so the user still gets an answer.
        ''' </summary>
        Public Shared ReadOnly Property AllowChatFallbackOnAgentFailure As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Host ActionHarness (e.g. WordActionHarness) may handle deterministic high-confidence
        ''' capabilities as a fast path. It should eventually report observe/explain back into the agent loop.
        ''' </summary>
        Public Shared ReadOnly Property AllowHostCapabilityFastPath As Boolean
            Get
                Return True
            End Get
        End Property

    End Class

End Namespace
