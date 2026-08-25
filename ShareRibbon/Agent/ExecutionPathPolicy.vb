' ShareRibbon\Agent\ExecutionPathPolicy.vb
' Single source of truth for product execution path selection.

Namespace Agent

    ''' <summary>
    ''' Product path policy for Office AI Agent.
    ''' Primary path: Analyze (AiNativeRuntime) → AgentKernel/LoopEngine → Tool/Capability Executor.
    '''
    ''' Legacy Ralph startLoop protocol is compatibility only and must not become a new
    ''' feature entry point. Smart-mode chat and tool use share the adaptive Agent path.
    ''' </summary>
    Public NotInheritable Class ExecutionPathPolicy

        Private Sub New()
        End Sub

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
