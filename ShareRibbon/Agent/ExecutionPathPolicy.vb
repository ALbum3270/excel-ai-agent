' ShareRibbon\Agent\ExecutionPathPolicy.vb
' Single source of truth for product execution path selection.

Imports System.Diagnostics

Namespace Agent

    ''' <summary>
    ''' Product path policy for Office AI Agent.
    ''' Primary path: Analyze (AiNativeRuntime) → AgentKernel/LoopEngine → Tool/Capability Executor.
    '''
    ''' Legacy paths (plain chat with intent, Ralph startLoop protocol, UseNewAgentKernel=false)
    ''' are compatibility only and must not become new feature entry points.
    ''' </summary>
    Public NotInheritable Class ExecutionPathPolicy

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Smart-mode product path always uses AgentKernel.
        ''' ConfigSettings.UseNewAgentKernel is deprecated and ignored for smart-mode routing.
        ''' </summary>
        Public Shared ReadOnly Property PreferAgentKernelForSmartMode As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' When AgentKernel fails to start, allow falling back to plain chat so the user still gets an answer.
        ''' </summary>
        Public Shared ReadOnly Property AllowChatFallbackOnAgentFailure As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' ExecuteJsonCommand / host JSON schemas are tool backends invoked by Agent tools or code cards,
        ''' not a parallel product entry for natural-language routing.
        ''' </summary>
        Public Shared ReadOnly Property JsonCommandIsToolBackendOnly As Boolean
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

        ''' <summary>
        ''' Logs once when obsolete UseNewAgentKernel=false is still set in process config.
        ''' </summary>
        Public Shared Sub WarnIfLegacyAgentKernelSwitchDisabled()
            If Not ConfigSettings.UseNewAgentKernel Then
                Debug.WriteLine("[ExecutionPathPolicy] UseNewAgentKernel=false is obsolete and ignored for smart-mode routing. Primary path remains AgentKernel.")
            End If
        End Sub

    End Class

End Namespace
