' ShareRibbon\Controls\Services\AgentInternalRequestPolicy.vb
' Provider options for machine-readable Agent planning/action/repair requests.

''' <summary>
''' Internal Agent calls respect the user's provider reasoning preference. They deliberately
''' do not impose a separate token budget or wall-clock timeout: reasoning models may need to
''' finish their reasoning before emitting the small machine-readable envelope.
''' </summary>
Public NotInheritable Class AgentInternalRequestPolicy
    Private Sub New()
    End Sub

    Public Shared ReadOnly Property MaxTokens As Integer?
        Get
            Return Nothing
        End Get
    End Property

    Public Shared ReadOnly Property RequestTimeout As System.TimeSpan?
        Get
            Return Nothing
        End Get
    End Property

    Public Shared Function CreateCancellationSource() As System.Threading.CancellationTokenSource
        Dim timeout = RequestTimeout
        If timeout.HasValue Then
            Return New System.Threading.CancellationTokenSource(timeout.Value)
        End If
        Return New System.Threading.CancellationTokenSource()
    End Function

    Public Shared ReadOnly Property ReasoningMode As String
        Get
            Return ResolveReasoningMode(ConfigSettings.ReasoningMode)
        End Get
    End Property

    Public Shared Function ResolveReasoningMode(configuredMode As String) As String
        Return ReasoningRequestHelper.NormalizeReasoningMode(configuredMode)
    End Function
End Class
