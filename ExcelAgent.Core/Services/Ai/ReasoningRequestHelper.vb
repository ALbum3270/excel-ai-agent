Imports Newtonsoft.Json.Linq

Public NotInheritable Class ReasoningRequestHelper
    Public Const ReasoningDefault As String = "default"
    Public Const ReasoningEnabled As String = "enabled"
    Public Const ReasoningDisabled As String = "disabled"

    Private Sub New()
    End Sub

    Public Shared Function NormalizeReasoningMode(mode As String) As String
        Select Case If(mode, "").Trim().ToLowerInvariant()
            Case ReasoningEnabled, "on", "true", "enable"
                Return ReasoningEnabled
            Case ReasoningDisabled, "off", "false", "disable"
                Return ReasoningDisabled
            Case Else
                Return ReasoningDefault
        End Select
    End Function

    Public Shared Sub ApplyReasoningOptions(requestObj As JObject,
                                            mode As String,
                                            Optional modelName As String = Nothing,
                                            Optional platform As String = Nothing,
                                            Optional apiUrl As String = Nothing)
        If requestObj Is Nothing Then Return
        Dim normalized = NormalizeReasoningMode(mode)
        If normalized = ReasoningDefault Then Return

        Dim identity = $"{If(modelName, "")} {If(platform, "")} {If(apiUrl, "")}".ToLowerInvariant()
        If identity.Contains("deepseek-v4") OrElse
           (identity.Contains("deepseek-reasoner") AndAlso identity.Contains("api.deepseek.com")) Then
            requestObj("thinking") = New JObject From {{"type", normalized}}
            Return
        End If
        requestObj("enable_thinking") = (normalized = ReasoningEnabled)
    End Sub
End Class
