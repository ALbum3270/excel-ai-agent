Public Class ChatSystemPromptResolver
    Public Function ResolveSystemPrompt(
        requestedSystemPrompt As String,
        appInfo As ApplicationInfo,
        intentResult As IntentResult,
        responseMode As String) As String

        If Not String.IsNullOrWhiteSpace(requestedSystemPrompt) Then
            Return requestedSystemPrompt
        End If

        Dim appType = If(appInfo IsNot Nothing, appInfo.Type.ToString(), "Excel")
        Dim context As New PromptContext With {
            .ApplicationType = appType,
            .IntentResult = intentResult,
            .FunctionMode = responseMode
        }

        Dim resolved = PromptManager.Instance.GetCombinedPrompt(context)
        If Not String.IsNullOrWhiteSpace(resolved) Then Return resolved

        If Not String.IsNullOrWhiteSpace(ConfigSettings.propmtContent) Then
            Return ConfigSettings.propmtContent
        End If

        Return "你是一个 Office AI 助手，请根据用户需求提供简洁、准确的回答。"
    End Function
End Class
