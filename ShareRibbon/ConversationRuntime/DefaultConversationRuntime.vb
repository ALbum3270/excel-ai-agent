Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Default OpenAI-compatible request builder used by the current HTTP streaming path.
''' </summary>
Public Class DefaultConversationRuntime
    Implements IConversationRuntime

    Private ReadOnly _contextComposer As IContextComposer
    Private ReadOnly _toolBroker As IToolBroker

    Public Sub New(contextComposer As IContextComposer, toolBroker As IToolBroker)
        _contextComposer = contextComposer
        _toolBroker = toolBroker
    End Sub

    Public Function BuildRequest(context As ChatRequestContext) As ChatRequestBuildResult Implements IConversationRuntime.BuildRequest
        Dim composition = _contextComposer.Compose(context)
        Dim requestObj As New JObject()

        requestObj("model") = context.ModelName
        requestObj("messages") = composition.Messages
        requestObj("stream") = context.Stream
        ReasoningRequestHelper.ApplyReasoningOptions(requestObj, context.ReasoningMode, context.ModelName, context.Platform, context.ApiUrl)

        Dim toolsArray = _toolBroker.GetTools(context)
        If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
            requestObj("tools") = toolsArray
        End If

        Return New ChatRequestBuildResult With {
            .RequestBody = requestObj.ToString(Formatting.None),
            .RagCount = composition.RagCount,
            .UsedContextBuilder = composition.UsedContextBuilder,
            .Trace = composition.Trace
        }
    End Function
End Class

Public Class ReasoningRequestHelper
    Public Const ReasoningDefault As String = "default"
    Public Const ReasoningEnabled As String = "enabled"
    Public Const ReasoningDisabled As String = "disabled"

    Public Shared Function NormalizeReasoningMode(mode As String) As String
        If String.IsNullOrWhiteSpace(mode) Then Return ReasoningDefault

        Select Case mode.Trim().ToLowerInvariant()
            Case ReasoningEnabled, "on", "true", "enable"
                Return ReasoningEnabled
            Case ReasoningDisabled, "off", "false", "disable"
                Return ReasoningDisabled
            Case Else
                Return ReasoningDefault
        End Select
    End Function

    Public Shared Sub ApplyReasoningOptions(requestObj As JObject, mode As String, Optional modelName As String = Nothing, Optional platform As String = Nothing, Optional apiUrl As String = Nothing)
        If requestObj Is Nothing Then Return

        Dim normalizedMode = NormalizeReasoningMode(mode)
        If normalizedMode = ReasoningDefault Then Return

        If UsesDeepSeekThinkingParameter(modelName, platform, apiUrl) Then
            requestObj("thinking") = New JObject() From {
                {"type", If(normalizedMode = ReasoningEnabled, "enabled", "disabled")}
            }
            Return
        End If

        Select Case normalizedMode
            Case ReasoningEnabled
                requestObj("enable_thinking") = True
            Case ReasoningDisabled
                requestObj("enable_thinking") = False
        End Select
    End Sub

    Private Shared Function UsesDeepSeekThinkingParameter(modelName As String, platform As String, apiUrl As String) As Boolean
        Dim text = $"{If(modelName, String.Empty)} {If(platform, String.Empty)} {If(apiUrl, String.Empty)}".ToLowerInvariant()

        If text.Contains("deepseek-v4") OrElse text.Contains("deepsee-v4") Then
            Return True
        End If

        If text.Contains("deepseek-reasoner") AndAlso text.Contains("api.deepseek.com") Then
            Return True
        End If

        Return False
    End Function
End Class
