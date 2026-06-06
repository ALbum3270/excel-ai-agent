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

        Dim toolsArray = _toolBroker.GetTools(context)
        If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
            requestObj("tools") = toolsArray
        End If

        Return New ChatRequestBuildResult With {
            .RequestBody = requestObj.ToString(Formatting.None),
            .RagCount = composition.RagCount,
            .UsedContextBuilder = composition.UsedContextBuilder
        }
    End Function
End Class
