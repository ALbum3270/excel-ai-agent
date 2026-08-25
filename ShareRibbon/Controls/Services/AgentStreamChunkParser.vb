Imports Newtonsoft.Json.Linq

''' <summary>
''' A provider-neutral delta extracted from one SSE data line used by Agent calls.
''' The Agent still receives the complete structured content after the stream ends,
''' while the UI can render reasoning and receive progress incrementally.
''' </summary>
Public NotInheritable Class AgentStreamDelta
    Public Property ReasoningDelta As String = ""
    Public Property ContentDelta As String = ""
    Public Property Done As Boolean
End Class

Public NotInheritable Class AgentStreamChunkParser
    Private Sub New()
    End Sub

    Public Shared Function ParseDataLine(line As String) As AgentStreamDelta
        Dim result As New AgentStreamDelta()
        If String.IsNullOrWhiteSpace(line) Then Return result

        Dim payload = line.Trim()
        If payload.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
            payload = payload.Substring(5).TrimStart()
        End If

        If String.Equals(payload, "[DONE]", StringComparison.OrdinalIgnoreCase) Then
            result.Done = True
            Return result
        End If

        If Not payload.StartsWith("{", StringComparison.Ordinal) Then Return result

        Try
            Dim obj = JObject.Parse(payload)

            ' OpenAI-compatible providers: Qwen, DeepSeek, OpenAI and compatible gateways.
            Dim delta = TryCast(obj.SelectToken("choices[0].delta"), JObject)
            If delta IsNot Nothing Then
                result.ReasoningDelta = FirstText(delta("reasoning_content"), delta("reasoning"))
                result.ContentDelta = FirstText(delta("content"), delta("text"))
                Return result
            End If

            ' OpenAI Responses API streaming events.
            Dim eventType = If(obj("type"), "").ToString()
            Select Case eventType
                Case "response.reasoning_summary_text.delta", "response.reasoning_text.delta"
                    result.ReasoningDelta = FirstText(obj("delta"))
                Case "response.output_text.delta"
                    result.ContentDelta = FirstText(obj("delta"))
                Case "response.completed", "response.failed", "response.incomplete"
                    result.Done = True
            End Select
            If result.Done OrElse result.ReasoningDelta.Length > 0 OrElse result.ContentDelta.Length > 0 Then Return result

            ' Anthropic Messages streaming events.
            If String.Equals(eventType, "content_block_delta", StringComparison.OrdinalIgnoreCase) Then
                Dim anthropicDelta = TryCast(obj("delta"), JObject)
                If anthropicDelta IsNot Nothing Then
                    Dim deltaType = If(anthropicDelta("type"), "").ToString()
                    If String.Equals(deltaType, "thinking_delta", StringComparison.OrdinalIgnoreCase) Then
                        result.ReasoningDelta = FirstText(anthropicDelta("thinking"))
                    Else
                        result.ContentDelta = FirstText(anthropicDelta("text"))
                    End If
                End If
            ElseIf String.Equals(eventType, "message_stop", StringComparison.OrdinalIgnoreCase) Then
                result.Done = True
            End If
        Catch
            ' Ignore keep-alives and provider metadata that are not content deltas.
        End Try

        Return result
    End Function

    Private Shared Function FirstText(ParamArray tokens() As JToken) As String
        If tokens Is Nothing Then Return ""
        For Each token In tokens
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Continue For
            If token.Type = JTokenType.String Then Return token.ToString()
            If token.Type = JTokenType.Array Then
                Dim parts As New List(Of String)()
                For Each item In DirectCast(token, JArray)
                    Dim text = If(item("text"), item("content"))
                    If text IsNot Nothing Then parts.Add(text.ToString())
                Next
                If parts.Count > 0 Then Return String.Join("", parts)
            End If
            Return token.ToString()
        Next
        Return ""
    End Function
End Class
