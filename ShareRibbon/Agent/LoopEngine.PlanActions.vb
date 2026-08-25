Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        ''' <summary>
        ''' Parses the canonical command envelope already emitted in a plan step. Legacy
        ''' action envelopes remain accepted; a multi-command envelope falls back to the
        ''' normal model action path until it has a dedicated batch execution contract.
        ''' </summary>
        Private Function ParsePlannedToolCall(code As String) As ToolCall
            If String.IsNullOrWhiteSpace(code) Then Return Nothing

            Try
                Dim jsonStr = ExtractJson(code)
                If String.IsNullOrWhiteSpace(jsonStr) Then Return Nothing
                Dim obj = JObject.Parse(jsonStr)

                If obj("action") IsNot Nothing Then Return ParseToolCall(jsonStr)

                Dim rootTool = obj("tool")?.ToString()
                If Not String.IsNullOrWhiteSpace(rootTool) Then
                    Dim rootParameters = TryCast(obj("parameters"), JObject)
                    If rootParameters Is Nothing Then rootParameters = TryCast(obj("params"), JObject)
                    If rootParameters Is Nothing Then rootParameters = New JObject()
                    Return New ToolCall With {
                        .ToolId = rootTool,
                        .Parameters = rootParameters
                    }
                End If

                Dim command = obj("command")?.ToString()
                If String.IsNullOrWhiteSpace(command) Then
                    Dim commands = TryCast(obj("commands"), JArray)
                    If commands Is Nothing OrElse commands.Count <> 1 Then Return Nothing
                    obj = TryCast(commands(0), JObject)
                    If obj Is Nothing Then Return Nothing
                    command = obj("command")?.ToString()
                End If
                If String.IsNullOrWhiteSpace(command) Then Return Nothing

                Dim parameters = TryCast(obj("params"), JObject)
                If parameters Is Nothing Then parameters = New JObject()
                Return New ToolCall With {
                    .ToolId = command,
                    .Parameters = parameters
                }
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

    End Class

End Namespace
