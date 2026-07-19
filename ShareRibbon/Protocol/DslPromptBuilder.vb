' ShareRibbon\Protocol\DslPromptBuilder.vb
' 保留旧 JSON/DSL 响应的格式检测；旧 Prompt Builder 已由 Agent PromptManager 取代。

''' <summary>
''' DSL 协议检测器。
''' </summary>
Public Class DslProtocolDetector

    Public Shared Function IsDslFormat(jsonText As String) As Boolean
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
            If json("version") IsNot Nothing AndAlso
               json("protocol") IsNot Nothing AndAlso
               json("instructions") IsNot Nothing Then
                Return True
            End If

            If json("operation") IsNot Nothing Then
                Dim operationName = json("operation").ToString().ToLowerInvariant()
                If operationName = "reformat" OrElse operationName = "proofread" Then
                    Return True
                End If
            End If
        Catch
        End Try

        Return False
    End Function

    Public Shared Function IsLegacyJsonCommandFormat(jsonText As String) As Boolean
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
            Return json("command") IsNot Nothing OrElse json("commands") IsNot Nothing
        Catch
            Return False
        End Try
    End Function

    Public Shared Function DetectFormat(jsonText As String) As InstructionFormat
        If String.IsNullOrWhiteSpace(jsonText) Then Return InstructionFormat.None
        If IsDslFormat(jsonText) Then Return InstructionFormat.DslJson
        If IsLegacyJsonCommandFormat(jsonText) Then Return InstructionFormat.LegacyJsonCommand

        If jsonText.Trim().StartsWith("[", StringComparison.Ordinal) Then
            Return InstructionFormat.ProofreadJson
        End If

        Return InstructionFormat.None
    End Function

End Class
