Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Binds outputs produced during execution to dependent tool inputs. Planning happens
    ''' before read/compute results exist, so literal preview payloads in a plan can never be
    ''' authoritative for a data pipeline.
    ''' </summary>
    Public NotInheritable Class AgentToolDataflow
        Private ReadOnly _latestSuccessfulResults As New Dictionary(Of String, JToken)(StringComparer.OrdinalIgnoreCase)

        Public Sub RecordSuccess(result As ToolResult)
            If result Is Nothing OrElse Not result.Success OrElse
               String.IsNullOrWhiteSpace(result.ToolId) OrElse result.Data Is Nothing Then Return

            Dim token = TryCast(result.Data, JToken)
            If token Is Nothing Then token = JToken.FromObject(result.Data)
            _latestSuccessfulResults(result.ToolId) = token.DeepClone()
        End Sub

        Public Sub BindInputs(toolCall As ToolCall)
            If toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(toolCall.ToolId) Then Return
            If toolCall.Parameters Is Nothing Then toolCall.Parameters = New JObject()

            Select Case toolCall.ToolId.Trim().ToLowerInvariant()
                Case "pythoncompute"
                    Dim readResult As JToken = Nothing
                    If _latestSuccessfulResults.TryGetValue("ReadRange", readResult) Then
                        toolCall.Parameters("input") = BuildPythonInput(readResult)
                    End If

                Case "writedata"
                    Dim computeResult As JToken = Nothing
                    If _latestSuccessfulResults.TryGetValue("PythonCompute", computeResult) Then
                        toolCall.Parameters("data") = BuildWritableData(computeResult)
                    End If
            End Select
        End Sub

        Private Shared Function BuildPythonInput(readResult As JToken) As JToken
            Dim obj = TryCast(readResult, JObject)
            If obj Is Nothing Then Return readResult.DeepClone()

            Dim values = TryCast(obj("values"), JArray)
            If values Is Nothing OrElse values.Count = 0 Then Return obj.DeepClone()

            Dim headers = TryCast(values(0), JArray)
            If headers Is Nothing Then Return obj.DeepClone()

            Dim rows As New JArray()
            For index = 1 To values.Count - 1
                rows.Add(values(index).DeepClone())
            Next

            Dim input As New JObject()
            CopyMetadata(obj, input, "workbook")
            CopyMetadata(obj, input, "sheet")
            CopyMetadata(obj, input, "address")
            input("headers") = headers.DeepClone()
            input("rows") = rows
            Return input
        End Function

        Private Shared Function BuildWritableData(computeResult As JToken) As JToken
            If TypeOf computeResult Is JArray Then Return computeResult.DeepClone()

            Dim obj = TryCast(computeResult, JObject)
            If obj Is Nothing Then Return computeResult.DeepClone()

            Dim data = TryCast(obj("data"), JArray)
            If data IsNot Nothing Then Return data.DeepClone()

            Dim rows = TryCast(obj("rows"), JArray)
            If rows Is Nothing Then Return obj.DeepClone()

            Dim writableRows As New JArray()
            Dim headers = TryCast(obj("headers"), JArray)
            If headers IsNot Nothing Then writableRows.Add(headers.DeepClone())
            For Each row In rows
                writableRows.Add(row.DeepClone())
            Next
            Return writableRows
        End Function

        Private Shared Sub CopyMetadata(source As JObject, target As JObject, propertyName As String)
            Dim value = source(propertyName)
            If value IsNot Nothing Then target(propertyName) = value.DeepClone()
        End Sub
    End Class

End Namespace
