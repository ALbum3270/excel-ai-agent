Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Binds outputs produced during execution to dependent tool inputs. Planning happens
    ''' before read/compute results exist, so literal preview payloads in a plan can never be
    ''' authoritative for a data pipeline.
    ''' </summary>
    Public NotInheritable Class AgentToolDataflow
        Private NotInheritable Class ProducedValue
            Public Property Data As JToken
            Public Property EvidenceId As String
        End Class

        Private ReadOnly _latestSuccessfulResults As New Dictionary(Of String, ProducedValue)(StringComparer.OrdinalIgnoreCase)

        Public Sub RecordSuccess(result As ToolResult)
            RecordSuccess(result, "")
        End Sub

        Public Sub RecordSuccess(result As ToolResult, evidenceId As String)
            If result Is Nothing OrElse Not result.Success OrElse
               String.IsNullOrWhiteSpace(result.ToolId) OrElse result.Data Is Nothing Then Return

            Dim token = TryCast(result.Data, JToken)
            If token Is Nothing Then token = JToken.FromObject(result.Data)
            _latestSuccessfulResults(result.ToolId) = New ProducedValue With {
                .Data = token.DeepClone(),
                .EvidenceId = If(evidenceId, "").Trim()
            }
        End Sub

        Public Sub BindInputs(toolCall As ToolCall)
            BindInputsWithDependencies(toolCall)
        End Sub

        ''' <summary>
        ''' Binds current runtime values and returns the evidence records that supplied those
        ''' values.  Completion verification uses this lineage to distinguish a persisted
        ''' computation result from an unrelated write that merely happened to succeed.
        ''' </summary>
        Public Function BindInputsWithDependencies(toolCall As ToolCall) As List(Of String)
            Dim dependencies As New List(Of String)()
            If toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(toolCall.ToolId) Then Return dependencies
            If toolCall.Parameters Is Nothing Then toolCall.Parameters = New JObject()

            Select Case toolCall.ToolId.Trim().ToLowerInvariant()
                Case "pythoncompute"
                    Dim readResult As ProducedValue = Nothing
                    If _latestSuccessfulResults.TryGetValue("ReadRange", readResult) Then
                        toolCall.Parameters("input") = BuildPythonInput(readResult.Data)
                        AddDependency(dependencies, readResult.EvidenceId)
                    End If

                Case "writedata"
                    Dim computeResult As ProducedValue = Nothing
                    If _latestSuccessfulResults.TryGetValue("PythonCompute", computeResult) Then
                        toolCall.Parameters("data") = BuildWritableData(computeResult.Data)
                        AddDependency(dependencies, computeResult.EvidenceId)
                    End If
            End Select

            Return dependencies
        End Function

        Private Shared Sub AddDependency(result As List(Of String), evidenceId As String)
            If result Is Nothing OrElse String.IsNullOrWhiteSpace(evidenceId) Then Return
            If Not result.Contains(evidenceId, StringComparer.OrdinalIgnoreCase) Then result.Add(evidenceId)
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
