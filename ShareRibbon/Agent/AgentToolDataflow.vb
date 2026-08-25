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
                        ' `$from: ReadRange` means the exact observed producer value. Invisible
                        ' reshaping made the model inspect a `values` payload while Python received
                        ' a different `headers/rows` object, so valid code failed at runtime.
                        toolCall.Parameters("input") = readResult.Data.DeepClone()
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

        Private Shared Function BuildWritableData(computeResult As JToken) As JToken
            If computeResult Is Nothing OrElse computeResult.Type = JTokenType.Null Then
                Return WrapRow(New JArray(JValue.CreateNull()))
            End If

            Dim array = TryCast(computeResult, JArray)
            If array IsNot Nothing Then Return BuildWritableArray(array)

            Dim obj = TryCast(computeResult, JObject)
            If obj Is Nothing Then
                Return WrapRow(New JArray(ToExcelCell(computeResult)))
            End If

            ' Structured Python results may expose their intended table explicitly.
            Dim data = TryCast(obj("data"), JArray)
            If data IsNot Nothing Then Return BuildWritableArray(data)

            Dim rows = TryCast(obj("rows"), JArray)
            If rows IsNot Nothing Then
                Dim headers = TryCast(obj("headers"), JArray)
                Return BuildWritableRows(rows, headers)
            End If

            ' A named scalar result (for example summary statistics) is a record, not a
            ' value Excel COM can assign to Range.Value2.  Project arbitrary JSON records
            ' deterministically to a two-column key/value table instead of passing JObject
            ' through to the host as an unsupported scalar.
            Dim recordRows As New JArray()
            recordRows.Add(New JArray("字段", "值"))
            For Each prop In obj.Properties()
                recordRows.Add(New JArray(prop.Name, ToExcelCell(prop.Value)))
            Next
            Return recordRows
        End Function

        Private Shared Function BuildWritableArray(array As JArray) As JArray
            If array.Count = 0 Then Return New JArray()

            If array.All(Function(item) TypeOf item Is JObject) Then
                Return BuildObjectRows(array.OfType(Of JObject)().ToList(), Nothing)
            End If

            If array.All(Function(item) TypeOf item Is JArray) Then
                Return Rectangularize(array.OfType(Of JArray)())
            End If

            Dim row As New JArray()
            For Each item In array
                row.Add(ToExcelCell(item))
            Next
            Return WrapRow(row)
        End Function

        Private Shared Function BuildWritableRows(rows As JArray, headers As JArray) As JArray
            If rows.Count > 0 AndAlso rows.All(Function(item) TypeOf item Is JObject) Then
                Return BuildObjectRows(rows.OfType(Of JObject)().ToList(), headers)
            End If

            Dim combined As New List(Of JArray)()
            If headers IsNot Nothing Then combined.Add(ToExcelRow(headers))
            For Each row In rows
                Dim rowArray = TryCast(row, JArray)
                If rowArray IsNot Nothing Then
                    combined.Add(ToExcelRow(rowArray))
                Else
                    combined.Add(New JArray(ToExcelCell(row)))
                End If
            Next
            Return Rectangularize(combined)
        End Function

        Private Shared Function BuildObjectRows(records As IList(Of JObject), headers As JArray) As JArray
            Dim columnNames As New List(Of String)()
            If headers IsNot Nothing Then
                For Each header In headers
                    Dim name = If(header?.ToString(), "")
                    If Not columnNames.Contains(name, StringComparer.Ordinal) Then columnNames.Add(name)
                Next
            End If
            For Each record In records
                For Each prop In record.Properties()
                    If Not columnNames.Contains(prop.Name, StringComparer.Ordinal) Then columnNames.Add(prop.Name)
                Next
            Next

            Dim result As New JArray()
            result.Add(New JArray(columnNames.Select(Function(name) CType(New JValue(name), JToken))))
            For Each record In records
                Dim row As New JArray()
                For Each name In columnNames
                    row.Add(ToExcelCell(record.GetValue(name, StringComparison.Ordinal)))
                Next
                result.Add(row)
            Next
            Return result
        End Function

        Private Shared Function Rectangularize(rows As IEnumerable(Of JArray)) As JArray
            Dim normalized = rows.Select(Function(row) ToExcelRow(row)).ToList()
            If normalized.Count = 0 Then Return New JArray()

            Dim columnCount = normalized.Max(Function(row) row.Count)
            Dim result As New JArray()
            For Each row In normalized
                While row.Count < columnCount
                    row.Add(JValue.CreateNull())
                End While
                result.Add(row)
            Next
            Return result
        End Function

        Private Shared Function ToExcelRow(row As JArray) As JArray
            Dim result As New JArray()
            If row Is Nothing Then Return result
            For Each value In row
                result.Add(ToExcelCell(value))
            Next
            Return result
        End Function

        Private Shared Function WrapRow(row As JArray) As JArray
            Dim result As New JArray()
            result.Add(row)
            Return result
        End Function

        Private Shared Function ToExcelCell(value As JToken) As JToken
            If value Is Nothing OrElse value.Type = JTokenType.Null OrElse value.Type = JTokenType.Undefined Then
                Return JValue.CreateNull()
            End If
            If TypeOf value Is JValue Then Return value.DeepClone()

            ' Nested JSON has no native Excel cell representation.  Preserve it losslessly
            ' as compact JSON text rather than passing a JArray/JObject into COM.
            Return New JValue(value.ToString(Newtonsoft.Json.Formatting.None))
        End Function

    End Class

End Namespace
