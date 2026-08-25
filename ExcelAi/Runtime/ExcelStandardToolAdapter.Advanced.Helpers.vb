Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent

Namespace OfficeRuntime

    Friend NotInheritable Partial Class ExcelStandardToolAdapter
        Private Shared Function BuildAdvancedMutationResult(toolId As String,
                                                            summary As String,
                                                            targetRefs As IEnumerable(Of String),
                                                            before As JObject,
                                                            after As JObject,
                                                            satisfied As Boolean,
                                                            Optional data As JObject = Nothing,
                                                            Optional verification As JObject = Nothing,
                                                            Optional artifacts As Object = Nothing) As ToolResult
            Dim refs As New JArray()
            For Each targetRef In If(targetRefs, Enumerable.Empty(Of String)())
                If Not String.IsNullOrWhiteSpace(targetRef) Then refs.Add(targetRef)
            Next
            Dim changed = Not JToken.DeepEquals(If(before, New JObject()), If(after, New JObject()))
            Dim observation As New JObject From {
                {"kind", "write"},
                {"writeExpected", True},
                {"adapter", "ExcelStandardToolAdapter"},
                {"sourceToolId", toolId},
                {"summary", summary},
                {"changed", changed},
                {"satisfied", satisfied},
                {"targetRefs", refs},
                {"before", If(before, New JObject())},
                {"after", If(after, New JObject())},
                {"diff", New JObject From {{"targetStateChanged", changed}}},
                {"warnings", New JArray()}
            }
            If verification IsNot Nothing Then observation("verification") = verification
            Dim resultData = If(data, New JObject())
            resultData("targetRefs") = refs.DeepClone()
            If satisfied Then Return ToolResult.Succeed(toolId, summary, resultData, observation, artifacts:=artifacts)
            observation("summary") = $"{summary}，但后置条件未满足"
            Return ToolResult.Failed(toolId,
                                     observation("summary").ToString(),
                                     data:=resultData,
                                     errorCode:=ExceptionClassifier.CodeVerifyFailed,
                                     userMessage:="Excel 操作已停止：实际结果未满足工具合同",
                                     recoverable:=False,
                                     observation:=observation,
                                     artifacts:=artifacts)
        End Function

        Private Shared Function OpenAdvancedRange(application As Object, rangeSpec As String) As AdvancedRangeHandle
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=False)
            Dim workbook As Object = Nothing
            Dim worksheets As Object = Nothing
            Dim worksheet As Object = Nothing
            Dim value As Object = Nothing
            Try
                workbook = application.ActiveWorkbook
                worksheets = workbook.Worksheets
                worksheet = worksheets.Item(descriptor.SheetName)
                value = worksheet.Range(descriptor.Address)
                Dim handle As New AdvancedRangeHandle(descriptor, workbook, worksheets, worksheet, value)
                workbook = Nothing
                worksheets = Nothing
                worksheet = Nothing
                value = Nothing
                Return handle
            Finally
                ReleaseCom(value)
                ReleaseCom(worksheet)
                ReleaseCom(worksheets)
                ReleaseCom(workbook)
            End Try
        End Function

        Private Shared Function ResolveAdvancedTargetSpec(source As AdvancedRangeHandle,
                                                          explicitTarget As String,
                                                          columnOffset As Integer) As String
            If Not String.IsNullOrWhiteSpace(explicitTarget) Then Return explicitTarget
            Dim topLeft As Object = Nothing
            Dim offsetCell As Object = Nothing
            Try
                topLeft = source.Value.Cells(1, 1)
                offsetCell = topLeft.Offset(0, columnOffset)
                Return QuoteSheet(source.Descriptor.SheetName) & "!" & CStr(offsetCell.Address(False, False))
            Finally
                ReleaseCom(offsetCell)
                ReleaseCom(topLeft)
            End Try
        End Function

        Private Shared Function CaptureAdvancedRangeState(value As Object) As JObject
            Return New JObject From {
                {"rowCount", CInt(value.Rows.Count)},
                {"columnCount", CInt(value.Columns.Count)},
                {"valueHash", ExcelOperationObserver.ComputeValueHash(value.Value2)},
                {"formulaHash", ExcelOperationObserver.ComputeValueHash(value.Formula)}
            }
        End Function

        Private Shared Function CaptureAdvancedConditionalFormats(value As Object) As JArray
            Dim result As New JArray()
            Dim rules As Object = Nothing
            Try
                rules = value.FormatConditions
                For index = 1 To CInt(rules.Count)
                    Dim rule As Object = Nothing
                    Dim interior As Object = Nothing
                    Dim appliesTo As Object = Nothing
                    Try
                        rule = rules.Item(index)
                        Dim captured As New JObject()
                        Try : captured("type") = CInt(rule.Type) : Catch : End Try
                        Try : captured("operator") = CInt(rule.Operator) : Catch : End Try
                        Try : captured("formula1") = CStr(If(rule.Formula1, "")) : Catch : End Try
                        Try : captured("formula2") = CStr(If(rule.Formula2, "")) : Catch : End Try
                        Try
                            interior = rule.Interior
                            captured("interiorColor") = CInt(interior.Color)
                        Catch
                        End Try
                        Try
                            appliesTo = rule.AppliesTo
                            captured("appliesTo") = CStr(appliesTo.Address(False, False))
                        Catch
                        End Try
                        result.Add(captured)
                    Finally
                        ReleaseCom(appliesTo)
                        ReleaseCom(interior)
                        ReleaseCom(rule)
                    End Try
                Next
            Finally
                ReleaseCom(rules)
            End Try
            Return result
        End Function

        Private Shared Sub ForEachCell(value As Object, action As Action(Of Object))
            Dim cells As Object = Nothing
            Try
                cells = value.Cells
                For index = 1 To CInt(cells.Count)
                    Dim cell As Object = Nothing
                    Try
                        cell = cells.Item(index)
                        action(cell)
                    Finally
                        ReleaseCom(cell)
                    End Try
                Next
            Finally
                ReleaseCom(cells)
            End Try
        End Sub

        Private Shared Function AllTextIsTrimmed(value As Object) As Boolean
            Dim satisfied = True
            ForEachCell(value,
                        Sub(cell)
                            Dim item = cell.Value2
                            If TypeOf item Is String AndAlso Not String.Equals(CStr(item), CStr(item).Trim(), StringComparison.Ordinal) Then satisfied = False
                        End Sub)
            Return satisfied
        End Function

        Private Shared Function CountEmptyCells(value As Object) As Integer
            Dim count As Integer = 0
            ForEachCell(value,
                        Sub(cell)
                            Dim item = cell.Value2
                            If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.ToString()) Then count += 1
                        End Sub)
            Return count
        End Function

        Private Shared Function RangeContainsText(value As Object, searchText As String) As Boolean
            Dim found = False
            ForEachCell(value,
                        Sub(cell)
                            Dim item = cell.Value2
                            If item IsNot Nothing AndAlso item.ToString().IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 Then found = True
                        End Sub)
            Return found
        End Function

        Private Shared Function BuildSummaryMatrix(source As Object) As Array
            Dim rows As New List(Of Object()) From {
                New Object() {"数据摘要", Nothing},
                New Object() {"行数", CInt(source.Rows.Count)},
                New Object() {"列数", CInt(source.Columns.Count)}
            }
            For column = 1 To CInt(source.Columns.Count)
                If ColumnHasNumericData(source, column) Then
                    Dim header = GetHeaderText(source, column, "列" & column.ToString(CultureInfo.InvariantCulture))
                    Dim sum As Double = 0
                    Dim count As Integer = 0
                    For row = 2 To CInt(source.Rows.Count)
                        Dim number As Double
                        If TryGetNumericCell(source, row, column, number) Then
                            sum += number
                            count += 1
                        End If
                    Next
                    rows.Add(New Object() {header & "合计", sum})
                    rows.Add(New Object() {header & "平均", If(count = 0, 0, sum / count)})
                End If
            Next
            Return MatrixFromRows(rows)
        End Function

        Private Shared Function BuildGroupByMatrix(source As Object, params As JObject) As Array
            Dim groupColumn = FindHeaderColumn(source, FirstText(params, "groupBy", "groupField"), 1)
            Dim valueColumn = FindHeaderColumn(source, FirstText(params, "valueField"), Math.Min(2, CInt(source.Columns.Count)))
            If groupColumn <= 0 OrElse valueColumn <= 0 Then Throw New FormatException("DataAnalysis groupby field was not found in the source headers")
            Dim aggregate = If(String.IsNullOrWhiteSpace(FirstText(params, "aggregate")), "sum", FirstText(params, "aggregate")).ToLowerInvariant()
            Dim order As New List(Of String)()
            Dim sums As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
            Dim counts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For row = 2 To CInt(source.Rows.Count)
                Dim key = GetCellText(source, row, groupColumn)
                If String.IsNullOrWhiteSpace(key) Then key = "(空)"
                If Not sums.ContainsKey(key) Then
                    sums(key) = 0
                    counts(key) = 0
                    order.Add(key)
                End If
                Dim number As Double
                If TryGetNumericCell(source, row, valueColumn, number) Then sums(key) += number
                counts(key) += 1
            Next
            Dim rows As New List(Of Object()) From {
                New Object() {"分组汇总", Nothing, Nothing},
                New Object() {GetHeaderText(source, groupColumn, "分组"), aggregate, "数量"}
            }
            For Each key In order
                Dim aggregateValue As Object
                Select Case aggregate
                    Case "count" : aggregateValue = counts(key)
                    Case "avg", "average" : aggregateValue = If(counts(key) = 0, 0, sums(key) / counts(key))
                    Case "sum" : aggregateValue = sums(key)
                    Case Else : Throw New FormatException($"Unsupported aggregate: {aggregate}")
                End Select
                rows.Add(New Object() {key, aggregateValue, counts(key)})
            Next
            Return MatrixFromRows(rows)
        End Function

        Private Shared Function BuildRankingMatrix(source As Object, params As JObject) As Array
            Dim labelColumn = FindHeaderColumn(source, FirstText(params, "labelField"), 1)
            Dim valueColumn = FindHeaderColumn(source, FirstText(params, "rankBy", "valueField"), Math.Min(2, CInt(source.Columns.Count)))
            If labelColumn <= 0 OrElse valueColumn <= 0 Then Throw New FormatException("DataAnalysis ranking field was not found in the source headers")
            Dim descending = If(params("descending")?.Value(Of Boolean)(), True)
            Dim topN = If(params("topN")?.Value(Of Integer)(), Math.Max(1, CInt(source.Rows.Count) - 1))
            Dim items As New List(Of Tuple(Of String, Double))()
            For row = 2 To CInt(source.Rows.Count)
                Dim number As Double
                If TryGetNumericCell(source, row, valueColumn, number) Then items.Add(Tuple.Create(GetCellText(source, row, labelColumn), number))
            Next
            items.Sort(Function(left, right) If(descending, right.Item2.CompareTo(left.Item2), left.Item2.CompareTo(right.Item2)))
            Dim rows As New List(Of Object()) From {
                New Object() {"排名分析", Nothing, Nothing},
                New Object() {"排名", GetHeaderText(source, labelColumn, "对象"), GetHeaderText(source, valueColumn, "数值")}
            }
            For index = 0 To Math.Min(Math.Max(1, topN), items.Count) - 1
                rows.Add(New Object() {index + 1, items(index).Item1, items(index).Item2})
            Next
            Return MatrixFromRows(rows)
        End Function

        Private Shared Function ReadRangeMatrix(value As Object) As Array
            Dim rows = CInt(value.Rows.Count)
            Dim columns = CInt(value.Columns.Count)
            Dim result = Array.CreateInstance(GetType(Object), New Integer() {rows, columns}, New Integer() {1, 1})
            For row = 1 To rows
                For column = 1 To columns
                    Dim cell As Object = Nothing
                    Try
                        cell = value.Cells(row, column)
                        result.SetValue(cell.Value2, row, column)
                    Finally
                        ReleaseCom(cell)
                    End Try
                Next
            Next
            Return result
        End Function

        Private Shared Function TransposeMatrix(source As Array) As Array
            Dim rows = source.GetLength(0)
            Dim columns = source.GetLength(1)
            Dim result = Array.CreateInstance(GetType(Object), New Integer() {columns, rows}, New Integer() {1, 1})
            For row = 1 To rows
                For column = 1 To columns
                    result.SetValue(source.GetValue(row, column), column, row)
                Next
            Next
            Return result
        End Function

        Private Shared Function SplitMatrix(source As Array, delimiter As String) As Array
            Dim rows = source.GetLength(0)
            Dim splitRows As New List(Of String())()
            Dim maxColumns = 1
            For row = 1 To rows
                Dim value = source.GetValue(row, 1)
                Dim parts = If(value Is Nothing, New String() {""}, value.ToString().Split(New String() {delimiter}, StringSplitOptions.None))
                splitRows.Add(parts)
                maxColumns = Math.Max(maxColumns, parts.Length)
            Next
            Dim result = Array.CreateInstance(GetType(Object), New Integer() {rows, maxColumns}, New Integer() {1, 1})
            For row = 1 To rows
                For column = 1 To splitRows(row - 1).Length
                    result.SetValue(splitRows(row - 1)(column - 1), row, column)
                Next
            Next
            Return result
        End Function

        Private Shared Function MergeMatrix(source As Array, delimiter As String) As Array
            Dim rows = source.GetLength(0)
            Dim columns = source.GetLength(1)
            Dim result = Array.CreateInstance(GetType(Object), New Integer() {rows, 1}, New Integer() {1, 1})
            For row = 1 To rows
                Dim values As New List(Of String)()
                For column = 1 To columns
                    Dim value = source.GetValue(row, column)
                    values.Add(If(value Is Nothing, "", value.ToString()))
                Next
                result.SetValue(String.Join(delimiter, values), row, 1)
            Next
            Return result
        End Function

        Private Shared Function MatrixFromRows(rows As IList(Of Object())) As Array
            Dim columnCount = If(rows.Count = 0, 1, rows.Max(Function(row) row.Length))
            Dim result = Array.CreateInstance(GetType(Object), New Integer() {Math.Max(1, rows.Count), columnCount}, New Integer() {1, 1})
            For row = 1 To rows.Count
                For column = 1 To rows(row - 1).Length
                    result.SetValue(rows(row - 1)(column - 1), row, column)
                Next
            Next
            Return result
        End Function

        Private Shared Function ResizeToMatrix(anchor As Object, matrix As Array) As Object
            Return anchor.Resize(matrix.GetLength(0), matrix.GetLength(1))
        End Function

        Private Shared Function RangeMatchesMatrix(value As Object, expected As Array) As Boolean
            If CInt(value.Rows.Count) <> expected.GetLength(0) OrElse CInt(value.Columns.Count) <> expected.GetLength(1) Then Return False
            For row = 1 To expected.GetLength(0)
                For column = 1 To expected.GetLength(1)
                    Dim cell As Object = Nothing
                    Try
                        cell = value.Cells(row, column)
                        If Not ScalarValuesEquivalent(cell.Value2, expected.GetValue(row, column)) Then Return False
                    Finally
                        ReleaseCom(cell)
                    End Try
                Next
            Next
            Return True
        End Function

        Private Shared Function ScalarValuesEquivalent(actual As Object, expected As Object) As Boolean
            If actual Is Nothing OrElse expected Is Nothing Then Return actual Is Nothing AndAlso expected Is Nothing
            Dim actualNumber As Decimal
            Dim expectedNumber As Decimal
            If Decimal.TryParse(actual.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, actualNumber) AndAlso
               Decimal.TryParse(expected.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, expectedNumber) Then Return actualNumber = expectedNumber
            Return String.Equals(actual.ToString(), expected.ToString(), StringComparison.Ordinal)
        End Function

        Private Shared Function FindHeaderColumn(source As Object, fieldName As String, fallback As Integer) As Integer
            If Not String.IsNullOrWhiteSpace(fieldName) Then
                For column = 1 To CInt(source.Columns.Count)
                    If String.Equals(GetHeaderText(source, column, ""), fieldName.Trim(), StringComparison.OrdinalIgnoreCase) Then Return column
                Next
                Return 0
            End If
            Return If(fallback >= 1 AndAlso fallback <= CInt(source.Columns.Count), fallback, 0)
        End Function

        Private Shared Function GetHeaderText(source As Object, column As Integer, fallback As String) As String
            Dim cell As Object = Nothing
            Try
                cell = source.Cells(1, column)
                Dim value = cell.Value2
                Return If(value Is Nothing OrElse String.IsNullOrWhiteSpace(value.ToString()), fallback, value.ToString().Trim())
            Finally
                ReleaseCom(cell)
            End Try
        End Function

        Private Shared Function GetCellText(source As Object, row As Integer, column As Integer) As String
            Dim cell As Object = Nothing
            Try
                cell = source.Cells(row, column)
                Return If(cell.Value2 Is Nothing, "", cell.Value2.ToString().Trim())
            Finally
                ReleaseCom(cell)
            End Try
        End Function

        Private Shared Function TryGetNumericCell(source As Object, row As Integer, column As Integer, ByRef number As Double) As Boolean
            Dim cell As Object = Nothing
            Try
                cell = source.Cells(row, column)
                Dim value = cell.Value2
                If value Is Nothing Then Return False
                If TypeOf value Is IConvertible Then
                    Try
                        number = Convert.ToDouble(value, CultureInfo.InvariantCulture)
                        Return True
                    Catch
                    End Try
                End If
                Return Double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, number) OrElse
                       Double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, number)
            Finally
                ReleaseCom(cell)
            End Try
        End Function

        Private Shared Function ColumnHasNumericData(source As Object, column As Integer) As Boolean
            For row = 2 To CInt(source.Rows.Count)
                Dim number As Double
                If TryGetNumericCell(source, row, column, number) Then Return True
            Next
            Return False
        End Function

        Private Shared Function ResolvePivotSourceBounds(source As Object,
                                                         ParamArray fieldGroups As JArray()) As Tuple(Of Integer, Integer)
            Dim indexes As New List(Of Integer)()
            For Each fields In If(fieldGroups, Array.Empty(Of JArray)())
                If fields Is Nothing Then Continue For
                For Each token In fields
                    Dim fieldName = PivotFieldName(token)
                    Dim column = FindHeaderColumn(source, fieldName, 0)
                    If column <= 0 Then Throw New FormatException($"Pivot field was not found in source headers: {fieldName}")
                    indexes.Add(column)
                Next
            Next
            If indexes.Count = 0 Then Throw New FormatException("Pivot table requires at least one source field")

            Dim firstColumn = indexes.Min()
            Dim lastColumn = indexes.Max()
            For column = firstColumn To lastColumn
                If String.IsNullOrWhiteSpace(GetHeaderText(source, column, "")) Then
                    Throw New FormatException(
                        $"Pivot source contains a blank header between required fields at source column {column}; choose a contiguous table range with non-empty headers")
                End If
            Next
            Return Tuple.Create(firstColumn, lastColumn)
        End Function

        Private Shared Sub ConfigurePivotFields(pivotTable As Object, fields As JArray, orientation As XlPivotFieldOrientation)
            If fields Is Nothing Then Return
            Dim position As Integer = 1
            For Each fieldToken In fields
                Dim field As PivotField = Nothing
                Try
                    field = DirectCast(DirectCast(pivotTable, PivotTable).PivotFields(PivotFieldName(fieldToken)), PivotField)
                    field.Orientation = orientation
                    field.Position = position
                    position += 1
                Finally
                    ReleaseCom(field)
                End Try
            Next
        End Sub

        Private Shared Sub ConfigurePivotValues(pivotTable As Object, fields As JArray)
            If fields Is Nothing Then Return
            For Each fieldToken In fields
                Dim field As PivotField = Nothing
                Dim dataField As PivotField = Nothing
                Try
                    Dim spec = ParsePivotValueField(fieldToken)
                    field = DirectCast(DirectCast(pivotTable, PivotTable).PivotFields(spec.Item1), PivotField)
                    dataField = DirectCast(DirectCast(pivotTable, PivotTable).AddDataField(field,
                                                        spec.Item3,
                                                        spec.Item2), PivotField)
                Finally
                    ReleaseCom(dataField)
                    ReleaseCom(field)
                End Try
            Next
        End Sub

        Private Shared Function PivotFieldName(token As JToken) As String
            If token Is Nothing Then Throw New FormatException("Pivot field cannot be empty")
            Dim obj = TryCast(token, JObject)
            Dim value = If(obj Is Nothing,
                           token.ToString(),
                           FirstText(obj, "field", "name", "fieldName"))
            If String.IsNullOrWhiteSpace(value) Then Throw New FormatException("Pivot field object requires field or name")
            Return value.Trim()
        End Function

        Private Shared Function ParsePivotValueField(token As JToken) As Tuple(Of String, XlConsolidationFunction, String)
            Dim fieldName = PivotFieldName(token)
            Dim obj = TryCast(token, JObject)
            Dim aggregate = If(obj Is Nothing, "sum", FirstText(obj, "aggregate", "function")).Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(aggregate) Then aggregate = "sum"
            Dim functionValue As XlConsolidationFunction
            Select Case aggregate
                Case "sum", "total"
                    functionValue = XlConsolidationFunction.xlSum
                Case "avg", "average", "mean"
                    functionValue = XlConsolidationFunction.xlAverage
                Case "count"
                    functionValue = XlConsolidationFunction.xlCount
                Case "max", "maximum"
                    functionValue = XlConsolidationFunction.xlMax
                Case "min", "minimum"
                    functionValue = XlConsolidationFunction.xlMin
                Case Else
                    Throw New FormatException($"Unsupported pivot aggregation: {aggregate}")
            End Select
            Dim caption = If(obj Is Nothing, "", FirstText(obj, "caption", "title"))
            If String.IsNullOrWhiteSpace(caption) Then caption = aggregate & " - " & fieldName
            Return Tuple.Create(fieldName, functionValue, caption)
        End Function

        Private Shared Function TryGetWorksheet(worksheets As Object, name As String, ByRef worksheet As Object) As Boolean
            worksheet = Nothing
            Try
                worksheet = worksheets.Item(name)
                Return worksheet IsNot Nothing
            Catch
                worksheet = Nothing
                Return False
            End Try
        End Function

        Private Shared Function UniqueSheetName(worksheets As Object, prefix As String) As String
            Dim baseName = prefix & "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            Dim candidate = baseName
            Dim index As Integer = 1
            Dim existing As Object = Nothing
            While TryGetWorksheet(worksheets, candidate, existing)
                ReleaseCom(existing)
                existing = Nothing
                candidate = baseName & "_" & index.ToString(CultureInfo.InvariantCulture)
                index += 1
            End While
            Return candidate
        End Function

        Private Shared Function GetChartCount(worksheet As Object) As Integer
            Dim charts As Object = Nothing
            Try
                charts = worksheet.ChartObjects()
                Return CInt(charts.Count)
            Finally
                ReleaseCom(charts)
            End Try
        End Function

        Private Shared Function BuildAdvancedRangeRef(sheetName As String, address As String) As String
            Return ExcelObjectResolver.BuildRangeRef(sheetName, address)
        End Function

        Private Shared Function TokenValue(token As JToken, fallback As Object) As Object
            Dim scalar = TryCast(token, JValue)
            If scalar Is Nothing OrElse scalar.Value Is Nothing Then Return fallback
            Return scalar.Value
        End Function

        Private NotInheritable Class AdvancedRangeHandle
            Implements IDisposable

            Public Sub New(descriptor As RangeDescriptor,
                           workbook As Object,
                           worksheets As Object,
                           worksheet As Object,
                           value As Object)
                Me.Descriptor = descriptor
                Me.Workbook = workbook
                Me.Worksheets = worksheets
                Me.Worksheet = worksheet
                Me.Value = value
            End Sub

            Public ReadOnly Property Descriptor As RangeDescriptor
            Public ReadOnly Property Workbook As Object
            Public ReadOnly Property Worksheets As Object
            Public ReadOnly Property Worksheet As Object
            Public ReadOnly Property Value As Object

            Public Sub Dispose() Implements IDisposable.Dispose
                ReleaseCom(Value)
                ReleaseCom(Worksheet)
                ReleaseCom(Worksheets)
                ReleaseCom(Workbook)
            End Sub
        End Class
    End Class

End Namespace
