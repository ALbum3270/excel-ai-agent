Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent

Namespace OfficeRuntime

    ''' <summary>
    ''' Safe, injectable Excel read primitive used by the Agent and live contract tests.
    ''' It returns complete structured values for the requested range and never mutates Excel.
    ''' </summary>
    Friend NotInheritable Class ExcelReadRangeAdapter
        Private Sub New()
        End Sub

        Public Shared Function Execute(application As Object, params As JObject) As ToolResult
            Dim workbook As Object = Nothing
            Dim worksheet As Object = Nothing
            Dim worksheets As Object = Nothing
            Dim target As Object = Nothing
            Dim targetRows As Object = Nothing
            Dim targetColumns As Object = Nothing
            Try
                Dim rangeSpec = If(params?("range")?.ToString(), "").Trim()
                If String.IsNullOrWhiteSpace(rangeSpec) Then
                    Return ToolResult.Failed("ReadRange",
                                             "ReadRange 缺少 range 参数",
                                             errorCode:=ExceptionClassifier.CodeArgument,
                                             userMessage:="请指定要读取的 Excel 范围",
                                             recoverable:=True)
                End If
                If application Is Nothing Then
                    Return ToolResult.Failed("ReadRange",
                                             "Excel application 不可用",
                                             errorCode:=ExceptionClassifier.CodeDocMissing,
                                             userMessage:="请先打开 Excel",
                                             recoverable:=False)
                End If

                workbook = application.ActiveWorkbook
                worksheet = application.ActiveSheet
                If workbook Is Nothing OrElse worksheet Is Nothing Then
                    Return ToolResult.Failed("ReadRange",
                                             "当前没有活动工作簿或工作表",
                                             errorCode:=ExceptionClassifier.CodeDocMissing,
                                             userMessage:="请先打开一个 Excel 工作簿",
                                             recoverable:=False)
                End If

                Dim address = rangeSpec
                Dim bangIndex = rangeSpec.LastIndexOf("!"c)
                If bangIndex > 0 Then
                    Dim sheetName = rangeSpec.Substring(0, bangIndex).Trim()
                    If sheetName.Length >= 2 AndAlso sheetName.StartsWith("'", StringComparison.Ordinal) AndAlso sheetName.EndsWith("'", StringComparison.Ordinal) Then
                        sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'")
                    End If
                    address = rangeSpec.Substring(bangIndex + 1).Trim()
                    ComObjectHelper.ReleaseComObject(worksheet)
                    worksheet = Nothing
                    worksheets = workbook.Worksheets
                    worksheet = worksheets(sheetName)
                End If

                target = worksheet.Range(address)
                targetRows = target.Rows
                targetColumns = target.Columns
                Dim rowCount = CInt(targetRows.Count)
                Dim columnCount = CInt(targetColumns.Count)
                Dim cellCount = CLng(rowCount) * CLng(columnCount)
                Dim maxCells = 5000
                If params?("maxCells") IsNot Nothing Then
                    Dim parsedMaxCells As Integer
                    If Integer.TryParse(params("maxCells").ToString(), parsedMaxCells) Then maxCells = parsedMaxCells
                End If
                maxCells = Math.Max(1, Math.Min(20000, maxCells))
                If cellCount > maxCells Then
                    Return ToolResult.Failed("ReadRange",
                                             $"范围包含 {cellCount} 个单元格，超过当前 {maxCells} 个限制",
                                             data:=New With {.address = rangeSpec, .cellCount = cellCount, .maxCells = maxCells},
                                             errorCode:=ExceptionClassifier.CodeArgument,
                                             userMessage:="读取范围过大，请缩小范围或分块读取",
                                             recoverable:=True)
                End If

                Dim includeFormulas = params?("includeFormulas")?.Value(Of Boolean)() = True
                Dim values = ExcelMatrixToJson(target.Value2, rowCount, columnCount)
                Dim formulas As JArray = Nothing
                If includeFormulas Then formulas = ExcelMatrixToJson(target.Formula, rowCount, columnCount)
                Dim actualAddress = CStr(target.Address(False, False))
                Dim sheetNameActual = CStr(worksheet.Name)

                Dim data As New JObject From {
                    {"workbook", CStr(workbook.Name)},
                    {"sheet", sheetNameActual},
                    {"address", actualAddress},
                    {"rowCount", rowCount},
                    {"columnCount", columnCount},
                    {"values", values}
                }
                If includeFormulas Then data("formulas") = formulas

                Dim observation As New JObject From {
                    {"kind", "read"},
                    {"summary", $"已读取 {sheetNameActual}!{actualAddress}（{rowCount} 行 x {columnCount} 列）"},
                    {"changed", False},
                    {"targetRefs", New JArray($"Excel:{sheetNameActual}!{actualAddress}")},
                    {"warnings", New JArray()}
                }
                Return ToolResult.Succeed("ReadRange", observation("summary").ToString(), data:=data, observation:=observation)
            Catch ex As Exception
                Return ToolResult.FromException("ReadRange", ex)
            Finally
                ComObjectHelper.ReleaseComObject(targetColumns)
                ComObjectHelper.ReleaseComObject(targetRows)
                ComObjectHelper.ReleaseComObject(target)
                ComObjectHelper.ReleaseComObject(worksheet)
                ComObjectHelper.ReleaseComObject(worksheets)
                ComObjectHelper.ReleaseComObject(workbook)
            End Try
        End Function

        Private Shared Function ExcelMatrixToJson(matrix As Object,
                                                  rowCount As Integer,
                                                  columnCount As Integer) As JArray
            Dim rows As New JArray()
            If rowCount <= 0 OrElse columnCount <= 0 Then Return rows

            For rowIndex = 1 To rowCount
                Dim row As New JArray()
                For columnIndex = 1 To columnCount
                    Dim value As Object = Nothing
                    If TypeOf matrix Is Object(,) Then
                        Dim values = DirectCast(matrix, Object(,))
                        value = values(values.GetLowerBound(0) + rowIndex - 1,
                                       values.GetLowerBound(1) + columnIndex - 1)
                    ElseIf rowIndex = 1 AndAlso columnIndex = 1 Then
                        value = matrix
                    End If

                    If value Is Nothing Then
                        row.Add(JValue.CreateNull())
                    Else
                        Try
                            row.Add(JToken.FromObject(value))
                        Catch
                            row.Add(value.ToString())
                        End Try
                    End If
                Next
                rows.Add(row)
            Next
            Return rows
        End Function
    End Class
End Namespace
