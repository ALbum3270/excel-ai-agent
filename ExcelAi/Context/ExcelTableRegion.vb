Imports Excel = Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports System.Globalization

Namespace Context

    ''' <summary>
    ''' 可序列化的 Excel 表区域画像。COM 对象不得跨 ContextProvider 调用保存到此模型。
    ''' </summary>
    Public Class ExcelTableRegion
        Public Property Sheet As String = ""
        Public Property Address As String = ""
        Public Property HasHeader As Boolean
        Public Property HeaderRow As Integer
        Public Property Headers As New List(Of String)()
        Public Property ColumnTypes As New List(Of String)()
        Public Property RowCount As Integer
        Public Property ColCount As Integer
        Public Property Source As String = ""
        Public Property ListObjectName As String = ""
        Public Property Confidence As Double
        Public Property Warnings As New List(Of String)()

        Public Function ToJson() As JObject
            Return New JObject From {
                {"sheet", Sheet},
                {"address", Address},
                {"hasHeader", HasHeader},
                {"headerRow", HeaderRow},
                {"headers", JArray.FromObject(Headers)},
                {"columnTypes", JArray.FromObject(ColumnTypes)},
                {"rowCount", RowCount},
                {"colCount", ColCount},
                {"source", Source},
                {"listObjectName", ListObjectName},
                {"confidence", Math.Round(Confidence, 2)},
                {"warnings", JArray.FromObject(Warnings)}
            }
        End Function

        Public Function ToPromptSummary() As String
            Dim headerText = If(Headers.Count = 0, "未识别", String.Join(", ", Headers.Take(12)))
            Dim typeText = If(ColumnTypes.Count = 0, "未识别", String.Join(", ", ColumnTypes.Take(12)))
            Return $"表区域: {Sheet}!{Address}; source={Source}; dataRows={RowCount}; columns={ColCount}; hasHeader={HasHeader}; headers=[{headerText}]; columnTypes=[{typeText}]; confidence={Confidence:F2}"
        End Function
    End Class

    ''' <summary>
    ''' 从当前 Excel UI 状态探测主要 TableRegion。必须在宿主 UI 线程调用。
    ''' </summary>
    Public Class ExcelTableRegionDetector
        Private Const MaxSampleRows As Integer = 200
        Private Const MaxAnalyzedColumns As Integer = 64

        Private ReadOnly _app As Excel.Application

        Public Sub New(app As Excel.Application)
            If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
            _app = app
        End Sub

        Public Function Detect() As ExcelTableRegion
            Dim warnings As New List(Of String)()
            Dim sheet = TryCast(_app.ActiveSheet, Excel.Worksheet)
            If sheet Is Nothing Then
                Return New ExcelTableRegion With {
                    .Source = "none",
                    .Confidence = 0,
                    .Warnings = New List(Of String) From {"当前活动对象不是工作表"}
                }
            End If

            Dim candidate As Excel.Range = Nothing
            Dim rawSelection As Excel.Range = Nothing
            Dim selection As Excel.Range = Nothing
            Dim activeCell As Excel.Range = Nothing
            Dim listObject As Excel.ListObject = Nothing
            Dim listRange As Excel.Range = Nothing
            Dim usedRange As Excel.Range = Nothing
            Dim currentRegion As Excel.Range = Nothing
            Dim source = ""
            Dim listObjectName = ""

            Try
                rawSelection = TryCast(_app.Selection, Excel.Range)
                selection = rawSelection
                If selection IsNot Nothing Then
                    selection = SelectLargestArea(selection, warnings)
                End If

                If selection IsNot Nothing AndAlso SafeCellCount(selection) > 1 AndAlso HasContent(selection) Then
                    candidate = selection
                    source = "selection"
                End If

                activeCell = TryCast(_app.ActiveCell, Excel.Range)
                If candidate Is Nothing AndAlso activeCell IsNot Nothing Then
                    Try
                        listObject = TryCast(activeCell.ListObject, Excel.ListObject)
                        If listObject IsNot Nothing Then listRange = TryCast(listObject.Range, Excel.Range)
                        If listRange IsNot Nothing Then
                            candidate = listRange
                            source = "list_object"
                            listObjectName = If(listObject.Name, "")
                        End If
                    Catch
                    End Try
                End If

                usedRange = sheet.UsedRange
                If candidate Is Nothing AndAlso activeCell IsNot Nothing AndAlso usedRange IsNot Nothing AndAlso IsInside(activeCell, usedRange) Then
                    Try
                        currentRegion = activeCell.CurrentRegion
                        If currentRegion IsNot Nothing AndAlso SafeCellCount(currentRegion) > 1 AndAlso HasContent(currentRegion) Then
                            candidate = currentRegion
                            source = "current_region"
                        End If
                    Catch ex As Exception
                        warnings.Add("连续区域探测失败: " & ex.Message)
                    End Try
                End If

                If candidate Is Nothing AndAlso usedRange IsNot Nothing AndAlso HasContent(usedRange) Then
                    candidate = usedRange
                    source = "used_range"
                End If

                If candidate Is Nothing Then
                    Return New ExcelTableRegion With {
                        .Sheet = sheet.Name,
                        .Source = "none",
                        .Confidence = 0.1,
                        .Warnings = New List(Of String) From {"当前工作表没有可探测的数据区域"}
                    }
                End If

                Return Analyze(candidate, source, listObjectName, warnings)
            Catch ex As Exception
                warnings.Add("TableRegion 探测失败: " & ex.Message)
                Return New ExcelTableRegion With {
                    .Sheet = SafeSheetName(sheet),
                    .Source = If(String.IsNullOrWhiteSpace(source), "none", source),
                    .Confidence = 0,
                    .Warnings = warnings
                }
            Finally
                ComObjectHelper.ReleaseComObject(currentRegion)
                ComObjectHelper.ReleaseComObject(usedRange)
                ComObjectHelper.ReleaseComObject(listRange)
                ComObjectHelper.ReleaseComObject(listObject)
                ComObjectHelper.ReleaseComObject(activeCell)
                If Not Object.ReferenceEquals(selection, rawSelection) Then
                    ComObjectHelper.ReleaseComObject(selection)
                End If
                ComObjectHelper.ReleaseComObject(rawSelection)
                ComObjectHelper.ReleaseComObject(sheet)
            End Try
        End Function

        Private Function Analyze(range As Excel.Range,
                                 source As String,
                                 listObjectName As String,
                                 warnings As List(Of String)) As ExcelTableRegion
            Dim rangeRows As Excel.Range = Nothing
            Dim rangeColumns As Excel.Range = Nothing
            Dim sampleRange As Excel.Range = Nothing
            Dim rangeSheet As Excel.Worksheet = Nothing
            Try
                rangeRows = range.Rows
                rangeColumns = range.Columns
                Dim totalRows = Math.Max(1, CInt(rangeRows.Count))
                Dim totalColumns = Math.Max(1, CInt(rangeColumns.Count))
                Dim analyzedColumns = Math.Min(totalColumns, MaxAnalyzedColumns)
                Dim sampleRows = Math.Min(totalRows, MaxSampleRows + 1)
                sampleRange = range.Resize(sampleRows, analyzedColumns)
                rangeSheet = TryCast(range.Worksheet, Excel.Worksheet)
                Dim values As Object = sampleRange.Value2
                Dim formulas As Object = sampleRange.Formula

                Dim region As New ExcelTableRegion With {
                    .Sheet = SafeSheetName(rangeSheet),
                    .Address = SafeAddress(range),
                    .ColCount = totalColumns,
                    .Source = source,
                    .ListObjectName = listObjectName,
                    .Warnings = warnings
                }

                If totalColumns > analyzedColumns Then
                    region.Warnings.Add($"列数超过 {MaxAnalyzedColumns}，仅分析前 {MaxAnalyzedColumns} 列类型")
                End If
                If totalRows > MaxSampleRows + 1 Then
                    region.Warnings.Add($"数据行较多，列类型仅抽样前 {MaxSampleRows} 行")
                End If

                region.HasHeader = InferHasHeader(values, sampleRows, analyzedColumns)
                region.HeaderRow = If(region.HasHeader, 1, 0)
                region.RowCount = Math.Max(0, totalRows - If(region.HasHeader, 1, 0))

                For columnIndex = 1 To analyzedColumns
                    Dim header = ""
                    If region.HasHeader Then header = SafeText(MatrixValue(values, 1, columnIndex)).Trim()
                    If String.IsNullOrWhiteSpace(header) Then header = ColumnLetter(range.Column + columnIndex - 1)
                    region.Headers.Add(header)
                    region.ColumnTypes.Add(InferColumnType(sampleRange, values, formulas, columnIndex, sampleRows, region.HasHeader))
                Next

                region.Confidence = BaseConfidence(source)
                If region.HasHeader Then region.Confidence += 0.04
                region.Confidence -= Math.Min(0.15, region.Warnings.Count * 0.03)
                region.Confidence = Math.Max(0, Math.Min(0.99, region.Confidence))
                Return region
            Finally
                ComObjectHelper.ReleaseComObject(rangeSheet)
                ComObjectHelper.ReleaseComObject(sampleRange)
                ComObjectHelper.ReleaseComObject(rangeColumns)
                ComObjectHelper.ReleaseComObject(rangeRows)
            End Try
        End Function

        Private Shared Function SelectLargestArea(selection As Excel.Range, warnings As List(Of String)) As Excel.Range
            Dim areas As Excel.Areas = Nothing
            Dim largest As Excel.Range = Nothing
            Try
                areas = selection.Areas
                If areas Is Nothing OrElse areas.Count <= 1 Then Return selection
                Dim largestCount As Long = -1
                For areaIndex = 1 To areas.Count
                    Dim area As Excel.Range = Nothing
                    area = TryCast(areas.Item(areaIndex), Excel.Range)
                    Dim count = SafeCellCount(area)
                    If count > largestCount Then
                        ComObjectHelper.ReleaseComObject(largest)
                        largest = area
                        largestCount = count
                    Else
                        ComObjectHelper.ReleaseComObject(area)
                    End If
                Next
                warnings.Add($"检测到多区域选区，使用最大矩形区域（共 {areas.Count} 个区域）")
                Return If(largest, selection)
            Catch
                ComObjectHelper.ReleaseComObject(largest)
                Return selection
            Finally
                ComObjectHelper.ReleaseComObject(areas)
            End Try
        End Function

        Private Function HasContent(range As Excel.Range) As Boolean
            Dim worksheetFunction As Excel.WorksheetFunction = Nothing
            Try
                worksheetFunction = _app.WorksheetFunction
                Return CDbl(worksheetFunction.CountA(range)) > 0
            Catch
                Try
                    Return range.Value2 IsNot Nothing
                Catch
                    Return False
                End Try
            Finally
                ComObjectHelper.ReleaseComObject(worksheetFunction)
            End Try
        End Function

        Private Function IsInside(cell As Excel.Range, container As Excel.Range) As Boolean
            Dim intersection As Excel.Range = Nothing
            Try
                intersection = _app.Intersect(cell, container)
                Return intersection IsNot Nothing
            Catch
                Return False
            Finally
                ComObjectHelper.ReleaseComObject(intersection)
            End Try
        End Function

        Private Shared Function InferHasHeader(values As Object, rowCount As Integer, columnCount As Integer) As Boolean
            If rowCount < 2 OrElse columnCount < 1 Then Return False

            Dim nonEmpty As Integer = 0
            Dim textValues As Integer = 0
            Dim unique As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For columnIndex = 1 To columnCount
                Dim value = MatrixValue(values, 1, columnIndex)
                Dim text = SafeText(value).Trim()
                If String.IsNullOrWhiteSpace(text) Then Continue For
                nonEmpty += 1
                unique.Add(text)
                If Not IsNumericValue(value) AndAlso Not IsDateValue(value) Then textValues += 1
            Next

            If nonEmpty = 0 Then Return False
            Dim coverage = nonEmpty / CDbl(columnCount)
            Dim textRatio = textValues / CDbl(nonEmpty)
            Dim uniqueRatio = unique.Count / CDbl(nonEmpty)
            Return coverage >= 0.6 AndAlso textRatio >= 0.6 AndAlso uniqueRatio >= 0.8
        End Function

        Private Shared Function InferColumnType(sampleRange As Excel.Range,
                                                values As Object,
                                                formulas As Object,
                                                columnIndex As Integer,
                                                rowCount As Integer,
                                                hasHeader As Boolean) As String
            Dim startRow = If(hasHeader, 2, 1)
            Dim populated As Integer = 0
            Dim formulaCount As Integer = 0
            Dim numericCount As Integer = 0
            Dim dateCount As Integer = 0

            For rowIndex = startRow To rowCount
                Dim value = MatrixValue(values, rowIndex, columnIndex)
                If value Is Nothing OrElse String.IsNullOrWhiteSpace(SafeText(value)) Then Continue For
                populated += 1

                Dim formula = SafeText(MatrixValue(formulas, rowIndex, columnIndex))
                If formula.StartsWith("=", StringComparison.Ordinal) Then formulaCount += 1
                If IsDateValue(value) Then
                    dateCount += 1
                ElseIf IsNumericValue(value) Then
                    numericCount += 1
                End If
            Next

            If populated = 0 Then Return "empty"
            If formulaCount / CDbl(populated) >= 0.5 Then Return "formula"
            If dateCount / CDbl(populated) >= 0.6 Then Return "date"
            If numericCount / CDbl(populated) >= 0.6 Then
                Dim sampleCell As Excel.Range = Nothing
                Try
                    sampleCell = TryCast(sampleRange.Cells(startRow, columnIndex), Excel.Range)
                    Dim numberFormat = SafeText(sampleCell?.NumberFormat).ToLowerInvariant()
                    If LooksLikeDateNumberFormat(numberFormat) Then Return "date"
                Catch
                Finally
                    ComObjectHelper.ReleaseComObject(sampleCell)
                End Try
                Return "number"
            End If
            Return "text"
        End Function

        Private Shared Function MatrixValue(matrix As Object, rowIndex As Integer, columnIndex As Integer) As Object
            If matrix Is Nothing Then Return Nothing
            If TypeOf matrix Is Object(,) Then
                Dim values = DirectCast(matrix, Object(,))
                Dim row = values.GetLowerBound(0) + rowIndex - 1
                Dim column = values.GetLowerBound(1) + columnIndex - 1
                If row > values.GetUpperBound(0) OrElse column > values.GetUpperBound(1) Then Return Nothing
                Return values(row, column)
            End If
            If rowIndex = 1 AndAlso columnIndex = 1 Then Return matrix
            Return Nothing
        End Function

        Private Shared Function IsNumericValue(value As Object) As Boolean
            If value Is Nothing OrElse TypeOf value Is Boolean Then Return False
            If TypeOf value Is Byte OrElse TypeOf value Is Short OrElse TypeOf value Is Integer OrElse
               TypeOf value Is Long OrElse TypeOf value Is Single OrElse TypeOf value Is Double OrElse
               TypeOf value Is Decimal Then Return True
            Dim number As Double
            Return Double.TryParse(SafeText(value), NumberStyles.Any, CultureInfo.CurrentCulture, number) OrElse
                   Double.TryParse(SafeText(value), NumberStyles.Any, CultureInfo.InvariantCulture, number)
        End Function

        Private Shared Function IsDateValue(value As Object) As Boolean
            If TypeOf value Is DateTime Then Return True
            If value Is Nothing OrElse IsNumericValue(value) Then Return False
            Dim parsed As DateTime
            Return DateTime.TryParse(SafeText(value), CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) OrElse
                   DateTime.TryParse(SafeText(value), CultureInfo.InvariantCulture, DateTimeStyles.None, parsed)
        End Function

        Private Shared Function LooksLikeDateNumberFormat(numberFormat As String) As Boolean
            If String.IsNullOrWhiteSpace(numberFormat) Then Return False
            Dim normalized = numberFormat.Replace("\\", "").Replace("""", "")
            Return normalized.Contains("yy") OrElse normalized.Contains("dd") OrElse
                   normalized.Contains("yyyy") OrElse normalized.Contains("m/d") OrElse
                   normalized.Contains("d/m")
        End Function

        Private Shared Function BaseConfidence(source As String) As Double
            Select Case source
                Case "list_object" : Return 0.94
                Case "selection" : Return 0.88
                Case "current_region" : Return 0.8
                Case "used_range" : Return 0.68
                Case Else : Return 0.4
            End Select
        End Function

        Private Shared Function SafeCellCount(range As Excel.Range) As Long
            Try
                Return CLng(range.CountLarge)
            Catch
                Try
                    Return CLng(range.Rows.Count) * CLng(range.Columns.Count)
                Catch
                    Return 0
                End Try
            End Try
        End Function

        Private Shared Function SafeAddress(range As Excel.Range) As String
            Try
                Return CStr(range.Address(False, False))
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeSheetName(sheet As Excel.Worksheet) As String
            Try
                Return If(sheet?.Name, "")
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeText(value As Object) As String
            If value Is Nothing Then Return ""
            Try
                Return Convert.ToString(value, CultureInfo.CurrentCulture)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function ColumnLetter(columnNumber As Integer) As String
            Dim dividend = Math.Max(1, columnNumber)
            Dim name = ""
            While dividend > 0
                Dim modulo = (dividend - 1) Mod 26
                name = ChrW(65 + modulo) & name
                dividend = (dividend - modulo) \ 26
            End While
            Return name
        End Function
    End Class

End Namespace
