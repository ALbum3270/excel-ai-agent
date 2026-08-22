Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Imports ShareRibbon.Agent
Imports ShareRibbon.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend NotInheritable Partial Class ExcelStandardToolAdapter
        Private Shared Function ExecuteWriteData(application As Object, params As JObject) As ToolResult
            Const toolId As String = "WriteData"
            Dim rangeSpec = FirstText(params, "targetRange", "startCell", "range")
            Dim targetSheet = FirstText(params, "targetSheet")
            If Not String.IsNullOrWhiteSpace(targetSheet) AndAlso rangeSpec.IndexOf("!"c) < 0 Then rangeSpec = QuoteSheet(targetSheet) & "!" & rangeSpec
            Dim data = If(params("data"), params("targetData"))
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse data Is Nothing Then Return Invalid(toolId, "WriteData requires targetRange and data")

            Dim rowCount As Integer = 1
            Dim columnCount As Integer = 1
            ValidateDataShape(data, rowCount, columnCount)
            Dim descriptor = ResolveRange(application, rangeSpec, rowCount, columnCount, captureValues:=False)
            Dim batch = NewBatch()
            AddOperation(batch,
                         "write-values",
                         descriptor.RangeRef,
                         "set",
                         "Range",
                         "Value2",
                         "property",
                         New JObject From {{"value", data.DeepClone()}},
                         New JObject From {
                             {"ValueHash", ExcelOperationObserver.ComputeValueHash(data)},
                             {"RowCount", rowCount},
                             {"ColumnCount", columnCount}
                         })
            Return ExecuteBatch(application, toolId, batch, $"已验证写入 {descriptor.SheetName}!{descriptor.Address}")
        End Function

        Private Shared Function ExecuteApplyFormula(application As Object, params As JObject) As ToolResult
            Const toolId As String = "ApplyFormula"
            Dim rangeSpec = FirstText(params, "targetRange", "range", "target")
            Dim formula = FirstText(params, "formula")
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse String.IsNullOrWhiteSpace(formula) Then Return Invalid(toolId, "ApplyFormula requires targetRange and formula")
            If Not formula.StartsWith("=", StringComparison.Ordinal) Then formula = "=" & formula
            Dim fillDown = If(params("fillDown")?.Value(Of Boolean?)(), True)
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=False)
            Dim batch = NewBatch()

            If fillDown AndAlso descriptor.CellCount > 1 Then
                AddOperation(batch,
                             "set-first-formula",
                             descriptor.TopLeftRef,
                             "set",
                             "Range",
                             "Formula",
                             "property",
                             New JObject From {{"value", formula}},
                             New JObject From {{"TopLeftFormula", formula}})
                AddOperation(batch,
                             "fill-formula",
                             descriptor.TopLeftRef,
                             "invoke",
                             "Range",
                             "AutoFill",
                             "method",
                             New JObject From {
                                 {"Destination", New JObject From {{"ref", descriptor.RangeRef}}},
                                 {"Type", CInt(XlAutoFillType.xlFillDefault)}
                             },
                             New JObject())
                batch.SuccessCriteria.Add(New OperationCriterion With {
                    .Id = "filled-formula-count",
                    .TargetRef = descriptor.RangeRef,
                    .PropertyName = "NonEmptyFormulaCount",
                    .Operator = "equals",
                    .ExpectedValue = New JValue(descriptor.CellCount),
                    .Required = True
                })
            Else
                AddOperation(batch,
                             "set-formula",
                             descriptor.RangeRef,
                             "set",
                             "Range",
                             "Formula",
                             "property",
                             New JObject From {{"value", formula}},
                             New JObject From {
                                 {"TopLeftFormula", formula},
                                 {"NonEmptyFormulaCount", descriptor.CellCount}
                             })
            End If
            Return ExecuteBatch(application, toolId, batch, $"已验证公式写入 {descriptor.SheetName}!{descriptor.Address}")
        End Function

        Private Shared Function ExecuteSortData(application As Object, params As JObject) As ToolResult
            Const toolId As String = "SortData"
            Dim rangeSpec = FirstText(params, "range")
            Dim sortColumn = If(params("sortColumn")?.Value(Of Integer?)(), 0)
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse sortColumn < 1 Then Return Invalid(toolId, "SortData requires range and a positive sortColumn")
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=True)
            If sortColumn > descriptor.ColumnCount Then Return Invalid(toolId, "sortColumn is outside the target range")
            Dim descending = String.Equals(FirstText(params, "order"), "desc", StringComparison.OrdinalIgnoreCase)
            Dim hasHeader = If(params("hasHeader")?.Value(Of Boolean?)(), True)
            Dim keyRef = ResolveRelativeColumnRef(application, descriptor, sortColumn)
            Dim batch = NewBatch()
            AddOperation(batch,
                         "sort-range",
                         descriptor.RangeRef,
                         "invoke",
                         "Range",
                         "Sort",
                         "method",
                         New JObject From {
                             {"Key1", New JObject From {{"ref", keyRef}}},
                             {"Order1", If(descending, CInt(XlSortOrder.xlDescending), CInt(XlSortOrder.xlAscending))},
                             {"Header", If(hasHeader, CInt(XlYesNoGuess.xlYes), CInt(XlYesNoGuess.xlNo))},
                             {"Orientation", CInt(XlSortOrientation.xlSortColumns)}
                         },
                         New JObject From {{"RowSetHash", ExcelOperationObserver.ComputeRowSetHash(descriptor.Values)}})
            Dim result = ExecuteBatch(application, toolId, batch, $"已验证排序 {descriptor.SheetName}!{descriptor.Address}")
            If result.Success AndAlso Not IsRangeSorted(application, descriptor, sortColumn, hasHeader, descending) Then
                Return SemanticFailure(toolId, result, "Excel returned from Sort, but the requested key order was not observed")
            End If
            Return result
        End Function

        Private Shared Function ExecuteFilterData(application As Object, params As JObject) As ToolResult
            Const toolId As String = "FilterData"
            Dim clearFilter = If(params("clearFilter")?.Value(Of Boolean?)(), False)
            Dim rangeSpec = FirstText(params, "range")
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=False)
            Dim worksheetRef = BuildWorksheetRef(descriptor.SheetName)
            Dim batch = NewBatch()
            If clearFilter Then
                AddOperation(batch,
                             "clear-filter",
                             worksheetRef,
                             "set",
                             "Worksheet",
                             "AutoFilterMode",
                             "property",
                             New JObject From {{"value", False}},
                             New JObject From {{"AutoFilterMode", False}})
            Else
                Dim column = If(params("column")?.Value(Of Integer?)(), 0)
                If column < 1 OrElse column > descriptor.ColumnCount Then Return Invalid(toolId, "FilterData column is outside the target range")
                Dim arguments As New JObject From {{"Field", column}}
                Dim criteria = params("criteria")
                If criteria IsNot Nothing AndAlso criteria.Type <> JTokenType.Null Then arguments("Criteria1") = criteria.DeepClone()
                AddOperation(batch,
                             "apply-filter",
                             descriptor.RangeRef,
                             "invoke",
                             "Range",
                             "AutoFilter",
                             "method",
                             arguments,
                             New JObject())
                batch.SuccessCriteria.Add(New OperationCriterion With {
                    .Id = "filter-state",
                    .TargetRef = worksheetRef,
                    .PropertyName = If(criteria Is Nothing OrElse String.IsNullOrWhiteSpace(criteria.ToString()), "AutoFilterMode", "FilterMode"),
                    .Operator = "equals",
                    .ExpectedValue = New JValue(True),
                    .Required = True
                })
            End If
            Return ExecuteBatch(application, toolId, batch, If(clearFilter, "已验证清除筛选", "已验证筛选状态"))
        End Function

        Private Shared Function ExecuteRemoveDuplicates(application As Object, params As JObject) As ToolResult
            Const toolId As String = "RemoveDuplicates"
            Dim rangeSpec = FirstText(params, "range")
            If String.IsNullOrWhiteSpace(rangeSpec) Then Return Invalid(toolId, "RemoveDuplicates requires range")
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=True)
            Dim columns = TryCast(params("columns"), JArray)
            If columns Is Nothing OrElse columns.Count = 0 Then
                columns = New JArray(Enumerable.Range(1, descriptor.ColumnCount))
            End If
            Dim hasHeader = If(params("hasHeader")?.Value(Of Boolean?)(), True)
            Dim batch = NewBatch()
            AddOperation(batch,
                         "remove-duplicates",
                         descriptor.RangeRef,
                         "invoke",
                         "Range",
                         "RemoveDuplicates",
                         "method",
                         New JObject From {
                             {"Columns", columns.DeepClone()},
                             {"Header", If(hasHeader, CInt(XlYesNoGuess.xlYes), CInt(XlYesNoGuess.xlNo))}
                         },
                         New JObject From {{"exists", True}})
            Dim result = ExecuteBatch(application, toolId, batch, $"已验证去重执行 {descriptor.SheetName}!{descriptor.Address}")
            If result.Success AndAlso Not VerifyNoDuplicates(application, descriptor, columns, hasHeader) Then
                Return SemanticFailure(toolId, result, "Duplicate rows remain in the requested key columns")
            End If
            Return result
        End Function

        Private Shared Function ExecuteMergeCells(application As Object, params As JObject) As ToolResult
            Const toolId As String = "MergeCells"
            Dim rangeSpec = FirstText(params, "range")
            If String.IsNullOrWhiteSpace(rangeSpec) Then Return Invalid(toolId, "MergeCells requires range")
            Dim unmerge = If(params("unmerge")?.Value(Of Boolean?)(), False)
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=False)
            Dim batch = NewBatch()
            AddOperation(batch,
                         If(unmerge, "unmerge-cells", "merge-cells"),
                         descriptor.RangeRef,
                         "invoke",
                         "Range",
                         If(unmerge, "UnMerge", "Merge"),
                         "method",
                         New JObject(),
                         New JObject From {{"MergeCells", Not unmerge}})
            Return ExecuteBatch(application, toolId, batch, If(unmerge, "已验证取消合并", "已验证合并单元格"))
        End Function

        Private Shared Function ExecuteAutoFit(application As Object, params As JObject) As ToolResult
            Const toolId As String = "AutoFit"
            Dim rangeSpec = FirstText(params, "range")
            Dim fitType = FirstText(params, "type").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse Not {"columns", "rows", "both"}.Contains(fitType) Then Return Invalid(toolId, "AutoFit requires range and type columns/rows/both")
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=False)

            Dim firstBatch = NewBatch()
            AddAutoFitOperations(firstBatch, descriptor.RangeRef, fitType, Nothing)
            Dim firstResult = ExecuteBatch(application, toolId, firstBatch, "Excel AutoFit first pass completed")
            If Not firstResult.Success Then Return firstResult

            Dim fixedPoint = CaptureAutoFitDimensions(application, descriptor.RangeRef, fitType)
            Dim verificationBatch = NewBatch()
            AddAutoFitOperations(verificationBatch, descriptor.RangeRef, fitType, fixedPoint)
            Return ExecuteBatch(application, toolId, verificationBatch, "已验证 AutoFit 达到稳定尺寸")
        End Function

        Private Shared Function ExecuteFindReplace(application As Object, params As JObject) As ToolResult
            Const toolId As String = "FindReplace"
            Dim rangeSpec = FirstText(params, "range")
            Dim findText = FirstText(params, "find")
            Dim replacement = FirstText(params, "replace")
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse String.IsNullOrEmpty(findText) Then Return Invalid(toolId, "FindReplace requires range and find")
            Dim matchCase = If(params("matchCase")?.Value(Of Boolean?)(), False)
            Dim whole = If(params("matchEntireCell")?.Value(Of Boolean?)(), False)
            Dim descriptor = ResolveRange(application, rangeSpec, captureValues:=True)
            Dim expectedFormulas = ReplaceValues(descriptor.Formulas, findText, replacement, matchCase, whole)
            Dim batch = NewBatch()
            AddOperation(batch,
                         "find-replace",
                         descriptor.RangeRef,
                         "invoke",
                         "Range",
                         "Replace",
                         "method",
                         New JObject From {
                             {"What", findText},
                             {"Replacement", replacement},
                             {"LookAt", If(whole, CInt(XlLookAt.xlWhole), CInt(XlLookAt.xlPart))},
                             {"MatchCase", matchCase}
                         },
                         New JObject From {{"FormulaHash", ExcelOperationObserver.ComputeValueHash(expectedFormulas)}})
            Return ExecuteBatch(application, toolId, batch, $"已验证查找替换 {descriptor.SheetName}!{descriptor.Address}")
        End Function

        Private Shared Sub AddAutoFitOperations(batch As OfficeOperationBatch,
                                                rangeRef As String,
                                                fitType As String,
                                                fixedPoint As JObject)
            If fitType = "columns" OrElse fitType = "both" Then
                Dim expected = New JObject From {{"exists", True}}
                If fixedPoint IsNot Nothing Then expected = New JObject From {{"ColumnWidth", fixedPoint("ColumnWidth").DeepClone()}}
                AddOperation(batch, "autofit-columns", rangeRef & "/columns", "invoke", "Range", "AutoFit", "method", New JObject(), expected)
            End If
            If fitType = "rows" OrElse fitType = "both" Then
                Dim expected = New JObject From {{"exists", True}}
                If fixedPoint IsNot Nothing Then expected = New JObject From {{"RowHeight", fixedPoint("RowHeight").DeepClone()}}
                AddOperation(batch, "autofit-rows", rangeRef & "/rows", "invoke", "Range", "AutoFit", "method", New JObject(), expected)
            End If
        End Sub

        Private Shared Function CaptureAutoFitDimensions(application As Object, rangeRef As String, fitType As String) As JObject
            Dim result As New JObject()
            If fitType = "columns" OrElse fitType = "both" Then
                Using resolved = ExcelObjectResolver.Resolve(application, rangeRef & "/columns")
                    result("ColumnWidth") = JToken.FromObject(resolved.Value.ColumnWidth)
                End Using
            End If
            If fitType = "rows" OrElse fitType = "both" Then
                Using resolved = ExcelObjectResolver.Resolve(application, rangeRef & "/rows")
                    result("RowHeight") = JToken.FromObject(resolved.Value.RowHeight)
                End Using
            End If
            Return result
        End Function

        Private Shared Function ResolveRelativeColumnRef(application As Object,
                                                         descriptor As RangeDescriptor,
                                                         sortColumn As Integer) As String
            Using resolved = ExcelObjectResolver.Resolve(application, descriptor.RangeRef)
                Dim cells As Object = Nothing
                Dim firstCell As Object = Nothing
                Dim column As Object = Nothing
                Try
                    cells = resolved.Value.Cells
                    firstCell = cells.Item(1, sortColumn)
                    column = firstCell.Resize(descriptor.RowCount, 1)
                    Return ExcelObjectResolver.BuildRangeRef(descriptor.SheetName, CStr(column.Address(False, False)))
                Finally
                    ReleaseCom(column)
                    ReleaseCom(firstCell)
                    ReleaseCom(cells)
                End Try
            End Using
        End Function

        Private Shared Function IsRangeSorted(application As Object,
                                              descriptor As RangeDescriptor,
                                              sortColumn As Integer,
                                              hasHeader As Boolean,
                                              descending As Boolean) As Boolean
            Dim current = ResolveRange(application, QuoteSheet(descriptor.SheetName) & "!" & descriptor.Address, captureValues:=True)
            Dim matrix = ToMatrix(current.Values, current.RowCount, current.ColumnCount)
            Dim startRow = If(hasHeader, 1, 0)
            For row = startRow + 1 To current.RowCount - 1
                Dim comparison = CompareExcelValues(matrix(row - 1, sortColumn - 1), matrix(row, sortColumn - 1))
                If (Not descending AndAlso comparison > 0) OrElse (descending AndAlso comparison < 0) Then Return False
            Next
            Return True
        End Function

        Private Shared Function CompareExcelValues(left As Object, right As Object) As Integer
            If left Is Nothing AndAlso right Is Nothing Then Return 0
            If left Is Nothing Then Return -1
            If right Is Nothing Then Return 1
            Dim leftNumber As Decimal
            Dim rightNumber As Decimal
            If Decimal.TryParse(left.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, leftNumber) AndAlso
               Decimal.TryParse(right.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, rightNumber) Then
                Return leftNumber.CompareTo(rightNumber)
            End If
            Return StringComparer.CurrentCultureIgnoreCase.Compare(left.ToString(), right.ToString())
        End Function

        Private Shared Function VerifyNoDuplicates(application As Object,
                                                   descriptor As RangeDescriptor,
                                                   columns As JArray,
                                                   hasHeader As Boolean) As Boolean
            Dim current = ResolveRange(application, QuoteSheet(descriptor.SheetName) & "!" & descriptor.Address, captureValues:=True)
            Dim matrix = ToMatrix(current.Values, current.RowCount, current.ColumnCount)
            Dim seen As New HashSet(Of String)(StringComparer.Ordinal)
            Dim startRow = If(hasHeader, 1, 0)
            For row = startRow To current.RowCount - 1
                If IsBlankRow(matrix, row, current.ColumnCount) Then Continue For
                Dim currentRow = row
                Dim key = String.Join(ChrW(30), columns.Select(Function(column) NormalizeCell(matrix(currentRow, column.Value(Of Integer)() - 1))))
                If Not seen.Add(key) Then Return False
            Next
            Return True
        End Function

        Private Shared Function ReplaceValues(value As Object,
                                              findText As String,
                                              replacement As String,
                                              matchCase As Boolean,
                                              whole As Boolean) As Object
            Dim token = If(value Is Nothing, JValue.CreateNull(), JToken.FromObject(value))
            Dim options = If(matchCase, RegexOptions.None, RegexOptions.IgnoreCase)
            Dim pattern = ExcelWildcardPattern(findText, whole)
            Dim scalarValues As New List(Of JValue)()
            CollectScalarValues(token, scalarValues)
            For Each scalar In scalarValues
                If scalar.Type <> JTokenType.String Then Continue For
                scalar.Value = Regex.Replace(scalar.ToString(), pattern, Function(match) replacement, options)
            Next
            Return token
        End Function

        Private Shared Function ExcelWildcardPattern(value As String, whole As Boolean) As String
            Dim builder As New System.Text.StringBuilder()
            Dim escaped As Boolean = False
            For Each ch In value
                If escaped Then
                    builder.Append(Regex.Escape(ch.ToString()))
                    escaped = False
                ElseIf ch = "~"c Then
                    escaped = True
                ElseIf ch = "*"c Then
                    builder.Append(".*")
                ElseIf ch = "?"c Then
                    builder.Append(".")
                Else
                    builder.Append(Regex.Escape(ch.ToString()))
                End If
            Next
            Dim pattern = builder.ToString()
            Return If(whole, "^" & pattern & "$", pattern)
        End Function

        Private Shared Function ToMatrix(value As Object, rows As Integer, columns As Integer) As Object(,)
            Dim result(Math.Max(0, rows - 1), Math.Max(0, columns - 1)) As Object
            If TypeOf value Is Array Then
                Dim source = DirectCast(value, Array)
                If source.Rank = 2 Then
                    Dim rowBase = source.GetLowerBound(0)
                    Dim columnBase = source.GetLowerBound(1)
                    For row = 0 To rows - 1
                        For column = 0 To columns - 1
                            result(row, column) = source.GetValue(row + rowBase, column + columnBase)
                        Next
                    Next
                    Return result
                End If
            End If
            result(0, 0) = value
            Return result
        End Function

        Private Shared Function FlattenValues(value As Object) As List(Of Object)
            Dim result As New List(Of Object)()
            If TypeOf value Is Array Then
                Dim array = DirectCast(value, Array)
                For Each item In array
                    result.Add(item)
                Next
            Else
                result.Add(value)
            End If
            Return result
        End Function

        Private Shared Function IsBlankRow(matrix As Object(,), row As Integer, columns As Integer) As Boolean
            For column = 0 To columns - 1
                If matrix(row, column) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(matrix(row, column).ToString()) Then Return False
            Next
            Return True
        End Function

        Private Shared Function NormalizeCell(value As Object) As String
            If value Is Nothing Then Return "<null>"
            Return value.GetType().FullName & ":" & value.ToString()
        End Function

        Private Shared Sub ValidateDataShape(data As JToken, ByRef rows As Integer, ByRef columns As Integer)
            If data.Type <> JTokenType.Array Then
                rows = 1
                columns = 1
                Return
            End If
            Dim array = DirectCast(data, JArray)
            If array.Count = 0 Then Throw New FormatException("WriteData data cannot be empty")
            If array.All(Function(item) item.Type = JTokenType.Array) Then
                rows = array.Count
                columns = DirectCast(array(0), JArray).Count
                Dim expectedColumns = columns
                If columns = 0 OrElse array.OfType(Of JArray)().Any(Function(row) row.Count <> expectedColumns) Then Throw New FormatException("WriteData requires a non-empty rectangular matrix")
            Else
                rows = 1
                columns = array.Count
            End If
        End Sub

    End Class

End Namespace
