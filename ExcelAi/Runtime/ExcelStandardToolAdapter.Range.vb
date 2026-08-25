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
            Dim writeValue = NormalizeWriteDataValue(data, rowCount, columnCount)
            Dim descriptor = ResolveRange(application, rangeSpec, rowCount, columnCount, captureValues:=False)
            Dim batch = NewBatch()
            AddOperation(batch,
                         "write-values",
                         descriptor.RangeRef,
                         "set",
                         "Range",
                         "Value2",
                         "property",
                         New JObject From {{"value", writeValue}},
                         New JObject From {
                             {"ValueHash", ExcelOperationObserver.ComputeValueHash(writeValue)},
                             {"RowCount", rowCount},
                             {"ColumnCount", columnCount}
                         })
            Dim result = ExecuteBatch(application, toolId, batch, $"已验证写入 {descriptor.SheetName}!{descriptor.Address}")
            Return AnnotateVerifiedRequestProjection(result, data)
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
            Dim result = ExecuteBatch(application, toolId, batch, $"已验证公式写入 {descriptor.SheetName}!{descriptor.Address}")
            result = VerifyFormulaPattern(application, descriptor, result, fillDown)
            Return AnnotateVerifiedRequestProjection(result, New JValue(formula), verifiedPropertyName:="FormulaPattern")
        End Function

        ''' <summary>
        ''' Verifies the semantic invariant produced by a scalar formula assignment. Fill-down
        ''' must preserve one relative R1C1 pattern across the destination; a direct multi-cell
        ''' assignment must preserve one literal A1 formula. A non-empty count alone is not proof
        ''' that every cell received the requested formula or the correct relative references.
        ''' </summary>
        Private Shared Function VerifyFormulaPattern(application As Object,
                                                     descriptor As RangeDescriptor,
                                                     source As ToolResult,
                                                     fillDown As Boolean) As ToolResult
            If source Is Nothing OrElse Not source.Success OrElse descriptor Is Nothing Then Return source
            Dim workbook As Object = Nothing
            Dim worksheets As Object = Nothing
            Dim worksheet As Object = Nothing
            Dim target As Object = Nothing
            Dim cells As Object = Nothing
            Dim topLeft As Object = Nothing
            Try
                workbook = application.ActiveWorkbook
                worksheets = workbook.Worksheets
                worksheet = worksheets.Item(descriptor.SheetName)
                target = worksheet.Range(descriptor.Address)
                cells = target.Cells
                topLeft = cells.Item(1, 1)
                Dim expectedPattern = If(fillDown,
                                         Convert.ToString(topLeft.FormulaR1C1, CultureInfo.InvariantCulture),
                                         Convert.ToString(topLeft.Formula, CultureInfo.InvariantCulture))
                Dim mismatch As String = ""
                For rowIndex = 1 To descriptor.RowCount
                    For columnIndex = 1 To descriptor.ColumnCount
                        Dim cell As Object = Nothing
                        Try
                            cell = cells.Item(rowIndex, columnIndex)
                            Dim actualPattern = If(fillDown,
                                                   Convert.ToString(cell.FormulaR1C1, CultureInfo.InvariantCulture),
                                                   Convert.ToString(cell.Formula, CultureInfo.InvariantCulture))
                            If Not String.Equals(actualPattern, expectedPattern, StringComparison.Ordinal) Then
                                mismatch = $"R{rowIndex}C{columnIndex}: {actualPattern}"
                                Exit For
                            End If
                        Finally
                            ReleaseCom(cell)
                        End Try
                    Next
                    If Not String.IsNullOrWhiteSpace(mismatch) Then Exit For
                Next

                Dim observation = TryCast(source.Observation, JObject)
                If observation Is Nothing Then observation = New JObject()
                Dim verification = TryCast(observation("verification"), JArray)
                If verification Is Nothing Then
                    verification = New JArray()
                    observation("verification") = verification
                End If
                Dim passed = String.IsNullOrWhiteSpace(mismatch)
                verification.Add(New JObject From {
                    {"id", "formula-pattern"},
                    {"required", True},
                    {"targetRef", descriptor.RangeRef},
                    {"effectType", "formula_state"},
                    {"property", "FormulaPattern"},
                    {"status", If(passed, "passed", "failed")},
                    {"expected", expectedPattern},
                    {"actual", If(passed, expectedPattern, mismatch)}
                })
                source.Observation = observation
                If passed Then Return source
                Return MarkCompositeMutationFailure(
                    SemanticFailure(source.ToolId, source, "Excel formula pattern did not match the requested fill semantics"),
                    descriptor.RangeRef)
            Catch ex As Exception
                Return MarkCompositeMutationFailure(
                    SemanticFailure(source.ToolId, source, $"Unable to verify the complete formula range: {ex.Message}"),
                    descriptor.RangeRef)
            Finally
                ReleaseCom(topLeft)
                ReleaseCom(cells)
                ReleaseCom(target)
                ReleaseCom(worksheet)
                ReleaseCom(worksheets)
                ReleaseCom(workbook)
            End Try
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
                Dim criteria1 As JToken = Nothing
                Dim criteria2 As JToken = Nothing
                Dim filterOperator As XlAutoFilterOperator = XlAutoFilterOperator.xlAnd
                NormalizeFilterCriteria(params, criteria1, criteria2, filterOperator)
                If criteria1 IsNot Nothing AndAlso criteria1.Type <> JTokenType.Null Then arguments("Criteria1") = criteria1.DeepClone()
                If criteria2 IsNot Nothing AndAlso criteria2.Type <> JTokenType.Null Then
                    arguments("Operator") = CInt(filterOperator)
                    arguments("Criteria2") = criteria2.DeepClone()
                End If
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
                    .PropertyName = If(criteria1 Is Nothing OrElse String.IsNullOrWhiteSpace(criteria1.ToString()), "AutoFilterMode", "FilterMode"),
                    .Operator = "equals",
                    .ExpectedValue = New JValue(True),
                    .Required = True
                })
            End If
            Dim result = ExecuteBatch(application, toolId, batch, If(clearFilter, "已验证清除筛选", "已验证筛选状态"))
            If clearFilter OrElse Not result.Success Then Return result

            Dim requestedCriteria1 As JToken = Nothing
            Dim requestedCriteria2 As JToken = Nothing
            Dim requestedOperator As XlAutoFilterOperator = XlAutoFilterOperator.xlAnd
            NormalizeFilterCriteria(params, requestedCriteria1, requestedCriteria2, requestedOperator)
            Return VerifyFilterRequest(application,
                                       descriptor,
                                       If(params("column")?.Value(Of Integer?)(), 0),
                                       requestedCriteria1,
                                       requestedCriteria2,
                                       requestedOperator,
                                       result)
        End Function

        ''' <summary>
        ''' Normalizes the public filter contract into Excel's two-criterion AutoFilter
        ''' shape.  The semicolon form remains supported for existing callers, while new
        ''' callers can send criteria + criteria2 + operator explicitly.  This is grammar
        ''' normalization only; it contains no business-value special cases.
        ''' </summary>
        Private Shared Sub NormalizeFilterCriteria(params As JObject,
                                                   ByRef criteria1 As JToken,
                                                   ByRef criteria2 As JToken,
                                                   ByRef filterOperator As XlAutoFilterOperator)
            criteria1 = params?("criteria")?.DeepClone()
            criteria2 = params?("criteria2")?.DeepClone()
            Dim operatorText = FirstText(params, "operator").Trim().ToLowerInvariant()
            filterOperator = If(operatorText = "or", XlAutoFilterOperator.xlOr, XlAutoFilterOperator.xlAnd)

            If criteria2 IsNot Nothing OrElse criteria1 Is Nothing OrElse criteria1.Type <> JTokenType.String Then Return
            Dim source = criteria1.ToString()
            Dim separator = source.IndexOf(";"c)
            If separator <= 0 OrElse separator >= source.Length - 1 OrElse
               source.IndexOf(";"c, separator + 1) >= 0 Then Return

            Dim left = source.Substring(0, separator).Trim()
            Dim right = source.Substring(separator + 1).Trim()
            If left.Length = 0 OrElse right.Length = 0 Then Return
            criteria1 = New JValue(left)
            criteria2 = New JValue(right)
        End Sub

        ''' <summary>
        ''' Reads Excel's live AutoFilter object after execution and proves the exact field,
        ''' criteria and Boolean operator that the host retained.  FilterMode=True alone is
        ''' not sufficient evidence for a user-requested predicate.
        ''' </summary>
        Private Shared Function VerifyFilterRequest(application As Object,
                                                    descriptor As RangeDescriptor,
                                                    column As Integer,
                                                    criteria1 As JToken,
                                                    criteria2 As JToken,
                                                    filterOperator As XlAutoFilterOperator,
                                                    source As ToolResult) As ToolResult
            If source Is Nothing OrElse Not source.Success Then Return source
            Dim workbook As Object = Nothing
            Dim worksheets As Object = Nothing
            Dim worksheet As Object = Nothing
            Dim autoFilter As Object = Nothing
            Dim filterRange As Object = Nothing
            Dim filters As Object = Nothing
            Dim filter As Object = Nothing
            Try
                workbook = application.ActiveWorkbook
                worksheets = workbook.Worksheets
                worksheet = worksheets.Item(descriptor.SheetName)
                autoFilter = worksheet.AutoFilter
                If autoFilter Is Nothing Then
                    Return SemanticFailure(source.ToolId, source, "Excel did not retain an AutoFilter object")
                End If
                filterRange = autoFilter.Range
                filters = autoFilter.Filters
                filter = filters.Item(column)

                Dim expectedRange = descriptor.Address.Replace("$", "")
                Dim actualRange = Convert.ToString(filterRange.Address(False, False, XlReferenceStyle.xlA1), CultureInfo.InvariantCulture).Replace("$", "")
                Dim actualOn = CBool(filter.On)
                Dim actualCriteria1 = ReadFilterCriterion(filter, "Criteria1")
                Dim actualCriteria2 = ReadFilterCriterion(filter, "Criteria2")
                Dim actualOperator = ReadFilterOperator(filter)
                Dim expectedCriteria1 = If(criteria1?.ToString(), "")
                Dim expectedCriteria2 = If(criteria2?.ToString(), "")
                Dim expectedOperator = If(criteria2 Is Nothing, 0, CInt(filterOperator))
                Dim passed = actualOn AndAlso
                    String.Equals(expectedRange, actualRange, StringComparison.OrdinalIgnoreCase) AndAlso
                    FilterCriteriaEquivalent(expectedCriteria1, actualCriteria1) AndAlso
                    FilterCriteriaEquivalent(expectedCriteria2, actualCriteria2) AndAlso
                    (criteria2 Is Nothing OrElse actualOperator = expectedOperator)

                Dim expected As New JObject From {
                    {"range", descriptor.RangeRef},
                    {"field", column},
                    {"criteria1", expectedCriteria1},
                    {"criteria2", expectedCriteria2},
                    {"operator", expectedOperator}
                }
                Dim actual As New JObject From {
                    {"range", descriptor.RangeRef},
                    {"field", column},
                    {"criteria1", actualCriteria1},
                    {"criteria2", actualCriteria2},
                    {"operator", actualOperator},
                    {"active", actualOn}
                }
                Dim observation = TryCast(source.Observation, JObject)
                If observation Is Nothing Then observation = New JObject()
                Dim verification = TryCast(observation("verification"), JArray)
                If verification Is Nothing Then
                    verification = New JArray()
                    observation("verification") = verification
                End If
                verification.Add(New JObject From {
                    {"id", "filter-request"},
                    {"required", True},
                    {"targetRef", BuildWorksheetRef(descriptor.SheetName)},
                    {"effectType", "filter_state"},
                    {"property", "FilterCriteria"},
                    {"status", If(passed, "passed", "failed")},
                    {"expected", expected},
                    {"actual", actual},
                    {"requestProperty", "FilterCriteria"},
                    {"requestExpected", expected.DeepClone()}
                })
                observation("satisfied") = passed AndAlso Not ExcelOperationObserver.HasRequiredVerificationFailure(verification)
                source.Observation = observation
                If passed Then Return source
                Return SemanticFailure(source.ToolId, source, "Excel retained a filter, but its live criteria did not match the requested predicate")
            Catch ex As Exception
                Return SemanticFailure(source.ToolId, source, $"Unable to verify Excel's live filter criteria: {ex.Message}")
            Finally
                ReleaseCom(filter)
                ReleaseCom(filters)
                ReleaseCom(filterRange)
                ReleaseCom(autoFilter)
                ReleaseCom(worksheet)
                ReleaseCom(worksheets)
                ReleaseCom(workbook)
            End Try
        End Function

        Private Shared Function ReadFilterCriterion(filter As Object, propertyName As String) As String
            Try
                Return Convert.ToString(Interaction.CallByName(filter, propertyName, CallType.Get), CultureInfo.InvariantCulture)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function ReadFilterOperator(filter As Object) As Integer
            Try
                Return Convert.ToInt32(Interaction.CallByName(filter, "Operator", CallType.Get), CultureInfo.InvariantCulture)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function FilterCriteriaEquivalent(expected As String, actual As String) As Boolean
            Dim left = If(expected, "").Trim()
            Dim right = If(actual, "").Trim()
            If String.Equals(left, right, StringComparison.OrdinalIgnoreCase) Then Return True
            ' Excel canonicalizes a plain text equality criterion to "=value" when it is
            ' read back from Filters.Item(...).Criteria1.
            Return left.Length > 0 AndAlso
                Not left.StartsWith("=", StringComparison.Ordinal) AndAlso
                String.Equals("=" & left, right, StringComparison.OrdinalIgnoreCase)
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
            Dim verificationResult = ExecuteBatch(application, toolId, verificationBatch, "已验证 AutoFit 达到稳定尺寸")
            If Not verificationResult.Success Then Return MarkCompositeMutationFailure(verificationResult, descriptor.RangeRef)
            Return verificationResult
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
                EnsureExcelCell(data)
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
                For Each row In array.OfType(Of JArray)()
                    For Each cell In row
                        EnsureExcelCell(cell)
                    Next
                Next
            Else
                If array.Any(Function(item) item.Type = JTokenType.Array) Then Throw New FormatException("WriteData does not accept mixed scalar and row values")
                For Each cell In array
                    EnsureExcelCell(cell)
                Next
                rows = 1
                columns = array.Count
            End If
        End Sub

        Private Shared Function NormalizeWriteDataValue(data As JToken,
                                                        rows As Integer,
                                                        columns As Integer) As JToken
            If rows = 1 AndAlso columns = 1 Then
                If data.Type <> JTokenType.Array Then Return data.DeepClone()
                Dim outer = DirectCast(data, JArray)
                If outer(0).Type <> JTokenType.Array Then Return outer(0).DeepClone()
                Return DirectCast(outer(0), JArray)(0).DeepClone()
            End If

            If data.Type = JTokenType.Array Then
                Dim array = DirectCast(data, JArray)
                If array.All(Function(item) item.Type = JTokenType.Array) Then Return data.DeepClone()
                Dim row As New JArray()
                For Each cell In array
                    row.Add(cell.DeepClone())
                Next
                Dim matrix As New JArray()
                matrix.Add(row)
                Return matrix
            End If
            Return data.DeepClone()
        End Function

        Private Shared Sub EnsureExcelCell(value As JToken)
            If value Is Nothing Then Return
            Select Case value.Type
                Case JTokenType.Object, JTokenType.Array, JTokenType.Property, JTokenType.Constructor
                    Throw New FormatException(
                        $"WriteData cell values must be scalar; received {value.Type}. Convert structured JSON to a rectangular array before writing.")
            End Select
        End Sub

    End Class

End Namespace
