Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace OfficeRuntime

    ' Structured adapters for high-semantic Excel tools. Each adapter executes the
    ' declared tool contract and verifies the resulting workbook state directly.
    ' Natural-language phrasing never reaches this layer.
    Friend NotInheritable Partial Class ExcelStandardToolAdapter
        Private Shared Function ExecuteCleanData(application As Object, params As JObject) As ToolResult
            Const toolId As String = "CleanData"
            Dim rangeSpec = FirstText(params, "range")
            Dim operation = FirstText(params, "operation").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse String.IsNullOrWhiteSpace(operation) Then
                Return Invalid(toolId, "CleanData requires range and operation")
            End If

            If operation = "removeduplicates" Then
                Dim removeParams = DirectCast(params.DeepClone(), JObject)
                Dim removeResult = ExecuteRemoveDuplicates(application, removeParams)
                If removeResult IsNot Nothing Then
                    removeResult.ToolId = toolId
                    Dim removeObservation = TryCast(removeResult.Observation, JObject)
                    If removeObservation IsNot Nothing Then
                        removeObservation("sourceToolId") = toolId
                        removeObservation("summary") = If(removeResult.Success,
                                                           "已验证数据清洗操作 removeduplicates",
                                                           removeObservation("summary"))
                    End If
                End If
                Return removeResult
            End If

            Using target = OpenAdvancedRange(application, rangeSpec)
                Dim before = CaptureAdvancedRangeState(target.Value)
                Dim operationSatisfied As Boolean
                Select Case operation
                    Case "trim"
                        ForEachCell(target.Value,
                                    Sub(cell)
                                        Dim value = cell.Value2
                                        If TypeOf value Is String Then cell.Value2 = CStr(value).Trim()
                                    End Sub)
                        operationSatisfied = AllTextIsTrimmed(target.Value)

                    Case "fillempty"
                        Dim fillValue = TokenValue(params("fillValue"), 0)
                        ForEachCell(target.Value,
                                    Sub(cell)
                                        Dim value = cell.Value2
                                        If value Is Nothing OrElse String.IsNullOrWhiteSpace(value.ToString()) Then cell.Value2 = fillValue
                                    End Sub)
                        operationSatisfied = CountEmptyCells(target.Value) = 0

                    Case "replace"
                        Dim findText = FirstText(params, "findText")
                        Dim replaceText = If(params("replaceText")?.ToString(), "")
                        If String.IsNullOrEmpty(findText) Then Return Invalid(toolId, "CleanData replace requires findText")
                        Dim replaceReported = CBool(target.Value.Replace(findText,
                                                                         replaceText,
                                                                         XlLookAt.xlPart,
                                                                         XlSearchOrder.xlByRows,
                                                                         False,
                                                                         False,
                                                                         False,
                                                                         False))
                        operationSatisfied = replaceReported OrElse Not RangeContainsText(target.Value, findText)

                    Case Else
                        Return Invalid(toolId, $"Unsupported CleanData operation: {operation}")
                End Select

                Dim after = CaptureAdvancedRangeState(target.Value)
                Return BuildAdvancedMutationResult(toolId,
                                                   $"已验证数据清洗操作 {operation}",
                                                   New String() {target.Descriptor.RangeRef},
                                                   before,
                                                   after,
                                                   operationSatisfied)
            End Using
        End Function

        Private Shared Function ExecuteConditionalFormat(application As Object, params As JObject) As ToolResult
            Const toolId As String = "ConditionalFormat"
            Dim rangeSpec = FirstText(params, "range")
            Dim rule = FirstText(params, "rule").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(rangeSpec) OrElse String.IsNullOrWhiteSpace(rule) Then
                Return Invalid(toolId, "ConditionalFormat requires range and rule")
            End If

            Using target = OpenAdvancedRange(application, rangeSpec)
                Dim beforeRules = CaptureAdvancedConditionalFormats(target.Value)
                Dim createdRule As Object = Nothing
                Dim rules As Object = Nothing
                Try
                    rules = target.Value.FormatConditions
                    Select Case rule
                        Case "highlight"
                            Dim condition = ExcelConditionalFormatContract.ParseHighlightCondition(FirstText(params, "condition"))
                            Dim color = ExcelConditionalFormatContract.ParseColor(
                                If(String.IsNullOrWhiteSpace(FirstText(params, "color")), "#FFC7CE", FirstText(params, "color")))
                            createdRule = rules.Add(CInt(XlFormatConditionType.xlCellValue),
                                                    CInt(condition.ExcelOperator),
                                                    condition.FormulaOperand)
                            Dim interior As Object = Nothing
                            Try
                                interior = createdRule.Interior
                                interior.Color = color
                            Finally
                                ReleaseCom(interior)
                            End Try
                        Case "databar"
                            createdRule = rules.AddDatabar()
                        Case "colorscale"
                            createdRule = rules.AddColorScale(3)
                        Case "iconset"
                            createdRule = rules.AddIconSetCondition()
                        Case Else
                            Return Invalid(toolId, $"Unsupported ConditionalFormat rule: {rule}")
                    End Select
                Finally
                    ReleaseCom(createdRule)
                    ReleaseCom(rules)
                End Try

                Dim afterRules = CaptureAdvancedConditionalFormats(target.Value)
                Dim verification = ExcelConditionalFormatContract.EvaluatePostState(
                    rule,
                    FirstText(params, "condition"),
                    FirstText(params, "color"),
                    afterRules)
                Dim before As New JObject From {
                    {"formatConditionCount", beforeRules.Count},
                    {"formatConditions", beforeRules}
                }
                Dim after As New JObject From {
                    {"formatConditionCount", afterRules.Count},
                    {"formatConditions", afterRules},
                    {"verification", verification}
                }
                Return BuildAdvancedMutationResult(toolId,
                                                   $"已验证条件格式规则 {rule}",
                                                   New String() {target.Descriptor.RangeRef},
                                                   before,
                                                   after,
                                                   If(verification("satisfied")?.Value(Of Boolean)(), False),
                                                   verification:=verification)
            End Using
        End Function

        Private Shared Function ExecuteDataAnalysis(application As Object, params As JObject) As ToolResult
            Const toolId As String = "DataAnalysis"
            Dim sourceSpec = FirstText(params, "sourceRange")
            Dim analysisType = FirstText(params, "type").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(sourceSpec) OrElse String.IsNullOrWhiteSpace(analysisType) Then
                Return Invalid(toolId, "DataAnalysis requires sourceRange and type")
            End If

            If analysisType = "pivot" Then
                Dim pivotParams = DirectCast(params.DeepClone(), JObject)
                pivotParams("targetCell") = If(String.IsNullOrWhiteSpace(FirstText(params, "targetRange")),
                                                 FirstText(params, "targetCell"),
                                                 FirstText(params, "targetRange"))
                Dim pivotResult = ExecuteCreatePivotTable(application, pivotParams)
                If pivotResult IsNot Nothing Then
                    pivotResult.ToolId = toolId
                    Dim pivotObservation = TryCast(pivotResult.Observation, JObject)
                    If pivotObservation IsNot Nothing Then pivotObservation("sourceToolId") = toolId
                End If
                Return pivotResult
            End If

            Using source = OpenAdvancedRange(application, sourceSpec)
                Dim output As Array
                Select Case analysisType
                    Case "summary"
                        output = BuildSummaryMatrix(source.Value)
                    Case "groupby"
                        output = BuildGroupByMatrix(source.Value, params)
                    Case "ranking"
                        output = BuildRankingMatrix(source.Value, params)
                    Case Else
                        Return Invalid(toolId, $"Unsupported DataAnalysis type: {analysisType}")
                End Select

                Dim targetSpec = ResolveAdvancedTargetSpec(source,
                                                           FirstText(params, "targetRange"),
                                                           source.Descriptor.ColumnCount + 2)
                Using target = OpenAdvancedRange(application, targetSpec)
                    Dim outputRange As Object = Nothing
                    Try
                        outputRange = ResizeToMatrix(target.Value, output)
                        Dim before = CaptureAdvancedRangeState(outputRange)
                        outputRange.Value2 = output
                        outputRange.Columns.AutoFit()
                        Dim after = CaptureAdvancedRangeState(outputRange)
                        Dim outputRef = BuildAdvancedRangeRef(target.Descriptor.SheetName, CStr(outputRange.Address(False, False)))
                        Dim satisfied = RangeMatchesMatrix(outputRange, output)
                        Dim data As New JObject From {
                            {"outputRange", outputRef},
                            {"rows", output.GetLength(0)},
                            {"columns", output.GetLength(1)}
                        }
                        Return BuildAdvancedMutationResult(toolId,
                                                           $"已验证 {analysisType} 分析结果",
                                                           New String() {outputRef},
                                                           before,
                                                           after,
                                                           satisfied,
                                                           data)
                    Finally
                        ReleaseCom(outputRange)
                    End Try
                End Using
            End Using
        End Function

        Private Shared Function ExecuteTransformData(application As Object, params As JObject) As ToolResult
            Const toolId As String = "TransformData"
            Dim sourceSpec = FirstText(params, "sourceRange")
            Dim operation = FirstText(params, "operation").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(sourceSpec) OrElse String.IsNullOrWhiteSpace(operation) Then
                Return Invalid(toolId, "TransformData requires sourceRange and operation")
            End If

            Using source = OpenAdvancedRange(application, sourceSpec)
                Dim sourceMatrix = ReadRangeMatrix(source.Value)
                Dim output As Array
                Select Case operation
                    Case "transpose"
                        output = TransposeMatrix(sourceMatrix)
                    Case "split"
                        If source.Descriptor.ColumnCount <> 1 Then Return Invalid(toolId, "TransformData split requires a single source column")
                        Dim delimiter = FirstText(params, "delimiter")
                        If String.IsNullOrEmpty(delimiter) Then Return Invalid(toolId, "TransformData split requires delimiter")
                        output = SplitMatrix(sourceMatrix, delimiter)
                    Case "merge"
                        output = MergeMatrix(sourceMatrix, If(params("delimiter")?.ToString(), " "))
                    Case Else
                        Return Invalid(toolId, $"Unsupported TransformData operation: {operation}")
                End Select

                Dim defaultOffset = If(operation = "split", 0, source.Descriptor.ColumnCount + 1)
                Dim targetSpec = ResolveAdvancedTargetSpec(source, FirstText(params, "targetRange"), defaultOffset)
                Using target = OpenAdvancedRange(application, targetSpec)
                    Dim outputRange As Object = Nothing
                    Try
                        outputRange = ResizeToMatrix(target.Value, output)
                        Dim before = CaptureAdvancedRangeState(outputRange)
                        outputRange.Value2 = output
                        Dim after = CaptureAdvancedRangeState(outputRange)
                        Dim outputRef = BuildAdvancedRangeRef(target.Descriptor.SheetName, CStr(outputRange.Address(False, False)))
                        Return BuildAdvancedMutationResult(toolId,
                                                           $"已验证数据转换操作 {operation}",
                                                           New String() {outputRef},
                                                           before,
                                                           after,
                                                           RangeMatchesMatrix(outputRange, output),
                                                           New JObject From {{"outputRange", outputRef}})
                    Finally
                        ReleaseCom(outputRange)
                    End Try
                End Using
            End Using
        End Function

        Private Shared Function ExecuteGenerateReport(application As Object, params As JObject) As ToolResult
            Const toolId As String = "GenerateReport"
            Dim sourceSpec = FirstText(params, "sourceRange")
            If String.IsNullOrWhiteSpace(sourceSpec) Then Return Invalid(toolId, "GenerateReport requires sourceRange")

            Using source = OpenAdvancedRange(application, sourceSpec)
                Dim workbook As Object = Nothing
                Dim worksheets As Object = Nothing
                Dim targetSheet As Object = Nothing
                Dim dataAnchor As Object = Nothing
                Dim dataRange As Object = Nothing
                Dim headerRange As Object = Nothing
                Dim titleRange As Object = Nothing
                Dim borders As Object = Nothing
                Try
                    workbook = application.ActiveWorkbook
                    worksheets = workbook.Worksheets
                    Dim targetName = FirstText(params, "targetSheet")
                    If String.IsNullOrWhiteSpace(targetName) Then targetName = UniqueSheetName(worksheets, "报表")
                    Dim existed = TryGetWorksheet(worksheets, targetName, targetSheet)
                    Dim beforeSheetCount = CInt(worksheets.Count)
                    If Not existed Then
                        targetSheet = worksheets.Add()
                        targetSheet.Name = targetName
                    End If

                    Dim title = FirstText(params, "title")
                    Dim dataStartRow = If(String.IsNullOrWhiteSpace(title), 1, 3)
                    dataAnchor = targetSheet.Range("A" & dataStartRow.ToString(CultureInfo.InvariantCulture))
                    dataRange = dataAnchor.Resize(source.Descriptor.RowCount, source.Descriptor.ColumnCount)
                    Dim before As New JObject From {
                        {"sheetExisted", existed},
                        {"sheetCount", beforeSheetCount},
                        {"data", CaptureAdvancedRangeState(dataRange)},
                        {"chartCount", GetChartCount(targetSheet)}
                    }

                    source.Value.Copy(dataRange)
                    headerRange = dataAnchor.Resize(1, source.Descriptor.ColumnCount)
                    headerRange.Font.Bold = True
                    headerRange.Interior.Color = RGB(68, 114, 196)
                    headerRange.Font.Color = RGB(255, 255, 255)
                    borders = dataRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = XlBorderWeight.xlThin
                    targetSheet.Columns.AutoFit()
                    If Not String.IsNullOrWhiteSpace(title) Then
                        titleRange = targetSheet.Range("A1")
                        titleRange.Value2 = title
                        titleRange.Font.Size = 16
                        titleRange.Font.Bold = True
                    End If

                    Dim chartResult As ToolResult = Nothing
                    If If(params("includeChart")?.Value(Of Boolean)(), False) Then
                        Dim chartPositionRow = dataStartRow + source.Descriptor.RowCount + 2
                        chartResult = ExecuteCreateChart(application,
                                                         New JObject From {
                                                             {"dataRange", QuoteSheet(targetName) & "!" & CStr(dataRange.Address(False, False))},
                                                             {"position", QuoteSheet(targetName) & "!A" & chartPositionRow.ToString(CultureInfo.InvariantCulture)},
                                                             {"title", If(String.IsNullOrWhiteSpace(title), "报表图表", title)}
                                                         })
                        If chartResult Is Nothing OrElse Not chartResult.Success Then
                            If chartResult IsNot Nothing Then chartResult.ToolId = toolId
                            Return chartResult
                        End If
                    End If

                    Dim after As New JObject From {
                        {"sheetExisted", True},
                        {"sheetCount", CInt(worksheets.Count)},
                        {"data", CaptureAdvancedRangeState(dataRange)},
                        {"title", If(titleRange Is Nothing, "", CStr(If(titleRange.Value2, "")))},
                        {"chartCount", GetChartCount(targetSheet)}
                    }
                    Dim sourceMatrix = ReadRangeMatrix(source.Value)
                    Dim satisfied = RangeMatchesMatrix(dataRange, sourceMatrix) AndAlso
                                    (String.IsNullOrWhiteSpace(title) OrElse String.Equals(CStr(titleRange.Value2), title, StringComparison.Ordinal)) AndAlso
                                    (Not If(params("includeChart")?.Value(Of Boolean)(), False) OrElse GetChartCount(targetSheet) > 0)
                    Dim worksheetRef = BuildWorksheetRef(targetName)
                    Dim outputRef = BuildAdvancedRangeRef(targetName, CStr(dataRange.Address(False, False)))
                    Return BuildAdvancedMutationResult(toolId,
                                                       $"已验证报表工作表 {targetName}",
                                                       New String() {worksheetRef, outputRef},
                                                       before,
                                                       after,
                                                       satisfied,
                                                       New JObject From {
                                                           {"targetSheet", targetName},
                                                           {"outputRange", outputRef}
                                                       },
                                                       artifacts:=chartResult?.Artifacts)
                Finally
                    ReleaseCom(borders)
                    ReleaseCom(titleRange)
                    ReleaseCom(headerRange)
                    ReleaseCom(dataRange)
                    ReleaseCom(dataAnchor)
                    ReleaseCom(targetSheet)
                    ReleaseCom(worksheets)
                    ReleaseCom(workbook)
                End Try
            End Using
        End Function

        Private Shared Function ExecuteCreatePivotTable(application As Object, params As JObject) As ToolResult
            Const toolId As String = "CreatePivotTable"
            Dim sourceSpec = FirstText(params, "sourceRange")
            Dim targetSpec = FirstText(params, "targetCell")
            If String.IsNullOrWhiteSpace(sourceSpec) OrElse String.IsNullOrWhiteSpace(targetSpec) Then
                Return Invalid(toolId, "CreatePivotTable requires sourceRange and targetCell")
            End If
            Dim rowFields = TryCast(params("rowFields"), JArray)
            Dim valueFields = TryCast(params("valueFields"), JArray)
            Dim columnFields = TryCast(params("columnFields"), JArray)
            If rowFields Is Nothing OrElse rowFields.Count = 0 OrElse valueFields Is Nothing OrElse valueFields.Count = 0 Then
                Return Invalid(toolId, "CreatePivotTable requires rowFields and valueFields")
            End If

            Using source = OpenAdvancedRange(application, sourceSpec)
                Using target = OpenAdvancedRange(application, targetSpec)
                    Dim workbook As Object = Nothing
                    Dim pivotCaches As Object = Nothing
                    Dim pivotCache As Object = Nothing
                    Dim pivotTables As Object = Nothing
                    Dim pivotTable As Object = Nothing
                    Dim tableRange As Object = Nothing
                    Try
                        workbook = application.ActiveWorkbook
                        pivotTables = target.Worksheet.PivotTables()
                        Dim beforeCount = CInt(pivotTables.Count)
                        Dim before As New JObject From {
                            {"pivotTableCount", beforeCount},
                            {"target", CaptureAdvancedRangeState(target.Value)}
                        }
                        Dim sourceAddress = DirectCast(source.Value, Range).Address(
                            RowAbsolute:=True,
                            ColumnAbsolute:=True,
                            ReferenceStyle:=XlReferenceStyle.xlR1C1,
                            External:=True)
                        pivotCaches = DirectCast(workbook, Workbook).PivotCaches()
                        pivotCache = DirectCast(pivotCaches, PivotCaches).Create(
                            SourceType:=XlPivotTableSourceType.xlDatabase,
                            SourceData:=sourceAddress)
                        Dim tableName = "PivotTable_" & Guid.NewGuid().ToString("N").Substring(0, 10)
                        pivotTable = DirectCast(pivotCache, PivotCache).CreatePivotTable(
                            TableDestination:=DirectCast(target.Value, Range),
                            TableName:=tableName)
                        ConfigurePivotFields(pivotTable, rowFields, XlPivotFieldOrientation.xlRowField)
                        ConfigurePivotFields(pivotTable, columnFields, XlPivotFieldOrientation.xlColumnField)
                        ConfigurePivotValues(pivotTable, valueFields)

                        ReleaseCom(pivotTables)
                        pivotTables = target.Worksheet.PivotTables()
                        Dim afterCount = CInt(pivotTables.Count)
                        tableRange = pivotTable.TableRange2
                        Dim tableAddress = CStr(tableRange.Address(False, False))
                        Dim topLeft As Object = Nothing
                        Dim topLeftAddress As String = ""
                        Try
                            topLeft = tableRange.Cells(1, 1)
                            topLeftAddress = CStr(topLeft.Address(False, False))
                        Finally
                            ReleaseCom(topLeft)
                        End Try
                        Dim after As New JObject From {
                            {"pivotTableCount", afterCount},
                            {"tableName", tableName},
                            {"tableAddress", tableAddress},
                            {"topLeft", topLeftAddress}
                        }
                        Dim pivotRef = BuildWorksheetRef(target.Descriptor.SheetName) & "/pivottables/" & ExcelObjectResolver.EncodeSegment(tableName)
                        Dim satisfied = afterCount = beforeCount + 1 AndAlso
                                        String.Equals(topLeftAddress,
                                                      target.Descriptor.Address.Split(":"c)(0),
                                                      StringComparison.OrdinalIgnoreCase)
                        Return BuildAdvancedMutationResult(toolId,
                                                           $"已验证数据透视表 {tableName}",
                                                           New String() {pivotRef},
                                                           before,
                                                           after,
                                                           satisfied,
                                                           New JObject From {
                                                               {"pivotTable", pivotRef},
                                                               {"targetSheet", target.Descriptor.SheetName}
                                                           },
                                                           artifacts:=New JObject From {{"pivotTable", pivotRef}})
                    Finally
                        ReleaseCom(tableRange)
                        ReleaseCom(pivotTable)
                        ReleaseCom(pivotTables)
                        ReleaseCom(pivotCache)
                        ReleaseCom(pivotCaches)
                        ReleaseCom(workbook)
                    End Try
                End Using
            End Using
        End Function

    End Class

End Namespace
