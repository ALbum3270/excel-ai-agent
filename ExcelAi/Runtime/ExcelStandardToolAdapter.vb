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

    ''' <summary>
    ''' Compatibility Adapter for stable, object-model Excel tools. Natural-language
    ''' variation stops at the tool schema; every mutation below is compiled into the
    ''' same catalogued, observed, and verified OfficeOperationBatch Module.
    ''' </summary>
    Friend NotInheritable Partial Class ExcelStandardToolAdapter
        Private Shared ReadOnly MigratedToolIds As New HashSet(Of String)(
            {
                "WriteData", "ApplyFormula", "SortData", "FilterData", "RemoveDuplicates",
                "MergeCells", "AutoFit", "FindReplace", "CreateSheet", "DeleteSheet",
                "RenameSheet", "CopySheet", "InsertRowCol", "DeleteRowCol", "HideRowCol",
                "ProtectSheet", "CreateChart", "CleanData", "ConditionalFormat",
                "CreatePivotTable", "DataAnalysis", "TransformData", "GenerateReport"
            },
            StringComparer.OrdinalIgnoreCase)

        Private Sub New()
        End Sub

        Public Shared Function TryExecute(application As Object,
                                          toolId As String,
                                          params As JObject,
                                          ByRef result As ToolResult) As Boolean
            result = Nothing
            toolId = NormalizeToolId(toolId)
            If String.IsNullOrWhiteSpace(toolId) OrElse Not MigratedToolIds.Contains(toolId) Then Return False
            Try
                If application Is Nothing Then
                    result = Failed(toolId, "Excel application is unavailable", ExceptionClassifier.CodeDocMissing, False)
                    Return True
                End If
                If params Is Nothing Then
                    result = Invalid(toolId, $"{toolId} params are required")
                    Return True
                End If

                Select Case toolId.Trim().ToLowerInvariant()
                    Case "writedata" : result = ExecuteWriteData(application, params)
                    Case "applyformula" : result = ExecuteApplyFormula(application, params)
                    Case "sortdata" : result = ExecuteSortData(application, params)
                    Case "filterdata" : result = ExecuteFilterData(application, params)
                    Case "removeduplicates" : result = ExecuteRemoveDuplicates(application, params)
                    Case "mergecells" : result = ExecuteMergeCells(application, params)
                    Case "autofit" : result = ExecuteAutoFit(application, params)
                    Case "findreplace" : result = ExecuteFindReplace(application, params)
                    Case "createsheet" : result = ExecuteCreateSheet(application, params)
                    Case "deletesheet" : result = ExecuteDeleteSheet(application, params)
                    Case "renamesheet" : result = ExecuteRenameSheet(application, params)
                    Case "copysheet" : result = ExecuteCopySheet(application, params)
                    Case "insertrowcol" : result = ExecuteInsertDeleteRowColumn(application, params, insert:=True)
                    Case "deleterowcol" : result = ExecuteInsertDeleteRowColumn(application, params, insert:=False)
                    Case "hiderowcol" : result = ExecuteHideRowColumn(application, params)
                    Case "protectsheet" : result = ExecuteProtectSheet(application, params)
                    Case "createchart" : result = ExecuteCreateChart(application, params)
                    Case "cleandata" : result = ExecuteCleanData(application, params)
                    Case "conditionalformat" : result = ExecuteConditionalFormat(application, params)
                    Case "createpivottable" : result = ExecuteCreatePivotTable(application, params)
                    Case "dataanalysis" : result = ExecuteDataAnalysis(application, params)
                    Case "transformdata" : result = ExecuteTransformData(application, params)
                    Case "generatereport" : result = ExecuteGenerateReport(application, params)
                End Select
            Catch ex As FormatException
                result = Invalid(toolId, ex.Message)
            Catch ex As Exception
                result = ToolResult.FromException(toolId, ex)
            End Try
            Return True
        End Function

        Private Shared Function NewBatch() As OfficeOperationBatch
            Return New OfficeOperationBatch With {
                .SchemaVersion = OfficeOperationValidation.CurrentSchemaVersion,
                .AppType = "Excel",
                .Atomic = True
            }
        End Function

        Private Shared Sub AddOperation(batch As OfficeOperationBatch,
                                        id As String,
                                        targetRef As String,
                                        action As String,
                                        typeName As String,
                                        memberName As String,
                                        memberKind As String,
                                        arguments As JObject,
                                        expectedEffects As JObject)
            Dim memberId = ExcelApiCatalogProvider.FindMemberId(typeName, memberName, memberKind)
            If String.IsNullOrWhiteSpace(memberId) Then Throw New FormatException($"Excel capability catalog does not expose {typeName}.{memberName}")
            batch.Operations.Add(New OfficeOperation With {
                .Id = id,
                .TargetRef = targetRef,
                .Action = action,
                .MemberId = memberId,
                .Arguments = If(arguments, New JObject()),
                .ExpectedEffects = If(expectedEffects, New JObject())
            })
        End Sub

        Private Shared Function ExecuteBatch(application As Object,
                                              toolId As String,
                                              batch As OfficeOperationBatch,
                                              successMessage As String) As ToolResult
            Dim result = ExcelOperationExecutor.Execute(application, New JObject From {{"batch", JObject.FromObject(batch)}})
            result.ToolId = toolId
            Dim observation = TryCast(result.Observation, JObject)
            If observation IsNot Nothing Then
                observation("adapter") = "ExcelStandardToolAdapter"
                observation("sourceToolId") = toolId
                observation("summary") = If(result.Success, successMessage, observation("summary"))
            End If
            If result.Success Then
                result.Message = successMessage
                result.UserMessage = successMessage
            End If
            Return result
        End Function

        Private Shared Function ResolveRange(application As Object,
                                             rangeSpec As String,
                                             Optional resizeRows As Integer = 0,
                                             Optional resizeColumns As Integer = 0,
                                             Optional captureValues As Boolean = False) As RangeDescriptor
            If String.IsNullOrWhiteSpace(rangeSpec) Then Throw New FormatException("Range is required")
            Dim workbook As Object = Nothing
            Dim worksheets As Object = Nothing
            Dim worksheet As Object = Nothing
            Dim usedRange As Object = Nothing
            Dim usedRows As Object = Nothing
            Dim target As Object = Nothing
            Dim resized As Object = Nothing
            Dim topLeft As Object = Nothing
            Try
                workbook = application.ActiveWorkbook
                worksheet = application.ActiveSheet
                If workbook Is Nothing OrElse worksheet Is Nothing Then Throw New FormatException("No active workbook or worksheet")
                Dim sheetName = CStr(worksheet.Name)
                Dim address = rangeSpec.Trim()
                Dim bangIndex = address.LastIndexOf("!"c)
                If bangIndex > 0 Then
                    sheetName = UnquoteSheet(address.Substring(0, bangIndex))
                    address = address.Substring(bangIndex + 1).Trim()
                    worksheets = workbook.Worksheets
                    ReleaseCom(worksheet)
                    worksheet = worksheets.Item(sheetName)
                End If
                If address.IndexOf("{lastRow}", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    usedRange = worksheet.UsedRange
                    usedRows = usedRange.Rows
                    Dim lastRow = CInt(usedRange.Row) + CInt(usedRows.Count) - 1
                    address = Regex.Replace(address, "\{lastRow\}", Math.Max(1, lastRow).ToString(CultureInfo.InvariantCulture), RegexOptions.IgnoreCase)
                End If
                target = worksheet.Range(address)
                If resizeRows > 0 AndAlso resizeColumns > 0 Then
                    topLeft = target.Cells(1, 1)
                    resized = topLeft.Resize(resizeRows, resizeColumns)
                    ReleaseCom(target)
                    target = resized
                    resized = Nothing
                End If
                Dim rows = CInt(target.Rows.Count)
                Dim columns = CInt(target.Columns.Count)
                Dim resolvedAddress = CStr(target.Address(False, False))
                topLeft = target.Cells(1, 1)
                Dim topLeftAddress = CStr(topLeft.Address(False, False))
                Return New RangeDescriptor With {
                    .SheetName = sheetName,
                    .Address = resolvedAddress,
                    .RangeRef = ExcelObjectResolver.BuildRangeRef(sheetName, resolvedAddress),
                    .TopLeftRef = ExcelObjectResolver.BuildRangeRef(sheetName, topLeftAddress),
                    .RowCount = rows,
                    .ColumnCount = columns,
                    .CellCount = rows * columns,
                    .Values = If(captureValues, target.Value2, Nothing),
                    .Formulas = If(captureValues, target.Formula, Nothing)
                }
            Catch ex As Exception
                If TypeOf ex Is FormatException Then Throw
                Throw New FormatException($"Unable to resolve Excel range '{rangeSpec}': {ex.Message}")
            Finally
                ReleaseCom(topLeft)
                ReleaseCom(resized)
                ReleaseCom(target)
                ReleaseCom(usedRows)
                ReleaseCom(usedRange)
                ReleaseCom(worksheet)
                ReleaseCom(worksheets)
                ReleaseCom(workbook)
            End Try
        End Function

        Private Shared Sub CollectScalarValues(token As JToken, values As List(Of JValue))
            If token Is Nothing Then Return
            Dim scalar = TryCast(token, JValue)
            If scalar IsNot Nothing Then
                values.Add(scalar)
                Return
            End If
            For Each child In token.Children()
                CollectScalarValues(child, values)
            Next
        End Sub

        Private Shared Function GetOperationResultRef(result As ToolResult, index As Integer) As String
            Dim data = TryCast(result?.Data, JObject)
            Return data?("operations")?(index)?("resultRef")?.ToString()
        End Function

        Private Shared Function FirstText(source As JObject, ParamArray names As String()) As String
            For Each name In names
                Dim token = source?.GetValue(name, StringComparison.OrdinalIgnoreCase)
                If token IsNot Nothing AndAlso token.Type <> JTokenType.Null AndAlso Not String.IsNullOrWhiteSpace(token.ToString()) Then Return token.ToString()
            Next
            Return ""
        End Function

        Private Shared Function NormalizeToolId(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "writedata", "write", "setvalue", "setvalues" : Return "WriteData"
                Case "applyformula", "formula", "calculatesum", "calculate", "range_operations" : Return "ApplyFormula"
                Case "sortdata", "sort" : Return "SortData"
                Case "filterdata", "filter" : Return "FilterData"
                Case "removeduplicates" : Return "RemoveDuplicates"
                Case "mergecells", "merge" : Return "MergeCells"
                Case "autofit" : Return "AutoFit"
                Case "findreplace" : Return "FindReplace"
                Case "createsheet" : Return "CreateSheet"
                Case "deletesheet" : Return "DeleteSheet"
                Case "renamesheet" : Return "RenameSheet"
                Case "copysheet" : Return "CopySheet"
                Case "insertrowcol" : Return "InsertRowCol"
                Case "deleterowcol" : Return "DeleteRowCol"
                Case "hiderowcol" : Return "HideRowCol"
                Case "protectsheet" : Return "ProtectSheet"
                Case "createchart", "chart" : Return "CreateChart"
                Case "cleandata", "clean" : Return "CleanData"
                Case "conditionalformat" : Return "ConditionalFormat"
                Case "createpivottable", "pivot" : Return "CreatePivotTable"
                Case "dataanalysis", "analyze" : Return "DataAnalysis"
                Case "transformdata", "transform" : Return "TransformData"
                Case "generatereport", "report" : Return "GenerateReport"
                Case Else : Return ""
            End Select
        End Function

        Private Shared Function BuildWorksheetRef(name As String) As String
            Return "Excel:workbooks/active/worksheets/" & ExcelObjectResolver.EncodeSegment(name)
        End Function

        Private Shared Function QuoteSheet(name As String) As String
            Return "'" & If(name, "").Replace("'", "''") & "'"
        End Function

        Private Shared Function UnquoteSheet(value As String) As String
            Dim result = If(value, "").Trim()
            If result.Length >= 2 AndAlso result.StartsWith("'", StringComparison.Ordinal) AndAlso result.EndsWith("'", StringComparison.Ordinal) Then
                result = result.Substring(1, result.Length - 2).Replace("''", "'")
            End If
            Return result
        End Function

        Private Shared Function GetColumnOffset(columnName As String, offset As Integer) As String
            Dim number = ColumnNumber(columnName) + offset
            If number < 1 OrElse number > 16384 Then Throw New FormatException("Column position is outside Excel limits")
            Dim result = ""
            While number > 0
                number -= 1
                result = ChrW(AscW("A"c) + (number Mod 26)) & result
                number \= 26
            End While
            Return result
        End Function

        Private Shared Function ColumnNumber(columnName As String) As Integer
            Dim result As Integer = 0
            For Each ch In If(columnName, "").Trim().ToUpperInvariant()
                If ch < "A"c OrElse ch > "Z"c Then Throw New FormatException($"Invalid column position '{columnName}'")
                result = result * 26 + AscW(ch) - AscW("A"c) + 1
            Next
            If result = 0 Then Throw New FormatException($"Invalid column position '{columnName}'")
            Return result
        End Function

        Private Shared Function SemanticFailure(toolId As String, source As ToolResult, message As String) As ToolResult
            Dim observation = TryCast(source?.Observation, JObject)
            If observation IsNot Nothing Then
                observation("satisfied") = False
                observation("summary") = message
            End If
            Return ToolResult.Failed(toolId,
                                     message,
                                     data:=source?.Data,
                                     errorCode:=ExceptionClassifier.CodeVerifyFailed,
                                     userMessage:="Excel 已执行操作，但实际结果未满足声明的后置条件",
                                     recoverable:=False,
                                     observation:=source?.Observation,
                                     artifacts:=source?.Artifacts)
        End Function

        Private Shared Function Invalid(toolId As String, message As String) As ToolResult
            Return Failed(toolId, message, ExceptionClassifier.CodeOperationSchemaInvalid, True)
        End Function

        Private Shared Function Failed(toolId As String,
                                       message As String,
                                       errorCode As String,
                                       recoverable As Boolean) As ToolResult
            Return ToolResult.Failed(toolId,
                                     message,
                                     errorCode:=errorCode,
                                     userMessage:="Excel 操作参数或当前对象状态无效",
                                     recoverable:=recoverable)
        End Function

        Private Shared Sub ReleaseCom(value As Object)
            If value Is Nothing Then Return
            Try
                If Marshal.IsComObject(value) Then Marshal.ReleaseComObject(value)
            Catch
            End Try
        End Sub

        Private Class RangeDescriptor
            Public Property SheetName As String
            Public Property Address As String
            Public Property RangeRef As String
            Public Property TopLeftRef As String
            Public Property RowCount As Integer
            Public Property ColumnCount As Integer
            Public Property CellCount As Integer
            Public Property StartRow As Integer
            Public Property StartColumn As Integer
            Public Property Values As Object
            Public Property Formulas As Object
        End Class
    End Class

End Namespace
