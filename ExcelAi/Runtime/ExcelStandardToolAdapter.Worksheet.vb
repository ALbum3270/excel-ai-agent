Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent
Imports ExcelAgent.Core.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend NotInheritable Partial Class ExcelStandardToolAdapter
        Private Shared Function ExecuteCreateSheet(application As Object, params As JObject) As ToolResult
            Const toolId As String = "CreateSheet"
            Dim name = FirstText(params, "name", "sheetName")
            ValidateWorksheetName(application, name, mustExist:=False)
            Dim worksheetsRef = "Excel:workbooks/active/worksheets"
            Dim addArguments As New JObject()
            Dim referenceSheet = FirstText(params, "referenceSheet")
            Dim position = FirstText(params, "position").Trim().ToLowerInvariant()
            If Not String.IsNullOrWhiteSpace(referenceSheet) Then
                ValidateWorksheetName(application, referenceSheet, mustExist:=True)
                addArguments(If(position = "before", "Before", "After")) = New JObject From {{"ref", BuildWorksheetRef(referenceSheet)}}
            End If

            Dim addBatch = NewBatch()
            AddOperation(addBatch, "add-sheet", worksheetsRef, "create", "Worksheets", "Add", "method", addArguments, New JObject From {{"exists", True}})
            Dim addResult = ExecuteBatch(application, toolId, addBatch, "Excel 已创建待命名工作表")
            If Not addResult.Success Then Return addResult
            Dim createdRef = GetOperationResultRef(addResult, 0)
            If String.IsNullOrWhiteSpace(createdRef) Then Return SemanticFailure(toolId, addResult, "Worksheets.Add returned no canonical object ref")

            Dim renameBatch = NewBatch()
            AddOperation(renameBatch, "name-sheet", createdRef, "set", "Worksheet", "Name", "property",
                         New JObject From {{"value", name}}, New JObject From {{"Name", name}})
            Dim renameResult = ExecuteBatch(application, toolId, renameBatch, $"已验证创建工作表 {name}")
            If Not renameResult.Success Then
                CleanupCreatedSheet(application, createdRef)
                Return MarkCompositeMutationFailure(renameResult, createdRef)
            End If
            Return renameResult
        End Function

        Private Shared Function ExecuteDeleteSheet(application As Object, params As JObject) As ToolResult
            Const toolId As String = "DeleteSheet"
            Dim name = FirstText(params, "name")
            ValidateWorksheetName(application, name, mustExist:=True)
            Dim batch = NewBatch()
            AddOperation(batch, "delete-sheet", BuildWorksheetRef(name), "delete", "Worksheet", "Delete", "method",
                         New JObject(), New JObject From {{"exists", False}})
            Dim previousAlerts = GetDisplayAlerts(application)
            Try
                SetDisplayAlerts(application, False)
                Return ExecuteBatch(application, toolId, batch, $"已验证删除工作表 {name}")
            Finally
                SetDisplayAlerts(application, previousAlerts)
            End Try
        End Function

        Private Shared Function ExecuteRenameSheet(application As Object, params As JObject) As ToolResult
            Const toolId As String = "RenameSheet"
            Dim oldName = FirstText(params, "oldName")
            Dim newName = FirstText(params, "newName")
            ValidateWorksheetName(application, oldName, mustExist:=True)
            ValidateWorksheetName(application, newName, mustExist:=False)
            Dim batch = NewBatch()
            AddOperation(batch, "rename-sheet", BuildWorksheetRef(oldName), "set", "Worksheet", "Name", "property",
                         New JObject From {{"value", newName}}, New JObject From {{"Name", newName}})
            Return ExecuteBatch(application, toolId, batch, $"已验证重命名 {oldName} → {newName}")
        End Function

        Private Shared Function ExecuteCopySheet(application As Object, params As JObject) As ToolResult
            Const toolId As String = "CopySheet"
            Dim sourceName = FirstText(params, "sourceName")
            Dim newName = FirstText(params, "newName")
            ValidateWorksheetName(application, sourceName, mustExist:=True)
            ValidateWorksheetName(application, newName, mustExist:=False)
            Dim sourceRef = BuildWorksheetRef(sourceName)
            Dim sourceHash = CaptureWorksheetUsedValueHash(application, sourceName)
            Dim countBefore = GetWorksheetCount(application)
            Dim copyBatch = NewBatch()
            AddOperation(copyBatch, "copy-sheet", sourceRef, "invoke", "Worksheet", "Copy", "method",
                         New JObject From {{"After", New JObject From {{"ref", sourceRef}}}}, New JObject())
            copyBatch.SuccessCriteria.Add(New OperationCriterion With {
                .Id = "sheet-count-increased", .TargetRef = "Excel:workbooks/active", .PropertyName = "WorksheetCount",
                .Operator = "gte", .ExpectedValue = New JValue(countBefore + 1), .Required = True
            })
            Dim copyResult = ExecuteBatch(application, toolId, copyBatch, "Excel 已复制工作表")
            If Not copyResult.Success Then Return copyResult
            Dim copiedRef = BuildWorksheetRef(GetActiveSheetName(application))
            Dim finalRef = BuildWorksheetRef(newName)
            Dim renameBatch = NewBatch()
            AddOperation(renameBatch, "name-copy", copiedRef, "set", "Worksheet", "Name", "property",
                         New JObject From {{"value", newName}}, New JObject From {{"Name", newName}})
            renameBatch.SuccessCriteria.Add(New OperationCriterion With {
                .Id = "copied-values", .TargetRef = finalRef, .PropertyName = "UsedValueHash",
                .Operator = "equals", .ExpectedValue = New JValue(sourceHash), .Required = True
            })
            Dim renameResult = ExecuteBatch(application, toolId, renameBatch, $"已验证复制工作表 {sourceName} → {newName}")
            If Not renameResult.Success Then Return MarkCompositeMutationFailure(renameResult, copiedRef)
            Return renameResult
        End Function

        Private Shared Function ExecuteInsertDeleteRowColumn(application As Object,
                                                              params As JObject,
                                                              insert As Boolean) As ToolResult
            Dim toolId = If(insert, "InsertRowCol", "DeleteRowCol")
            Dim kind = FirstText(params, "type").Trim().ToLowerInvariant()
            Dim position = FirstText(params, "position").Trim()
            Dim count = Math.Max(1, If(params("count")?.Value(Of Integer?)(), 1))
            If Not {"row", "column"}.Contains(kind) OrElse String.IsNullOrWhiteSpace(position) Then Return Invalid(toolId, $"{toolId} requires type row/column and position")
            Dim sheetName = GetActiveSheetName(application)
            Dim beforeMatrix = CaptureWorksheetValueMatrix(application, sheetName)
            Dim targetAddress As String
            If kind = "row" Then
                Dim rowNumber As Integer
                If Not Integer.TryParse(position, rowNumber) OrElse rowNumber < 1 Then Return Invalid(toolId, "row position must be a positive integer")
                targetAddress = rowNumber.ToString(CultureInfo.InvariantCulture) & ":" & (rowNumber + count - 1).ToString(CultureInfo.InvariantCulture)
            Else
                targetAddress = position & ":" & GetColumnOffset(position, count - 1)
            End If
            Dim target = ResolveRange(application, QuoteSheet(sheetName) & "!" & targetAddress, captureValues:=False)
            Dim batch = NewBatch()
            Dim arguments As New JObject()
            If kind = "row" Then
                arguments("Shift") = If(insert, CInt(XlInsertShiftDirection.xlShiftDown), CInt(XlDeleteShiftDirection.xlShiftUp))
            Else
                arguments("Shift") = If(insert, CInt(XlInsertShiftDirection.xlShiftToRight), CInt(XlDeleteShiftDirection.xlShiftToLeft))
            End If
            AddOperation(batch,
                         If(insert, "insert-row-column", "delete-row-column"),
                         target.RangeRef,
                         If(insert, "create", "delete"),
                         "Range",
                         If(insert, "Insert", "Delete"),
                         "method",
                         arguments,
                         New JObject From {{"exists", True}})
            Dim result = ExecuteBatch(application, toolId, batch, If(insert, "Excel 已插入行列", "Excel 已删除行列"))
            If result.Success AndAlso Not VerifyRowColumnShift(application, sheetName, beforeMatrix, kind, position, count, insert) Then
                Return SemanticFailure(toolId, result, "The worksheet value layout did not match the requested row/column shift")
            End If
            Return result
        End Function

        Private Shared Function ExecuteHideRowColumn(application As Object, params As JObject) As ToolResult
            Const toolId As String = "HideRowCol"
            Dim kind = FirstText(params, "type").Trim().ToLowerInvariant()
            Dim position = FirstText(params, "position").Trim()
            Dim unhide = If(params("unhide")?.Value(Of Boolean?)(), False)
            If Not {"row", "column"}.Contains(kind) OrElse String.IsNullOrWhiteSpace(position) Then Return Invalid(toolId, "HideRowCol requires type row/column and position")
            Dim address = If(kind = "row", position & ":" & position, position & ":" & position)
            Dim descriptor = ResolveRange(application, QuoteSheet(GetActiveSheetName(application)) & "!" & address, captureValues:=False)
            Dim hideTargetRef = descriptor.RangeRef & If(kind = "row", "/entirerow", "/entirecolumn")
            Dim batch = NewBatch()
            AddOperation(batch, "set-hidden", hideTargetRef, "set", "Range", "Hidden", "property",
                         New JObject From {{"value", Not unhide}}, New JObject From {{"Hidden", Not unhide}})
            Return ExecuteBatch(application, toolId, batch, If(unhide, "已验证显示行列", "已验证隐藏行列"))
        End Function

        Private Shared Function ExecuteProtectSheet(application As Object, params As JObject) As ToolResult
            Const toolId As String = "ProtectSheet"
            Dim sheetName = FirstText(params, "sheetName")
            If String.IsNullOrWhiteSpace(sheetName) Then sheetName = GetActiveSheetName(application)
            ValidateWorksheetName(application, sheetName, mustExist:=True)
            Dim unprotect = If(params("unprotect")?.Value(Of Boolean?)(), False)
            Dim arguments As New JObject()
            If params("password") IsNot Nothing Then arguments("Password") = params("password").DeepClone()
            Dim batch = NewBatch()
            AddOperation(batch,
                         If(unprotect, "unprotect-sheet", "protect-sheet"),
                         BuildWorksheetRef(sheetName),
                         "invoke",
                         "Worksheet",
                         If(unprotect, "Unprotect", "Protect"),
                         "method",
                         arguments,
                         New JObject From {{"ProtectContents", Not unprotect}})
            Return ExecuteBatch(application, toolId, batch, If(unprotect, "已验证取消工作表保护", "已验证工作表保护"))
        End Function

        Private Shared Function CaptureWorksheetUsedValueHash(application As Object, sheetName As String) As String
            Dim descriptor = ResolveUsedRange(application, sheetName)
            Return ExcelOperationObserver.ComputeValueHash(descriptor.Values)
        End Function

        Private Shared Function CaptureWorksheetValueMatrix(application As Object, sheetName As String) As WorksheetMatrix
            Dim descriptor = ResolveUsedRange(application, sheetName)
            Return New WorksheetMatrix With {
                .StartRow = descriptor.StartRow,
                .StartColumn = descriptor.StartColumn,
                .Values = JToken.FromObject(descriptor.Values)
            }
        End Function

        Private Shared Function ResolveUsedRange(application As Object, sheetName As String) As RangeDescriptor
            Dim worksheet As Object = Nothing
            Dim usedRange As Object = Nothing
            Try
                worksheet = application.ActiveWorkbook.Worksheets(sheetName)
                usedRange = worksheet.UsedRange
                Dim address = CStr(usedRange.Address(False, False))
                Dim descriptor = ResolveRange(application, QuoteSheet(sheetName) & "!" & address, captureValues:=True)
                descriptor.StartRow = CInt(usedRange.Row)
                descriptor.StartColumn = CInt(usedRange.Column)
                Return descriptor
            Finally
                ReleaseCom(usedRange)
                ReleaseCom(worksheet)
            End Try
        End Function

        Private Shared Function VerifyRowColumnShift(application As Object,
                                                     sheetName As String,
                                                     before As WorksheetMatrix,
                                                     kind As String,
                                                     position As String,
                                                     count As Integer,
                                                     insert As Boolean) As Boolean
            Dim expected = TransformWorksheetMatrix(before, kind, position, count, insert)
            Dim actual = CaptureWorksheetValueMatrix(application, sheetName)
            Return String.Equals(HashNormalizedMatrix(expected.Values), HashNormalizedMatrix(actual.Values), StringComparison.Ordinal)
        End Function

        Private Shared Function TransformWorksheetMatrix(before As WorksheetMatrix,
                                                         kind As String,
                                                         position As String,
                                                         count As Integer,
                                                         insert As Boolean) As WorksheetMatrix
            Dim rows = NormalizeMatrix(before.Values)
            If rows.Count = 0 Then Return New WorksheetMatrix With {.Values = rows}
            Dim columnCount = DirectCast(rows(0), JArray).Count
            If kind = "row" Then
                Dim absoluteRow = Integer.Parse(position, CultureInfo.InvariantCulture)
                Dim index = Math.Max(0, Math.Min(rows.Count, absoluteRow - before.StartRow))
                If insert Then
                    For offset = 1 To count
                        rows.Insert(index, New JArray(Enumerable.Range(1, columnCount).Select(Function(ignored) JValue.CreateNull())))
                    Next
                Else
                    For offset = 1 To count
                        If index < rows.Count Then rows.RemoveAt(index)
                    Next
                End If
            Else
                Dim absoluteColumn = ColumnNumber(position)
                Dim index = Math.Max(0, Math.Min(columnCount, absoluteColumn - before.StartColumn))
                For Each row In rows.OfType(Of JArray)()
                    If insert Then
                        For offset = 1 To count
                            row.Insert(index, JValue.CreateNull())
                        Next
                    Else
                        For offset = 1 To count
                            If index < row.Count Then row.RemoveAt(index)
                        Next
                    End If
                Next
            End If
            Return New WorksheetMatrix With {.Values = rows}
        End Function

        Private Shared Function NormalizeMatrix(value As JToken) As JArray
            Dim rows As JArray
            If value Is Nothing OrElse value.Type = JTokenType.Null Then
                rows = New JArray()
            ElseIf value.Type = JTokenType.Array AndAlso DirectCast(value, JArray).All(Function(item) item.Type = JTokenType.Array) Then
                rows = DirectCast(value.DeepClone(), JArray)
            Else
                rows = New JArray(New JArray(value.DeepClone()))
            End If
            While rows.Count > 0 AndAlso DirectCast(rows(rows.Count - 1), JArray).All(AddressOf IsBlankToken)
                rows.RemoveAt(rows.Count - 1)
            End While
            If rows.Count = 0 Then Return rows
            Dim maxColumns = rows.OfType(Of JArray)().Max(Function(row) row.Count)
            While maxColumns > 0 AndAlso rows.OfType(Of JArray)().All(Function(row) row.Count < maxColumns OrElse IsBlankToken(row(maxColumns - 1)))
                maxColumns -= 1
            End While
            For Each row In rows.OfType(Of JArray)()
                While row.Count > maxColumns
                    row.RemoveAt(row.Count - 1)
                End While
            Next
            Return rows
        End Function

        Private Shared Function HashNormalizedMatrix(value As JToken) As String
            Return ExcelOperationObserver.ComputeValueHash(NormalizeMatrix(value))
        End Function

        Private Shared Function IsBlankToken(value As JToken) As Boolean
            Return value Is Nothing OrElse value.Type = JTokenType.Null OrElse String.IsNullOrEmpty(value.ToString())
        End Function

        Private Shared Function GetWorksheetCount(application As Object) As Integer
            Dim worksheets As Object = Nothing
            Try
                worksheets = application.ActiveWorkbook.Worksheets
                Return CInt(worksheets.Count)
            Finally
                ReleaseCom(worksheets)
            End Try
        End Function

        Private Shared Function GetActiveSheetName(application As Object) As String
            Dim worksheet As Object = Nothing
            Try
                worksheet = application.ActiveSheet
                If worksheet Is Nothing Then Throw New FormatException("No active worksheet")
                Return CStr(worksheet.Name)
            Finally
                ReleaseCom(worksheet)
            End Try
        End Function

        Private Shared Sub ValidateWorksheetName(application As Object, name As String, mustExist As Boolean)
            If String.IsNullOrWhiteSpace(name) Then Throw New FormatException("Worksheet name is required")
            If name.Length > 31 OrElse name.IndexOfAny(New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}) >= 0 Then Throw New FormatException($"Invalid worksheet name '{name}'")
            Dim exists = WorksheetExists(application, name)
            If mustExist AndAlso Not exists Then Throw New FormatException($"Worksheet '{name}' does not exist")
            If Not mustExist AndAlso exists Then Throw New FormatException($"Worksheet '{name}' already exists")
        End Sub

        Private Shared Function WorksheetExists(application As Object, name As String) As Boolean
            Dim worksheet As Object = Nothing
            Try
                worksheet = application.ActiveWorkbook.Worksheets(name)
                Return worksheet IsNot Nothing
            Catch
                Return False
            Finally
                ReleaseCom(worksheet)
            End Try
        End Function

        Private Shared Sub CleanupCreatedSheet(application As Object, createdRef As String)
            If String.IsNullOrWhiteSpace(createdRef) Then Return
            Dim batch = NewBatch()
            AddOperation(batch, "cleanup-sheet", createdRef, "delete", "Worksheet", "Delete", "method", New JObject(), New JObject From {{"exists", False}})
            Dim alerts = GetDisplayAlerts(application)
            Try
                SetDisplayAlerts(application, False)
                ExcelOperationExecutor.Execute(application, New JObject From {{"batch", JObject.FromObject(batch)}})
            Catch
            Finally
                SetDisplayAlerts(application, alerts)
            End Try
        End Sub

        Private Shared Function GetDisplayAlerts(application As Object) As Boolean
            Try
                Return CBool(application.DisplayAlerts)
            Catch
                Return True
            End Try
        End Function

        Private Shared Sub SetDisplayAlerts(application As Object, value As Boolean)
            Try
                application.DisplayAlerts = value
            Catch
            End Try
        End Sub

        Private Class WorksheetMatrix
            Public Property StartRow As Integer
            Public Property StartColumn As Integer
            Public Property Values As JToken
        End Class
    End Class

End Namespace
