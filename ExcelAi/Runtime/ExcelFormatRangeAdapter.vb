Imports Microsoft.Office.Interop.Excel
Imports System.Linq
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent
Imports ExcelAgent.Core.Agent.OfficeOperations

Namespace OfficeRuntime

    ''' <summary>
    ''' Compatibility adapter from the high-level FormatRange tool to the declarative
    ''' Excel object-operation Module. It contains no business-column vocabulary.
    ''' </summary>
    Friend NotInheritable Class ExcelFormatRangeAdapter
        Private Sub New()
        End Sub

        Public Shared Function Execute(application As Object, params As JObject) As ToolResult
            Const toolId As String = "FormatRange"
            Try
                If application Is Nothing Then
                    Return ToolResult.Failed(toolId,
                                             "Excel application is unavailable",
                                             errorCode:=ExceptionClassifier.CodeDocMissing,
                                             userMessage:="当前没有可用的 Excel 应用",
                                             recoverable:=False)
                End If
                If params Is Nothing Then
                    Return Invalid(toolId, "FormatRange params are required")
                End If

                Dim rangeSpec = If(params("range")?.ToString(), "").Trim()
                If String.IsNullOrWhiteSpace(rangeSpec) Then Return Invalid(toolId, "FormatRange requires range")
                Dim resolvedSpec = ResolveRangeSpec(application, rangeSpec)
                Dim rangeRef = ExcelObjectResolver.BuildRangeRef(resolvedSpec.SheetName, resolvedSpec.Address)
                Dim batch As New OfficeOperationBatch With {
                    .SchemaVersion = OfficeOperationValidation.CurrentSchemaVersion,
                    .AppType = "Excel",
                    .Atomic = True
                }

                AddPropertyOperation(batch, rangeRef, "Range", "NumberFormat", params("numberFormat"))
                AddPropertyOperation(batch, rangeRef, "Range", "HorizontalAlignment",
                                     ConvertHorizontalAlignment(params("horizontalAlignment")))
                AddPropertyOperation(batch, rangeRef, "Range", "VerticalAlignment",
                                     ConvertVerticalAlignment(params("verticalAlignment")))
                AddPropertyOperation(batch, rangeRef, "Range", "WrapText", params("wrapText"))

                Dim fontRef = rangeRef & "/font"
                AddPropertyOperation(batch, fontRef, "Font", "Bold", params("bold"))
                AddPropertyOperation(batch, fontRef, "Font", "Italic", params("italic"))
                AddPropertyOperation(batch, fontRef, "Font", "Size", params("fontSize"))
                AddPropertyOperation(batch, fontRef, "Font", "Color", ConvertColor(params("fontColor")))

                Dim interiorRef = rangeRef & "/interior"
                AddPropertyOperation(batch, interiorRef, "Interior", "Color", ConvertColor(params("backgroundColor")))
                AddBorderOperations(batch, rangeRef, params("borders"))

                If batch.Operations.Count = 0 Then
                    Return Invalid(toolId,
                                   "FormatRange must declare at least one explicit format property; preset style names are not executable contracts")
                End If

                Dim envelope As New JObject From {
                    {"batch", JObject.FromObject(batch)}
                }
                Dim result = ExcelOperationExecutor.Execute(application, envelope)
                result.ToolId = toolId
                Dim observation = TryCast(result.Observation, JObject)
                If observation IsNot Nothing Then
                    AnnotateVerifiedRequestFields(TryCast(observation("verification"), JArray), params)
                    observation("adapter") = toolId
                    observation("requestedRange") = resolvedSpec.SheetName & "!" & resolvedSpec.Address
                    observation("summary") = If(result.Success,
                                                $"已验证 {resolvedSpec.SheetName}!{resolvedSpec.Address} 的明确格式属性",
                                                $"未能验证 {resolvedSpec.SheetName}!{resolvedSpec.Address} 的全部格式属性")
                    result.Message = observation("summary").ToString()
                    result.UserMessage = If(result.Success, result.Message, result.UserMessage)
                End If
                Return result
            Catch ex As FormatException
                Return Invalid(toolId, ex.Message)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            End Try
        End Function

        Private Shared Sub AnnotateVerifiedRequestFields(verification As JArray,
                                                          params As JObject)
            If verification Is Nothing OrElse params Is Nothing Then Return
            For Each item In verification.OfType(Of JObject)()
                Dim targetRef = If(item("targetRef")?.ToString(), "").Trim().ToLowerInvariant()
                Dim propertyName = If(item("property")?.ToString(), "").Trim()
                Dim requestProperty As String = ""
                Select Case propertyName.ToLowerInvariant()
                    Case "numberformat" : requestProperty = "numberFormat"
                    Case "horizontalalignment" : requestProperty = "horizontalAlignment"
                    Case "verticalalignment" : requestProperty = "verticalAlignment"
                    Case "wraptext" : requestProperty = "wrapText"
                    Case "bold" : requestProperty = "bold"
                    Case "italic" : requestProperty = "italic"
                    Case "size" : requestProperty = "fontSize"
                    Case "color"
                        If targetRef.EndsWith("/font", StringComparison.OrdinalIgnoreCase) Then
                            requestProperty = "fontColor"
                        ElseIf targetRef.EndsWith("/interior", StringComparison.OrdinalIgnoreCase) Then
                            requestProperty = "backgroundColor"
                        End If
                    Case "linestyle" : requestProperty = "borders"
                End Select
                If String.IsNullOrWhiteSpace(requestProperty) OrElse params(requestProperty) Is Nothing Then Continue For
                item("requestProperty") = requestProperty
                item("requestExpected") = params(requestProperty).DeepClone()
            Next
        End Sub

        Private Shared Sub AddPropertyOperation(batch As OfficeOperationBatch,
                                                targetRef As String,
                                                typeName As String,
                                                propertyName As String,
                                                value As JToken)
            If value Is Nothing OrElse value.Type = JTokenType.Null Then Return
            Dim memberId = ExcelApiCatalogProvider.FindMemberId(typeName, propertyName, "property")
            If String.IsNullOrWhiteSpace(memberId) Then
                Throw New FormatException($"Excel capability catalog does not expose {typeName}.{propertyName}")
            End If
            Dim operationId = "format-" & (batch.Operations.Count + 1).ToString()
            batch.Operations.Add(New OfficeOperation With {
                .Id = operationId,
                .TargetRef = targetRef,
                .Action = "set",
                .MemberId = memberId,
                .Arguments = New JObject From {{"value", value.DeepClone()}},
                .ExpectedEffects = New JObject From {{propertyName, value.DeepClone()}}
            })
        End Sub

        Private Shared Sub AddBorderOperations(batch As OfficeOperationBatch,
                                               rangeRef As String,
                                               borderToken As JToken)
            If borderToken Is Nothing OrElse borderToken.Type = JTokenType.Null Then Return
            Dim normalized = borderToken.ToString().Trim().ToLowerInvariant()
            Dim lineStyle As Integer
            Select Case normalized
                Case "all", "thin", "true", "全部", "细边框"
                    lineStyle = CInt(XlLineStyle.xlContinuous)
                Case "none", "false", "无", "无边框"
                    lineStyle = CInt(XlLineStyle.xlLineStyleNone)
                Case "outline", "outside", "外框", "外边框"
                    For Each borderIndex In New Integer() {
                        CInt(XlBordersIndex.xlEdgeLeft),
                        CInt(XlBordersIndex.xlEdgeTop),
                        CInt(XlBordersIndex.xlEdgeBottom),
                        CInt(XlBordersIndex.xlEdgeRight)
                    }
                        AddPropertyOperation(batch,
                                             rangeRef & "/borders/" & borderIndex.ToString(),
                                             "Border",
                                             "LineStyle",
                                             New JValue(CInt(XlLineStyle.xlContinuous)))
                    Next
                    Return
                Case Else
                    Throw New FormatException($"Unsupported borders value: {borderToken}")
            End Select
            AddPropertyOperation(batch,
                                 rangeRef & "/borders",
                                 "Borders",
                                 "LineStyle",
                                 New JValue(lineStyle))
        End Sub

        Private Shared Function ConvertColor(token As JToken) As JToken
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return Nothing
            If token.Type = JTokenType.Integer Then Return token.DeepClone()
            Return New JValue(ExcelConditionalFormatContract.ParseColor(token.ToString()))
        End Function

        Private Shared Function ConvertHorizontalAlignment(token As JToken) As JToken
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return Nothing
            If token.Type = JTokenType.Integer Then Return token.DeepClone()
            Select Case Normalize(token.ToString())
                Case "left", "左", "左对齐" : Return New JValue(CInt(XlHAlign.xlHAlignLeft))
                Case "center", "centre", "居中", "水平居中" : Return New JValue(CInt(XlHAlign.xlHAlignCenter))
                Case "right", "右", "右对齐" : Return New JValue(CInt(XlHAlign.xlHAlignRight))
                Case "general", "常规" : Return New JValue(CInt(XlHAlign.xlHAlignGeneral))
                Case Else : Throw New FormatException($"Unsupported horizontalAlignment: {token}")
            End Select
        End Function

        Private Shared Function ConvertVerticalAlignment(token As JToken) As JToken
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return Nothing
            If token.Type = JTokenType.Integer Then Return token.DeepClone()
            Select Case Normalize(token.ToString())
                Case "top", "顶部", "顶端" : Return New JValue(CInt(XlVAlign.xlVAlignTop))
                Case "center", "centre", "middle", "居中", "垂直居中" : Return New JValue(CInt(XlVAlign.xlVAlignCenter))
                Case "bottom", "底部", "底端" : Return New JValue(CInt(XlVAlign.xlVAlignBottom))
                Case Else : Throw New FormatException($"Unsupported verticalAlignment: {token}")
            End Select
        End Function

        Private Shared Function ResolveRangeSpec(application As Object, rangeSpec As String) As ResolvedRangeSpec
            Dim workbook As Object = Nothing
            Dim worksheet As Object = Nothing
            Dim worksheets As Object = Nothing
            Dim usedRange As Object = Nothing
            Dim usedRows As Object = Nothing
            Try
                workbook = application.ActiveWorkbook
                worksheet = application.ActiveSheet
                If workbook Is Nothing OrElse worksheet Is Nothing Then Throw New FormatException("No active workbook or worksheet")

                Dim sheetName = CStr(worksheet.Name)
                Dim address = rangeSpec.Trim()
                Dim bangIndex = address.LastIndexOf("!"c)
                If bangIndex > 0 Then
                    sheetName = address.Substring(0, bangIndex).Trim()
                    If sheetName.Length >= 2 AndAlso sheetName.StartsWith("'", StringComparison.Ordinal) AndAlso sheetName.EndsWith("'", StringComparison.Ordinal) Then
                        sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'")
                    End If
                    address = address.Substring(bangIndex + 1).Trim()
                    worksheets = workbook.Worksheets
                    ComObjectHelper.ReleaseComObject(worksheet)
                    worksheet = worksheets.Item(sheetName)
                End If

                If address.IndexOf("{lastRow}", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    usedRange = worksheet.UsedRange
                    usedRows = usedRange.Rows
                    Dim lastRow = CInt(usedRange.Row) + CInt(usedRows.Count) - 1
                    address = address.Replace("{lastRow}", Math.Max(1, lastRow).ToString())
                End If
                Return New ResolvedRangeSpec With {.SheetName = sheetName, .Address = address}
            Finally
                ComObjectHelper.ReleaseComObject(usedRows)
                ComObjectHelper.ReleaseComObject(usedRange)
                ComObjectHelper.ReleaseComObject(worksheet)
                ComObjectHelper.ReleaseComObject(worksheets)
                ComObjectHelper.ReleaseComObject(workbook)
            End Try
        End Function

        Private Shared Function Normalize(value As String) As String
            Return If(value, "").Trim().Replace(" ", "").Replace("_", "").Replace("-", "").ToLowerInvariant()
        End Function

        Private Shared Function Invalid(toolId As String, message As String) As ToolResult
            Return ToolResult.Failed(toolId,
                                     message,
                                     errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                     userMessage:="格式化参数无效；请使用明确的格式属性",
                                     recoverable:=True)
        End Function

        Private Class ResolvedRangeSpec
            Public Property SheetName As String
            Public Property Address As String
        End Class
    End Class

End Namespace
