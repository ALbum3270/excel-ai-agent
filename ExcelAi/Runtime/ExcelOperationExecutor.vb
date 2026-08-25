Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent
Imports ExcelAgent.Core.Agent.OfficeOperations

Namespace OfficeRuntime

    ''' <summary>
    ''' Executes only members returned by ExcelApiCatalogProvider, then verifies the
    ''' declared effects against live Excel state before returning success.
    ''' </summary>
    Friend NotInheritable Class ExcelOperationExecutor
        Private Sub New()
        End Sub

        Public Shared Function Execute(application As Object, params As JObject) As ToolResult
            Const toolId As String = "OfficeObjectOperation"
            Dim batch As OfficeOperationBatch = Nothing
            Try
                Dim batchToken = params?("batch")
                If batchToken Is Nothing OrElse batchToken.Type <> JTokenType.Object Then
                    Return ToolResult.Failed(toolId,
                                             "OfficeObjectOperation requires a batch object",
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:="Excel 声明式操作缺少 batch",
                                             recoverable:=True)
                End If

                batch = batchToken.ToObject(Of OfficeOperationBatch)()
                Dim validation = OfficeOperationValidation.ValidateBatch(batch)
                If Not validation.IsValid Then
                    Return ToolResult.Failed(toolId,
                                             validation.ToErrorMessage(),
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:="Excel 声明式操作未通过结构校验",
                                             recoverable:=True)
                End If
                If Not String.Equals(OfficeObjectRef.NormalizeAppType(batch.AppType), "Excel", StringComparison.OrdinalIgnoreCase) Then
                    Return ToolResult.Failed(toolId,
                                             $"Unsupported batch appType {batch.AppType}",
                                             errorCode:=ExceptionClassifier.CodeHostUnsupported,
                                             userMessage:="当前 Excel 宿主不能执行其他 Office 应用的操作",
                                             recoverable:=False)
                End If

                Dim beforeState = ExcelOperationObserver.CaptureState(application, batch)
                Dim operationResults As New JArray()
                Dim targetRefs As New List(Of String)()
                Dim warnings As New List(Of String)()
                Dim succeededCount As Integer = 0

                For Each operation In batch.Operations
                    Try
                        Dim execution = ExecuteOperation(application, operation)
                        succeededCount += 1
                        targetRefs.Add(operation.TargetRef)
                        If Not String.IsNullOrWhiteSpace(execution.ResultRef) Then targetRefs.Add(execution.ResultRef)
                        operationResults.Add(New JObject From {
                            {"id", operation.Id},
                            {"status", "succeeded"},
                            {"memberId", operation.MemberId},
                            {"targetRef", operation.TargetRef},
                            {"resultRef", If(execution.ResultRef, "")},
                            {"data", If(execution.Data, JValue.CreateNull())}
                        })
                    Catch ex As ExcelOperationException
                        operationResults.Add(BuildFailureResult(operation, ex.ErrorCode, ex.Message))
                        If batch.Atomic AndAlso succeededCount > 0 Then warnings.Add("Atomic batch partially applied; Excel COM rollback was not guaranteed")
                        Return BuildExecutionFailure(application,
                                                     batch,
                                                     operationResults,
                                                     beforeState,
                                                     targetRefs,
                                                     warnings,
                                                     ex.Message,
                                                     If(succeededCount > 0, ExceptionClassifier.CodePartialApply, ex.ErrorCode),
                                                     If(succeededCount > 0, False, ex.Recoverable))
                    Catch ex As Exception
                        Dim baseException = If(TypeOf ex Is TargetInvocationException AndAlso ex.InnerException IsNot Nothing,
                                               ex.InnerException,
                                               ex)
                        Dim classified = ExceptionClassifier.Classify(baseException)
                        operationResults.Add(BuildFailureResult(operation, classified.ErrorCode, classified.UserMessage))
                        If batch.Atomic AndAlso succeededCount > 0 Then warnings.Add("Atomic batch partially applied")
                        Return BuildExecutionFailure(application,
                                                     batch,
                                                     operationResults,
                                                     beforeState,
                                                     targetRefs,
                                                     warnings,
                                                     classified.DebugDetail,
                                                     If(succeededCount > 0, ExceptionClassifier.CodePartialApply, classified.ErrorCode),
                                                     If(succeededCount > 0, False, classified.Recoverable))
                    End Try
                Next

                Dim afterState = ExcelOperationObserver.CaptureState(application, batch, targetRefs)
                Dim observation = ExcelOperationObserver.BuildObservation(batch,
                                                                           operationResults,
                                                                           beforeState,
                                                                           afterState,
                                                                           targetRefs,
                                                                           warnings)
                Dim verification = ExcelOperationObserver.VerifyExpectedEffects(batch, operationResults, afterState)
                observation("verification") = verification
                ' An empty verification set means no semantic postcondition was declared. It
                ' must not be upgraded to satisfied=true merely because no check failed.
                observation("satisfied") = verification.Count > 0 AndAlso
                    Not ExcelOperationObserver.HasRequiredVerificationFailure(verification)
                Dim data As New JObject From {
                    {"schemaVersion", batch.SchemaVersion},
                    {"targetRefs", JArray.FromObject(targetRefs.Distinct(StringComparer.OrdinalIgnoreCase))},
                    {"operations", operationResults.DeepClone()}
                }

                Dim captureError = afterState?("captureError")?.ToString()
                If ExcelOperationObserver.HasMutatingOperations(batch) AndAlso Not String.IsNullOrWhiteSpace(captureError) Then
                    Return ToolResult.Failed(toolId,
                                             $"Excel operation may have completed, but observation failed: {captureError}",
                                             data:=data,
                                             errorCode:=ExceptionClassifier.CodeObservationFailed,
                                             userMessage:="Excel 操作可能已经执行，但无法安全验证；已停止自动重试",
                                             recoverable:=False,
                                             observation:=observation)
                End If

                If ExcelOperationObserver.HasRequiredVerificationFailure(verification) Then
                    Return ToolResult.Failed(toolId,
                                             "Excel operation verification failed",
                                             data:=data,
                                             errorCode:=ExceptionClassifier.CodeVerifyFailed,
                                             userMessage:="Excel 已执行操作，但真实结果不符合声明的成功条件；已停止自动重试",
                                             recoverable:=False,
                                             observation:=observation)
                End If
                Return ToolResult.Succeed(toolId, observation("summary").ToString(), data:=data, observation:=observation)
            Catch ex As ExcelOperationException
                Return ToolResult.Failed(toolId,
                                         ex.Message,
                                         errorCode:=ex.ErrorCode,
                                         userMessage:="Excel 对象操作失败",
                                         recoverable:=ex.Recoverable)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            End Try
        End Function

        Private Shared Function ExecuteOperation(application As Object,
                                                 operation As OfficeOperation) As OperationExecutionResult
            Dim capability As OfficeCapabilityMember = Nothing
            Dim reflectedMember As MemberInfo = Nothing
            If Not ExcelApiCatalogProvider.TryGetMemberBinding(operation.MemberId, capability, reflectedMember) Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeCapabilityNotFound,
                                                  $"Catalog member was not found: {operation.MemberId}")
            End If
            If Not capability.Executable Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeMemberNotExecutable,
                                                  If(capability.UnsupportedReason, $"Member is not executable: {operation.MemberId}"),
                                                  recoverable:=False)
            End If

            Using resolved = ExcelObjectResolver.Resolve(application, operation.TargetRef)
                EnsureTargetCompatibility(resolved, reflectedMember)
                Dim boundScopes As New List(Of IDisposable)()
                Dim returnValue As Object = Nothing
                Try
                    If TypeOf reflectedMember Is PropertyInfo Then
                        returnValue = ExecuteProperty(application,
                                                      operation,
                                                      resolved.Value,
                                                      DirectCast(reflectedMember, PropertyInfo),
                                                      boundScopes)
                    ElseIf TypeOf reflectedMember Is MethodInfo Then
                        returnValue = ExecuteMethod(application,
                                                    operation,
                                                    resolved.Value,
                                                    DirectCast(reflectedMember, MethodInfo),
                                                    boundScopes)
                    Else
                        Throw New ExcelOperationException(ExceptionClassifier.CodeMemberNotExecutable,
                                                          "Only catalog properties and methods are executable",
                                                          recoverable:=False)
                    End If

                    Return New OperationExecutionResult With {
                        .ResultRef = BuildResultRef(operation,
                                                    resolved.ObjectKind,
                                                    resolved.Value,
                                                    reflectedMember,
                                                    returnValue),
                        .Data = ConvertReturnValue(returnValue)
                    }
                Finally
                    If returnValue IsNot Nothing AndAlso Marshal.IsComObject(returnValue) AndAlso
                       Not Object.ReferenceEquals(returnValue, resolved.Value) Then ReleaseCom(returnValue)
                    For index = boundScopes.Count - 1 To 0 Step -1
                        boundScopes(index).Dispose()
                    Next
                End Try
            End Using
        End Function

        Private Shared Function ExecuteProperty(application As Object,
                                                operation As OfficeOperation,
                                                target As Object,
                                                prop As PropertyInfo,
                                                boundScopes As List(Of IDisposable)) As Object
            Dim action = operation.Action.Trim().ToLowerInvariant()
            Select Case action
                Case "get", "collection_item"
                    If Not prop.CanRead Then ThrowMemberAction(operation, "Property is not readable")
                    Return prop.GetValue(target, BindParameters(application, prop.GetIndexParameters(), operation.Arguments, boundScopes))
                Case "set"
                    If Not prop.CanWrite Then ThrowMemberAction(operation, "Property is not writable")
                    Dim valueToken = GetArgument(operation.Arguments, "value")
                    If valueToken Is Nothing Then
                        Throw New ExcelOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                          $"Operation {operation.Id} requires arguments.value")
                    End If
                    Dim value = ConvertArgument(application, valueToken, prop.PropertyType, boundScopes)
                    prop.SetValue(target,
                                  value,
                                  BindParameters(application, prop.GetIndexParameters(), operation.Arguments, boundScopes))
                    Return Nothing
                Case Else
                    ThrowMemberAction(operation, $"Action {operation.Action} cannot execute a property")
                    Return Nothing
            End Select
        End Function

        Private Shared Function ExecuteMethod(application As Object,
                                              operation As OfficeOperation,
                                              target As Object,
                                              method As MethodInfo,
                                              boundScopes As List(Of IDisposable)) As Object
            Dim action = operation.Action.Trim().ToLowerInvariant()
            If action <> "invoke" AndAlso action <> "create" AndAlso action <> "delete" AndAlso
               action <> "collection_item" AndAlso action <> "get" Then
                ThrowMemberAction(operation, $"Action {operation.Action} cannot execute a method")
            End If
            Return method.Invoke(target, BindParameters(application, method.GetParameters(), operation.Arguments, boundScopes))
        End Function

        Private Shared Function BindParameters(application As Object,
                                               parameters As ParameterInfo(),
                                               arguments As JObject,
                                               boundScopes As List(Of IDisposable)) As Object()
            If parameters Is Nothing OrElse parameters.Length = 0 Then Return Array.Empty(Of Object)()
            Dim result(parameters.Length - 1) As Object
            For index = 0 To parameters.Length - 1
                Dim parameter = parameters(index)
                Dim token = GetArgument(arguments, parameter.Name)
                If token Is Nothing Then
                    If parameter.IsOptional Then
                        result(index) = Type.Missing
                        Continue For
                    End If
                    Throw New ExcelOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                      $"Missing required argument '{parameter.Name}'")
                End If
                result(index) = ConvertArgument(application, token, parameter.ParameterType, boundScopes)
            Next
            Return result
        End Function

        Private Shared Function ConvertArgument(application As Object,
                                                token As JToken,
                                                expectedType As Type,
                                                boundScopes As List(Of IDisposable)) As Object
            If expectedType.IsByRef Then expectedType = expectedType.GetElementType()
            If token Is Nothing OrElse token.Type = JTokenType.Null Then
                If Not expectedType.IsValueType OrElse Nullable.GetUnderlyingType(expectedType) IsNot Nothing Then Return Nothing
                Throw New ExcelOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                  $"Null is not valid for {expectedType.FullName}")
            End If

            If token.Type = JTokenType.Object Then
                Dim refToken = DirectCast(token, JObject).GetValue("ref", StringComparison.OrdinalIgnoreCase)
                If refToken Is Nothing Then refToken = DirectCast(token, JObject).GetValue("$ref", StringComparison.OrdinalIgnoreCase)
                If refToken IsNot Nothing Then
                    Dim scope = ExcelObjectResolver.Resolve(application, refToken.ToString())
                    boundScopes.Add(scope)
                    Return scope.Value
                End If
            End If

            If expectedType Is GetType(Object) Then Return ConvertUntypedValue(token)
            If expectedType Is GetType(String) Then Return token.ToString()
            If expectedType Is GetType(Boolean) Then Return token.Value(Of Boolean)()
            If expectedType Is GetType(Integer) Then Return token.Value(Of Integer)()
            If expectedType Is GetType(Long) Then Return token.Value(Of Long)()
            If expectedType Is GetType(Short) Then Return token.Value(Of Short)()
            If expectedType Is GetType(Byte) Then Return token.Value(Of Byte)()
            If expectedType Is GetType(Single) Then Return Single.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType Is GetType(Double) Then Return Double.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType Is GetType(Decimal) Then Return Decimal.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType Is GetType(DateTime) Then Return DateTime.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType.IsEnum Then
                If token.Type = JTokenType.Integer Then Return [Enum].ToObject(expectedType, token.Value(Of Integer)())
                Return [Enum].Parse(expectedType, token.ToString(), ignoreCase:=True)
            End If
            If expectedType.IsAssignableFrom(GetType(Object())) AndAlso token.Type = JTokenType.Array Then
                Return DirectCast(token, JArray).Select(Function(item) ConvertUntypedValue(item)).ToArray()
            End If
            Try
                Return token.ToObject(expectedType)
            Catch ex As Exception
                Throw New ExcelOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                  $"Unsupported argument type {expectedType.FullName}: {ex.Message}")
            End Try
        End Function

        Private Shared Function ConvertUntypedValue(token As JToken) As Object
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return Nothing
            If token.Type = JTokenType.Array Then
                Dim rows = DirectCast(token, JArray)
                If rows.Count > 0 AndAlso rows.All(Function(item) item.Type = JTokenType.Array) Then
                    Dim columnCount = rows.OfType(Of JArray)().Max(Function(row) row.Count)
                    Dim matrix(rows.Count - 1, Math.Max(0, columnCount - 1)) As Object
                    For rowIndex = 0 To rows.Count - 1
                        Dim row = DirectCast(rows(rowIndex), JArray)
                        For columnIndex = 0 To row.Count - 1
                            matrix(rowIndex, columnIndex) = ConvertUntypedValue(row(columnIndex))
                        Next
                    Next
                    Return matrix
                End If
                Return rows.Select(Function(item) ConvertUntypedValue(item)).ToArray()
            End If
            If TypeOf token Is JValue Then Return DirectCast(token, JValue).Value
            Return token.ToObject(Of Object)()
        End Function

        Private Shared Function BuildResultRef(operation As OfficeOperation,
                                               targetKind As String,
                                               target As Object,
                                               member As MemberInfo,
                                               returnValue As Object) As String
            Try
                ' A Name set invalidates the old canonical ref. Return the new identity so
                ' the observer verifies the same COM object at its post-operation ref.
                If String.Equals(operation.Action, "set", StringComparison.OrdinalIgnoreCase) AndAlso
                   String.Equals(member?.Name, "Name", StringComparison.OrdinalIgnoreCase) Then
                    Select Case If(targetKind, "").ToLowerInvariant()
                        Case "worksheet"
                            Return "Excel:workbooks/active/worksheets/" & ExcelObjectResolver.EncodeSegment(CStr(target.Name))
                        Case "chartobject", "listobject", "pivottable"
                            Dim slash = operation.TargetRef.LastIndexOf("/"c)
                            If slash > 0 Then Return operation.TargetRef.Substring(0, slash + 1) & ExcelObjectResolver.EncodeSegment(CStr(target.Name))
                    End Select
                End If

                If returnValue Is Nothing OrElse Not Marshal.IsComObject(returnValue) Then Return ""
                If targetKind = "Worksheets" OrElse targetKind = "Sheets" Then
                    Return "Excel:workbooks/active/worksheets/" & ExcelObjectResolver.EncodeSegment(CStr(returnValue.Name))
                End If
                If targetKind = "ChartObjects" OrElse targetKind = "ListObjects" OrElse targetKind = "PivotTables" Then
                    Return operation.TargetRef.TrimEnd("/"c) & "/" & ExcelObjectResolver.EncodeSegment(CStr(returnValue.Name))
                End If

                Dim returnTypeName = ""
                If TypeOf member Is PropertyInfo Then returnTypeName = DirectCast(member, PropertyInfo).PropertyType.Name.TrimStart("_"c)
                If TypeOf member Is MethodInfo Then returnTypeName = DirectCast(member, MethodInfo).ReturnType.Name.TrimStart("_"c)
                If String.Equals(returnTypeName, "Worksheet", StringComparison.OrdinalIgnoreCase) Then
                    Return "Excel:workbooks/active/worksheets/" & ExcelObjectResolver.EncodeSegment(CStr(returnValue.Name))
                End If
                If String.Equals(returnTypeName, "Range", StringComparison.OrdinalIgnoreCase) Then
                    Dim worksheet As Object = Nothing
                    Try
                        worksheet = returnValue.Worksheet
                        Return ExcelObjectResolver.BuildRangeRef(CStr(worksheet.Name), CStr(returnValue.Address(False, False)))
                    Finally
                        ReleaseCom(worksheet)
                    End Try
                End If
            Catch
            End Try
            Return ""
        End Function

        Private Shared Function ConvertReturnValue(value As Object) As JToken
            If value Is Nothing OrElse Marshal.IsComObject(value) Then Return JValue.CreateNull()
            If TypeOf value Is Array Then Return JArray.FromObject(value)
            Try
                Return JToken.FromObject(value)
            Catch
                Return New JValue(value.ToString())
            End Try
        End Function

        Private Shared Sub EnsureTargetCompatibility(resolved As ResolvedExcelObject, member As MemberInfo)
            If resolved Is Nothing OrElse resolved.Value Is Nothing OrElse member?.DeclaringType Is Nothing Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeObjectTypeMismatch, "Operation target is unavailable")
            End If
            If member.DeclaringType.IsInstanceOfType(resolved.Value) Then Return
            Dim expected = member.DeclaringType.Name.TrimStart("_"c)
            If String.Equals(expected, "Sheets", StringComparison.OrdinalIgnoreCase) Then expected = "Worksheets"
            If String.Equals(expected, resolved.ObjectKind, StringComparison.OrdinalIgnoreCase) Then Return
            Throw New ExcelOperationException(ExceptionClassifier.CodeObjectTypeMismatch,
                                              $"Member {member.Name} requires {expected}, target is {resolved.ObjectKind}")
        End Sub

        Private Shared Function GetArgument(arguments As JObject, name As String) As JToken
            If arguments Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return Nothing
            Return arguments.GetValue(name, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Sub ThrowMemberAction(operation As OfficeOperation, message As String)
            Throw New ExcelOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                              $"Operation {operation.Id}: {message}")
        End Sub

        Private Shared Function BuildFailureResult(operation As OfficeOperation,
                                                   errorCode As String,
                                                   message As String) As JObject
            Return New JObject From {
                {"id", operation.Id},
                {"status", "failed"},
                {"memberId", operation.MemberId},
                {"targetRef", operation.TargetRef},
                {"errorCode", errorCode},
                {"message", message}
            }
        End Function

        Private Shared Function BuildExecutionFailure(application As Object,
                                                      batch As OfficeOperationBatch,
                                                      operationResults As JArray,
                                                      beforeState As JObject,
                                                      targetRefs As List(Of String),
                                                      warnings As List(Of String),
                                                      message As String,
                                                      errorCode As String,
                                                      recoverable As Boolean) As ToolResult
            Dim afterFailure = ExcelOperationObserver.CaptureState(application, batch, targetRefs)
            Dim observation = ExcelOperationObserver.BuildObservation(batch,
                                                                       operationResults,
                                                                       beforeState,
                                                                       afterFailure,
                                                                       targetRefs,
                                                                       warnings)
            Return ToolResult.Failed("OfficeObjectOperation",
                                     message,
                                     data:=New JObject From {{"targetRefs", JArray.FromObject(targetRefs)}},
                                     errorCode:=errorCode,
                                     userMessage:="Excel 声明式操作未全部完成",
                                     recoverable:=recoverable,
                                     observation:=observation)
        End Function

        Private Shared Sub ReleaseCom(value As Object)
            If value Is Nothing Then Return
            Try
                If Marshal.IsComObject(value) Then Marshal.ReleaseComObject(value)
            Catch
            End Try
        End Sub

        Private Class OperationExecutionResult
            Public Property ResultRef As String
            Public Property Data As JToken
        End Class
    End Class

End Namespace
