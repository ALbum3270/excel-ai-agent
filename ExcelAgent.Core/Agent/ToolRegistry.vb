Imports System.IO
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Enum ToolRegistrationProvenance
        Custom = 0
        HostManifest = 1
        AgentInternal = 2
    End Enum

    Public Class ToolDescriptor
        Private _registrationProvenance As ToolRegistrationProvenance = ToolRegistrationProvenance.Custom
        Private _registrationOwnerId As String = ""

        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property AppType As String
        Public Property Category As String
        Public Property RiskLevel As String = "safe"
        Public Property AccessMode As String = "write"
        Public Property OutcomeEffects As New List(Of String)()
        Public Property AvailabilityStatus As String = "available"
        Public Property LastError As String = ""
        Public Property Parameters As New List(Of ToolParam)()

        <JsonIgnore>
        Public ReadOnly Property RegistrationProvenance As ToolRegistrationProvenance
            Get
                Return _registrationProvenance
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property RegistrationOwnerId As String
            Get
                Return _registrationOwnerId
            End Get
        End Property

        Friend Sub AssignRegistrationTrust(provenance As ToolRegistrationProvenance, ownerId As String)
            _registrationProvenance = provenance
            _registrationOwnerId = If(ownerId, "").Trim()
        End Sub
    End Class

    Public Class ToolParam
        Public Property Name As String
        Public Property Type As String
        Public Property Required As Boolean
        Public Property Description As String
        Public Property DefaultValue As Object
    End Class

    Public Class AgentVisualEvidence
        Public Property MimeType As String
        Public Property DataUrl As String
        Public Property Source As String
        Public Property ItemIndex As Integer
        Public Property Width As Integer
        Public Property Height As Integer
        Public Property ByteLength As Integer
    End Class

    Public Class ToolResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Data As Object
        Public Property ToolId As String
        Public Property ElapsedMs As Long
        Public Property Observation As Object
        <JsonIgnore>
        Public Property VisualEvidence As New List(Of AgentVisualEvidence)()
        Public Property UndoPointId As String = ""
        Public Property Artifacts As Object
        Public Property ErrorCode As String = ""
        Public Property UserMessage As String = ""
        Public Property DebugDetail As String = ""
        Public Property Recoverable As Boolean = True
        Public Property Retryable As Boolean
        Public Property TaskFatal As Boolean
        Public Property SessionFatal As Boolean

        Public Shared Function Succeed(toolId As String,
                                       Optional message As String = "",
                                       Optional data As Object = Nothing,
                                       Optional observation As Object = Nothing,
                                       Optional undoPointId As String = "",
                                       Optional artifacts As Object = Nothing) As ToolResult
            Return New ToolResult With {
                .Success = True,
                .ToolId = toolId,
                .Message = message,
                .UserMessage = message,
                .Data = data,
                .Observation = observation,
                .UndoPointId = undoPointId,
                .Artifacts = artifacts
            }
        End Function

        Public Shared Function Failed(toolId As String,
                                      message As String,
                                      Optional data As Object = Nothing,
                                      Optional errorCode As String = Nothing,
                                      Optional userMessage As String = Nothing,
                                      Optional debugDetail As String = Nothing,
                                      Optional recoverable As Boolean = True,
                                      Optional observation As Object = Nothing,
                                      Optional artifacts As Object = Nothing,
                                      Optional taskFatal As Boolean = False,
                                      Optional sessionFatal As Boolean = False,
                                      Optional retryable As Boolean = False) As ToolResult
            Dim code = If(String.IsNullOrWhiteSpace(errorCode), ExceptionClassifier.CodeUnknown, errorCode)
            Return New ToolResult With {
                .Success = False,
                .ToolId = toolId,
                .Message = message,
                .UserMessage = If(String.IsNullOrWhiteSpace(userMessage), message, userMessage),
                .DebugDetail = AppLogger.Redact(If(String.IsNullOrWhiteSpace(debugDetail), message, debugDetail)),
                .Data = data,
                .Observation = observation,
                .Artifacts = artifacts,
                .ErrorCode = code,
                .Recoverable = recoverable,
                .Retryable = If(taskFatal OrElse sessionFatal, False, retryable),
                .TaskFatal = taskFatal OrElse sessionFatal,
                .SessionFatal = sessionFatal
            }
        End Function

        Public Shared Function FromException(toolId As String, ex As Exception, Optional data As Object = Nothing) As ToolResult
            Dim classified = ExceptionClassifier.Classify(ex)
            Return Failed(toolId,
                          classified.DebugDetail,
                          data,
                          classified.ErrorCode,
                          classified.UserMessage,
                          classified.DebugDetail,
                          classified.Recoverable,
                          taskFatal:=classified.TaskFatal,
                          sessionFatal:=classified.SessionFatal,
                          retryable:=classified.Retryable)
        End Function

        Public Function ToObserveSummary() As String
            If Not Success Then
                Return $"[{If(ErrorCode, ExceptionClassifier.CodeUnknown)}] retryable={Retryable} taskFatal={TaskFatal} sessionFatal={SessionFatal}: {If(UserMessage, Message)}"
            End If
            If Observation IsNot Nothing Then
                Try
                    Dim token = TryCast(Observation, JToken)
                    If token IsNot Nothing AndAlso token("summary") IsNot Nothing Then Return token("summary").ToString()
                    Dim propertyInfo = Observation.GetType().GetProperty("summary")
                    If propertyInfo Is Nothing Then propertyInfo = Observation.GetType().GetProperty("Summary")
                    If propertyInfo IsNot Nothing Then
                        Dim value = propertyInfo.GetValue(Observation, Nothing)
                        If value IsNot Nothing Then Return value.ToString()
                    End If
                Catch
                End Try
            End If
            Return If(String.IsNullOrWhiteSpace(Message), "ok", Message)
        End Function
    End Class

    Public Class ToolCall
        Public Property ToolId As String
        Public Property Parameters As JObject
        Public Property RequiresApproval As Boolean
    End Class

    ''' <summary>
    ''' Excel-only tool registry. The model may call registered manifests, controlled Python,
    ''' or the Excel host executor. Arbitrary VBA, MCP, memory tools, and Skill scripts do not
    ''' exist in this runtime.
    ''' </summary>
    Public Class ToolRegistry
        Private Const HostOwnerId As String = "excel-host"
        Private ReadOnly _tools As New Dictionary(Of String, ToolDescriptor)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _safetyGate As New Execution.SafetyGate()

        Public Property ExecuteCodeWithToolResult As Func(Of String, String, Boolean, ToolResult)

        Public Sub New()
        End Sub

        Public Sub LoadFromDirectory(directoryPath As String)
            If Not Directory.Exists(directoryPath) Then Return
            For Each filePath In Directory.GetFiles(directoryPath, "*.json", SearchOption.AllDirectories)
                Try
                    Dim tool = JsonConvert.DeserializeObject(Of ToolDescriptor)(File.ReadAllText(filePath))
                    If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.Id) Then Continue For
                    If String.Equals(tool.Id, "ExecuteVBA", StringComparison.OrdinalIgnoreCase) Then Continue For
                    RegisterTrustedTool(tool, ToolRegistrationProvenance.HostManifest, HostOwnerId)
                Catch ex As Exception
                    AppLogger.Warn("ToolRegistry", "Cannot load tool manifest: " & ex.Message)
                End Try
            Next
        End Sub

        Public Function LoadFromRuntimeDirectories(Optional preferredBaseDirectory As String = Nothing) As Integer
            Dim before = ToolCount
            Dim candidates As New List(Of String)()
            For Each root In {preferredBaseDirectory, Path.GetDirectoryName(GetType(ToolRegistry).Assembly.Location), AppDomain.CurrentDomain.BaseDirectory}
                If String.IsNullOrWhiteSpace(root) Then Continue For
                Dim current = root
                While Not String.IsNullOrWhiteSpace(current)
                    AddCandidate(candidates, Path.Combine(current, "Tools"))
                    AddCandidate(candidates, Path.Combine(current, "ExcelAgent.Core", "Tools"))
                    current = Path.GetDirectoryName(current)
                End While
            Next
            For Each candidate In candidates
                LoadFromDirectory(candidate)
            Next
            Return Math.Max(0, ToolCount - before)
        End Function

        Private Shared Sub AddCandidate(candidates As List(Of String), candidate As String)
            Dim fullPath = candidate
            Try
                fullPath = Path.GetFullPath(candidate)
            Catch
            End Try
            If Not candidates.Any(Function(item) String.Equals(item, fullPath, StringComparison.OrdinalIgnoreCase)) Then candidates.Add(fullPath)
        End Sub

        Public Sub RegisterTool(tool As ToolDescriptor)
            RegisterOrMerge(tool, ToolRegistrationProvenance.Custom, "")
        End Sub

        Private Sub RegisterTrustedTool(tool As ToolDescriptor, provenance As ToolRegistrationProvenance, ownerId As String)
            RegisterOrMerge(tool, provenance, ownerId)
        End Sub

        Private Sub RegisterOrMerge(tool As ToolDescriptor, provenance As ToolRegistrationProvenance, ownerId As String)
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.Id) Then Return
            Dim existing As ToolDescriptor = Nothing
            If Not _tools.TryGetValue(tool.Id, existing) Then
                tool.AssignRegistrationTrust(provenance, ownerId)
                _tools(tool.Id) = tool
                Return
            End If
            If provenance = ToolRegistrationProvenance.Custom OrElse
               existing.RegistrationProvenance <> provenance OrElse
               Not String.Equals(existing.RegistrationOwnerId, ownerId, StringComparison.OrdinalIgnoreCase) Then Return

            existing.AppType = MergeValues(existing.AppType, tool.AppType)
            If String.IsNullOrWhiteSpace(existing.Name) Then existing.Name = tool.Name
            If String.IsNullOrWhiteSpace(existing.Description) Then existing.Description = tool.Description
            If existing.Parameters Is Nothing OrElse existing.Parameters.Count = 0 Then existing.Parameters = tool.Parameters
            For Each effect In If(tool.OutcomeEffects, New List(Of String)())
                If Not existing.OutcomeEffects.Contains(effect, StringComparer.OrdinalIgnoreCase) Then existing.OutcomeEffects.Add(effect)
            Next
        End Sub

        Private Shared Function MergeValues(left As String, right As String) As String
            Dim values = ($"{If(left, "")},{If(right, "")}").Split(","c).
                Select(Function(item) item.Trim()).Where(Function(item) item.Length > 0).
                Distinct(StringComparer.OrdinalIgnoreCase)
            Return String.Join(",", values)
        End Function

        Public Function GetAvailableTools(appType As String) As List(Of ToolDescriptor)
            Return _tools.Values.Where(Function(tool) SupportsApp(tool, appType)).ToList()
        End Function

        Public Function GetVisibleTools(appType As String, executionContext As ToolExecutionContext) As List(Of ToolDescriptor)
            Dim tools = GetAvailableTools(appType)
            If executionContext Is Nothing OrElse Not executionContext.HasPrimarySkillGate() Then Return tools
            Return tools.Where(Function(tool) executionContext.IsToolAllowed(tool)).ToList()
        End Function

        Private Shared Function SupportsApp(tool As ToolDescriptor, appType As String) As Boolean
            If tool Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(tool.AppType) Then Return True
            Dim requested = NormalizeAppType(appType)
            Return tool.AppType.Split({","c, ";"c, "|"c}, StringSplitOptions.RemoveEmptyEntries).
                Any(Function(item)
                        Dim value = item.Trim()
                        Return String.Equals(value, "common", StringComparison.OrdinalIgnoreCase) OrElse
                               String.Equals(NormalizeAppType(value), requested, StringComparison.OrdinalIgnoreCase)
                    End Function)
        End Function

        Private Shared Function NormalizeAppType(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "xls", "xlsx", "excel"
                    Return "excel"
                Case Else
                    Return If(value, "").Trim().ToLowerInvariant()
            End Select
        End Function

        Public Function GetTool(toolId As String) As ToolDescriptor
            Dim result As ToolDescriptor = Nothing
            If Not String.IsNullOrWhiteSpace(toolId) Then _tools.TryGetValue(toolId, result)
            Return result
        End Function

        Public Function TryNormalizeToolCall(appType As String, toolCall As ToolCall, ByRef message As String) As Boolean
            Dim ignored = ""
            Return TryNormalizeToolCall(appType, toolCall, Nothing, message, ignored)
        End Function

        Public Function TryNormalizeToolCall(appType As String,
                                             toolCall As ToolCall,
                                             executionContext As ToolExecutionContext,
                                             ByRef message As String) As Boolean
            Dim ignored = ""
            Return TryNormalizeToolCall(appType, toolCall, executionContext, message, ignored)
        End Function

        Public Function TryNormalizeToolCall(appType As String,
                                             toolCall As ToolCall,
                                             executionContext As ToolExecutionContext,
                                             ByRef message As String,
                                             ByRef errorCode As String) As Boolean
            message = ""
            errorCode = ExceptionClassifier.CodeNotFound
            If toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(toolCall.ToolId) Then
                message = "工具调用为空"
                Return False
            End If
            If toolCall.Parameters Is Nothing Then toolCall.Parameters = New JObject()

            Dim original = toolCall.ToolId.Trim()
            Dim direct = GetTool(original)
            If direct Is Nothing Then
                Dim key = NormalizeToolKey(original)
                Dim matches = GetVisibleTools(appType, executionContext).
                    Where(Function(tool) NormalizeToolKey(tool.Id) = key OrElse NormalizeToolKey(tool.Name) = key).ToList()
                If matches.Count = 1 Then direct = matches(0)
            End If
            If direct Is Nothing OrElse Not SupportsApp(direct, appType) Then
                message = "未找到 Excel 工具: " & original
                Return False
            End If
            If executionContext IsNot Nothing AndAlso Not executionContext.IsToolAllowed(direct) Then
                errorCode = ExceptionClassifier.CodeToolNotAllowed
                message = "工具未获得当前任务授权: " & direct.Id
                Return False
            End If

            toolCall.ToolId = direct.Id
            If String.Equals(direct.Id, "CreateSheet", StringComparison.OrdinalIgnoreCase) AndAlso
               toolCall.Parameters("name") Is Nothing AndAlso toolCall.Parameters("sheetName") IsNot Nothing Then
                toolCall.Parameters("name") = toolCall.Parameters("sheetName").DeepClone()
            End If
            errorCode = ""
            Return True
        End Function

        Private Shared Function NormalizeToolKey(value As String) As String
            Dim sb As New StringBuilder()
            For Each character In If(value, "")
                If Char.IsLetterOrDigit(character) Then sb.Append(Char.ToLowerInvariant(character))
            Next
            Return sb.ToString()
        End Function

        Public Async Function ExecuteToolAsync(toolId As String,
                                               params As JObject,
                                               Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResult)
            Return Await ExecuteToolAsync(Nothing, toolId, params, cancellationToken)
        End Function

        Public Async Function ExecuteToolAsync(executionContext As ToolExecutionContext,
                                               toolId As String,
                                               params As JObject,
                                               Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResult)
            Dim stopwatch = Diagnostics.Stopwatch.StartNew()
            cancellationToken.ThrowIfCancellationRequested()
            If params Is Nothing Then params = New JObject()
            Dim tool = GetTool(toolId)
            If tool Is Nothing Then Return ToolResult.Failed(toolId, "未找到工具: " & toolId, errorCode:=ExceptionClassifier.CodeNotFound)

            If executionContext IsNot Nothing AndAlso Not SupportsApp(tool, executionContext.AppType) Then
                Return ToolResult.Failed(toolId, "工具不支持当前 Excel 宿主", errorCode:=ExceptionClassifier.CodeHostUnsupported, recoverable:=False)
            End If
            If executionContext IsNot Nothing AndAlso Not executionContext.IsToolAllowed(tool) Then
                Return ToolResult.Failed(toolId, "工具未获得当前任务授权", errorCode:=ExceptionClassifier.CodeToolNotAllowed)
            End If
            If executionContext IsNot Nothing AndAlso
               String.Equals(tool.Id, "OfficeObjectOperation", StringComparison.OrdinalIgnoreCase) AndAlso
               Not executionContext.IsOfficeObjectOperationReady() Then
                Return ToolResult.Failed(toolId,
                                         "请先调用 DiscoverOfficeCapability，再执行长尾 Office 对象操作",
                                         errorCode:=ExceptionClassifier.CodeToolNotAllowed)
            End If

            Dim decision = _safetyGate.Evaluate(tool, params)
            If decision.Action = Execution.SafetyAction.RequireApproval AndAlso
               executionContext IsNot Nothing AndAlso executionContext.ConsumeToolApproval(tool.Id, params) Then
                decision = Execution.SafetyDecision.Allow(decision.RiskLevel)
            End If
            If decision.Action <> Execution.SafetyAction.Allow Then Return BuildSafetyFailure(tool, decision)

            If String.Equals(tool.Id, "PythonCompute", StringComparison.OrdinalIgnoreCase) Then
                Dim compute = Await Services.Python.PythonComputeService.ExecuteAsync(params, cancellationToken)
                If compute.Success Then
                    Dim result = ToolResult.Succeed(tool.Id,
                                                    "Python 计算完成",
                                                    compute.Data,
                                                    New JObject From {
                                                        {"kind", "compute"},
                                                        {"summary", "Python 计算完成；工作簿尚未修改"},
                                                        {"changed", False}
                                                    })
                    result.ElapsedMs = compute.ElapsedMs
                    Return result
                End If
                Return ToolResult.Failed(tool.Id,
                                         compute.ErrorMessage,
                                         errorCode:=compute.ErrorCode,
                                         debugDetail:=compute.DebugDetail,
                                         observation:=New JObject From {{"kind", "compute"}, {"changed", False}})
            End If

            If ExecuteCodeWithToolResult Is Nothing Then
                Return ToolResult.Failed(tool.Id, "Excel 宿主执行器未初始化", recoverable:=False)
            End If
            Try
                Dim command = New JObject From {{"command", tool.Id}, {"params", params}}.ToString(Formatting.None)
                Dim result = ExecuteCodeWithToolResult.Invoke(command, "json", False)
                stopwatch.Stop()
                If result Is Nothing Then Return ToolResult.Failed(tool.Id, "Excel 宿主未返回执行结果")
                If String.IsNullOrWhiteSpace(result.ToolId) Then result.ToolId = tool.Id
                If result.ElapsedMs <= 0 Then result.ElapsedMs = stopwatch.ElapsedMilliseconds
                If result.Success AndAlso executionContext IsNot Nothing Then executionContext.RecordSuccessfulTool(tool.Id)
                Return result
            Catch ex As Exception
                stopwatch.Stop()
                Return ToolResult.FromException(tool.Id, ex, New With {.elapsedMs = stopwatch.ElapsedMilliseconds})
            End Try
        End Function

        Private Shared Function BuildSafetyFailure(tool As ToolDescriptor, decision As Execution.SafetyDecision) As ToolResult
            Dim code = If(String.IsNullOrWhiteSpace(decision.ErrorCode), ExceptionClassifier.CodeSafetyBlocked, decision.ErrorCode)
            Dim message = If(String.IsNullOrWhiteSpace(decision.UserMessage), decision.Reason, decision.UserMessage)
            Return ToolResult.Failed(tool.Id,
                                     message,
                                     errorCode:=code,
                                     debugDetail:=decision.Reason,
                                     recoverable:=String.Equals(code, ExceptionClassifier.CodeOperationSchemaInvalid, StringComparison.OrdinalIgnoreCase))
        End Function

        Public ReadOnly Property ToolCount As Integer
            Get
                Return _tools.Count
            End Get
        End Property

        Public Sub Clear()
            _tools.Clear()
        End Sub
    End Class
End Namespace
