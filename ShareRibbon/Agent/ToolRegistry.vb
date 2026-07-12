Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
' StreamJsonRpcMCPClient 和 MCPToolInfo 在根命名空间

Namespace Agent

    ''' <summary>
    ''' 工具描述符 - 描述一个可调用工具的结构
    ''' </summary>
    Public Class ToolDescriptor
        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property AppType As String              ' "excel" / "word" / "powerpoint" / "common"
        Public Property Category As String             ' "基础操作" / "数据操作" / "高级功能"
        Public Property RiskLevel As String = "safe"   ' "safe" / "medium" / "risky"
        Public Property AvailabilityStatus As String = "available" ' "available" / "unavailable" / "error"
        Public Property LastError As String = ""
        Public Property IsVbaFallback As Boolean = False
        Public Property Parameters As New List(Of ToolParam)()
    End Class

    ''' <summary>
    ''' 工具参数描述
    ''' </summary>
    Public Class ToolParam
        Public Property Name As String
        Public Property Type As String              ' "string" / "integer" / "boolean" / "array" / "object"
        Public Property Required As Boolean = False
        Public Property Description As String
        Public Property DefaultValue As Object = Nothing
    End Class

    ''' <summary>
    ''' 工具调用结果（P0-4 统一错误契约）。
    ''' Success/Message 保持兼容；新增 ErrorCode/UserMessage/DebugDetail/Recoverable 供 observe/repair 使用。
    ''' </summary>
    Public Class ToolResult
        Public Property Success As Boolean
        ''' <summary>技术/观察用消息（可含简要错误）；兼容旧调用方。</summary>
        Public Property Message As String
        Public Property Data As Object              ' 执行结果的原始数据
        Public Property ToolId As String
        Public Property ElapsedMs As Long
        ''' <summary>稳定错误码，如 NETWORK_ERROR / COM_ERROR / NOT_FOUND。</summary>
        Public Property ErrorCode As String = ""
        ''' <summary>面向用户的简短说明（已脱敏）。</summary>
        Public Property UserMessage As String = ""
        ''' <summary>调试细节（已脱敏），勿直接展示给用户。</summary>
        Public Property DebugDetail As String = ""
        ''' <summary>是否适合 Agent 自动 repair / 重试。</summary>
        Public Property Recoverable As Boolean = True

        Public Shared Function Succeed(toolId As String, Optional message As String = "",
                                       Optional data As Object = Nothing) As ToolResult
            Return New ToolResult With {
                .Success = True,
                .ToolId = toolId,
                .Message = message,
                .UserMessage = message,
                .Data = data,
                .ErrorCode = "",
                .Recoverable = True
            }
        End Function

        Public Shared Function Failed(toolId As String, message As String,
                                      Optional data As Object = Nothing,
                                      Optional errorCode As String = Nothing,
                                      Optional userMessage As String = Nothing,
                                      Optional debugDetail As String = Nothing,
                                      Optional recoverable As Boolean = True) As ToolResult
            Dim code = If(String.IsNullOrWhiteSpace(errorCode), ExceptionClassifier.CodeUnknown, errorCode)
            Dim userMsg = If(String.IsNullOrWhiteSpace(userMessage), message, userMessage)
            Dim detail = If(String.IsNullOrWhiteSpace(debugDetail), message, debugDetail)
            AppLogger.Warn("ToolRegistry", $"Tool failed toolId={toolId} code={code}: {AppLogger.Redact(message)}")
            Return New ToolResult With {
                .Success = False,
                .ToolId = toolId,
                .Message = message,
                .Data = data,
                .ErrorCode = code,
                .UserMessage = userMsg,
                .DebugDetail = AppLogger.Redact(detail),
                .Recoverable = recoverable
            }
        End Function

        Public Shared Function FromException(toolId As String, ex As Exception,
                                             Optional data As Object = Nothing) As ToolResult
            Dim classified = ExceptionClassifier.Classify(ex)
            Return Failed(toolId,
                          classified.DebugDetail,
                          data,
                          classified.ErrorCode,
                          classified.UserMessage,
                          classified.DebugDetail,
                          classified.Recoverable)
        End Function

        ''' <summary>供 Loop observe/repair 使用的紧凑摘要。</summary>
        Public Function ToObserveSummary() As String
            If Success Then
                Return If(String.IsNullOrWhiteSpace(Message), "ok", Message)
            End If
            Dim code = If(String.IsNullOrWhiteSpace(ErrorCode), ExceptionClassifier.CodeUnknown, ErrorCode)
            Dim um = If(String.IsNullOrWhiteSpace(UserMessage), Message, UserMessage)
            Return $"[{code}] recoverable={Recoverable}: {um}"
        End Function
    End Class

    ''' <summary>
    ''' 工具调用请求
    ''' </summary>
    Public Class ToolCall
        Public Property ToolId As String
        Public Property Parameters As JObject
        Public Property RequiresApproval As Boolean = False
    End Class

    ''' <summary>
    ''' 工具注册表 - 统一管理 MCP 工具和原生 Office 命令
    ''' </summary>
    Public Class ToolRegistry
        Private ReadOnly _tools As New Dictionary(Of String, ToolDescriptor)(StringComparer.OrdinalIgnoreCase)
        Private _mcpClient As StreamJsonRpcMCPClient
        Private ReadOnly _executeCodeCallback As Action(Of String, String, Boolean)

        ''' <summary>
        ''' 代码执行委托（用于原生 Office 工具）
        ''' </summary>
        Public Property ExecuteCode As Action(Of String, String, Boolean)
        Public Property ExecuteCodeWithResult As Func(Of String, String, Boolean, Boolean)

        ''' <summary>
        ''' MCP 客户端（用于远程工具调用）
        ''' </summary>
        Public Property McpClient As StreamJsonRpcMCPClient
            Get
                Return _mcpClient
            End Get
            Set(value As StreamJsonRpcMCPClient)
                _mcpClient = value
            End Set
        End Property

        Public Sub New(Optional mcpClient As StreamJsonRpcMCPClient = Nothing)
            _mcpClient = mcpClient
            RegisterBuiltInTools()
        End Sub

        Private Sub RegisterBuiltInTools()
            RegisterTool(New ToolDescriptor With {
                .Id = "memory.search",
                .Name = "检索长期记忆",
                .Description = "按关键词检索当前宿主可用的长期记忆。仅在 memory.enable_agentic_search 开启时注入给 Agent。",
                .AppType = "common",
                .Category = "记忆工具",
                .RiskLevel = "safe",
                .Parameters = New List(Of ToolParam) From {
                    New ToolParam With {.Name = "keyword", .Type = "string", .Required = True, .Description = "要检索的关键词或自然语言问题"},
                    New ToolParam With {.Name = "appType", .Type = "string", .Required = False, .Description = "Office 宿主类型，如 Excel/Word/PowerPoint"},
                    New ToolParam With {.Name = "topN", .Type = "integer", .Required = False, .Description = "最多返回条数，默认使用 MemoryConfig.RagTopN"}
                }
            })

            RegisterTool(New ToolDescriptor With {
                .Id = "memory.list_recent",
                .Name = "查看近期长期记忆",
                .Description = "列出近期长期记忆，供 Agent 解释当前可用记忆上下文。",
                .AppType = "common",
                .Category = "记忆工具",
                .RiskLevel = "safe",
                .Parameters = New List(Of ToolParam) From {
                    New ToolParam With {.Name = "appType", .Type = "string", .Required = False, .Description = "Office 宿主类型，如 Excel/Word/PowerPoint"},
                    New ToolParam With {.Name = "limit", .Type = "integer", .Required = False, .Description = "最多返回条数，默认 10"}
                }
            })

            RegisterTool(New ToolDescriptor With {
                .Id = "memory.promote",
                .Name = "晋升长期记忆",
                .Description = "将指定记忆晋升为长期记忆，用于保留用户偏好、事实和可复用工作经验。",
                .AppType = "common",
                .Category = "记忆工具",
                .RiskLevel = "medium",
                .Parameters = New List(Of ToolParam) From {
                    New ToolParam With {.Name = "memoryId", .Type = "integer", .Required = True, .Description = "要晋升的 atomic_memory id"}
                }
            })

            RegisterTool(New ToolDescriptor With {
                .Id = "memory.promote_session",
                .Name = "晋升当前会话高价值记忆",
                .Description = "将指定会话中高重要性的短期记忆批量晋升为长期记忆。",
                .AppType = "common",
                .Category = "记忆工具",
                .RiskLevel = "medium",
                .Parameters = New List(Of ToolParam) From {
                    New ToolParam With {.Name = "sessionId", .Type = "string", .Required = True, .Description = "会话 ID"},
                    New ToolParam With {.Name = "threshold", .Type = "integer", .Required = False, .Description = "重要性阈值，默认 0.65"},
                    New ToolParam With {.Name = "limit", .Type = "integer", .Required = False, .Description = "最多晋升条数，默认 20"}
                }
            })
        End Sub

        ''' <summary>
        ''' 从目录加载原生工具定义（JSON 文件）
        ''' </summary>
        Public Sub LoadFromDirectory(dir As String)
            If Not Directory.Exists(dir) Then Return
            Dim loaded As Integer = 0
            For Each file In Directory.GetFiles(dir, "*.json", SearchOption.AllDirectories)
                Try
                    Dim json = System.IO.File.ReadAllText(file)
                    Dim tool = JsonConvert.DeserializeObject(Of ToolDescriptor)(json)
                    If tool IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tool.Id) Then
                        RegisterOrMergeTool(tool)
                        loaded += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ToolRegistry] 加载工具失败 {file}: {ex.Message}")
                End Try
            Next
            Debug.WriteLine($"[ToolRegistry] 从目录加载原生工具 {loaded} 个，注册表当前 {ToolCount} 个: {dir}")
        End Sub

        ''' <summary>
        ''' 注册单个工具
        ''' </summary>
        Public Sub RegisterTool(tool As ToolDescriptor)
            RegisterOrMergeTool(tool)
        End Sub

        Private Sub RegisterOrMergeTool(tool As ToolDescriptor)
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.Id) Then Return

            Dim existing As ToolDescriptor = Nothing
            If _tools.TryGetValue(tool.Id, existing) Then
                existing.AppType = MergeAppTypes(existing.AppType, tool.AppType)
                If String.IsNullOrWhiteSpace(existing.Description) Then existing.Description = tool.Description
                If String.IsNullOrWhiteSpace(existing.Name) Then existing.Name = tool.Name
                If existing.Parameters Is Nothing OrElse existing.Parameters.Count = 0 Then existing.Parameters = tool.Parameters
                If String.IsNullOrWhiteSpace(existing.Category) OrElse existing.Category = "基础操作" Then existing.Category = tool.Category
                If String.Equals(existing.RiskLevel, "safe", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(tool.RiskLevel, "safe", StringComparison.OrdinalIgnoreCase) Then
                    existing.RiskLevel = tool.RiskLevel
                End If
            Else
                _tools(tool.Id) = tool
            End If
        End Sub

        Private Shared Function MergeAppTypes(left As String, right As String) As String
            Dim values As New List(Of String)()
            For Each raw In {left, right}
                If String.IsNullOrWhiteSpace(raw) Then Continue For
                For Each part In raw.Split({","c, ";"c, "|"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim item = part.Trim()
                    If item.Length = 0 Then Continue For
                    If Not values.Any(Function(v) String.Equals(v, item, StringComparison.OrdinalIgnoreCase)) Then
                        values.Add(item)
                    End If
                Next
            Next
            If values.Count = 0 Then Return ""
            Return String.Join(",", values)
        End Function

        ''' <summary>
        ''' 从 Skills 目录加载所有脚本作为工具
        ''' </summary>
        Public Sub LoadSkillScriptsAsTools()
            Try
                Dim allSkills = SkillsDirectoryService.GetAllSkills()
                For Each skill In allSkills
                    If skill.Scripts IsNot Nothing AndAlso skill.Scripts.Count > 0 Then
                        For Each script In skill.Scripts
                            Dim toolId = $"skill_script.{skill.Name}.{script.FileName}"
                            Dim toolDesc = If(Not String.IsNullOrEmpty(script.Description), script.Description, "")
                            Dim tool As New ToolDescriptor() With {
                                .Id = toolId,
                                .Name = $"{skill.Name}/{script.FileName}",
                                .Description = $"执行 Skill '{skill.Name}' 的脚本 {script.FileName} ({script.ScriptType})" &
                                              If(toolDesc <> "", vbCrLf & toolDesc, ""),
                                .AppType = "common",
                                .Category = "Skill 脚本",
                                .RiskLevel = "medium",
                                .Parameters = New List(Of ToolParam)()
                            }

                            ' 添加参数说明
                            If Not String.IsNullOrEmpty(script.ArgsHint) Then
                                tool.Parameters.Add(New ToolParam() With {
                                    .Name = "args",
                                    .Type = "string",
                                    .Required = False,
                                    .Description = script.ArgsHint
                                })
                            End If

                            _tools(toolId) = tool
                        Next
                    End If
                Next
                Dim totalScripts = allSkills.Sum(Function(s) If(s.Scripts IsNot Nothing, s.Scripts.Count, 0))
                Debug.WriteLine($"[ToolRegistry] 从 Skills 加载了 {totalScripts} 个脚本工具")
            Catch ex As Exception
                Debug.WriteLine($"[ToolRegistry] 加载 Skill 脚本工具失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 获取指定应用可用的工具
        ''' </summary>
        Public Function GetAvailableTools(appType As String) As List(Of ToolDescriptor)
            Dim result = _tools.Values.Where(Function(t)
                If t.Id.StartsWith("memory.", StringComparison.OrdinalIgnoreCase) AndAlso Not MemoryConfig.EnableAgenticSearch Then
                    Return False
                End If

                Return SupportsApp(t, appType)
            End Function).ToList()
            Return result
        End Function

        Private Shared Function SupportsApp(tool As ToolDescriptor, appType As String) As Boolean
            If tool Is Nothing Then Return False
            Dim raw = If(tool.AppType, "")
            If String.IsNullOrWhiteSpace(raw) Then Return True
            For Each part In raw.Split({","c, ";"c, "|"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim item = part.Trim()
                If String.Equals(item, "common", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(item, appType, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' 获取工具描述
        ''' </summary>
        Public Function GetTool(toolId As String) As ToolDescriptor
            If _tools.ContainsKey(toolId) Then Return _tools(toolId)
            Return Nothing
        End Function

        ''' <summary>
        ''' 检查工具是否存在
        ''' </summary>
        Public Function HasTool(toolId As String) As Boolean
            Return _tools.ContainsKey(toolId)
        End Function

        Public Function TryNormalizeToolCall(appType As String, toolCall As ToolCall, ByRef message As String) As Boolean
            message = ""
            If toolCall Is Nothing OrElse String.IsNullOrWhiteSpace(toolCall.ToolId) Then
                message = "工具调用为空"
                Return False
            End If

            Dim original = toolCall.ToolId.Trim()
            If toolCall.Parameters Is Nothing Then toolCall.Parameters = New JObject()
            Dim direct = GetTool(original)
            If direct IsNot Nothing AndAlso SupportsApp(direct, appType) Then
                toolCall.ToolId = direct.Id
                Return True
            End If

            Dim aliasId = ResolveBuiltInAlias(original, toolCall.Parameters)
            If Not String.IsNullOrWhiteSpace(aliasId) Then
                Dim aliasTool = GetTool(aliasId)
                If aliasTool IsNot Nothing AndAlso SupportsApp(aliasTool, appType) Then
                    toolCall.ToolId = aliasTool.Id
                    Return True
                End If
            End If

            Dim normalized = NormalizeToolKey(original)
            Dim matches = GetAvailableTools(appType).
                Where(Function(t) NormalizeToolKey(t.Id) = normalized OrElse NormalizeToolKey(t.Name) = normalized).
                ToList()
            If matches.Count = 1 Then
                toolCall.ToolId = matches(0).Id
                message = $"已将工具 {original} 规范化为 {toolCall.ToolId}"
                Return True
            End If

            Dim available = String.Join(", ", GetAvailableTools(appType).Select(Function(t) t.Id).OrderBy(Function(id) id).Take(30))
            message = $"未找到工具: {original}。只能使用当前 {If(appType, "Office")} 已注册工具，例如: {available}"
            Return False
        End Function

        Private Shared Function NormalizeToolKey(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim sb As New StringBuilder()
            For Each ch In value
                If Char.IsLetterOrDigit(ch) Then
                    sb.Append(Char.ToLowerInvariant(ch))
                End If
            Next
            Return sb.ToString()
        End Function

        Private Shared Function ResolveBuiltInAlias(toolId As String, params As JObject) As String
            Dim key = NormalizeToolKey(toolId)
            Select Case key
                Case "cleardocument", "cleardoc", "deletedocument", "deletealltext"
                    If params Is Nothing Then params = New JObject()
                    If params("range") Is Nothing Then params("range") = "all"
                    Return "DeleteText"
                Case "deletetext", "removetext"
                    Return "DeleteText"
                Case "replacetext"
                    Return "ReplaceText"
                Case "inserttext", "writetext", "appendtext"
                    Return "InsertText"
                Case "setparagraphformat", "setparagraphproperties"
                    Return "SetParagraphFormat"
            End Select
            Return Nothing
        End Function

        ''' <summary>
        ''' 自动生成工具描述文本（注入 LLM Prompt）
        ''' </summary>
        Public Function GenerateToolDescriptions(appType As String) As String
            Dim tools = GetAvailableTools(appType)
            Dim sb As New StringBuilder()
            sb.AppendLine($"【已注册工具 - 共 {tools.Count} 个】")
            sb.AppendLine()

            Dim grouped = tools.GroupBy(Function(t) t.Category).OrderBy(Function(g) g.Key)
            For Each group In grouped
                sb.AppendLine($"=== {group.Key} ({group.Count()}个) ===")
                For Each tool In group.OrderBy(Function(t) t.Id)
                sb.AppendLine($"{tool.Id} - {tool.Name}: {tool.Description}")
                    If Not String.IsNullOrWhiteSpace(tool.AvailabilityStatus) AndAlso Not String.Equals(tool.AvailabilityStatus, "available", StringComparison.OrdinalIgnoreCase) Then
                        sb.AppendLine($"  - 状态: {tool.AvailabilityStatus}{If(String.IsNullOrWhiteSpace(tool.LastError), "", "，错误: " & tool.LastError)}")
                    End If
                    For Each param In tool.Parameters
                        Dim reqMark = If(param.Required, "必需", "可选")
                        Dim defaultHint = If(param.DefaultValue IsNot Nothing, $", 默认: {param.DefaultValue}", "")
                        sb.AppendLine($"  - {param.Name}({param.Type}, {reqMark}{defaultHint}): {param.Description}")
                    Next
                    sb.AppendLine()
                Next
            Next

            sb.AppendLine()
            sb.AppendLine("【命令格式要求】")
            sb.AppendLine("每个步骤的 code 字段必须是完整 JSON 对象字符串，格式如下：")
            sb.AppendLine("单命令: {""command"":""命令名"",""params"":{...}}")
            sb.AppendLine("多命令: {""commands"":[{""command"":""命令名"",""params"":{...}},...]}")
            sb.AppendLine()
            sb.AppendLine("【绝对禁止】")
            sb.AppendLine("- 禁止使用 actions/operations 数组")
            sb.AppendLine("- 禁止省略 params 包装")
            sb.AppendLine("- 禁止自创未注册的命令")
            sb.AppendLine("- 禁止返回不带代码块的裸 JSON")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 从 MCP 服务器加载远程工具
        ''' </summary>
        Public Async Function LoadMcpToolsAsync() As Task
            If _mcpClient Is Nothing OrElse Not _mcpClient.IsInitialized Then
                Debug.WriteLine("[ToolRegistry] MCP 客户端未初始化，跳过加载远程工具")
                Return
            End If

            Try
                Dim mcpTools = Await _mcpClient.ListToolsAsync()
                If mcpTools Is Nothing Then Return

                For Each mcpTool In mcpTools
                    If String.IsNullOrWhiteSpace(mcpTool.Name) Then Continue For
                    Dim descriptor = ConvertMcpToDescriptor(mcpTool)
                    _tools(descriptor.Id) = descriptor
                Next

                Debug.WriteLine($"[ToolRegistry] 从 MCP 服务器加载了 {mcpTools.Count} 个工具")
            Catch ex As Exception
                Debug.WriteLine($"[ToolRegistry] 加载 MCP 工具失败: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' 将 MCP 工具信息转换为 ToolDescriptor
        ''' </summary>
        Private Function ConvertMcpToDescriptor(mcpTool As MCPToolInfo) As ToolDescriptor
            Dim descriptor As New ToolDescriptor With {
                .Id = $"mcp.{mcpTool.Name}",
                .Name = mcpTool.Name,
                .Description = If(mcpTool.Description, $"MCP 工具: {mcpTool.Name}"),
                .AppType = "common",
                .Category = "MCP 工具",
                .RiskLevel = "medium",
                .AvailabilityStatus = "available",
                .LastError = ""
            }

            ' 解析 InputSchema 中的参数
            Try
                If mcpTool.InputSchema IsNot Nothing Then
                    Dim schema = JObject.FromObject(mcpTool.InputSchema)
                    Dim props = TryCast(schema("properties"), JObject)
                    Dim requiredArray = TryCast(schema("required"), JArray)
                    Dim requiredSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    If requiredArray IsNot Nothing Then
                        For Each r In requiredArray
                            requiredSet.Add(r.ToString())
                        Next
                    End If

                    If props IsNot Nothing Then
                        For Each prop In props.Properties()
                            Dim paramType = "string"
                            Dim propType = prop.Value("type")?.ToString()
                            If Not String.IsNullOrEmpty(propType) Then
                                Select Case propType.ToLower()
                                    Case "integer", "number"
                                        paramType = "integer"
                                    Case "boolean"
                                        paramType = "boolean"
                                    Case "array"
                                        paramType = "array"
                                    Case "object"
                                        paramType = "object"
                                    Case Else
                                        paramType = "string"
                                End Select
                            End If

                            descriptor.Parameters.Add(New ToolParam With {
                                .Name = prop.Name,
                                .Type = paramType,
                                .Required = requiredSet.Contains(prop.Name),
                                .Description = prop.Value("description")?.ToString()
                            })
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[ToolRegistry] 解析 MCP 工具参数失败 {mcpTool.Name}: {ex.Message}")
            End Try

            Return descriptor
        End Function

        ''' <summary>
        ''' 执行工具调用
        ''' </summary>
        Public Async Function ExecuteToolAsync(toolId As String, params As JObject) As Task(Of ToolResult)
            Dim sw = Diagnostics.Stopwatch.StartNew()

            Dim tool = GetTool(toolId)
            If tool Is Nothing Then
                sw.Stop()
                Return ToolResult.Failed(toolId, $"未找到工具: {toolId}")
            End If

            If toolId.StartsWith("memory.", StringComparison.OrdinalIgnoreCase) Then
                Dim memoryResult = Await ExecuteMemoryToolAsync(toolId, params)
                sw.Stop()
                memoryResult.ElapsedMs = sw.ElapsedMilliseconds
                Return memoryResult
            End If

            ' Skill 脚本工具调用（以 skill_script. 开头）
            If toolId.StartsWith("skill_script.") Then
                Dim scriptResult = Await ExecuteSkillScriptAsync(toolId, params)
                sw.Stop()
                scriptResult.ElapsedMs = sw.ElapsedMilliseconds
                Return scriptResult
            End If

            ' MCP 工具调用（以 mcp. 开头）
            If toolId.StartsWith("mcp.") Then
                If _mcpClient Is Nothing OrElse Not _mcpClient.IsInitialized Then
                    sw.Stop()
                    Dim failureMessage = "MCP 客户端未初始化"
                    MarkToolHealth(tool, "unavailable", failureMessage)
                    Return ToolResult.Failed(toolId, failureMessage,
                        New With {
                            .mcpToolName = toolId.Substring(4),
                            .mcpStatus = "unavailable",
                            .failureReason = failureMessage,
                            .elapsedMs = sw.ElapsedMilliseconds
                        })
                End If

                Try
                    Dim actualToolName = toolId.Substring(4)
                    Dim mcpResult = Await _mcpClient.CallToolAsync(actualToolName, params)
                    sw.Stop()

                    If mcpResult.IsError Then
                        Dim failureMessage = If(mcpResult.ErrorMessage, "MCP 工具执行失败")
                        MarkToolHealth(tool, "error", failureMessage)
                        Return ToolResult.Failed(toolId, failureMessage,
                            New With {
                                .mcpToolName = actualToolName,
                                .mcpStatus = "error",
                                .failureReason = failureMessage,
                                .elapsedMs = sw.ElapsedMilliseconds
                            })
                    End If

                    Dim outputText As String = ""
                    If mcpResult.Content IsNot Nothing AndAlso mcpResult.Content.Count > 0 Then
                        Dim sb As New StringBuilder()
                        For Each content In mcpResult.Content
                            If content.Type = "text" AndAlso Not String.IsNullOrEmpty(content.Text) Then
                                sb.AppendLine(content.Text)
                            End If
                        Next
                        outputText = sb.ToString().Trim()
                    End If

                    MarkToolHealth(tool, "available", "")
                    Return ToolResult.Succeed(toolId, If(String.IsNullOrEmpty(outputText), "执行成功", outputText),
                                               New With {
                                                   .elapsedMs = sw.ElapsedMilliseconds,
                                                   .mcpToolName = actualToolName,
                                                   .mcpStatus = "available",
                                                   .contentCount = If(mcpResult.Content Is Nothing, 0, mcpResult.Content.Count)
                                               })
                Catch ex As Exception
                    sw.Stop()
                    Dim classified = ExceptionClassifier.Classify(ex)
                    Dim failureMessage = $"MCP 调用异常: {classified.DebugDetail}"
                    MarkToolHealth(tool, "error", failureMessage)
                    AppLogger.Error("ToolRegistry", $"MCP tool exception toolId={toolId}", ex)
                    Return ToolResult.FromException(toolId, ex,
                        New With {
                            .mcpToolName = toolId.Substring(4),
                            .mcpStatus = "error",
                            .failureReason = failureMessage,
                            .elapsedMs = sw.ElapsedMilliseconds,
                            .errorCode = classified.ErrorCode
                        })
                End Try
            End If

            ' 原生 Office 工具，通过 ExecuteCode 回调执行
            If tool.IsVbaFallback OrElse Not toolId.StartsWith("mcp.") Then
                If ExecuteCodeWithResult Is Nothing AndAlso ExecuteCode Is Nothing Then
                    sw.Stop()
                    Return ToolResult.Failed(toolId, "ExecuteCode 回调未设置")
                End If

                ' 构建完整的 JSON 命令
                Dim command As String
                If params.ContainsKey("commands") Then
                    command = params.ToString(Formatting.None)
                Else
                    Dim wrapped = New JObject From {
                        {"command", toolId},
                        {"params", params}
                    }
                    command = wrapped.ToString(Formatting.None)
                End If

                Try
                    ' 调用现有执行逻辑
                    Dim hostSuccess As Boolean = True
                    If ExecuteCodeWithResult IsNot Nothing Then
                        hostSuccess = ExecuteCodeWithResult.Invoke(command, "json", False)
                    Else
                        ExecuteCode.Invoke(command, "json", False)
                    End If
                    sw.Stop()
                    If hostSuccess Then
                        Return ToolResult.Succeed(toolId, "执行成功", New With {.elapsedMs = sw.ElapsedMilliseconds})
                    End If
                    Dim failureMessage = $"宿主执行器返回失败: {toolId}"
                    Return ToolResult.Failed(toolId,
                                             failureMessage,
                                             New With {.elapsedMs = sw.ElapsedMilliseconds, .command = command},
                                             ExceptionClassifier.CodeUnknown,
                                             failureMessage,
                                             failureMessage,
                                             recoverable:=True)
                Catch ex As Exception
                    sw.Stop()
                    AppLogger.Error("ToolRegistry", $"Native tool execute failed toolId={toolId}", ex)
                    Return ToolResult.FromException(toolId, ex, New With {.elapsedMs = sw.ElapsedMilliseconds})
                End Try
            End If

            sw.Stop()
            Return ToolResult.Failed(toolId, "未知的工具类型",
                                    errorCode:=ExceptionClassifier.CodeNotFound,
                                    userMessage:="未识别的工具类型",
                                    recoverable:=False)
        End Function

        Private Async Function ExecuteMemoryToolAsync(toolId As String, params As JObject) As Task(Of ToolResult)
            Await Task.Yield()

            If Not MemoryConfig.EnableAgenticSearch Then
                Return ToolResult.Failed(toolId, "memory.enable_agentic_search 未开启，Agent 不能主动检索或修改记忆")
            End If

            Select Case toolId.ToLowerInvariant()
                Case "memory.search"
                    Dim keyword = GetStringParam(params, "keyword")
                    If String.IsNullOrWhiteSpace(keyword) Then
                        Return ToolResult.Failed(toolId, "缺少 keyword 参数")
                    End If

                    Dim appType = GetStringParam(params, "appType")
                    Dim topN = GetIntegerParam(params, "topN", MemoryConfig.RagTopN)
                    Dim searchResults = MemoryService.SearchMemories(keyword, Nothing, Nothing, appType)
                    Dim memories = searchResults.Take(Math.Max(1, topN)).
                        Select(Function(m) ToMemoryToolPayload(m)).ToList()
                    Return ToolResult.Succeed(toolId, $"找到 {memories.Count} 条长期记忆", memories)

                Case "memory.list_recent"
                    Dim appType = GetStringParam(params, "appType")
                    Dim limit = GetIntegerParam(params, "limit", 10)
                    Dim recentMemories = MemoryRepository.ListAtomicMemories(Math.Max(1, limit * 3), 0, appType)
                    Dim memories = recentMemories.
                        Where(Function(m) String.Equals(m.MemoryType, "long_term", StringComparison.OrdinalIgnoreCase)).
                        Take(Math.Max(1, limit)).
                        Select(Function(m) ToMemoryToolPayload(m)).ToList()
                    Return ToolResult.Succeed(toolId, $"返回 {memories.Count} 条近期长期记忆", memories)

                Case "memory.promote"
                    Dim memoryId = GetLongParam(params, "memoryId", 0)
                    If memoryId <= 0 Then
                        Return ToolResult.Failed(toolId, "缺少有效的 memoryId 参数")
                    End If

                    Dim changed = MemoryService.PromoteMemoryToLongTerm(memoryId)
                    Return ToolResult.Succeed(toolId, If(changed, $"已晋升记忆 {memoryId}", $"记忆 {memoryId} 已是长期记忆或不存在"),
                                               New With {.memoryId = memoryId, .changed = changed})

                Case "memory.promote_session"
                    Dim sessionId = GetStringParam(params, "sessionId")
                    If String.IsNullOrWhiteSpace(sessionId) Then
                        Return ToolResult.Failed(toolId, "缺少 sessionId 参数")
                    End If

                    Dim threshold = GetDoubleParam(params, "threshold", 0.65R)
                    Dim limit = GetIntegerParam(params, "limit", 20)
                    Dim promoted = MemoryService.PromoteImportantShortTermMemories(sessionId, threshold, limit)
                    Return ToolResult.Succeed(toolId, $"已晋升 {promoted} 条会话记忆",
                                               New With {.sessionId = sessionId, .promoted = promoted, .threshold = threshold, .limit = limit})
            End Select

            Return ToolResult.Failed(toolId, $"未知的记忆工具: {toolId}")
        End Function

        Private Function ToMemoryToolPayload(memory As AtomicMemoryRecord) As Object
            Return New With {
                .id = memory.Id,
                .content = memory.Content,
                .memoryType = memory.MemoryType,
                .importance = memory.Importance,
                .sourceType = memory.SourceType,
                .sessionId = memory.SessionId,
                .createdAt = memory.CreateTime,
                .similarity = memory.SimilarityScore,
                .accessCount = memory.AccessCount
            }
        End Function

        Private Sub MarkToolHealth(tool As ToolDescriptor, status As String, errorMessage As String)
            If tool Is Nothing Then Return
            tool.AvailabilityStatus = If(String.IsNullOrWhiteSpace(status), "available", status)
            tool.LastError = If(errorMessage, "")
        End Sub

        Private Function GetStringParam(params As JObject, name As String, Optional defaultValue As String = "") As String
            If params Is Nothing OrElse params(name) Is Nothing Then Return defaultValue
            Return params(name).ToString()
        End Function

        Private Function GetIntegerParam(params As JObject, name As String, defaultValue As Integer) As Integer
            If params Is Nothing OrElse params(name) Is Nothing Then Return defaultValue
            Dim value As Integer
            If Integer.TryParse(params(name).ToString(), value) Then Return value
            Return defaultValue
        End Function

        Private Function GetLongParam(params As JObject, name As String, defaultValue As Long) As Long
            If params Is Nothing OrElse params(name) Is Nothing Then Return defaultValue
            Dim value As Long
            If Long.TryParse(params(name).ToString(), value) Then Return value
            Return defaultValue
        End Function

        Private Function GetDoubleParam(params As JObject, name As String, defaultValue As Double) As Double
            If params Is Nothing OrElse params(name) Is Nothing Then Return defaultValue
            Dim value As Double
            If Double.TryParse(params(name).ToString(), value) Then Return value
            Return defaultValue
        End Function

        ''' <summary>
        ''' 执行 Skill 脚本工具
        ''' </summary>
        Private Async Function ExecuteSkillScriptAsync(toolId As String, params As JObject) As Task(Of ToolResult)
            ' toolId 格式: skill_script.{skillName}.{scriptFileName}
            Const prefix As String = "skill_script."
            If Not toolId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) OrElse toolId.Length <= prefix.Length Then
                Return ToolResult.Failed(toolId, $"无效的 Skill 脚本工具 ID: {toolId}")
            End If

            ' 查找 Skill 和脚本
            Dim allSkills = SkillsDirectoryService.GetAllSkills()
            Dim skillName As String = ""
            Dim scriptFileName As String = ""
            Dim skill As SkillFileDefinition = Nothing

            Dim remainder = toolId.Substring(prefix.Length)
            For Each candidate In allSkills
                If candidate Is Nothing OrElse String.IsNullOrWhiteSpace(candidate.Name) Then Continue For
                Dim candidatePrefix = candidate.Name & "."
                If remainder.StartsWith(candidatePrefix, StringComparison.OrdinalIgnoreCase) Then
                    skill = candidate
                    skillName = candidate.Name
                    scriptFileName = remainder.Substring(candidatePrefix.Length)
                    Exit For
                End If
            Next

            If skill Is Nothing Then
                Dim parts = remainder.Split("."c)
                If parts.Length >= 2 Then
                    skillName = parts(0)
                    scriptFileName = String.Join(".", parts.Skip(1))
                    skill = allSkills.FirstOrDefault(Function(s) s.Name.Equals(skillName, StringComparison.OrdinalIgnoreCase))
                End If
            End If

            If skill Is Nothing Then
                Dim failed = ToolResult.Failed(toolId, $"未找到 Skill: {If(String.IsNullOrWhiteSpace(skillName), remainder, skillName)}",
                    New With {.skillName = skillName, .scriptFileName = scriptFileName, .failureReason = "skill_not_found"})
                SkillsService.RecordSkillExecution(If(String.IsNullOrWhiteSpace(skillName), remainder, skillName), False, failed.Message, "Skill 脚本工具解析")
                Return failed
            End If

            Dim script = skill.Scripts.FirstOrDefault(Function(s) s.FileName.Equals(scriptFileName, StringComparison.OrdinalIgnoreCase))
            If script Is Nothing Then
                Dim failed = ToolResult.Failed(toolId, $"未找到脚本: {scriptFileName} (在 Skill {skillName} 中)",
                    New With {.skillName = skillName, .scriptFileName = scriptFileName, .failureReason = "script_not_found"})
                SkillsService.RecordSkillExecution(skillName, False, failed.Message, $"执行脚本 {scriptFileName}")
                Return failed
            End If

            ' 解析参数
            Dim args As New Dictionary(Of String, String)()
            If params IsNot Nothing Then
                For Each prop In params.Properties()
                    args(prop.Name) = prop.Value?.ToString()
                Next
            End If

            ' 执行脚本
            Try
                Dim result = Await SkillScriptExecutor.ExecuteScriptAsync(script, args, skill.FilePath)

                If result.Success Then
                    Dim output = result.StdOut
                    If String.IsNullOrEmpty(output) Then output = "脚本执行成功（无输出）"
                    SkillsService.RecordSkillExecution(skillName, True, "", $"执行脚本 {scriptFileName}")
                    Return ToolResult.Succeed(toolId, output,
                        New With {
                            .elapsedMs = result.ElapsedMs,
                            .exitCode = result.ExitCode,
                            .skillName = skillName,
                            .scriptFileName = scriptFileName
                        })
                Else
                    Dim failureMessage = $"脚本执行失败 (退出码: {result.ExitCode})" &
                        If(Not String.IsNullOrEmpty(result.ErrorMessage), $": {result.ErrorMessage}", "")
                    SkillsService.RecordSkillExecution(skillName, False, failureMessage, $"执行脚本 {scriptFileName}")
                    Return ToolResult.Failed(toolId,
                        failureMessage,
                        New With {
                            .elapsedMs = result.ElapsedMs,
                            .exitCode = result.ExitCode,
                            .stdErr = result.StdErr,
                            .skillName = skillName,
                            .scriptFileName = scriptFileName,
                            .failureReason = failureMessage
                        })
                End If
            Catch ex As Exception
                Dim failureMessage = $"脚本执行异常: {ex.Message}"
                SkillsService.RecordSkillExecution(skillName, False, failureMessage, $"执行脚本 {scriptFileName}")
                Return ToolResult.Failed(toolId, failureMessage,
                    New With {
                        .skillName = skillName,
                        .scriptFileName = scriptFileName,
                        .failureReason = failureMessage
                    })
            End Try
        End Function

        ''' <summary>
        ''' 获取所有已注册工具数量
        ''' </summary>
        Public ReadOnly Property ToolCount As Integer
            Get
                Return _tools.Count
            End Get
        End Property

        ''' <summary>
        ''' 清空所有工具
        ''' </summary>
        Public Sub Clear()
            _tools.Clear()
            RegisterBuiltInTools()
        End Sub
    End Class

End Namespace
