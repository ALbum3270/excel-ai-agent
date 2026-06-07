Imports System.Linq
Imports Newtonsoft.Json.Linq

''' <summary>
''' Builds tool definitions from the enabled MCP connection settings.
''' </summary>
Public Class McpToolBroker
    Implements IToolBroker

    Public Function GetTools(context As ChatRequestContext) As JArray Implements IToolBroker.GetTools
        Dim appInfo = If(context IsNot Nothing, context.AppInfo, Nothing)
        Dim chatSettings As New ChatSettings(appInfo)

        If chatSettings.EnabledMcpList Is Nothing OrElse chatSettings.EnabledMcpList.Count = 0 Then
            Return Nothing
        End If

        Dim tools As New JArray()
        Dim connections = MCPConnectionManager.LoadConnections()

        For Each mcpName In chatSettings.EnabledMcpList
            Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpName AndAlso c.IsActive)
            If connection Is Nothing Then Continue For

            If connection.Tools IsNot Nothing AndAlso connection.Tools.Count > 0 Then
                For Each toolObj In connection.Tools
                    tools.Add(toolObj)
                Next
                Debug.WriteLine($"[McpToolBroker] Loaded {connection.Tools.Count} tools from {connection.Name}")
            Else
                tools.Add(CreateGenericMcpCallTool(connection.Name))
                Debug.WriteLine($"[McpToolBroker] Connection {connection.Name} has no cached tools; using generic mcp_call")
            End If
        Next

        If tools.Count = 0 Then Return Nothing
        Return tools
    End Function

    Private Function CreateGenericMcpCallTool(connectionName As String) As JObject
        Dim toolObj As New JObject()
        toolObj("type") = "function"

        Dim functionObj As New JObject()
        functionObj("name") = "mcp_call"
        functionObj("description") = $"Call MCP tool through {connectionName} connection"

        Dim parametersObj As New JObject()
        parametersObj("type") = "object"

        Dim propertiesObj As New JObject()
        propertiesObj("tool_name") = New JObject From {
            {"type", "string"},
            {"description", "The name of the MCP tool to call"}
        }
        propertiesObj("arguments") = New JObject From {
            {"type", "object"},
            {"description", "The arguments to pass to the MCP tool"}
        }

        parametersObj("properties") = propertiesObj
        parametersObj("required") = New JArray({"tool_name", "arguments"})

        functionObj("parameters") = parametersObj
        toolObj("function") = functionObj

        Return toolObj
    End Function
End Class
