Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Exposes the same registered native Office tools to ordinary model requests that the
''' Agent runtime uses. Chat versus Agent is a presentation choice, not a capability wall.
''' </summary>
Public Class OfficeToolBroker
    Implements IToolBroker

    Private Shared ReadOnly FunctionNamePattern As New Regex("^[A-Za-z0-9_-]+$", RegexOptions.CultureInvariant)
    Private ReadOnly _registry As Agent.ToolRegistry

    Public Sub New(registry As Agent.ToolRegistry)
        If registry Is Nothing Then Throw New ArgumentNullException(NameOf(registry))
        _registry = registry
    End Sub

    Public Function GetTools(context As ChatRequestContext) As JArray Implements IToolBroker.GetTools
        Dim appType = If(context?.AppInfo?.Name, "")
        Return BuildTools(_registry, appType)
    End Function

    Public Shared Function BuildTools(registry As Agent.ToolRegistry, appType As String) As JArray
        Dim result As New JArray()
        If registry Is Nothing Then Return result

        For Each tool In registry.GetAvailableTools(appType).
            Where(Function(item) IsModelCallable(item)).
            OrderBy(Function(item) item.Id, StringComparer.OrdinalIgnoreCase)

            Dim properties As New JObject()
            Dim required As New JArray()
            For Each parameter In If(tool.Parameters, New List(Of Agent.ToolParam)())
                If parameter Is Nothing OrElse String.IsNullOrWhiteSpace(parameter.Name) Then Continue For
                Dim schema As New JObject From {
                    {"type", NormalizeJsonType(parameter.Type)}
                }
                If Not String.IsNullOrWhiteSpace(parameter.Description) Then
                    schema("description") = parameter.Description
                End If
                If parameter.DefaultValue IsNot Nothing Then
                    schema("default") = JToken.FromObject(parameter.DefaultValue)
                End If
                If String.Equals(schema("type")?.ToString(), "array", StringComparison.OrdinalIgnoreCase) Then
                    schema("items") = New JObject()
                End If
                properties(parameter.Name) = schema
                If parameter.Required Then required.Add(parameter.Name)
            Next

            Dim parameters As New JObject From {
                {"type", "object"},
                {"properties", properties},
                {"additionalProperties", False}
            }
            If required.Count > 0 Then parameters("required") = required

            result.Add(New JObject From {
                {"type", "function"},
                {"function", New JObject From {
                    {"name", tool.Id},
                    {"description", If(String.IsNullOrWhiteSpace(tool.Description), tool.Name, tool.Description)},
                    {"parameters", parameters}
                }}
            })
        Next
        Return result
    End Function

    Private Shared Function IsModelCallable(tool As Agent.ToolDescriptor) As Boolean
        If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.Id) Then Return False
        If Not FunctionNamePattern.IsMatch(tool.Id) Then Return False
        If tool.Id.StartsWith("memory.", StringComparison.OrdinalIgnoreCase) OrElse
           tool.Id.StartsWith("mcp.", StringComparison.OrdinalIgnoreCase) OrElse
           tool.Id.StartsWith("skill_script.", StringComparison.OrdinalIgnoreCase) Then Return False
        Return String.Equals(If(tool.AvailabilityStatus, "available"),
                             "available",
                             StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function NormalizeJsonType(rawType As String) As String
        Select Case If(rawType, "").Trim().ToLowerInvariant()
            Case "integer", "int", "long"
                Return "integer"
            Case "number", "double", "decimal", "float"
                Return "number"
            Case "boolean", "bool"
                Return "boolean"
            Case "array", "list"
                Return "array"
            Case "object", "json"
                Return "object"
            Case Else
                Return "string"
        End Select
    End Function
End Class

''' <summary>Merges tool sources while keeping one canonical function name.</summary>
Public Class CompositeToolBroker
    Implements IToolBroker

    Private ReadOnly _brokers As List(Of IToolBroker)

    Public Sub New(ParamArray brokers As IToolBroker())
        _brokers = If(brokers, Array.Empty(Of IToolBroker)()).Where(Function(item) item IsNot Nothing).ToList()
    End Sub

    Public Function GetTools(context As ChatRequestContext) As JArray Implements IToolBroker.GetTools
        Dim result As New JArray()
        Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each broker In _brokers
            Dim tools = broker.GetTools(context)
            If tools Is Nothing Then Continue For
            For Each token In tools
                Dim name = token?("function")?("name")?.ToString()
                If String.IsNullOrWhiteSpace(name) OrElse Not names.Add(name) Then Continue For
                result.Add(token.DeepClone())
            Next
        Next
        Return result
    End Function
End Class

''' <summary>
''' Parses only an assistant response whose whole content is a structured tool-call envelope.
''' Prose that merely quotes JSON is deliberately ignored to prevent accidental execution.
''' </summary>
Public NotInheritable Class ChatToolCallParser
    Private Sub New()
    End Sub

    Public Shared Function TryParse(content As String, ByRef toolCall As Agent.ToolCall) As Boolean
        toolCall = Nothing
        Dim json = UnwrapJson(content)
        If String.IsNullOrWhiteSpace(json) Then Return False

        Dim root As JObject
        Try
            root = JObject.Parse(json)
        Catch
            Return False
        End Try

        Dim envelope = root
        If root("action") IsNot Nothing AndAlso root("action").Type = JTokenType.Object Then
            envelope = DirectCast(root("action"), JObject)
        End If

        Dim toolId = FirstText(envelope, "tool", "command", "name")
        If String.IsNullOrWhiteSpace(toolId) Then Return False
        Dim parameters = TryCast(envelope("params"), JObject)
        If parameters Is Nothing Then parameters = TryCast(envelope("parameters"), JObject)
        If parameters Is Nothing Then parameters = TryCast(envelope("arguments"), JObject)
        If parameters Is Nothing Then parameters = New JObject()

        toolCall = New Agent.ToolCall With {
            .ToolId = toolId,
            .Parameters = parameters
        }
        Return True
    End Function

    Private Shared Function UnwrapJson(content As String) As String
        Dim value = If(content, "").Trim()
        If value.StartsWith("```", StringComparison.Ordinal) Then
            Dim firstLine = value.IndexOf(vbLf)
            Dim closing = value.LastIndexOf("```", StringComparison.Ordinal)
            If firstLine < 0 OrElse closing <= firstLine Then Return ""
            value = value.Substring(firstLine + 1, closing - firstLine - 1).Trim()
        End If
        If Not value.StartsWith("{", StringComparison.Ordinal) OrElse
           Not value.EndsWith("}", StringComparison.Ordinal) Then Return ""
        Return value
    End Function

    Private Shared Function FirstText(obj As JObject, ParamArray names As String()) As String
        If obj Is Nothing Then Return ""
        For Each name In names
            Dim value = obj(name)?.ToString()
            If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
        Next
        Return ""
    End Function
End Class

''' <summary>Executes chat-originated calls through ToolRegistry and its SafetyGate.</summary>
Public Class ChatToolCallRuntime
    Private ReadOnly _registry As Agent.ToolRegistry

    Public Sub New(registry As Agent.ToolRegistry)
        If registry Is Nothing Then Throw New ArgumentNullException(NameOf(registry))
        _registry = registry
    End Sub

    Public Async Function ExecuteAsync(appType As String,
                                       toolId As String,
                                       parameters As JObject) As Task(Of Agent.ToolResult)
        Dim callInfo As New Agent.ToolCall With {
            .ToolId = toolId,
            .Parameters = If(parameters, New JObject())
        }
        Dim normalizationMessage As String = ""
        If Not _registry.TryNormalizeToolCall(appType, callInfo, normalizationMessage) Then Return Nothing

        Dim executionContext As New Agent.ToolExecutionContext With {
            .AppType = appType,
            .EnforceAllowedTools = False
        }
        Return Await _registry.ExecuteToolAsync(executionContext, callInfo.ToolId, callInfo.Parameters)
    End Function
End Class
