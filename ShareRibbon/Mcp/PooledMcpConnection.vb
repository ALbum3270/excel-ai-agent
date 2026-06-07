Imports System.Threading
Imports System.Threading.Tasks

Public Class PooledMcpConnection
    Implements IDisposable

    Private ReadOnly _client As StreamJsonRpcMCPClient
    Private ReadOnly _callSemaphore As New SemaphoreSlim(1, 1)
    Private _isInitialized As Boolean
    Private _isDisposed As Boolean
    Private _tools As List(Of MCPToolInfo)

    Public ReadOnly Property ConnectionName As String
    Public ReadOnly Property Config As MCPConnectionConfig
    Public Property LastAccessUtc As DateTime

    Public Sub New(connectionName As String, config As MCPConnectionConfig)
        Me.ConnectionName = connectionName
        Me.Config = config
        _client = New StreamJsonRpcMCPClient()
        LastAccessUtc = DateTime.UtcNow
    End Sub

    Public Async Function InitializeAsync() As Task
        If _isInitialized Then Return

        Await _client.ConfigureAsync(Config.Url)
        Dim initResult = Await _client.InitializeAsync()
        If initResult Is Nothing OrElse Not initResult.Success Then
            Dim message = If(initResult?.ErrorMessage, "unknown error")
            Throw New InvalidOperationException($"MCP connection '{ConnectionName}' initialization failed: {message}")
        End If

        _tools = Await _client.ListToolsAsync()
        _isInitialized = True
        Touch()
    End Function

    Public Function IsHealthy() As Boolean
        Return _isInitialized AndAlso Not _isDisposed
    End Function

    Public Function GetTools() As List(Of MCPToolInfo)
        Touch()
        If _tools Is Nothing Then Return New List(Of MCPToolInfo)()
        Return New List(Of MCPToolInfo)(_tools)
    End Function

    Public Async Function CallToolAsync(toolName As String, arguments As Object) As Task(Of MCPToolResult)
        Await _callSemaphore.WaitAsync()
        Try
            Touch()
            Return Await _client.CallToolAsync(toolName, arguments)
        Finally
            _callSemaphore.Release()
        End Try
    End Function

    Public Sub Touch()
        LastAccessUtc = DateTime.UtcNow
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If _isDisposed Then Return
        _isDisposed = True
        _callSemaphore.Dispose()
        _client.Dispose()
    End Sub
End Class
