Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Threading.Tasks

Public NotInheritable Class McpConnectionPool
    Private Shared ReadOnly _instance As New Lazy(Of McpConnectionPool)(Function() New McpConnectionPool())

    Private ReadOnly _connections As New ConcurrentDictionary(Of String, PooledMcpConnection)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _gate As New SemaphoreSlim(1, 1)

    Private Sub New()
    End Sub

    Public Shared ReadOnly Property Instance As McpConnectionPool
        Get
            Return _instance.Value
        End Get
    End Property

    Public Async Function GetOrCreateConnectionAsync(config As MCPConnectionConfig) As Task(Of PooledMcpConnection)
        If config Is Nothing Then Throw New ArgumentNullException(NameOf(config))
        Dim key = BuildKey(config)

        Dim existing As PooledMcpConnection = Nothing
        If _connections.TryGetValue(key, existing) AndAlso existing.IsHealthy() Then
            existing.Touch()
            Return existing
        End If

        Await _gate.WaitAsync()
        Try
            If _connections.TryGetValue(key, existing) Then
                If existing.IsHealthy() Then
                    existing.Touch()
                    Return existing
                End If

                Dim removed As PooledMcpConnection = Nothing
                If _connections.TryRemove(key, removed) Then
                    removed.Dispose()
                End If
            End If

            Dim pooled = New PooledMcpConnection(config.Name, config)
            Await pooled.InitializeAsync()
            _connections(key) = pooled
            Return pooled
        Finally
            _gate.Release()
        End Try
    End Function

    Public Sub Invalidate(config As MCPConnectionConfig)
        If config Is Nothing Then Return
        Dim key = BuildKey(config)
        Dim removed As PooledMcpConnection = Nothing
        If _connections.TryRemove(key, removed) Then
            removed.Dispose()
        End If
    End Sub

    Public Sub CleanupIdleConnections(idleTimeout As TimeSpan)
        Dim now = DateTime.UtcNow
        For Each kvp In _connections
            If now - kvp.Value.LastAccessUtc > idleTimeout Then
                Dim removed As PooledMcpConnection = Nothing
                If _connections.TryRemove(kvp.Key, removed) Then
                    removed.Dispose()
                End If
            End If
        Next
    End Sub

    Private Shared Function BuildKey(config As MCPConnectionConfig) As String
        Dim name = If(config.Name, String.Empty).Trim()
        Dim url = If(config.Url, String.Empty).Trim()
        Return $"{name}|{url}"
    End Function
End Class
