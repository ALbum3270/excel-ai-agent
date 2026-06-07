Imports System.Collections.Concurrent
Imports System.Net.Http
Imports System.Threading

Public NotInheritable Class HttpClientPool
    Private Shared ReadOnly _clients As New ConcurrentDictionary(Of String, HttpClient)(StringComparer.OrdinalIgnoreCase)

    Private Sub New()
    End Sub

    Public Shared Function GetClient(apiUrl As String) As HttpClient
        Dim key = GetClientKey(apiUrl)
        Return _clients.GetOrAdd(key, Function(clientKey)
                                          Dim client = New HttpClient()
                                          client.Timeout = Timeout.InfiniteTimeSpan
                                          Return client
                                      End Function)
    End Function

    Private Shared Function GetClientKey(apiUrl As String) As String
        If String.IsNullOrWhiteSpace(apiUrl) Then Return "default"

        Try
            Dim uri = New Uri(apiUrl)
            Return $"{uri.Scheme}://{uri.Host}:{uri.Port}"
        Catch
            Return apiUrl.Trim()
        End Try
    End Function
End Class
