Imports System.Collections.Concurrent
Imports System.Diagnostics
Imports System.IO
Imports System.Threading.Tasks
Imports Microsoft.Web.WebView2.Core

Public NotInheritable Class WebView2EnvironmentCache
    Private Shared ReadOnly _environments As New ConcurrentDictionary(Of String, Lazy(Of Task(Of CoreWebView2Environment)))(StringComparer.OrdinalIgnoreCase)

    Private Sub New()
    End Sub

    Public Shared Function GetOrCreateAsync(userDataFolder As String, Optional options As CoreWebView2EnvironmentOptions = Nothing) As Task(Of CoreWebView2Environment)
        If String.IsNullOrWhiteSpace(userDataFolder) Then
            Throw New ArgumentException("WebView2 userDataFolder cannot be empty.", NameOf(userDataFolder))
        End If

        Directory.CreateDirectory(userDataFolder)
        Dim key = BuildKey(userDataFolder, options)

        Return _environments.GetOrAdd(
            key,
            Function(cacheKey)
                Return New Lazy(Of Task(Of CoreWebView2Environment))(
                    Function() CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, options))
            End Function).Value
    End Function

    Public Shared Async Function PrewarmDefaultAsync() As Task
        Try
            Dim defaultFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MyAppWebView2Cache")
            Await GetOrCreateAsync(defaultFolder)
        Catch ex As Exception
            Debug.WriteLine($"[WebView2EnvironmentCache] Prewarm default environment failed: {ex.Message}")
        End Try
    End Function

    Private Shared Function BuildKey(userDataFolder As String, options As CoreWebView2EnvironmentOptions) As String
        Dim fullPath = Path.GetFullPath(userDataFolder)
        Dim args = If(options?.AdditionalBrowserArguments, String.Empty)
        Return $"{fullPath}|{args}"
    End Function
End Class
