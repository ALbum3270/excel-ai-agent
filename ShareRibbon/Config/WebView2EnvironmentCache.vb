Imports System.Collections.Concurrent
Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
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
                    Function()
                        ' 确保在 STA 线程上创建 WebView2 环境
                        If SynchronizationContext.Current Is Nothing OrElse
                           Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
                            ' 当前不在 UI 线程，切换到 UI 线程
                            Dim tcs As New TaskCompletionSource(Of CoreWebView2Environment)()

                            ' 使用 Control.Invoke 确保在 UI 线程上创建
                            Dim dummyControl As Control = Nothing
                            Try
                                ' 查找任何现有的窗体或控件
                                If Application.OpenForms.Count > 0 Then
                                    dummyControl = Application.OpenForms(0)
                                End If
                            Catch
                                ' 忽略异常
                            End Try

                            If dummyControl IsNot Nothing AndAlso dummyControl.IsHandleCreated Then
                                dummyControl.BeginInvoke(
                                    Async Sub()
                                        Try
                                            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, options)
                                            tcs.SetResult(env)
                                        Catch ex As Exception
                                            tcs.SetException(ex)
                                        End Try
                                    End Sub)
                                Return tcs.Task
                            End If
                        End If

                        ' 已经在 STA 线程，直接创建
                        Return CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, options)
                    End Function)
            End Function).Value
    End Function

    Private Shared Function BuildKey(userDataFolder As String, options As CoreWebView2EnvironmentOptions) As String
        Dim fullPath = Path.GetFullPath(userDataFolder)
        Dim args = If(options?.AdditionalBrowserArguments, String.Empty)
        Return $"{fullPath}|{args}"
    End Function
End Class
