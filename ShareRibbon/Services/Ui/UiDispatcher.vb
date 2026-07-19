' ShareRibbon\Services\Ui\UiDispatcher.vb
' Centralized WinForms UI-thread dispatch helper.

Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class UiDispatcher
    Private Const DefaultHandleTimeoutMs As Integer = 5000

    ''' <summary>
    ''' Synchronous UI marshal for background threads that must touch WinForms/WebView2/COM.
    ''' Prefer InvokeAsync. Do not call long network work inside the action.
    ''' Safe when already on the UI thread (runs inline). Uses Control.Invoke when marshaling,
    ''' never GetAwaiter().GetResult() on an async UI dispatch (avoids STA deadlocks).
    ''' </summary>
    Public Shared Sub InvokeSync(control As Control, action As Action, Optional handleTimeoutMs As Integer = DefaultHandleTimeoutMs)
        If action Is Nothing Then Return
        If control Is Nothing OrElse control.IsDisposed Then Return

        If Not control.InvokeRequired Then
            action()
            Return
        End If

        If Not WaitForHandleSync(control, handleTimeoutMs) Then
            Throw New InvalidOperationException("Cannot dispatch to UI thread before the control handle is created.")
        End If
        If control.IsDisposed Then Return

        control.Invoke(New MethodInvoker(Sub()
                                             If control.IsDisposed Then Return
                                             action()
                                         End Sub))
    End Sub

    ''' <summary>
    ''' Synchronous UI marshal with a return value. This is the single entry used by
    ''' background Agent execution before it touches WinForms/WebView2/Office COM.
    ''' </summary>
    Public Shared Function InvokeSync(Of T)(control As Control,
                                            func As Func(Of T),
                                            Optional handleTimeoutMs As Integer = DefaultHandleTimeoutMs) As T
        If func Is Nothing Then Return Nothing
        If control Is Nothing OrElse control.IsDisposed Then Return Nothing

        If Not control.InvokeRequired Then
            Return func()
        End If

        If Not WaitForHandleSync(control, handleTimeoutMs) Then
            Throw New InvalidOperationException("Cannot dispatch to UI thread before the control handle is created.")
        End If
        If control.IsDisposed Then Return Nothing

        Dim result As Object = control.Invoke(New Func(Of T)(Function()
                                                                  If control.IsDisposed Then Return Nothing
                                                                  Return func()
                                                              End Function))
        If result Is Nothing Then Return Nothing
        Return CType(result, T)
    End Function

    Private Shared Function WaitForHandleSync(control As Control, timeoutMs As Integer) As Boolean
        If control Is Nothing OrElse control.IsDisposed Then Return False
        If control.IsHandleCreated Then Return True

        Dim waited As Integer = 0
        Dim delayMs As Integer = 25
        Dim maxWait As Integer = If(timeoutMs > 0, timeoutMs, DefaultHandleTimeoutMs)
        While waited < maxWait
            If control Is Nothing OrElse control.IsDisposed Then Return False
            If control.IsHandleCreated Then Return True
            Threading.Thread.Sleep(delayMs)
            waited += delayMs
        End While
        Debug.WriteLine($"[UiDispatcher] Sync wait for handle timed out: {control.GetType().Name}")
        Return control IsNot Nothing AndAlso Not control.IsDisposed AndAlso control.IsHandleCreated
    End Function

    Public Shared Async Function InvokeAsync(control As Control, action As Action, Optional handleTimeoutMs As Integer = DefaultHandleTimeoutMs) As Task
        If action Is Nothing Then Return
        If control Is Nothing OrElse control.IsDisposed Then Return

        If Not control.InvokeRequired Then
            action()
            Return
        End If

        Await WaitForHandleAsync(control, handleTimeoutMs)
        If control.IsDisposed Then Return
        If Not control.IsHandleCreated Then
            Throw New InvalidOperationException("Cannot dispatch to UI thread before the control handle is created.")
        End If

        Dim tcs As New TaskCompletionSource(Of Boolean)()
        control.BeginInvoke(New MethodInvoker(Sub()
                                                  Try
                                                      If control.IsDisposed Then
                                                          tcs.TrySetResult(True)
                                                          Return
                                                      End If
                                                      action()
                                                      tcs.TrySetResult(True)
                                                  Catch ex As Exception
                                                      tcs.TrySetException(ex)
                                                  End Try
                                              End Sub))
        Await tcs.Task
    End Function

    Public Shared Async Function InvokeAsync(control As Control, asyncAction As Func(Of Task), Optional handleTimeoutMs As Integer = DefaultHandleTimeoutMs) As Task
        If asyncAction Is Nothing Then Return
        If control Is Nothing OrElse control.IsDisposed Then Return

        If Not control.InvokeRequired Then
            Await asyncAction()
            Return
        End If

        Await WaitForHandleAsync(control, handleTimeoutMs)
        If control.IsDisposed Then Return
        If Not control.IsHandleCreated Then
            Throw New InvalidOperationException("Cannot dispatch to UI thread before the control handle is created.")
        End If

        Dim tcs As New TaskCompletionSource(Of Boolean)()
        control.BeginInvoke(New MethodInvoker(Async Sub()
                                                  Try
                                                      If control.IsDisposed Then
                                                          tcs.TrySetResult(True)
                                                          Return
                                                      End If
                                                      Await asyncAction()
                                                      tcs.TrySetResult(True)
                                                  Catch ex As Exception
                                                      tcs.TrySetException(ex)
                                                  End Try
                                              End Sub))
        Await tcs.Task
    End Function

    Public Shared Async Function InvokeAsync(Of T)(control As Control, asyncFunc As Func(Of Task(Of T)), Optional handleTimeoutMs As Integer = DefaultHandleTimeoutMs) As Task(Of T)
        If asyncFunc Is Nothing Then Return Nothing
        If control Is Nothing OrElse control.IsDisposed Then Return Nothing

        If Not control.InvokeRequired Then
            Return Await asyncFunc()
        End If

        Await WaitForHandleAsync(control, handleTimeoutMs)
        If control.IsDisposed Then Return Nothing
        If Not control.IsHandleCreated Then
            Throw New InvalidOperationException("Cannot dispatch to UI thread before the control handle is created.")
        End If

        Dim tcs As New TaskCompletionSource(Of T)()
        control.BeginInvoke(New MethodInvoker(Async Sub()
                                                  Try
                                                      If control.IsDisposed Then
                                                          tcs.TrySetResult(Nothing)
                                                          Return
                                                      End If
                                                      Dim result = Await asyncFunc()
                                                      tcs.TrySetResult(result)
                                                  Catch ex As Exception
                                                      tcs.TrySetException(ex)
                                                  End Try
                                              End Sub))
        Return Await tcs.Task
    End Function

    Public Shared Async Function WaitForHandleAsync(control As Control, Optional timeoutMs As Integer = DefaultHandleTimeoutMs) As Task(Of Boolean)
        If control Is Nothing OrElse control.IsDisposed Then Return False
        If control.IsHandleCreated Then Return True

        Dim waited As Integer = 0
        Dim delayMs As Integer = 25
        Dim maxWait As Integer = If(timeoutMs > 0, timeoutMs, DefaultHandleTimeoutMs)

        While waited < maxWait
            If control Is Nothing OrElse control.IsDisposed Then Return False
            If control.IsHandleCreated Then Return True
            Await Task.Delay(delayMs)
            waited += delayMs
        End While

        Debug.WriteLine($"[UiDispatcher] Timed out waiting for control handle: {control.GetType().Name}")
        Return control IsNot Nothing AndAlso Not control.IsDisposed AndAlso control.IsHandleCreated
    End Function
End Class
