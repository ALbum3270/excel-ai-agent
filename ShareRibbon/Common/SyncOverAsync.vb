' ShareRibbon\Common\SyncOverAsync.vb
' Safe sync-over-async helpers for VSTO/WinForms (P0-3).
' Always runs the async body on the thread pool so it does not capture the UI SynchronizationContext.

Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks

''' <summary>
''' Prefer async APIs. Use these helpers only at true sync boundaries
''' (ExcelDna UDF, legacy sync interfaces, fire-and-forget bridges).
''' Never call from the UI thread for long network work without a timeout.
''' </summary>
Public NotInheritable Class SyncOverAsync
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Runs an async function on the thread pool and blocks the caller until completion or timeout.
    ''' Does not deadlock WinForms/VSTO UI SynchronizationContext.
    ''' </summary>
    Public Shared Function Run(Of T)(asyncFunc As Func(Of Task(Of T)), Optional timeoutMs As Integer = 30000) As T
        If asyncFunc Is Nothing Then Return Nothing

        Dim work = Task.Run(Async Function()
                                Return Await asyncFunc().ConfigureAwait(False)
                            End Function)

        If timeoutMs > 0 Then
            If Not work.Wait(timeoutMs) Then
                Debug.WriteLine($"[SyncOverAsync] Timed out after {timeoutMs}ms")
                Return Nothing
            End If
        Else
            work.Wait()
        End If

        If work.IsFaulted Then
            Dim baseEx = work.Exception?.GetBaseException()
            If baseEx IsNot Nothing Then Throw baseEx
            Throw work.Exception
        End If

        Return work.Result
    End Function

    ''' <summary>
    ''' Runs an async action on the thread pool and blocks until completion or timeout.
    ''' </summary>
    Public Shared Sub Run(asyncAction As Func(Of Task), Optional timeoutMs As Integer = 30000)
        If asyncAction Is Nothing Then Return

        Dim work = Task.Run(Async Function()
                                Await asyncAction().ConfigureAwait(False)
                                Return True
                            End Function)

        If timeoutMs > 0 Then
            If Not work.Wait(timeoutMs) Then
                Debug.WriteLine($"[SyncOverAsync] Timed out after {timeoutMs}ms")
                Return
            End If
        Else
            work.Wait()
        End If

        If work.IsFaulted Then
            Dim baseEx = work.Exception?.GetBaseException()
            If baseEx IsNot Nothing Then Throw baseEx
            Throw work.Exception
        End If
    End Sub
End Class
