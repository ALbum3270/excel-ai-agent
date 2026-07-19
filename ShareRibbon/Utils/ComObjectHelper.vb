Imports System.Diagnostics
Imports System.Runtime.InteropServices

Public NotInheritable Class ComObjectHelper
    Private Sub New()
    End Sub

    Public Shared Sub ReleaseComObject(ByRef comObject As Object)
        If comObject Is Nothing Then Return

        Try
            If Marshal.IsComObject(comObject) Then
                Marshal.ReleaseComObject(comObject)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[ComObjectHelper] ReleaseComObject failed: {ex.Message}")
        Finally
            comObject = Nothing
        End Try
    End Sub

    Public Shared Sub ReleaseComObject(Of T As Class)(ByRef comObject As T)
        If comObject Is Nothing Then Return

        Try
            Dim boxed As Object = comObject
            If Marshal.IsComObject(boxed) Then
                Marshal.ReleaseComObject(boxed)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[ComObjectHelper] ReleaseComObject failed: {ex.Message}")
        Finally
            comObject = Nothing
        End Try
    End Sub

End Class
