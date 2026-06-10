Imports System.Runtime.InteropServices

Namespace Common

    ''' <summary>
    ''' COM 对象辅助类 - 统一管理 Office Interop COM 对象的生命周期
    ''' </summary>
    Public Class ComObjectHelper

        ''' <summary>
        ''' 安全释放 COM 对象
        ''' </summary>
        ''' <param name="obj">要释放的 COM 对象</param>
        Public Shared Sub ReleaseComObject(ByRef obj As Object)
            Try
                If obj IsNot Nothing Then
                    ' 释放 COM 对象
                    Marshal.ReleaseComObject(obj)
                    obj = Nothing
                End If
            Catch ex As Exception
                ' 释放失败时记录日志，但不抛出异常
                System.Diagnostics.Debug.WriteLine(String.Format("[ComObjectHelper] 释放 COM 对象失败: {0}", ex.Message))
            Finally
                obj = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' 批量释放 COM 对象
        ''' </summary>
        ''' <param name="objects">要释放的 COM 对象数组</param>
        Public Shared Sub ReleaseComObjects(ParamArray objects As Object())
            If objects Is Nothing Then
                Return
            End If

            For Each obj In objects
                If obj IsNot Nothing Then
                    ReleaseComObject(obj)
                End If
            Next
        End Sub

        ''' <summary>
        ''' 强制垃圾回收（谨慎使用）
        ''' </summary>
        ''' <remarks>
        ''' 仅在释放大量 COM 对象后使用
        ''' 频繁调用会影响性能
        ''' </remarks>
        Public Shared Sub ForceGarbageCollection()
            Try
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
                System.Diagnostics.Debug.WriteLine("[ComObjectHelper] 强制垃圾回收完成")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[ComObjectHelper] 垃圾回收失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 使用 Using 模式管理 COM 对象（适用于实现了 IDisposable 的对象）
        ''' </summary>
        Public Class ComObjectScope
            Implements IDisposable

            Private _objects As List(Of Object)
            Private _disposed As Boolean = False

            Public Sub New()
                _objects = New List(Of Object)()
            End Sub

            ''' <summary>
            ''' 跟踪 COM 对象
            ''' </summary>
            Public Function Track(obj As Object) As Object
                If obj IsNot Nothing Then
                    _objects.Add(obj)
                End If
                Return obj
            End Function

            ''' <summary>
            ''' 释放所有跟踪的 COM 对象
            ''' </summary>
            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub

            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not _disposed Then
                    If disposing Then
                        ' 释放所有跟踪的 COM 对象
                        For Each obj In _objects
                            ReleaseComObject(obj)
                        Next
                        _objects.Clear()
                    End If
                    _disposed = True
                End If
            End Sub

            Protected Overrides Sub Finalize()
                Dispose(False)
            End Sub
        End Class

        ''' <summary>
        ''' Excel Range 专用释放（处理多层嵌套）
        ''' </summary>
        Public Shared Sub ReleaseExcelRange(ByRef range As Object)
            Try
                If range IsNot Nothing Then
                    ' Excel Range 可能包含嵌套的 COM 对象
                    ' 需要特殊处理
                    Marshal.FinalReleaseComObject(range)
                    range = Nothing
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[ComObjectHelper] 释放 Excel Range 失败: {0}", ex.Message))
            Finally
                range = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' 检查对象是否为 COM 对象
        ''' </summary>
        Public Shared Function IsComObject(obj As Object) As Boolean
            If obj Is Nothing Then
                Return False
            End If

            Try
                Return Marshal.IsComObject(obj)
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取 COM 对象的引用计数（用于调试）
        ''' </summary>
        Public Shared Function GetComReferenceCount(obj As Object) As Integer
            If obj Is Nothing OrElse Not IsComObject(obj) Then
                Return 0
            End If

            Try
                ' 增加引用计数
                Dim count As Integer = Marshal.AddRef(Marshal.GetIUnknownForObject(obj))
                ' 减少引用计数（还原）
                Marshal.Release(Marshal.GetIUnknownForObject(obj))
                Return count - 1
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[ComObjectHelper] 获取引用计数失败: {0}", ex.Message))
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' COM 对象使用统计（用于性能监控）
        ''' </summary>
        Public Class ComObjectStats
            Private Shared _createdCount As Integer = 0
            Private Shared _releasedCount As Integer = 0
            Private Shared _lockObj As New Object()

            ''' <summary>
            ''' 记录创建
            ''' </summary>
            Public Shared Sub RecordCreated()
                SyncLock _lockObj
                    _createdCount += 1
                End SyncLock
            End Sub

            ''' <summary>
            ''' 记录释放
            ''' </summary>
            Public Shared Sub RecordReleased()
                SyncLock _lockObj
                    _releasedCount += 1
                End SyncLock
            End Sub

            ''' <summary>
            ''' 获取统计信息
            ''' </summary>
            Public Shared Function GetStats() As String
                SyncLock _lockObj
                    Dim leaked As Integer = _createdCount - _releasedCount
                    Return String.Format("COM对象统计 - 创建: {0}, 释放: {1}, 疑似泄漏: {2}", _createdCount, _releasedCount, leaked)
                End SyncLock
            End Function

            ''' <summary>
            ''' 重置统计
            ''' </summary>
            Public Shared Sub Reset()
                SyncLock _lockObj
                    _createdCount = 0
                    _releasedCount = 0
                End SyncLock
            End Sub
        End Class

    End Class

End Namespace
