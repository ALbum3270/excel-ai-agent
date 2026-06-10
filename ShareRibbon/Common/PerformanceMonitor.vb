Imports System.Diagnostics

Namespace Common

    ''' <summary>
    ''' 性能监控器 - 监控操作性能和资源使用
    ''' </summary>
    Public Class PerformanceMonitor

        ''' <summary>
        ''' 性能计数器
        ''' </summary>
        Public Class PerformanceCounter
            Public Property Name As String
            Public Property StartTime As DateTime
            Public Property EndTime As DateTime
            Public Property Duration As TimeSpan
            Public Property MemoryBefore As Long
            Public Property MemoryAfter As Long
            Public Property MemoryDelta As Long
            Public Property Success As Boolean
            Public Property ErrorMessage As String

            Public ReadOnly Property DurationMs As Double
                Get
                    Return Duration.TotalMilliseconds
                End Get
            End Property
        End Class

        Private Shared _counters As New List(Of PerformanceCounter)()
        Private Shared _lockObj As New Object()
        Private Shared _maxCounters As Integer = 1000
        Private Shared _enableMonitoring As Boolean = True

        ''' <summary>
        ''' 启用/禁用性能监控
        ''' </summary>
        Public Shared Property EnableMonitoring As Boolean
            Get
                Return _enableMonitoring
            End Get
            Set(value As Boolean)
                _enableMonitoring = value
            End Set
        End Property

        ''' <summary>
        ''' 开始监控操作
        ''' </summary>
        Public Shared Function StartOperation(operationName As String) As PerformanceCounter
            If Not _enableMonitoring Then
                Return Nothing
            End If

            Dim counter As New PerformanceCounter With {
                .Name = operationName,
                .StartTime = DateTime.Now,
                .MemoryBefore = GC.GetTotalMemory(False)
            }

            Return counter
        End Function

        ''' <summary>
        ''' 结束监控操作
        ''' </summary>
        Public Shared Sub EndOperation(counter As PerformanceCounter, Optional success As Boolean = True, Optional errorMessage As String = Nothing)
            If counter Is Nothing OrElse Not _enableMonitoring Then
                Return
            End If

            counter.EndTime = DateTime.Now
            counter.Duration = counter.EndTime - counter.StartTime
            counter.MemoryAfter = GC.GetTotalMemory(False)
            counter.MemoryDelta = counter.MemoryAfter - counter.MemoryBefore
            counter.Success = success
            counter.ErrorMessage = errorMessage

            SyncLock _lockObj
                _counters.Add(counter)

                ' 限制计数器数量
                If _counters.Count > _maxCounters Then
                    _counters.RemoveAt(0)
                End If
            End SyncLock

            ' 记录日志
            Debug.WriteLine(String.Format("[PerformanceMonitor] {0}: {1:F2}ms, Memory: {2:F2}KB",
                counter.Name,
                counter.DurationMs,
                counter.MemoryDelta / 1024.0))
        End Sub

        ''' <summary>
        ''' 监控操作（自动计时）
        ''' </summary>
        Public Shared Function MonitorOperation(Of T)(operationName As String, operation As Func(Of T)) As T
            If Not _enableMonitoring Then
                Return operation()
            End If

            Dim counter As PerformanceCounter = StartOperation(operationName)
            Dim result As T = Nothing
            Dim success As Boolean = True
            Dim errorMsg As String = Nothing

            Try
                result = operation()
            Catch ex As Exception
                success = False
                errorMsg = ex.Message
                Throw
            Finally
                EndOperation(counter, success, errorMsg)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 监控操作（无返回值）
        ''' </summary>
        Public Shared Sub MonitorOperation(operationName As String, operation As Action)
            If Not _enableMonitoring Then
                operation()
                Return
            End If

            Dim counter As PerformanceCounter = StartOperation(operationName)
            Dim success As Boolean = True
            Dim errorMsg As String = Nothing

            Try
                operation()
            Catch ex As Exception
                success = False
                errorMsg = ex.Message
                Throw
            Finally
                EndOperation(counter, success, errorMsg)
            End Try
        End Sub

        ''' <summary>
        ''' 获取性能统计
        ''' </summary>
        Public Shared Function GetStatistics(Optional operationName As String = Nothing) As String
            SyncLock _lockObj
                Dim relevantCounters As List(Of PerformanceCounter)

                If String.IsNullOrEmpty(operationName) Then
                    relevantCounters = _counters
                Else
                    relevantCounters = _counters.Where(Function(c) c.Name = operationName).ToList()
                End If

                If relevantCounters.Count = 0 Then
                    Return "没有性能数据"
                End If

                Dim avgDuration As Double = relevantCounters.Average(Function(c) c.DurationMs)
                Dim maxDuration As Double = relevantCounters.Max(Function(c) c.DurationMs)
                Dim minDuration As Double = relevantCounters.Min(Function(c) c.DurationMs)
                Dim avgMemory As Double = relevantCounters.Average(Function(c) c.MemoryDelta / 1024.0)

                Dim successCount As Integer = 0
                For Each c In relevantCounters
                    If c.Success Then
                        successCount += 1
                    End If
                Next
                Dim successRate As Double = successCount / relevantCounters.Count * 100

                Dim stats As String = String.Format(
                    "性能统计 - 操作: {0}, 次数: {1}" & vbCrLf &
                    "  平均耗时: {2:F2}ms, 最大: {3:F2}ms, 最小: {4:F2}ms" & vbCrLf &
                    "  平均内存: {5:F2}KB, 成功率: {6:F1}%",
                    If(String.IsNullOrEmpty(operationName), "所有", operationName),
                    relevantCounters.Count,
                    avgDuration,
                    maxDuration,
                    minDuration,
                    avgMemory,
                    successRate)

                Return stats
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取慢操作列表
        ''' </summary>
        Public Shared Function GetSlowOperations(thresholdMs As Double) As List(Of PerformanceCounter)
            SyncLock _lockObj
                Return _counters.Where(Function(c) c.DurationMs > thresholdMs).OrderByDescending(Function(c) c.DurationMs).ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取失败操作列表
        ''' </summary>
        Public Shared Function GetFailedOperations() As List(Of PerformanceCounter)
            SyncLock _lockObj
                Return _counters.Where(Function(c) Not c.Success).ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 清除性能数据
        ''' </summary>
        Public Shared Sub ClearData()
            SyncLock _lockObj
                _counters.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 性能监控作用域（使用 Using 模式）
        ''' </summary>
        Public Class PerformanceScope
            Implements IDisposable

            Private _counter As PerformanceCounter
            Private _disposed As Boolean = False
            Private _success As Boolean = True

            Public Sub New(operationName As String)
                _counter = PerformanceMonitor.StartOperation(operationName)
            End Sub

            ''' <summary>
            ''' 标记操作失败
            ''' </summary>
            Public Sub MarkFailed(errorMessage As String)
                _success = False
                If _counter IsNot Nothing Then
                    _counter.ErrorMessage = errorMessage
                End If
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub

            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not _disposed Then
                    If disposing Then
                        PerformanceMonitor.EndOperation(_counter, _success)
                    End If
                    _disposed = True
                End If
            End Sub
        End Class

        ''' <summary>
        ''' 批处理性能监控
        ''' </summary>
        Public Class BatchMonitor

            Private _batchName As String
            Private _items As New List(Of PerformanceCounter)()
            Private _startTime As DateTime
            Private _endTime As DateTime

            Public Sub New(batchName As String)
                _batchName = batchName
                _startTime = DateTime.Now
            End Sub

            ''' <summary>
            ''' 添加子操作
            ''' </summary>
            Public Sub AddItem(counter As PerformanceCounter)
                If counter IsNot Nothing Then
                    _items.Add(counter)
                End If
            End Sub

            ''' <summary>
            ''' 完成批处理
            ''' </summary>
            Public Function Complete() As String
                _endTime = DateTime.Now
                Dim totalDuration As TimeSpan = _endTime - _startTime

                Dim avgDuration As Double = 0
                If _items.Count > 0 Then
                    avgDuration = _items.Average(Function(i) i.DurationMs)
                End If

                Dim successCount As Integer = 0
                Dim failCount As Integer = 0
                For Each item In _items
                    If item.Success Then
                        successCount += 1
                    Else
                        failCount += 1
                    End If
                Next

                Dim stats As String = String.Format(
                    "批处理 {0} 完成" & vbCrLf &
                    "  总耗时: {1:F2}ms" & vbCrLf &
                    "  项目数: {2}" & vbCrLf &
                    "  平均耗时: {3:F2}ms" & vbCrLf &
                    "  成功: {4}, 失败: {5}",
                    _batchName,
                    totalDuration.TotalMilliseconds,
                    _items.Count,
                    avgDuration,
                    successCount,
                    failCount)

                Debug.WriteLine(String.Format("[PerformanceMonitor] {0}", stats))
                Return stats
            End Function

        End Class

    End Class

End Namespace
