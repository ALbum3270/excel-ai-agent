Imports System.Text
Imports System.Threading

Namespace Agent

    ''' <summary>
    ''' 流式响应处理器 - 渐进式显示 AI 响应内容
    ''' </summary>
    Public Class StreamingResponseHandler

        ''' <summary>
        ''' 流式响应事件参数
        ''' </summary>
        Public Class StreamingEventArgs
            Inherits EventArgs

            ''' <summary>
            ''' 当前接收到的内容片段
            ''' </summary>
            Public Property Chunk As String

            ''' <summary>
            ''' 累积的完整内容
            ''' </summary>
            Public Property AccumulatedContent As String

            ''' <summary>
            ''' 是否已完成
            ''' </summary>
            Public Property IsComplete As Boolean

            ''' <summary>
            ''' 错误信息（如果有）
            ''' </summary>
            Public Property ErrorMessage As String
        End Class

        ''' <summary>
        ''' 流式内容接收事件
        ''' </summary>
        Public Event ContentReceived As EventHandler(Of StreamingEventArgs)

        ''' <summary>
        ''' 流式响应完成事件
        ''' </summary>
        Public Event StreamingCompleted As EventHandler(Of StreamingEventArgs)

        ''' <summary>
        ''' 流式响应错误事件
        ''' </summary>
        Public Event StreamingError As EventHandler(Of StreamingEventArgs)

        Private _accumulatedContent As StringBuilder
        Private _isStreaming As Boolean
        Private _cancellationTokenSource As CancellationTokenSource
        Private _lockObj As New Object()

        Public Sub New()
            _accumulatedContent = New StringBuilder()
            _isStreaming = False
        End Sub

        ''' <summary>
        ''' 是否正在流式传输
        ''' </summary>
        Public ReadOnly Property IsStreaming As Boolean
            Get
                SyncLock _lockObj
                    Return _isStreaming
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' 开始流式响应
        ''' </summary>
        Public Sub StartStreaming()
            SyncLock _lockObj
                If _isStreaming Then
                    Throw New InvalidOperationException("流式响应已在进行中")
                End If

                _isStreaming = True
                _accumulatedContent.Clear()
                _cancellationTokenSource = New CancellationTokenSource()
            End SyncLock

            System.Diagnostics.Debug.WriteLine("[StreamingResponseHandler] 开始流式响应")
        End Sub

        ''' <summary>
        ''' 处理接收到的内容片段
        ''' </summary>
        Public Sub ProcessChunk(chunk As String)
            SyncLock _lockObj
                If Not _isStreaming Then
                    Return
                End If

                If _cancellationTokenSource.IsCancellationRequested Then
                    Return
                End If

                ' 累积内容
                _accumulatedContent.Append(chunk)

                ' 触发事件
                Dim args As New StreamingEventArgs With {
                    .Chunk = chunk,
                    .AccumulatedContent = _accumulatedContent.ToString(),
                    .IsComplete = False
                }

                RaiseEvent ContentReceived(Me, args)
            End SyncLock
        End Sub

        ''' <summary>
        ''' 完成流式响应
        ''' </summary>
        Public Sub CompleteStreaming()
            SyncLock _lockObj
                If Not _isStreaming Then
                    Return
                End If

                _isStreaming = False

                Dim args As New StreamingEventArgs With {
                    .Chunk = String.Empty,
                    .AccumulatedContent = _accumulatedContent.ToString(),
                    .IsComplete = True
                }

                RaiseEvent StreamingCompleted(Me, args)
                System.Diagnostics.Debug.WriteLine("[StreamingResponseHandler] 流式响应完成")
            End SyncLock
        End Sub

        ''' <summary>
        ''' 流式响应出错
        ''' </summary>
        Public Sub ErrorStreaming(errorMessage As String)
            SyncLock _lockObj
                If Not _isStreaming Then
                    Return
                End If

                _isStreaming = False

                Dim args As New StreamingEventArgs With {
                    .Chunk = String.Empty,
                    .AccumulatedContent = _accumulatedContent.ToString(),
                    .IsComplete = False,
                    .ErrorMessage = errorMessage
                }

                RaiseEvent StreamingError(Me, args)
                System.Diagnostics.Debug.WriteLine(String.Format("[StreamingResponseHandler] 流式响应错误: {0}", errorMessage))
            End SyncLock
        End Sub

        ''' <summary>
        ''' 取消流式响应
        ''' </summary>
        Public Sub CancelStreaming()
            SyncLock _lockObj
                If Not _isStreaming Then
                    Return
                End If

                If _cancellationTokenSource IsNot Nothing Then
                    _cancellationTokenSource.Cancel()
                End If

                _isStreaming = False
                System.Diagnostics.Debug.WriteLine("[StreamingResponseHandler] 流式响应已取消")
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取当前累积的内容
        ''' </summary>
        Public Function GetAccumulatedContent() As String
            SyncLock _lockObj
                Return _accumulatedContent.ToString()
            End SyncLock
        End Function

        ''' <summary>
        ''' 重置处理器
        ''' </summary>
        Public Sub Reset()
            SyncLock _lockObj
                _isStreaming = False
                _accumulatedContent.Clear()

                If _cancellationTokenSource IsNot Nothing Then
                    _cancellationTokenSource.Dispose()
                    _cancellationTokenSource = Nothing
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 流式响应模拟器（用于测试）
        ''' </summary>
        Public Class StreamingSimulator

            Private _handler As StreamingResponseHandler
            Private _content As String
            Private _chunkSize As Integer
            Private _delayMs As Integer

            Public Sub New(handler As StreamingResponseHandler, content As String, Optional chunkSize As Integer = 10, Optional delayMs As Integer = 50)
                _handler = handler
                _content = content
                _chunkSize = chunkSize
                _delayMs = delayMs
            End Sub

            ''' <summary>
            ''' 模拟流式传输
            ''' </summary>
            Public Sub Simulate()
                _handler.StartStreaming()

                Dim position As Integer = 0
                While position < _content.Length
                    ' 获取当前片段
                    Dim remaining As Integer = _content.Length - position
                    Dim currentChunkSize As Integer = Math.Min(_chunkSize, remaining)
                    Dim chunk As String = _content.Substring(position, currentChunkSize)

                    ' 处理片段
                    _handler.ProcessChunk(chunk)

                    ' 延迟
                    Thread.Sleep(_delayMs)

                    position += currentChunkSize
                End While

                _handler.CompleteStreaming()
            End Sub

            ''' <summary>
            ''' 异步模拟流式传输
            ''' </summary>
            Public Async Function SimulateAsync() As Task
                Await Task.Run(Sub() Simulate())
            End Function

        End Class

        ''' <summary>
        ''' 流式响应缓冲器（优化显示性能）
        ''' </summary>
        Public Class StreamingBuffer

            Private _buffer As StringBuilder
            Private _flushThreshold As Integer
            Private _flushInterval As Integer
            Private _lastFlushTime As DateTime
            Private _lockObj As New Object()

            Public Event BufferFlushed As EventHandler(Of String)

            Public Sub New(Optional flushThreshold As Integer = 100, Optional flushIntervalMs As Integer = 200)
                _buffer = New StringBuilder()
                _flushThreshold = flushThreshold
                _flushInterval = flushIntervalMs
                _lastFlushTime = DateTime.Now
            End Sub

            ''' <summary>
            ''' 添加内容到缓冲区
            ''' </summary>
            Public Sub Add(chunk As String)
                SyncLock _lockObj
                    _buffer.Append(chunk)

                    ' 检查是否需要刷新
                    Dim shouldFlush As Boolean = False

                    ' 条件1: 缓冲区大小超过阈值
                    If _buffer.Length >= _flushThreshold Then
                        shouldFlush = True
                    End If

                    ' 条件2: 距离上次刷新超过时间间隔
                    Dim elapsed As TimeSpan = DateTime.Now - _lastFlushTime
                    If elapsed.TotalMilliseconds >= _flushInterval Then
                        shouldFlush = True
                    End If

                    If shouldFlush Then
                        Flush()
                    End If
                End SyncLock
            End Sub

            ''' <summary>
            ''' 强制刷新缓冲区
            ''' </summary>
            Public Sub Flush()
                SyncLock _lockObj
                    If _buffer.Length > 0 Then
                        Dim content As String = _buffer.ToString()
                        _buffer.Clear()
                        _lastFlushTime = DateTime.Now

                        RaiseEvent BufferFlushed(Me, content)
                    End If
                End SyncLock
            End Sub

            ''' <summary>
            ''' 重置缓冲区
            ''' </summary>
            Public Sub Reset()
                SyncLock _lockObj
                    _buffer.Clear()
                    _lastFlushTime = DateTime.Now
                End SyncLock
            End Sub

        End Class

    End Class

End Namespace
