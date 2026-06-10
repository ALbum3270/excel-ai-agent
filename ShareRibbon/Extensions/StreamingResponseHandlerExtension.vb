Imports ShareRibbon.Agent
Imports System.Text

Namespace Extensions

    ''' <summary>
    ''' StreamingResponseHandler 集成扩展
    ''' 为 AI 响应提供流式处理支持
    ''' </summary>
    Public Module StreamingResponseHandlerExtension

        ''' <summary>
        ''' 创建流式响应处理器
        ''' </summary>
        Public Function CreateStreamingHandler() As StreamingResponseHandler
            Dim handler As New StreamingResponseHandler()

            ' 添加默认事件处理
            AddHandler handler.ContentReceived, AddressOf OnContentReceived
            AddHandler handler.StreamingCompleted, AddressOf OnStreamingCompleted
            AddHandler handler.StreamingError, AddressOf OnStreamingError

            Return handler
        End Function

        ''' <summary>
        ''' 内容接收事件处理（默认）
        ''' </summary>
        Private Sub OnContentReceived(sender As Object, e As StreamingResponseHandler.StreamingEventArgs)
            System.Diagnostics.Debug.WriteLine("[StreamingHandler] 接收内容片段: " & e.Chunk.Length & " 字符")
        End Sub

        ''' <summary>
        ''' 流式完成事件处理（默认）
        ''' </summary>
        Private Sub OnStreamingCompleted(sender As Object, e As StreamingResponseHandler.StreamingEventArgs)
            System.Diagnostics.Debug.WriteLine("[StreamingHandler] 流式响应完成，总长度: " & e.AccumulatedContent.Length & " 字符")
        End Sub

        ''' <summary>
        ''' 流式错误事件处理（默认）
        ''' </summary>
        Private Sub OnStreamingError(sender As Object, e As StreamingResponseHandler.StreamingEventArgs)
            System.Diagnostics.Debug.WriteLine("[StreamingHandler] 流式响应错误: " & e.ErrorMessage)
        End Sub

        ''' <summary>
        ''' 处理流式响应
        ''' </summary>
        Public Sub ProcessStreamingResponse(
            handler As StreamingResponseHandler,
            chunks As IEnumerable(Of String),
            Optional onChunk As Action(Of String) = Nothing,
            Optional onComplete As Action(Of String) = Nothing,
            Optional onError As Action(Of String) = Nothing)

            Try
                ' 开始流式传输
                handler.StartStreaming()

                ' 处理每个片段
                For Each chunk In chunks
                    handler.ProcessChunk(chunk)

                    ' 调用自定义处理器
                    If onChunk IsNot Nothing Then
                        onChunk(chunk)
                    End If
                Next

                ' 完成流式传输
                handler.CompleteStreaming()

                ' 调用完成处理器
                If onComplete IsNot Nothing Then
                    onComplete(handler.IsStreaming.ToString())
                End If

            Catch ex As Exception
                ' 处理错误
                handler.ErrorStreaming(ex.Message)

                ' 调用错误处理器
                If onError IsNot Nothing Then
                    onError(ex.Message)
                End If
            End Try
        End Sub

        ''' <summary>
        ''' 检查是否支持流式响应
        ''' </summary>
        Public Function IsStreamingSupported() As Boolean
            ' 当前所有 Office 应用都支持流式响应
            Return True
        End Function

        ''' <summary>
        ''' 获取流式响应建议
        ''' </summary>
        Public Function GetStreamingRecommendation(contentLength As Integer) As String
            If contentLength < 100 Then
                Return "内容较短，建议使用普通响应模式"
            ElseIf contentLength < 1000 Then
                Return "内容适中，流式和普通模式均可"
            Else
                Return "内容较长，建议使用流式响应模式以提升用户体验"
            End If
        End Function

    End Module

End Namespace
