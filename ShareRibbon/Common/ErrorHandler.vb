Imports System.Windows.Forms

Namespace Common

    ''' <summary>
    ''' 错误处理辅助类 - 统一的错误处理和用户提示
    ''' </summary>
    Public Class ErrorHandler

        ''' <summary>
        ''' 错误级别
        ''' </summary>
        Public Enum ErrorLevel
            Info = 0        ' 信息
            Warning = 1     ' 警告
            [Error] = 2     ' 错误
            Critical = 3    ' 严重错误
        End Enum

        ''' <summary>
        ''' 错误上下文
        ''' </summary>
        Public Class ErrorContext
            Public Property Level As ErrorLevel
            Public Property Message As String
            Public Property Exception As Exception
            Public Property ModuleName As String
            Public Property Operation As String
            Public Property Timestamp As DateTime
            Public Property UserMessage As String

            Public Sub New()
                Timestamp = DateTime.Now
            End Sub
        End Class

        Private Shared _enableLogging As Boolean = True
        Private Shared _showUserMessages As Boolean = True
        Private Shared _errorHistory As New List(Of ErrorContext)()
        Private Shared _maxHistorySize As Integer = 100
        Private Shared _lockObj As New Object()

        ''' <summary>
        ''' 启用/禁用日志记录
        ''' </summary>
        Public Shared Property EnableLogging As Boolean
            Get
                Return _enableLogging
            End Get
            Set(value As Boolean)
                _enableLogging = value
            End Set
        End Property

        ''' <summary>
        ''' 启用/禁用用户消息
        ''' </summary>
        Public Shared Property ShowUserMessages As Boolean
            Get
                Return _showUserMessages
            End Get
            Set(value As Boolean)
                _showUserMessages = value
            End Set
        End Property

        ''' <summary>
        ''' 处理错误
        ''' </summary>
        Public Shared Sub Handle(ex As Exception, moduleName As String, operationName As String, Optional userMessage As String = Nothing)
            Dim context As New ErrorContext With {
                .Level = ErrorLevel.Error,
                .Exception = ex,
                .ModuleName = moduleName,
                .Operation = operationName,
                .Message = ex.Message,
                .UserMessage = userMessage
            }

            ProcessError(context)
        End Sub

        ''' <summary>
        ''' 处理警告
        ''' </summary>
        Public Shared Sub HandleWarning(message As String, moduleName As String, Optional userMessage As String = Nothing)
            Dim context As New ErrorContext With {
                .Level = ErrorLevel.Warning,
                .ModuleName = moduleName,
                .Message = message,
                .UserMessage = userMessage
            }

            ProcessError(context)
        End Sub

        ''' <summary>
        ''' 处理信息
        ''' </summary>
        Public Shared Sub HandleInfo(message As String, moduleName As String, Optional userMessage As String = Nothing)
            Dim context As New ErrorContext With {
                .Level = ErrorLevel.Info,
                .ModuleName = moduleName,
                .Message = message,
                .UserMessage = userMessage
            }

            ProcessError(context)
        End Sub

        ''' <summary>
        ''' 处理严重错误
        ''' </summary>
        Public Shared Sub HandleCritical(ex As Exception, moduleName As String, operationName As String, Optional userMessage As String = Nothing)
            Dim context As New ErrorContext With {
                .Level = ErrorLevel.Critical,
                .Exception = ex,
                .ModuleName = moduleName,
                .Operation = operationName,
                .Message = ex.Message,
                .UserMessage = userMessage
            }

            ProcessError(context)
        End Sub

        ''' <summary>
        ''' 处理错误上下文
        ''' </summary>
        Private Shared Sub ProcessError(context As ErrorContext)
            ' 记录日志
            If _enableLogging Then
                LogError(context)
            End If

            ' 添加到历史
            AddToHistory(context)

            ' 显示用户消息
            If _showUserMessages AndAlso Not String.IsNullOrEmpty(context.UserMessage) Then
                ShowUserMessage(context)
            End If
        End Sub

        ''' <summary>
        ''' 记录错误日志
        ''' </summary>
        Private Shared Sub LogError(context As ErrorContext)
            Try
                Dim logMessage As String = String.Format("[{0}] [{1}] {2}.{3}: {4}",
                    context.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                    context.Level.ToString(),
                    context.ModuleName,
                    If(String.IsNullOrEmpty(context.Operation), "Unknown", context.Operation),
                    context.Message)

                System.Diagnostics.Debug.WriteLine(logMessage)

                ' 如果有异常，记录堆栈跟踪
                If context.Exception IsNot Nothing Then
                    System.Diagnostics.Debug.WriteLine(String.Format("  StackTrace: {0}", context.Exception.StackTrace))
                End If

            Catch ex As Exception
                ' 日志记录失败，不抛出异常
                System.Diagnostics.Debug.WriteLine(String.Format("[ErrorHandler] 日志记录失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 显示用户消息
        ''' </summary>
        Private Shared Sub ShowUserMessage(context As ErrorContext)
            Try
                Dim icon As MessageBoxIcon

                Select Case context.Level
                    Case ErrorLevel.Info
                        icon = MessageBoxIcon.Information
                    Case ErrorLevel.Warning
                        icon = MessageBoxIcon.Warning
                    Case ErrorLevel.Error
                        icon = MessageBoxIcon.Error
                    Case ErrorLevel.Critical
                        icon = MessageBoxIcon.Error
                    Case Else
                        icon = MessageBoxIcon.Information
                End Select

                Dim title As String = String.Format("AiHelper - {0}", context.Level.ToString())
                MessageBox.Show(context.UserMessage, title, MessageBoxButtons.OK, icon)

            Catch ex As Exception
                ' 消息框显示失败，不抛出异常
                System.Diagnostics.Debug.WriteLine(String.Format("[ErrorHandler] 显示消息失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 添加到错误历史
        ''' </summary>
        Private Shared Sub AddToHistory(context As ErrorContext)
            SyncLock _lockObj
                _errorHistory.Add(context)

                ' 限制历史大小
                If _errorHistory.Count > _maxHistorySize Then
                    _errorHistory.RemoveAt(0)
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取错误历史
        ''' </summary>
        Public Shared Function GetErrorHistory(Optional level As ErrorLevel? = Nothing) As List(Of ErrorContext)
            SyncLock _lockObj
                If level.HasValue Then
                    Return _errorHistory.Where(Function(e) e.Level = level.Value).ToList()
                Else
                    Return New List(Of ErrorContext)(_errorHistory)
                End If
            End SyncLock
        End Function

        ''' <summary>
        ''' 清除错误历史
        ''' </summary>
        Public Shared Sub ClearHistory()
            SyncLock _lockObj
                _errorHistory.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取错误统计
        ''' </summary>
        Public Shared Function GetErrorStats() As String
            SyncLock _lockObj
                Dim infoCount As Integer = 0
                Dim warningCount As Integer = 0
                Dim errorCount As Integer = 0
                Dim criticalCount As Integer = 0

                For Each e In _errorHistory
                    Select Case e.Level
                        Case ErrorLevel.Info
                            infoCount += 1
                        Case ErrorLevel.Warning
                            warningCount += 1
                        Case ErrorLevel.Error
                            errorCount += 1
                        Case ErrorLevel.Critical
                            criticalCount += 1
                    End Select
                Next

                Return String.Format("错误统计 - 信息: {0}, 警告: {1}, 错误: {2}, 严重: {3}",
                    infoCount, warningCount, errorCount, criticalCount)
            End SyncLock
        End Function

        ''' <summary>
        ''' 安全执行操作（带错误处理）
        ''' </summary>
        Public Shared Function SafeExecute(Of T)(action As Func(Of T), moduleName As String, operationName As String, Optional defaultValue As T = Nothing, Optional userErrorMessage As String = Nothing) As T
            Try
                Return action()
            Catch ex As Exception
                Handle(ex, moduleName, operationName, userErrorMessage)
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 安全执行操作（无返回值）
        ''' </summary>
        Public Shared Sub SafeExecute(action As Action, moduleName As String, operationName As String, Optional userErrorMessage As String = Nothing)
            Try
                action()
            Catch ex As Exception
                Handle(ex, moduleName, operationName, userErrorMessage)
            End Try
        End Sub

        ''' <summary>
        ''' 用户友好的错误消息生成
        ''' </summary>
        Public Class UserFriendlyMessages

            ''' <summary>
            ''' 网络错误消息
            ''' </summary>
            Public Shared Function NetworkError() As String
                Return "网络连接失败。请检查您的网络连接后重试。"
            End Function

            ''' <summary>
            ''' API 错误消息
            ''' </summary>
            Public Shared Function ApiError(Optional statusCode As Integer? = Nothing) As String
                If statusCode.HasValue Then
                    Select Case statusCode.Value
                        Case 401
                            Return "API 认证失败。请检查您的 API Key 是否正确。"
                        Case 429
                            Return "API 请求频率超限。请稍后重试。"
                        Case 500, 502, 503
                            Return "AI 服务暂时不可用。请稍后重试。"
                        Case Else
                            Return String.Format("AI 服务返回错误（错误码: {0}）。请稍后重试。", statusCode.Value)
                    End Select
                Else
                    Return "AI 服务调用失败。请检查网络连接和 API 配置。"
                End If
            End Function

            ''' <summary>
            ''' Office 操作错误消息
            ''' </summary>
            Public Shared Function OfficeOperationError(appType As String) As String
                Return String.Format("执行 {0} 操作失败。请确保文档处于可编辑状态。", appType)
            End Function

            ''' <summary>
            ''' 数据验证错误消息
            ''' </summary>
            Public Shared Function ValidationError(fieldName As String) As String
                Return String.Format("输入验证失败：{0} 格式不正确。", fieldName)
            End Function

            ''' <summary>
            ''' COM 对象错误消息
            ''' </summary>
            Public Shared Function ComObjectError() As String
                Return "Office 组件访问失败。请尝试重启 Office 应用。"
            End Function

            ''' <summary>
            ''' 文件操作错误消息
            ''' </summary>
            Public Shared Function FileOperationError(operation As String) As String
                Return String.Format("文件{0}失败。请检查文件权限和磁盘空间。", operation)
            End Function

        End Class

    End Class

End Namespace
