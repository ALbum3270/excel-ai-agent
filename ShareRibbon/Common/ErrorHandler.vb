Imports System.Windows.Forms

Namespace Common

    ''' <summary>
    ''' 错误处理辅助类 - 统一的错误处理和用户提示
    ''' </summary>
    Public Class ErrorHandler

        ''' <summary>
        ''' 错误上下文
        ''' </summary>
        Public Class ErrorContext
            Public Property Message As String
            Public Property Exception As Exception
            Public Property ModuleName As String
            Public Property Operation As String
            Public Property UserMessage As String
        End Class

        ''' <summary>
        ''' 处理错误
        ''' </summary>
        Public Shared Sub Handle(ex As Exception, moduleName As String, operationName As String, Optional userMessage As String = Nothing)
            Dim context As New ErrorContext With {
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
            LogError(context)

            ' 显示用户消息
            If Not String.IsNullOrEmpty(context.UserMessage) Then
                ShowUserMessage(context)
            End If
        End Sub

        ''' <summary>
        ''' 记录错误日志（P0-4：走 AppLogger，自动脱敏）
        ''' </summary>
        Private Shared Sub LogError(context As ErrorContext)
            Try
                Dim moduleName = If(String.IsNullOrEmpty(context.ModuleName), "ErrorHandler", context.ModuleName)
                Dim op = If(String.IsNullOrEmpty(context.Operation), "Unknown", context.Operation)
                Dim msg = $"{op}: {context.Message}"
                AppLogger.Error(moduleName, msg, context.Exception)
            Catch
                ' 日志记录失败，不抛出异常
            End Try
        End Sub

        ''' <summary>
        ''' 显示用户消息
        ''' </summary>
        Private Shared Sub ShowUserMessage(context As ErrorContext)
            Try
                MessageBox.Show(context.UserMessage, "AiHelper - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Catch ex As Exception
                ' 消息框显示失败，不抛出异常
                System.Diagnostics.Debug.WriteLine(String.Format("[ErrorHandler] 显示消息失败: {0}", ex.Message))
            End Try
        End Sub

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

    End Class

End Namespace
