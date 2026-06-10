Imports ShareRibbon.Common
Imports System.Windows.Forms

Namespace Extensions

    ''' <summary>
    ''' ErrorHandler 集成扩展
    ''' 为 AI 操作提供统一的错误处理
    ''' </summary>
    Public Module ErrorHandlerExtension

        ''' <summary>
        ''' 安全执行操作，自动捕获并处理错误
        ''' </summary>
        Public Function SafeExecute(Of T)(
            action As Func(Of T),
            moduleName As String,
            operationName As String,
            Optional defaultValue As T = Nothing,
            Optional userErrorMessage As String = Nothing) As T

            Try
                Return action()

            Catch ex As Exception
                ' 使用 ErrorHandler 处理错误
                Dim friendlyMessage As String = If(userErrorMessage, GetFriendlyErrorMessage(operationName))
                ErrorHandler.Handle(ex, moduleName, operationName, friendlyMessage)

                ' 返回默认值
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 安全执行操作（无返回值）
        ''' </summary>
        Public Sub SafeExecute(
            action As Action,
            moduleName As String,
            operationName As String,
            Optional userErrorMessage As String = Nothing)

            Try
                action()

            Catch ex As Exception
                ' 使用 ErrorHandler 处理错误
                Dim friendlyMessage As String = If(userErrorMessage, GetFriendlyErrorMessage(operationName))
                ErrorHandler.Handle(ex, moduleName, operationName, friendlyMessage)
            End Try
        End Sub

        ''' <summary>
        ''' 获取友好的错误消息
        ''' </summary>
        Private Function GetFriendlyErrorMessage(operationName As String) As String
            Select Case operationName.ToLower()
                Case "公式应用", "formula"
                    Return "应用公式时出错，请检查公式格式是否正确。"

                Case "幻灯片生成", "slide generation"
                    Return "生成幻灯片时出错，请检查大纲格式是否正确。"

                Case "撤销", "undo"
                    Return "撤销操作失败，请尝试手动撤销（Ctrl+Z）。"

                Case "内容生成", "content generation"
                    Return "生成内容时出错，请重试。"

                Case Else
                    Return String.Format("{0}操作失败，请重试。", operationName)
            End Select
        End Function

        ''' <summary>
        ''' 显示友好的错误消息
        ''' </summary>
        Public Sub ShowError(message As String, Optional title As String = "错误")
            Try
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Catch
                System.Diagnostics.Debug.WriteLine("[ErrorHandlerExtension] 错误: " & message)
            End Try
        End Sub

        ''' <summary>
        ''' 显示友好的警告消息
        ''' </summary>
        Public Sub ShowWarning(message As String, Optional title As String = "警告")
            Try
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Catch
                System.Diagnostics.Debug.WriteLine("[ErrorHandlerExtension] 警告: " & message)
            End Try
        End Sub

        ''' <summary>
        ''' 记录警告（不显示给用户）
        ''' </summary>
        Public Sub LogWarning(message As String, moduleName As String)
            ErrorHandler.HandleWarning(message, moduleName)
        End Sub

        ''' <summary>
        ''' 尝试执行操作并返回是否成功
        ''' </summary>
        Public Function TryExecute(
            action As Action,
            moduleName As String,
            operationName As String,
            ByRef errorMessage As String) As Boolean

            Try
                action()
                errorMessage = String.Empty
                Return True

            Catch ex As Exception
                errorMessage = ex.Message
                ErrorHandler.Handle(ex, moduleName, operationName)
                Return False
            End Try
        End Function

    End Module

End Namespace
