Imports ShareRibbon.Common

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

    End Module

End Namespace
