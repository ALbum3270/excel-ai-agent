''' <summary>
''' Excel Agent 当前进程的模型配置。持久化由 Excel 宿主负责，这里只保留运行时状态。
''' </summary>
Public Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String
    Public Shared Property ReasoningMode As String = "default"
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String
    Public Const OfficeAiAppDataFolder As String = "ExcelAiAgent"
End Class
