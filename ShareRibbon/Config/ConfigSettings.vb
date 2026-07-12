' 存储配置的api大模型和api key
Public Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String
    Public Shared Property mcpable As Boolean
    Public Shared Property ReasoningMode As String = "default"

    ' Embedding 模型配置
    Public Shared Property EmbeddingModel As String = ""

    ' FIM (Fill-In-the-Middle) 补全能力
    Public Shared Property fimSupported As Boolean = False
    Public Shared Property fimUrl As String = ""

    ' 提示词相关配置
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String

    ' OBSOLETE for smart-mode routing (P0-2): product path is always AgentKernel via ExecutionPathPolicy.
    ' Kept for binary/config compatibility; ChatRoutingOrchestrator ignores false for smart-mode.
    ' Prefer reading Agent.ExecutionPathPolicy.PreferAgentKernelForSmartMode instead of this flag.
    Public Shared Property UseNewAgentKernel As Boolean = True

    ' Loop 框架配置
    Public Shared Property UseLoopFramework As Boolean = True
    Public Shared Property UseAdvancedPlanning As Boolean = True
    Public Shared Property MaxLoopIterations As Integer = 3

    Public Const OfficeAiAppDataFolder As String = "OfficeAiAppData"
End Class
