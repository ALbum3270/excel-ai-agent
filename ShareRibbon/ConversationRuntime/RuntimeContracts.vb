Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' Builds the model message list for a chat request.
''' </summary>
Public Interface IContextComposer
    Function Compose(context As ChatRequestContext) As ChatContextCompositionResult
End Interface

''' <summary>
''' Exposes enabled tools, including MCP tools, to the chat request builder.
''' </summary>
Public Interface IToolBroker
    Function GetTools(context As ChatRequestContext) As JArray
End Interface

''' <summary>
''' Future adapter point for Semantic Kernel, Agent Framework sidecar, or another open-source runtime.
''' </summary>
Public Interface IConversationRuntime
    Function BuildRequest(context As ChatRequestContext) As ChatRequestBuildResult
End Interface

Public Class ChatRequestContext
    Public Property RequestUuid As String
    Public Property Question As String
    Public Property SystemPrompt As String
    Public Property AddHistory As Boolean
    Public Property ModelName As String
    Public Property Platform As String
    Public Property ApiUrl As String
    Public Property ReasoningMode As String
    Public Property Stream As Boolean = True
    Public Property AppInfo As ApplicationInfo
    Public Property HistoryMessages As List(Of HistoryMessage)
    Public Property SelectionPendingMap As Dictionary(Of String, SelectionInfo)
    Public Property UseContextBuilder As Boolean
    Public Property EnableMemory As Boolean
End Class

Public Class ChatContextCompositionResult
    Public Property Messages As JArray
    Public Property RagCount As Integer
    Public Property UsedContextBuilder As Boolean
    Public Property Trace As ChatContextTrace
End Class

Public Class ChatRequestBuildResult
    Public Property RequestBody As String
    Public Property RagCount As Integer
    Public Property UsedContextBuilder As Boolean
    Public Property Trace As ChatContextTrace
End Class

Public Class ChatContextTrace
    Public Property Query As String
    Public Property AppType As String
    Public Property IntentType As String
    Public Property IntentDescription As String
    Public Property OfficeContext As String
    Public Property UserProfileInjected As Boolean
    Public Property Memories As New List(Of ChatContextMemoryTrace)()
    Public Property RecentSessions As New List(Of ChatContextSessionTrace)()
    Public Property Skills As New List(Of ChatContextSkillTrace)()
    Public Property Tools As New List(Of ChatContextToolTrace)()
    Public Property ExecutionPlan As ChatContextPlanTrace
    Public Property TaskSpec As ChatContextTaskSpecTrace
End Class

Public Class ChatContextMemoryTrace
    Public Property Id As String
    Public Property Source As String
    Public Property MemoryType As String
    Public Property Content As String
    Public Property Score As Double
    Public Property Importance As Double
End Class

Public Class ChatContextSessionTrace
    Public Property SessionId As String
    Public Property Title As String
    Public Property Snippet As String
End Class

Public Class ChatContextSkillTrace
    Public Property Name As String
    Public Property Source As String
    Public Property Reason As String
End Class

Public Class ChatContextToolTrace
    Public Property Id As String
    Public Property Name As String
    Public Property Category As String
    Public Property RiskLevel As String
    Public Property AvailabilityStatus As String
    Public Property LastError As String
End Class

Public Class ChatContextPlanTrace
    Public Property Summary As String
    Public Property Understanding As String
    Public Property Steps As New List(Of ChatContextPlanStepTrace)()
End Class

Public Class ChatContextPlanStepTrace
    Public Property StepNumber As Integer
    Public Property Description As String
    Public Property ToolOrCode As String
    Public Property Language As String
End Class

Public Class ChatContextTaskSpecTrace
    Public Property Goal As String
    Public Property TargetObject As String
    Public Property Complexity As String
    Public Property RiskLevel As String
    Public Property Constraints As New List(Of String)()
    Public Property SuccessCriteria As New List(Of String)()
    Public Property RequiredTools As New List(Of String)()
End Class
