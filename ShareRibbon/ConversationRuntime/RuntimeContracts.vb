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

''' <summary>
''' Conversation persistence boundary. Existing ChatStateService/ConversationRepository can implement this later.
''' </summary>
Public Interface IConversationStore
    Sub AddMessage(role As String, content As String)
    Function GetRecentMessages(limit As Integer) As List(Of HistoryMessage)
End Interface

''' <summary>
''' Retrieval boundary for RAG providers.
''' </summary>
Public Interface IRetrievalService
    Function Retrieve(query As String, appType As String, topN As Integer) As List(Of RetrievalHit)
End Interface

''' <summary>
''' Skills boundary for local skills or open-source skill registries.
''' </summary>
Public Interface ISkillCatalog
    Function GetSkillPrompt(query As String, appType As String) As String
End Interface

Public Class RetrievalHit
    Public Property Id As String
    Public Property Content As String
    Public Property Source As String
    Public Property Score As Double
End Class

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
End Class

Public Class ChatRequestBuildResult
    Public Property RequestBody As String
    Public Property RagCount As Integer
    Public Property UsedContextBuilder As Boolean
End Class
