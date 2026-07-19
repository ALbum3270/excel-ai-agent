Imports System.Collections.Generic

''' <summary>
''' 原始交互事件。事件不可丢，后续记忆和画像都应能从事件重建。
''' </summary>
Public Class ConversationEventRecord
    Public Property EventId As String
    Public Property SessionId As String
    Public Property AppType As String
    Public Property DocumentId As String
    Public Property EventType As String
    Public Property Role As String
    Public Property Content As String
    Public Property MetadataJson As String
    Public Property CreatedAt As String
    Public Property ProcessedAt As String
End Class

''' <summary>
''' 结构化记忆，不等同于聊天原文。
''' </summary>
Public Class MemoryItemRecord
    Public Property MemoryId As String
    Public Property SourceEventId As String
    Public Property Scope As String
    Public Property AppType As String
    Public Property DocumentId As String
    Public Property ProjectId As String
    Public Property MemoryType As String
    Public Property Content As String
    Public Property Summary As String
    Public Property Confidence As Double
    Public Property Importance As Double
    Public Property Status As String
    Public Property ExpiresAt As String
    Public Property LastVerifiedAt As String
    Public Property CreatedAt As String
    Public Property UpdatedAt As String
    Public Property Score As Double
End Class

''' <summary>
''' 记忆向量，与 memory_item 分表，便于切换 embedding 模型和重建索引。
''' </summary>
Public Class MemoryEmbeddingRecord
    Public Property EmbeddingId As String
    Public Property MemoryId As String
    Public Property EmbeddingModel As String
    Public Property EmbeddingDim As Integer
    Public Property EmbeddingJson As String
    Public Property VectorStatus As String
    Public Property LastError As String
    Public Property UpdatedAt As String
End Class

''' <summary>
''' 增量记忆任务。
''' </summary>
Public Class MemoryJobRecord
    Public Property JobId As String
    Public Property JobType As String
    Public Property TargetId As String
    Public Property PayloadJson As String
    Public Property Status As String
    Public Property AttemptCount As Integer
    Public Property LastError As String
    Public Property CreatedAt As String
    Public Property UpdatedAt As String
    Public Property NextRunAt As String
End Class

''' <summary>
''' Skill 索引项。请求时先召回索引，命中后再加载完整 Skill 内容。
''' </summary>
Public Class SkillRegistryRecord
    Public Property SkillName As String
    Public Property FilePath As String
    Public Property AppScope As String
    Public Property IntentTypes As String
    Public Property TriggerKeywords As String
    Public Property Description As String
    Public Property EmbeddingJson As String
    Public Property UsageCount As Integer
    Public Property SuccessCount As Integer
    Public Property LastIndexedAt As String
    Public Property Enabled As Boolean = True
    Public Property Score As Double
End Class

Public Interface IMemoryEventStore
    Function InsertConversationEvent(record As ConversationEventRecord) As String
    Function EnqueueMemoryJob(record As MemoryJobRecord) As String
End Interface

Public Interface IMemoryExtractor
    Function ExtractMemories(events As List(Of ConversationEventRecord), Optional contextJson As String = Nothing) As List(Of MemoryItemRecord)
End Interface

Public Interface ISkillSelector
    Function SelectSkills(query As String, intentType As String, appType As String, topN As Integer) As List(Of SkillRegistryRecord)
End Interface
