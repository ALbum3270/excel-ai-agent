Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class MemoryTurnRecorder
    Public Sub RecordConversationTurn(
        userQuestion As String,
        assistantAnswer As String,
        sessionId As String,
        responseMode As String,
        addHistory As Boolean,
        appType As String)

        If String.IsNullOrWhiteSpace(userQuestion) AndAlso String.IsNullOrWhiteSpace(assistantAnswer) Then Return

        Dim capturedUserQuestion = If(userQuestion, "")
        Dim capturedAssistantAnswer = If(assistantAnswer, "")
        Dim capturedSessionId = If(String.IsNullOrWhiteSpace(sessionId), Guid.NewGuid().ToString(), sessionId)
        Dim capturedResponseMode = If(responseMode, "")
        Dim capturedAppType = If(appType, "")

        Task.Run(Sub()
                     Try
                         Dim userMetadata As New JObject()
                         userMetadata("responseMode") = capturedResponseMode
                         userMetadata("addHistory") = addHistory

                         Dim userEventId = AgentMemoryRepository.AppendConversationEvent(New ConversationEventRecord With {
                             .SessionId = capturedSessionId,
                             .AppType = capturedAppType,
                             .EventType = "chat_message",
                             .Role = "user",
                             .Content = capturedUserQuestion,
                             .MetadataJson = userMetadata.ToString(Formatting.None)
                         })

                         Dim assistantMetadata As New JObject()
                         assistantMetadata("responseMode") = capturedResponseMode
                         assistantMetadata("sourceUserEventId") = userEventId

                         Dim assistantEventId = AgentMemoryRepository.AppendConversationEvent(New ConversationEventRecord With {
                             .SessionId = capturedSessionId,
                             .AppType = capturedAppType,
                             .EventType = "chat_message",
                             .Role = "assistant",
                             .Content = capturedAssistantAnswer,
                             .MetadataJson = assistantMetadata.ToString(Formatting.None)
                         })

                         Dim payload As New JObject()
                         payload("user_event_id") = userEventId
                         payload("assistant_event_id") = assistantEventId
                         payload("session_id") = capturedSessionId
                         payload("app_type") = capturedAppType
                         payload("response_mode") = capturedResponseMode

                         AgentMemoryRepository.EnqueueJob(New MemoryJobRecord With {
                             .JobType = "extract_memory",
                             .TargetId = assistantEventId,
                             .PayloadJson = payload.ToString(Formatting.None),
                             .Status = "pending"
                         })
                         AgentMemoryPipelineService.KickoffPendingJobs()
                     Catch ex As Exception
                         Debug.WriteLine($"MemoryTurnRecorder failed: {ex.Message}")
                     End Try
                 End Sub)
    End Sub
End Class
