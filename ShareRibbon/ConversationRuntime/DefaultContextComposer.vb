Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

''' <summary>
''' Default message composer backed by the existing ChatContextBuilder.
''' </summary>
Public Class DefaultContextComposer
    Implements IContextComposer

    Public Function Compose(context As ChatRequestContext) As ChatContextCompositionResult Implements IContextComposer.Compose
        If context Is Nothing Then Throw New ArgumentNullException(NameOf(context))

        Dim ragCount As Integer = 0
        Dim usedContextBuilder As Boolean = False
        Dim messages As JArray = Nothing
        Dim trace As ChatContextTrace = Nothing

        If context.AddHistory AndAlso context.UseContextBuilder Then
            Try
                messages = BuildLayeredMessages(context, ragCount)
                trace = ChatContextBuilder.LastTrace
                usedContextBuilder = True
            Catch ex As Exception
                Debug.WriteLine("DefaultContextComposer fallback: " & ex.Message)
                Debug.WriteLine("DefaultContextComposer stack: " & ex.StackTrace)
                usedContextBuilder = False
            End Try
        End If

        If Not usedContextBuilder Then
            messages = BuildFallbackMessages(context)
        End If

        Return New ChatContextCompositionResult With {
            .Messages = messages,
            .RagCount = ragCount,
            .UsedContextBuilder = usedContextBuilder,
            .Trace = trace
        }
    End Function

    Private Function BuildLayeredMessages(context As ChatRequestContext, ByRef ragCount As Integer) As JArray
        Dim appType = GetAppType(context)
        Dim scenario = appType.ToLowerInvariant()
        Dim sessionMessages As New List(Of HistoryMessage)()

        If context.HistoryMessages IsNot Nothing Then
            For Each message In context.HistoryMessages
                If message.role <> "system" AndAlso Not String.IsNullOrEmpty(message.content) Then
                    sessionMessages.Add(message)
                End If
            Next
        End If

        Dim vars As New Dictionary(Of String, String)()
        If context.SelectionPendingMap IsNot Nothing AndAlso Not String.IsNullOrEmpty(context.RequestUuid) AndAlso context.SelectionPendingMap.ContainsKey(context.RequestUuid) Then
            Dim sel = context.SelectionPendingMap(context.RequestUuid)
            If sel IsNot Nothing AndAlso Not String.IsNullOrEmpty(sel.SelectedText) Then
                vars("选中内容") = sel.SelectedText
            End If
        End If

        Debug.WriteLine($"[DefaultContextComposer] UseContextBuilder={context.UseContextBuilder}, EnableMemory={context.EnableMemory}")
        Debug.WriteLine($"[DefaultContextComposer] Session message count: {sessionMessages.Count}")

        Dim built = ChatContextBuilder.BuildMessages(
            scenario,
            appType,
            context.Question,
            sessionMessages,
            context.Question,
            context.SystemPrompt,
            vars,
            context.EnableMemory,
            ragCount)

        Return ToJArray(built)
    End Function

    Private Function BuildFallbackMessages(context As ChatRequestContext) As JArray
        Dim messages As New List(Of HistoryMessage)()

        If context.AddHistory Then
            If context.HistoryMessages IsNot Nothing Then
                For Each message In context.HistoryMessages
                    If message.role <> "system" Then
                        messages.Add(message)
                    End If
                Next
            End If

            messages.Insert(0, New HistoryMessage With {
                .role = "system",
                .content = context.SystemPrompt
            })
            messages.Add(New HistoryMessage With {
                .role = "user",
                .content = context.Question
            })
        Else
            messages.Add(New HistoryMessage With {
                .role = "system",
                .content = If(context.SystemPrompt, String.Empty)
            })
            messages.Add(New HistoryMessage With {
                .role = "user",
                .content = If(context.Question, String.Empty)
            })
        End If

        Return ToJArray(messages)
    End Function

    Private Function ToJArray(messages As List(Of HistoryMessage)) As JArray
        Dim result As New JArray()
        If messages Is Nothing Then Return result

        For Each message In messages
            Dim item As New JObject()
            item("role") = message.role
            item("content") = If(message.content, String.Empty)
            result.Add(item)
        Next

        Return result
    End Function

    Private Function GetAppType(context As ChatRequestContext) As String
        If context IsNot Nothing AndAlso context.AppInfo IsNot Nothing Then
            Return context.AppInfo.Type.ToString()
        End If

        Return "Excel"
    End Function
End Class
