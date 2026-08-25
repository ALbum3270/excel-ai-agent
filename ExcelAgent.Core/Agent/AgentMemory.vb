Imports System.Linq

Namespace Agent

    Public Class SessionMessage
        Public Property Role As String
        Public Property Content As String
        Public Property Timestamp As DateTime = DateTime.Now
    End Class

    ''' <summary>
    ''' Per-run working memory and bounded in-process conversation history. Nothing is written
    ''' to disk, so a new Excel session starts without legacy state.
    ''' </summary>
    Public Class AgentMemory
        Private ReadOnly _workingContext As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _sessionHistory As New List(Of SessionMessage)()
        Private ReadOnly _taskContextSnapshots As New List(Of String)()
        Private ReadOnly _lock As New Object()
        Private Const MaxHistoryMessages As Integer = 30
        Private Const MaxTaskContextSnapshots As Integer = 3
        Private Const MaxTaskContextChars As Integer = 12000

        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))

        Public Sub SetWorking(key As String, value As Object)
            SyncLock _lock
                _workingContext(key) = value
            End SyncLock
        End Sub

        Public Function GetWorking(key As String) As Object
            SyncLock _lock
                Dim value As Object = Nothing
                If _workingContext.TryGetValue(key, value) Then Return value
                Return Nothing
            End SyncLock
        End Function

        Public Function GetWorkingString(key As String) As String
            Dim value = GetWorking(key)
            Return If(value Is Nothing, "", value.ToString())
        End Function

        Public Sub ClearWorking()
            SyncLock _lock
                _workingContext.Clear()
                _taskContextSnapshots.Clear()
            End SyncLock
        End Sub

        Public Sub BeginTaskContext(initialContext As Context.ContextPack)
            SyncLock _lock
                _taskContextSnapshots.Clear()
                AddContextSnapshot(initialContext)
            End SyncLock
        End Sub

        Public Sub ObserveTaskContext(contextPack As Context.ContextPack)
            SyncLock _lock
                AddContextSnapshot(contextPack)
            End SyncLock
        End Sub

        Public Function GetPriorTaskContexts(currentContext As Context.ContextPack) As String
            Dim current = RenderContext(currentContext)
            SyncLock _lock
                Return String.Join(vbCrLf & vbCrLf,
                    _taskContextSnapshots.Where(Function(item) Not String.Equals(item, current, StringComparison.Ordinal)))
            End SyncLock
        End Function

        Private Sub AddContextSnapshot(contextPack As Context.ContextPack)
            Dim snapshot = RenderContext(contextPack)
            If String.IsNullOrWhiteSpace(snapshot) OrElse _taskContextSnapshots.Contains(snapshot) Then Return
            If _taskContextSnapshots.Count >= MaxTaskContextSnapshots Then _taskContextSnapshots.RemoveAt(0)
            _taskContextSnapshots.Add(snapshot)
        End Sub

        Private Shared Function RenderContext(contextPack As Context.ContextPack) As String
            If contextPack Is Nothing Then Return ""
            Dim value = If(contextPack.ToPromptText(), "").Trim()
            Return If(value.Length <= MaxTaskContextChars, value, value.Substring(0, MaxTaskContextChars))
        End Function

        Public Sub AddSessionMessage(role As String, content As String)
            SyncLock _lock
                _sessionHistory.Add(New SessionMessage With {.Role = role, .Content = content})
                While _sessionHistory.Count > MaxHistoryMessages
                    _sessionHistory.RemoveAt(0)
                End While
            End SyncLock
        End Sub

        Public Function GetRecentMessages(count As Integer) As List(Of HistoryMessage)
            SyncLock _lock
                Return _sessionHistory.Skip(Math.Max(0, _sessionHistory.Count - Math.Max(0, count))).
                    Select(Function(item) New HistoryMessage With {.role = item.Role, .content = item.Content}).ToList()
            End SyncLock
        End Function

        Public Sub AddTaskRecord(result As AgentResult)
            ' The result already lives on AgentSession; no second persistence channel is used.
        End Sub

        Public Function Search(query As String, topK As Integer) As List(Of String)
            Return New List(Of String)()
        End Function
    End Class
End Namespace
