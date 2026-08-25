Namespace Core

    ''' <summary>
    ''' Records the logical Excel action shown in execution explanations. Excel itself owns
    ''' the actual Ctrl+Z stack; the Agent never claims a programmable rollback succeeded.
    ''' </summary>
    Public Class UndoManager
        Public Class UndoPoint
            Public Property Name As String
            Public Property CreatedAt As DateTime
            Public Property AppType As String
            Public Property Description As String
            Public Property CanUndo As Boolean
        End Class

        Private ReadOnly _undoStack As New Stack(Of UndoPoint)()
        Private ReadOnly _maxUndoLevels As Integer

        Public Sub New(Optional maxUndoLevels As Integer = 10)
            _maxUndoLevels = Math.Max(1, maxUndoLevels)
        End Sub

        Public Function CreateUndoPoint(appType As String, operationName As String, description As String) As UndoPoint
            If _undoStack.Count >= _maxUndoLevels Then
                Dim retained = _undoStack.Reverse().Take(_maxUndoLevels - 1).Reverse().ToList()
                _undoStack.Clear()
                For Each item In retained
                    _undoStack.Push(item)
                Next
            End If
            Dim point As New UndoPoint With {
                .Name = operationName,
                .AppType = "Excel",
                .Description = description,
                .CreatedAt = DateTime.Now,
                .CanUndo = False
            }
            _undoStack.Push(point)
            Return point
        End Function

        Public Function GetUndoHistory() As List(Of UndoPoint)
            Return _undoStack.ToList()
        End Function

        Public Sub ClearHistory()
            _undoStack.Clear()
        End Sub

        Public Function GetUndoHint(appType As String) As String
            If _undoStack.Count = 0 Then Return ""
            Return "可尝试按 Ctrl+Z 撤销: " & _undoStack.Peek().Description
        End Function
    End Class
End Namespace
