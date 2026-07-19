Imports System.Collections.Generic

Namespace Core

    ''' <summary>
    ''' 撤销点管理器 - 在 AI 操作前创建撤销点，支持回退
    ''' </summary>
    Public Class UndoManager

        ''' <summary>
        ''' 撤销点信息
        ''' </summary>
        Public Class UndoPoint
            ''' <summary>
            ''' 撤销点名称
            ''' </summary>
            Public Property Name As String

            ''' <summary>
            ''' 创建时间
            ''' </summary>
            Public Property CreatedAt As DateTime

            ''' <summary>
            ''' Office 应用类型
            ''' </summary>
            Public Property AppType As String

            ''' <summary>
            ''' 操作描述
            ''' </summary>
            Public Property Description As String

            ''' <summary>
            ''' 是否可以撤销
            ''' </summary>
            Public Property CanUndo As Boolean

            Public Sub New(name As String, appType As String, description As String)
                Me.Name = name
                Me.AppType = appType
                Me.Description = description
                Me.CreatedAt = DateTime.Now
                Me.CanUndo = True
            End Sub
        End Class

        Private _undoStack As Stack(Of UndoPoint)
        Private _maxUndoLevels As Integer

        Public Sub New(Optional maxUndoLevels As Integer = 10)
            _undoStack = New Stack(Of UndoPoint)()
            _maxUndoLevels = maxUndoLevels
        End Sub

        ''' <summary>
        ''' 创建撤销点（Word/PowerPoint）
        ''' </summary>
        Public Function CreateUndoPoint(appType As String, operationName As String, description As String) As UndoPoint
            Try
                Dim undoPoint As New UndoPoint(operationName, appType, description)

                ' 限制撤销栈大小
                If _undoStack.Count >= _maxUndoLevels Then
                    Dim tempList As New List(Of UndoPoint)(_undoStack)
                    tempList.RemoveAt(tempList.Count - 1) ' 移除最旧的
                    _undoStack.Clear()
                    For i As Integer = tempList.Count - 1 To 0 Step -1
                        _undoStack.Push(tempList(i))
                    Next
                End If

                _undoStack.Push(undoPoint)

                System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] 创建撤销点: {0} ({1})", operationName, appType))

                Return undoPoint

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] 创建撤销点失败: {0}", ex.Message))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Word - 创建撤销点
        ''' </summary>
        Public Function CreateWordUndoPoint(app As Object, operationName As String, description As String) As UndoPoint
            Try
                ' Word 支持 UndoRecord，但需要通过 COM 调用
                ' 这里我们记录撤销点信息，实际撤销由 Word 自己的 Undo 机制处理
                Return CreateUndoPoint("Word", operationName, description)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] Word 创建撤销点失败: {0}", ex.Message))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Excel - 创建撤销点（Excel 不直接支持编程创建撤销点）
        ''' </summary>
        Public Function CreateExcelUndoPoint(app As Object, operationName As String, description As String) As UndoPoint
            Try
                ' Excel 没有直接的 UndoRecord API
                ' 我们记录操作信息，建议用户手动撤销
                Dim undoPoint As UndoPoint = CreateUndoPoint("Excel", operationName, description)
                undoPoint.CanUndo = False ' Excel 不支持编程撤销
                Return undoPoint

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] Excel 创建撤销点失败: {0}", ex.Message))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' PowerPoint - 创建撤销点
        ''' </summary>
        Public Function CreatePowerPointUndoPoint(app As Object, operationName As String, description As String) As UndoPoint
            Try
                ' PowerPoint 的 Undo 机制自动记录操作
                Return CreateUndoPoint("PowerPoint", operationName, description)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] PowerPoint 创建撤销点失败: {0}", ex.Message))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 获取所有撤销点历史
        ''' </summary>
        Public Function GetUndoHistory() As List(Of UndoPoint)
            Return New List(Of UndoPoint)(_undoStack)
        End Function

        ''' <summary>
        ''' 清除撤销历史
        ''' </summary>
        Public Sub ClearHistory()
            _undoStack.Clear()
            System.Diagnostics.Debug.WriteLine("[UndoManager] 清除撤销历史")
        End Sub

        ''' <summary>
        ''' 执行撤销（调用 Office 的原生 Undo）
        ''' </summary>
        Public Function Undo(app As Object, appType As String) As Boolean
            Try
                If _undoStack.Count = 0 Then
                    Return False
                End If

                Dim lastPoint As UndoPoint = _undoStack.Pop()

                Select Case appType.ToLower()
                    Case "word"
                        ' Word.Application.Undo
                        Dim wordApp As Object = app
                        wordApp.Undo()
                        System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] 撤销 Word 操作: {0}", lastPoint.Name))
                        Return True

                    Case "powerpoint"
                        ' PowerPoint.Application.Undo
                        Dim pptApp As Object = app
                        pptApp.Undo()
                        System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] 撤销 PowerPoint 操作: {0}", lastPoint.Name))
                        Return True

                    Case "excel"
                        ' Excel 不支持编程 Undo，需要提示用户手动 Ctrl+Z
                        System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] Excel 不支持编程撤销，请手动按 Ctrl+Z 撤销: {0}", lastPoint.Name))
                        Return False

                    Case Else
                        Return False
                End Select

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[UndoManager] 撤销操作失败: {0}", ex.Message))
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取撤销提示文本
        ''' </summary>
        Public Function GetUndoHint(appType As String) As String
            If _undoStack.Count = 0 Then
                Return "没有可撤销的操作"
            End If

            Dim lastPoint As UndoPoint = _undoStack.Peek()

            If appType.ToLower() = "excel" Then
                Return String.Format("按 Ctrl+Z 撤销: {0}", lastPoint.Description)
            Else
                Return String.Format("可以撤销: {0}", lastPoint.Description)
            End If
        End Function

    End Class

End Namespace
