Imports ShareRibbon.Core

Namespace Extensions

    ''' <summary>
    ''' UndoManager 集成扩展
    ''' 为 AI 操作提供撤销支持
    ''' </summary>
    Public Module UndoManagerExtension

        ' 全局 UndoManager 实例（每个应用共享）
        Private _wordUndoManager As UndoManager
        Private _excelUndoManager As UndoManager
        Private _pptUndoManager As UndoManager

        ''' <summary>
        ''' 获取或创建 Word UndoManager
        ''' </summary>
        Private Function GetWordUndoManager() As UndoManager
            If _wordUndoManager Is Nothing Then
                _wordUndoManager = New UndoManager()
            End If
            Return _wordUndoManager
        End Function

        ''' <summary>
        ''' 获取或创建 Excel UndoManager
        ''' </summary>
        Private Function GetExcelUndoManager() As UndoManager
            If _excelUndoManager Is Nothing Then
                _excelUndoManager = New UndoManager()
            End If
            Return _excelUndoManager
        End Function

        ''' <summary>
        ''' 获取或创建 PowerPoint UndoManager
        ''' </summary>
        Private Function GetPowerPointUndoManager() As UndoManager
            If _pptUndoManager Is Nothing Then
                _pptUndoManager = New UndoManager()
            End If
            Return _pptUndoManager
        End Function

        ''' <summary>
        ''' 在 AI 操作前创建撤销点
        ''' </summary>
        Public Function CreateAIOperationUndoPoint(appType As String, app As Object, operationName As String, description As String) As UndoManager.UndoPoint
            Try
                Dim undoManager As UndoManager = Nothing
                Dim undoPoint As UndoManager.UndoPoint = Nothing

                Select Case appType.ToLower()
                    Case "word"
                        undoManager = GetWordUndoManager()
                        undoPoint = undoManager.CreateWordUndoPoint(app, operationName, description)

                    Case "excel"
                        undoManager = GetExcelUndoManager()
                        undoPoint = undoManager.CreateExcelUndoPoint(app, operationName, description)

                    Case "powerpoint"
                        undoManager = GetPowerPointUndoManager()
                        undoPoint = undoManager.CreatePowerPointUndoPoint(app, operationName, description)

                    Case Else
                        System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 不支持的应用类型: " & appType)
                        Return Nothing
                End Select

                If undoPoint IsNot Nothing Then
                    System.Diagnostics.Debug.WriteLine(String.Format("[UndoManagerExtension] 创建撤销点: {0} - {1}", operationName, description))
                End If

                Return undoPoint

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 创建撤销点失败: " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 获取撤销提示信息
        ''' </summary>
        Public Function GetUndoHint(appType As String) As String
            Try
                Dim undoManager As UndoManager = Nothing

                Select Case appType.ToLower()
                    Case "word"
                        undoManager = GetWordUndoManager()
                    Case "excel"
                        undoManager = GetExcelUndoManager()
                    Case "powerpoint"
                        undoManager = GetPowerPointUndoManager()
                    Case Else
                        Return "不支持的应用类型"
                End Select

                If undoManager IsNot Nothing Then
                    Return undoManager.GetUndoHint(appType)
                End If

                Return "没有可撤销的操作"

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 获取撤销提示失败: " & ex.Message)
                Return "撤销提示获取失败"
            End Try
        End Function

        ''' <summary>
        ''' 执行撤销操作
        ''' </summary>
        Public Function ExecuteUndo(appType As String, app As Object) As Boolean
            Try
                Dim undoManager As UndoManager = Nothing

                Select Case appType.ToLower()
                    Case "word"
                        undoManager = GetWordUndoManager()
                    Case "excel"
                        undoManager = GetExcelUndoManager()
                    Case "powerpoint"
                        undoManager = GetPowerPointUndoManager()
                    Case Else
                        Return False
                End Select

                If undoManager IsNot Nothing Then
                    Return undoManager.Undo(app, appType)
                End If

                Return False

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 执行撤销失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示撤销提示（状态栏或消息框）
        ''' </summary>
        Public Sub ShowUndoHint(appType As String, Optional showInMessageBox As Boolean = False)
            Try
                Dim hint As String = GetUndoHint(appType)

                If showInMessageBox Then
                    System.Windows.Forms.MessageBox.Show(
                        hint,
                        "撤销提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information)
                Else
                    ' 显示在状态栏
                    System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 撤销提示: " & hint)
                    ' TODO: 如果有状态栏，在这里显示
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 显示撤销提示失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 获取撤销历史
        ''' </summary>
        Public Function GetUndoHistory(appType As String) As List(Of UndoManager.UndoPoint)
            Try
                Dim undoManager As UndoManager = Nothing

                Select Case appType.ToLower()
                    Case "word"
                        undoManager = GetWordUndoManager()
                    Case "excel"
                        undoManager = GetExcelUndoManager()
                    Case "powerpoint"
                        undoManager = GetPowerPointUndoManager()
                    Case Else
                        Return New List(Of UndoManager.UndoPoint)()
                End Select

                If undoManager IsNot Nothing Then
                    Return undoManager.GetUndoHistory()
                End If

                Return New List(Of UndoManager.UndoPoint)()

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[UndoManagerExtension] 获取撤销历史失败: " & ex.Message)
                Return New List(Of UndoManager.UndoPoint)()
            End Try
        End Function

    End Module

End Namespace
