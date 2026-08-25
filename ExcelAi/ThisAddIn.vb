Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Core

Public Class ThisAddIn
    Private _chatPane As Microsoft.Office.Tools.CustomTaskPane
    Private _chatControl As ChatControl
    Private _host As ExcelAgentHost
    Private _uiContext As SynchronizationContext

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        _uiContext = SynchronizationContext.Current
        If _uiContext Is Nothing Then
            _uiContext = New WindowsFormsSynchronizationContext()
            SynchronizationContext.SetSynchronizationContext(_uiContext)
        End If
        AgentSettingsStore.Load()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        If _chatPane IsNot Nothing Then CustomTaskPanes.Remove(_chatPane)
        _chatControl?.Dispose()
        _chatControl = Nothing
        _chatPane = Nothing
        _host = Nothing
    End Sub

    Public Sub ShowChatTaskPane()
        EnsureChatTaskPane()
        _chatPane.Visible = True
        _chatControl.FocusInput()
    End Sub

    Public Sub ShowSettings()
        Using form As New AgentSettingsForm()
            form.ShowDialog()
        End Using
    End Sub

    Public Async Function AnalyzeSelectionAsync() As Task
        ShowChatTaskPane()
        Await _chatControl.SubmitPromptAsync(
            "分析当前选区或当前数据表，识别数据结构和质量问题，并执行最合适的统计、汇总、公式、透视表或图表操作。完成后根据 Excel 的实际结果验证目标。")
    End Function

    Private Sub EnsureChatTaskPane()
        If _chatPane IsNot Nothing Then Return
        _host = New ExcelAgentHost(Application, _uiContext)
        _chatControl = New ChatControl(_host)
        _chatPane = CustomTaskPanes.Add(_chatControl, "Excel Agent")
        _chatPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
        _chatPane.Width = 420
    End Sub
End Class
