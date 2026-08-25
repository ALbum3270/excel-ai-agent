Imports Microsoft.Office.Tools.Ribbon

Partial Public Class Ribbon1
    Private Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ChatButton.Click
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub

    Private Async Sub AnalyzeButton_Click(sender As Object, e As RibbonControlEventArgs) Handles AnalyzeButton.Click
        Await Globals.ThisAddIn.AnalyzeSelectionAsync()
    End Sub

    Private Sub SettingsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingsButton.Click
        Globals.ThisAddIn.ShowSettings()
    End Sub
End Class
