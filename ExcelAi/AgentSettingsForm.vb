Imports System.Drawing
Imports System.Windows.Forms
Imports ExcelAgent.Core

Public Class AgentSettingsForm
    Inherits Form

    Private ReadOnly _platform As New TextBox()
    Private ReadOnly _apiUrl As New TextBox()
    Private ReadOnly _apiKey As New TextBox()
    Private ReadOnly _model As New TextBox()
    Private ReadOnly _reasoning As New ComboBox()
    Private ReadOnly _prompt As New TextBox()

    Public Sub New()
        Text = "Excel Agent 设置"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(560, 470)
        MinimumSize = New Size(520, 430)
        FormBorderStyle = FormBorderStyle.Sizable

        _apiKey.UseSystemPasswordChar = True
        _reasoning.DropDownStyle = ComboBoxStyle.DropDownList
        _reasoning.Items.AddRange(New Object() {"default", "enabled", "disabled"})
        _prompt.Multiline = True
        _prompt.ScrollBars = ScrollBars.Vertical

        Dim layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(16),
            .ColumnCount = 2,
            .RowCount = 7
        }
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 105))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        For index = 0 To 4
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 38))
        Next
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 46))

        AddField(layout, 0, "平台", _platform)
        AddField(layout, 1, "API URL", _apiUrl)
        AddField(layout, 2, "API Key", _apiKey)
        AddField(layout, 3, "模型", _model)
        AddField(layout, 4, "推理模式", _reasoning)
        AddField(layout, 5, "自定义提示词", _prompt)

        Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft}
        Dim saveButton As New Button With {.Text = "保存", .AutoSize = True}
        Dim cancelButton As New Button With {.Text = "取消", .AutoSize = True, .DialogResult = DialogResult.Cancel}
        AddHandler saveButton.Click, AddressOf SaveSettings
        buttons.Controls.Add(saveButton)
        buttons.Controls.Add(cancelButton)
        layout.Controls.Add(buttons, 0, 6)
        layout.SetColumnSpan(buttons, 2)
        Controls.Add(layout)
        AcceptButton = saveButton
        CancelButton = cancelButton

        Dim settings = AgentSettingsStore.Load()
        _platform.Text = settings.Platform
        _apiUrl.Text = settings.ApiUrl
        _apiKey.Text = AgentSettingsStore.GetApiKey(settings)
        _model.Text = settings.ModelName
        _reasoning.SelectedItem = ReasoningRequestHelper.NormalizeReasoningMode(settings.ReasoningMode)
        If _reasoning.SelectedIndex < 0 Then _reasoning.SelectedIndex = 0
        _prompt.Text = settings.PromptContent
    End Sub

    Private Shared Sub AddField(layout As TableLayoutPanel, row As Integer, labelText As String, control As Control)
        Dim label As New Label With {.Text = labelText, .TextAlign = ContentAlignment.MiddleLeft, .Dock = DockStyle.Fill}
        control.Dock = DockStyle.Fill
        layout.Controls.Add(label, 0, row)
        layout.Controls.Add(control, 1, row)
    End Sub

    Private Sub SaveSettings(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(_apiUrl.Text) OrElse String.IsNullOrWhiteSpace(_apiKey.Text) OrElse String.IsNullOrWhiteSpace(_model.Text) Then
            MessageBox.Show(Me, "API URL、API Key 和模型不能为空。", "Excel Agent", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim settings As New ExcelAgentSettings With {
            .Platform = _platform.Text.Trim(),
            .ApiUrl = _apiUrl.Text.Trim(),
            .ModelName = _model.Text.Trim(),
            .ReasoningMode = If(TryCast(_reasoning.SelectedItem, String), "default"),
            .PromptName = "custom",
            .PromptContent = _prompt.Text
        }
        AgentSettingsStore.Save(settings, _apiKey.Text)
        DialogResult = DialogResult.OK
        Close()
    End Sub
End Class
