Imports Microsoft.Office.Tools.Ribbon

<System.ComponentModel.ToolboxItemAttribute(False)>
Partial Class Ribbon1
    Inherits RibbonBase

    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())
        InitializeComponent()
    End Sub

    Public Sub New(container As System.ComponentModel.IContainer)
        Me.New()
        If container IsNot Nothing Then container.Add(Me)
    End Sub

    Private components As System.ComponentModel.IContainer

    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then components.Dispose()
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private Sub InitializeComponent()
        Me.TabExcelAgent = Me.Factory.CreateRibbonTab()
        Me.AgentGroup = Me.Factory.CreateRibbonGroup()
        Me.ChatButton = Me.Factory.CreateRibbonButton()
        Me.AnalyzeButton = Me.Factory.CreateRibbonButton()
        Me.SettingsButton = Me.Factory.CreateRibbonButton()

        Me.TabExcelAgent.ControlId.ControlIdType = RibbonControlIdType.Office
        Me.TabExcelAgent.Label = "Excel Agent"
        Me.TabExcelAgent.Name = "TabExcelAgent"
        Me.AgentGroup.Label = "智能操作"
        Me.AgentGroup.Name = "AgentGroup"

        Me.ChatButton.Label = "打开助手"
        Me.ChatButton.Name = "ChatButton"
        Me.ChatButton.ScreenTip = "打开 Excel Agent"
        Me.AnalyzeButton.Label = "分析选区"
        Me.AnalyzeButton.Name = "AnalyzeButton"
        Me.AnalyzeButton.ScreenTip = "分析并处理当前选区"
        Me.SettingsButton.Label = "设置"
        Me.SettingsButton.Name = "SettingsButton"

        Me.AgentGroup.Items.Add(Me.ChatButton)
        Me.AgentGroup.Items.Add(Me.AnalyzeButton)
        Me.AgentGroup.Items.Add(Me.SettingsButton)
        Me.TabExcelAgent.Groups.Add(Me.AgentGroup)
        Me.Tabs.Add(Me.TabExcelAgent)
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
    End Sub

    Friend WithEvents TabExcelAgent As RibbonTab
    Friend WithEvents AgentGroup As RibbonGroup
    Friend WithEvents ChatButton As RibbonButton
    Friend WithEvents AnalyzeButton As RibbonButton
    Friend WithEvents SettingsButton As RibbonButton
End Class

Partial Class ThisRibbonCollection
    Friend ReadOnly Property Ribbon1 As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
