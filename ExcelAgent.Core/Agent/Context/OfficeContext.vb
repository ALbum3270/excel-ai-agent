' ExcelAgent.Core/Agent/Context/OfficeContext.vb
' Excel 上下文基础类

Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Agent.Context

    ''' <summary>
    ''' Excel 上下文 - 封装当前工作簿的选区和结构状态
    ''' </summary>
    Public Class OfficeContext
        ''' <summary>应用类型</summary>
        Public Property AppType As String

        ''' <summary>当前选区信息</summary>
        Public Property Selection As SelectionInfo

        ''' <summary>文档结构信息（可选）</summary>
        Public Property DocStructure As DocumentStructure

        ''' <summary>
        ''' Excel 宿主的结构化上下文。核心层只负责传递和预算控制。
        ''' </summary>
        Public Property HostData As New JObject()

        ''' <summary>
        ''' 转换为 Prompt 文本，自动注入到 AI 系统提示词
        ''' </summary>
        Public Function ToPromptText() As String
            Dim sb As New StringBuilder()

            sb.AppendLine("## 当前 Excel 环境")
            sb.AppendLine("应用: " & AppType)

            If Selection IsNot Nothing Then
                sb.AppendLine()
                sb.AppendLine("### 当前选区")
                sb.AppendLine("- 位置: " & Selection.Address)
                sb.AppendLine("- 数量: " & Selection.ItemCount.ToString() & " 项")
                sb.AppendLine("- 类型: " & Selection.DataType)

                If Not String.IsNullOrEmpty(Selection.Preview) Then
                    sb.AppendLine()
                    sb.AppendLine("### 数据预览")
                    sb.AppendLine(Selection.Preview)
                End If
            End If

            If DocStructure IsNot Nothing AndAlso Not String.IsNullOrEmpty(DocStructure.Summary) Then
                sb.AppendLine()
                sb.AppendLine("### 文档结构")
                sb.AppendLine(DocStructure.Summary)
            End If

            Return sb.ToString()
        End Function
    End Class

    ''' <summary>
    ''' 选区信息
    ''' </summary>
    Public Class SelectionInfo
        ''' <summary>选区地址（如 A1:F100）</summary>
        Public Property Address As String

        ''' <summary>选区中的单元格数</summary>
        Public Property ItemCount As Integer

        ''' <summary>数据预览（前几行数据的文本表示）</summary>
        Public Property Preview As String

        ''' <summary>数据类型（如"表格(含表头)"、"纯文本"）</summary>
        Public Property DataType As String
    End Class

    ''' <summary>
    ''' 文档结构信息
    ''' </summary>
    Public Class DocumentStructure
        ''' <summary>工作表结构摘要</summary>
        Public Property Summary As String
    End Class

End Namespace
