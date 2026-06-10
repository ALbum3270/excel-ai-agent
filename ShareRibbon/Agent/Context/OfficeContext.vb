' ShareRibbon/Agent/Context/OfficeContext.vb
' Office 上下文基础类 - 统一封装当前 Office 应用的状态

Imports System.Text

Namespace Agent.Context

    ''' <summary>
    ''' Office 上下文 - 封装当前应用的状态（选区、文档结构等）
    ''' </summary>
    Public Class OfficeContext
        ''' <summary>应用类型</summary>
        Public Property AppType As String

        ''' <summary>当前选区信息</summary>
        Public Property Selection As SelectionInfo

        ''' <summary>文档结构信息（可选）</summary>
        Public Property DocStructure As DocumentStructure

        ''' <summary>
        ''' 转换为 Prompt 文本，自动注入到 AI 系统提示词
        ''' </summary>
        Public Function ToPromptText() As String
            Dim sb As New StringBuilder()

            sb.AppendLine("## 当前 Office 环境")
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

        ''' <summary>项目数量（单元格数、段落数等）</summary>
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
        ''' <summary>结构摘要（如标题层级、章节数量）</summary>
        Public Property Summary As String

        ''' <summary>是否有标题</summary>
        Public Property HasHeadings As Boolean

        ''' <summary>标题数量</summary>
        Public Property HeadingCount As Integer
    End Class

End Namespace
