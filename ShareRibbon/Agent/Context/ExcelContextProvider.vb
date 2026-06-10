' ShareRibbon/Agent/Context/ExcelContextProvider.vb
' Excel 上下文提供者 - 复用现有 ExcelContextService

Imports ShareRibbon.Controls.Services

Namespace Agent.Context

    ''' <summary>
    ''' Excel 上下文提供者 - 获取当前 Excel 应用的上下文
    ''' </summary>
    Public Class ExcelContextProvider
        Private ReadOnly _contextService As ExcelContextService
        Private ReadOnly _app As Object ' Excel.Application

        Public Sub New(excelApp As Object)
            _app = excelApp
            _contextService = New ExcelContextService()
        End Sub

        ''' <summary>
        ''' 获取当前 Excel 上下文
        ''' </summary>
        Public Function GetContext() As OfficeContext
            Dim ctx As New OfficeContext With {.AppType = "Excel"}

            Try
                ' 获取当前选区
                Dim sel As Object = _app.Selection
                If sel IsNot Nothing Then
                    ctx.Selection = GetSelectionInfo(sel)
                End If
            Catch ex As Exception
                Debug.WriteLine("[ExcelContextProvider] 获取选区失败: " & ex.Message)
            End Try

            Return ctx
        End Function

        ''' <summary>
        ''' 获取选区信息（使用后期绑定避免 Excel Interop 依赖）
        ''' </summary>
        Private Function GetSelectionInfo(sel As Object) As SelectionInfo
            Dim info As New SelectionInfo()

            Try
                ' 基本信息（使用后期绑定）
                info.Address = sel.Address(False, False).ToString()
                info.ItemCount = CInt(sel.Count)

                ' 获取数据
                Dim data = sel.Value

                ' 复用现有的数据结构检测
                Dim dataStructure = _contextService.DetectDataStructure(data)

                ' 数据类型推断
                If dataStructure.HasHeader Then
                    info.DataType = "表格(含表头)"
                Else
                    info.DataType = "数据区域"
                End If

                ' 生成数据预览（前3行）
                Dim topRows = _contextService.GetTopRows(data, 3)
                info.Preview = _contextService.FormatAsMarkdownTable(topRows, dataStructure.HasHeader)

            Catch ex As Exception
                Debug.WriteLine("[ExcelContextProvider] 获取选区详情失败: " & ex.Message)
                info.DataType = "未知"
            End Try

            Return info
        End Function
    End Class

End Namespace
