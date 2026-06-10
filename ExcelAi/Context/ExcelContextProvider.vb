Imports Excel = Microsoft.Office.Interop.Excel
Imports ShareRibbon.Agent.Context
Imports System.Diagnostics

Namespace Context
    Public Class ExcelContextProvider
        Implements IContextProvider

        Private ReadOnly _app As Excel.Application

        Public Sub New(app As Excel.Application)
            _app = app
        End Sub

        Public Function GetContext() As OfficeContext Implements IContextProvider.GetContext
            Dim ctx As New OfficeContext With {.AppType = "Excel"}

            Try
                ' 基础上下文信息
                If _app.Selection IsNot Nothing Then
                    ctx.Selection = New SelectionInfo With {
                        .DataType = "Excel数据",
                        .Address = "当前选区"
                    }
                End If

                ' 文档结构信息
                ctx.DocStructure = New DocumentStructure With {
                    .Summary = "Excel工作簿"
                }
            Catch ex As Exception
                Debug.WriteLine("获取Excel上下文失败: " & ex.Message)
            End Try

            Return ctx
        End Function
    End Class
End Namespace
