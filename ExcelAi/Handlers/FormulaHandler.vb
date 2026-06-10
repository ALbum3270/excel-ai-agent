Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions

Namespace Handlers

    ''' <summary>
    ''' Excel 公式处理器 - 将 AI 指令转换为 Excel 公式
    ''' </summary>
    Public Class FormulaHandler

        Private ReadOnly _app As Excel.Application

        Public Sub New(app As Excel.Application)
            _app = app
        End Sub

        ''' <summary>
        ''' 根据 AI 响应生成并应用公式
        ''' </summary>
        ''' <param name="aiResponse">AI 返回的公式描述</param>
        ''' <param name="targetRange">目标单元格区域</param>
        Public Function ApplyFormulaFromAI(aiResponse As String, targetRange As Excel.Range) As Boolean
            Try
                ' 从 AI 响应中提取公式
                Dim formula As String = ExtractFormula(aiResponse)
                If String.IsNullOrEmpty(formula) Then
                    Return False
                End If

                ' 验证公式语法
                If Not ValidateFormula(formula) Then
                    Return False
                End If

                ' 应用公式到目标区域
                targetRange.Formula = formula
                Return True

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[FormulaHandler] 应用公式失败: {0}", ex.Message))
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 从 AI 响应中提取 Excel 公式
        ''' </summary>
        Private Function ExtractFormula(aiResponse As String) As String
            Try
                ' 匹配 Excel 公式模式（以 = 开头）
                Dim formulaPattern As String = "=[\w\s\(\)\+\-\*/,.:$]+(?=\r|\n|$|\s)"
                Dim match As Match = Regex.Match(aiResponse, formulaPattern)

                If match.Success Then
                    Return match.Value.Trim()
                End If

                ' 如果没有找到公式，尝试匹配代码块中的公式
                Dim codeBlockPattern As String = "```(?:excel)?\s*(=.+?)```"
                match = Regex.Match(aiResponse, codeBlockPattern, RegexOptions.Singleline)

                If match.Success AndAlso match.Groups.Count > 1 Then
                    Return match.Groups(1).Value.Trim()
                End If

                Return Nothing

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[FormulaHandler] 提取公式失败: {0}", ex.Message))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 验证 Excel 公式语法
        ''' </summary>
        Private Function ValidateFormula(formula As String) As Boolean
            Try
                ' 基本验证：必须以 = 开头
                If Not formula.StartsWith("=") Then
                    Return False
                End If

                ' 验证括号匹配
                Dim openCount As Integer = 0
                For Each c As Char In formula
                    If c = "("c Then
                        openCount = openCount + 1
                    ElseIf c = ")"c Then
                        openCount = openCount - 1
                        If openCount < 0 Then
                            Return False
                        End If
                    End If
                Next

                If openCount <> 0 Then
                    Return False
                End If

                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 生成常用公式的辅助方法
        ''' </summary>
        Public Class FormulaBuilder

            ''' <summary>
            ''' 生成 SUM 公式
            ''' </summary>
            Public Shared Function Sum(range As String) As String
                Return String.Format("=SUM({0})", range)
            End Function

            ''' <summary>
            ''' 生成 AVERAGE 公式
            ''' </summary>
            Public Shared Function Average(range As String) As String
                Return String.Format("=AVERAGE({0})", range)
            End Function

            ''' <summary>
            ''' 生成 COUNT 公式
            ''' </summary>
            Public Shared Function Count(range As String) As String
                Return String.Format("=COUNT({0})", range)
            End Function

            ''' <summary>
            ''' 生成 IF 公式
            ''' </summary>
            Public Shared Function IfFormula(condition As String, valueIfTrue As String, valueIfFalse As String) As String
                Return String.Format("=IF({0},{1},{2})", condition, valueIfTrue, valueIfFalse)
            End Function

            ''' <summary>
            ''' 生成 VLOOKUP 公式
            ''' </summary>
            Public Shared Function VLookup(lookupValue As String, tableArray As String, columnIndex As Integer, Optional rangeLookup As Boolean = False) As String
                Return String.Format("=VLOOKUP({0},{1},{2},{3})", lookupValue, tableArray, columnIndex, If(rangeLookup, "TRUE", "FALSE"))
            End Function

            ''' <summary>
            ''' 生成 CONCATENATE 公式
            ''' </summary>
            Public Shared Function Concatenate(ParamArray values As String()) As String
                Return String.Format("=CONCATENATE({0})", String.Join(",", values))
            End Function

            ''' <summary>
            ''' 生成 SUMIF 公式
            ''' </summary>
            Public Shared Function SumIf(range As String, criteria As String, Optional sumRange As String = Nothing) As String
                If String.IsNullOrEmpty(sumRange) Then
                    Return String.Format("=SUMIF({0},{1})", range, criteria)
                Else
                    Return String.Format("=SUMIF({0},{1},{2})", range, criteria, sumRange)
                End If
            End Function

            ''' <summary>
            ''' 生成 COUNTIF 公式
            ''' </summary>
            Public Shared Function CountIf(range As String, criteria As String) As String
                Return String.Format("=COUNTIF({0},{1})", range, criteria)
            End Function

        End Class

        ''' <summary>
        ''' 智能单元格引用 - 将相对描述转换为单元格引用
        ''' </summary>
        Public Function ResolveReference(description As String, currentCell As Excel.Range) As String
            Try
                ' 处理相对引用
                Dim lowerDesc As String = description.ToLower().Trim()

                ' "上方" -> 上一行
                If lowerDesc.Contains("上方") OrElse lowerDesc.Contains("上面") Then
                    Return currentCell.Offset(-1, 0).Address(False, False)
                End If

                ' "下方" -> 下一行
                If lowerDesc.Contains("下方") OrElse lowerDesc.Contains("下面") Then
                    Return currentCell.Offset(1, 0).Address(False, False)
                End If

                ' "左边" -> 左一列
                If lowerDesc.Contains("左边") OrElse lowerDesc.Contains("左侧") Then
                    Return currentCell.Offset(0, -1).Address(False, False)
                End If

                ' "右边" -> 右一列
                If lowerDesc.Contains("右边") OrElse lowerDesc.Contains("右侧") Then
                    Return currentCell.Offset(0, 1).Address(False, False)
                End If

                ' "本行" -> 当前行
                If lowerDesc.Contains("本行") OrElse lowerDesc.Contains("当前行") Then
                    Dim startCell As Excel.Range = _app.Cells(currentCell.Row, 1)
                    Dim endCell As Excel.Range = _app.Cells(currentCell.Row, currentCell.Column - 1)
                    Return String.Format("{0}:{1}", startCell.Address(False, False), endCell.Address(False, False))
                End If

                ' "本列" -> 当前列
                If lowerDesc.Contains("本列") OrElse lowerDesc.Contains("当前列") Then
                    Dim startCell As Excel.Range = _app.Cells(1, currentCell.Column)
                    Dim endCell As Excel.Range = _app.Cells(currentCell.Row - 1, currentCell.Column)
                    Return String.Format("{0}:{1}", startCell.Address(False, False), endCell.Address(False, False))
                End If

                ' 默认返回原始描述（可能已经是单元格引用）
                Return description

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[FormulaHandler] 解析引用失败: {0}", ex.Message))
                Return description
            End Try
        End Function

    End Class

End Namespace
