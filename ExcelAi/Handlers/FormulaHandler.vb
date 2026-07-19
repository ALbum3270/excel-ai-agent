Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions

Namespace Handlers

    ''' <summary>
    ''' Excel 公式处理器 - 将 AI 指令转换为 Excel 公式
    ''' </summary>
    Public Class FormulaHandler

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

    End Class

End Namespace
