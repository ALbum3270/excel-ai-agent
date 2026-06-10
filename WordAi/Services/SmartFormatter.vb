' WordAi/Services/SmartFormatter.vb
' 智能格式化器 - 根据查询直接设置格式，无需先选中

Imports Word = Microsoft.Office.Interop.Word
Imports System.Diagnostics

Namespace Services

    ''' <summary>
    ''' 智能格式化器 - 根据自然语言描述直接格式化内容
    ''' </summary>
    Public Class SmartFormatter

        Private ReadOnly _app As Word.Application
        Private ReadOnly _locator As ContentLocator

        Public Sub New(app As Word.Application)
            _app = app
            _locator = New ContentLocator(app)
        End Sub

        ''' <summary>
        ''' 根据查询设置格式
        ''' </summary>
        ''' <param name="query">目标查询，如"标题"、"第3段"</param>
        ''' <param name="action">操作描述，如"调大"、"加粗"、"改为红色"</param>
        Public Function FormatByQuery(query As String, action As String) As Boolean
            Try
                Debug.WriteLine($"[SmartFormatter] 查询: {query}, 操作: {action}")

                ' 1. 先定位到目标
                If Not _locator.FindAndSelect(query) Then
                    Debug.WriteLine($"[SmartFormatter] 无法定位到: {query}")
                    Return False
                End If

                ' 2. 获取当前选区
                Dim sel As Word.Selection = _app.Selection
                If sel Is Nothing Then
                    Return False
                End If

                ' 3. 解析并执行操作
                Return ExecuteFormatAction(sel, action)

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] FormatByQuery 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 执行格式化操作
        ''' </summary>
        Private Function ExecuteFormatAction(sel As Word.Selection, action As String) As Boolean
            Try
                Dim actionLower As String = action.ToLower().Trim()

                ' 字号调整
                If actionLower.Contains("调大") OrElse actionLower.Contains("增大") OrElse actionLower.Contains("变大") Then
                    Return IncreaseFontSize(sel)
                End If

                If actionLower.Contains("调小") OrElse actionLower.Contains("缩小") OrElse actionLower.Contains("变小") Then
                    Return DecreaseFontSize(sel)
                End If

                ' 加粗/斜体/下划线
                If actionLower.Contains("加粗") OrElse actionLower.Contains("粗体") Then
                    sel.Font.Bold = -1
                    Debug.WriteLine("[SmartFormatter] 设置加粗")
                    Return True
                End If

                If actionLower.Contains("取消加粗") OrElse actionLower.Contains("不加粗") Then
                    sel.Font.Bold = 0
                    Debug.WriteLine("[SmartFormatter] 取消加粗")
                    Return True
                End If

                If actionLower.Contains("斜体") Then
                    sel.Font.Italic = -1
                    Debug.WriteLine("[SmartFormatter] 设置斜体")
                    Return True
                End If

                If actionLower.Contains("下划线") Then
                    sel.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                    Debug.WriteLine("[SmartFormatter] 设置下划线")
                    Return True
                End If

                ' 颜色
                If actionLower.Contains("红色") Then
                    sel.Font.Color = Word.WdColor.wdColorRed
                    Debug.WriteLine("[SmartFormatter] 设置红色")
                    Return True
                End If

                If actionLower.Contains("蓝色") Then
                    sel.Font.Color = Word.WdColor.wdColorBlue
                    Debug.WriteLine("[SmartFormatter] 设置蓝色")
                    Return True
                End If

                If actionLower.Contains("黑色") Then
                    sel.Font.Color = Word.WdColor.wdColorBlack
                    Debug.WriteLine("[SmartFormatter] 设置黑色")
                    Return True
                End If

                ' 对齐
                If actionLower.Contains("居中") Then
                    sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Debug.WriteLine("[SmartFormatter] 设置居中")
                    Return True
                End If

                If actionLower.Contains("左对齐") Then
                    sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    Debug.WriteLine("[SmartFormatter] 设置左对齐")
                    Return True
                End If

                If actionLower.Contains("右对齐") Then
                    sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    Debug.WriteLine("[SmartFormatter] 设置右对齐")
                    Return True
                End If

                Debug.WriteLine($"[SmartFormatter] 无法识别操作: {action}")
                Return False

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] ExecuteFormatAction 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 增大字号（智能递增）
        ''' </summary>
        Private Function IncreaseFontSize(sel As Word.Selection) As Boolean
            Try
                Dim currentSize As Single = sel.Font.Size
                Dim newSize As Single

                ' 智能递增：小字号+2，大字号+4
                If currentSize < 12 Then
                    newSize = currentSize + 2
                ElseIf currentSize < 18 Then
                    newSize = currentSize + 2
                ElseIf currentSize < 24 Then
                    newSize = currentSize + 4
                Else
                    newSize = currentSize + 4
                End If

                sel.Font.Size = newSize
                Debug.WriteLine($"[SmartFormatter] 字号 {currentSize}pt → {newSize}pt")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] IncreaseFontSize 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 减小字号（智能递减）
        ''' </summary>
        Private Function DecreaseFontSize(sel As Word.Selection) As Boolean
            Try
                Dim currentSize As Single = sel.Font.Size
                Dim newSize As Single

                ' 智能递减：大字号-4，小字号-2
                If currentSize <= 12 Then
                    newSize = Math.Max(8, currentSize - 2)
                ElseIf currentSize <= 18 Then
                    newSize = currentSize - 2
                ElseIf currentSize <= 24 Then
                    newSize = currentSize - 4
                Else
                    newSize = currentSize - 4
                End If

                sel.Font.Size = newSize
                Debug.WriteLine($"[SmartFormatter] 字号 {currentSize}pt → {newSize}pt")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] DecreaseFontSize 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 直接设置格式（不需要查询定位）
        ''' </summary>
        Public Function ApplyFormatToSelection(action As String) As Boolean
            Try
                Dim sel As Word.Selection = _app.Selection
                If sel Is Nothing Then
                    Return False
                End If

                Return ExecuteFormatAction(sel, action)

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] ApplyFormatToSelection 失败: {ex.Message}")
                Return False
            End Try
        End Function

    End Class

End Namespace
