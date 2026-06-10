' WordAi/Services/ContentLocator.vb
' 内容定位器 - 根据自然语言描述找到文档中的对象

Imports Word = Microsoft.Office.Interop.Word
Imports System.Text.RegularExpressions
Imports System.Diagnostics

Namespace Services

    ''' <summary>
    ''' 内容定位器 - 自然语言定位文档对象
    ''' </summary>
    Public Class ContentLocator

        Private ReadOnly _app As Word.Application

        Public Sub New(app As Word.Application)
            _app = app
        End Sub

        ''' <summary>
        ''' 根据描述查找并选中内容
        ''' </summary>
        ''' <param name="query">查询描述，如"标题"、"第3段"、"包含产品的段落"</param>
        Public Function FindAndSelect(query As String) As Boolean
            Try
                Dim doc As Word.Document = _app.ActiveDocument
                If doc Is Nothing Then Return False

                Dim queryLower As String = query.ToLower().Trim()

                ' 1. 匹配"标题" - 找到第一个标题样式
                If queryLower.Contains("标题") OrElse queryLower.Contains("heading") Then
                    Return FindHeading(doc, queryLower)
                End If

                ' 2. 匹配"第N段" - 找到指定段落
                Dim paraMatch As Match = Regex.Match(queryLower, "第\s*(\d+)\s*段")
                If paraMatch.Success Then
                    Dim paraNum As Integer = Integer.Parse(paraMatch.Groups(1).Value)
                    Return SelectParagraph(doc, paraNum)
                End If

                ' 3. 匹配"下一段"、"前一段"
                If queryLower.Contains("下一段") OrElse queryLower.Contains("下一个段落") Then
                    Return SelectNextParagraph(doc)
                End If

                If queryLower.Contains("前一段") OrElse queryLower.Contains("上一段") OrElse queryLower.Contains("前一个段落") Then
                    Return SelectPreviousParagraph(doc)
                End If

                ' 4. 匹配"包含XXX的段落" - 文本搜索
                Dim containsMatch As Match = Regex.Match(queryLower, "包含[""''](.+?)[""'']")
                If containsMatch.Success Then
                    Dim searchText As String = containsMatch.Groups(1).Value
                    Return FindParagraphContaining(doc, searchText)
                End If

                ' 5. 通用文本搜索
                If queryLower.Contains("找") OrElse queryLower.Contains("查找") Then
                    ' 提取引号中的内容
                    Dim textMatch As Match = Regex.Match(query, "[""''](.+?)[""'']")
                    If textMatch.Success Then
                        Return FindText(doc, textMatch.Groups(1).Value)
                    End If
                End If

                Debug.WriteLine($"[ContentLocator] 无法解析查询: {query}")
                Return False

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] FindAndSelect 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 查找并选中标题
        ''' </summary>
        Private Function FindHeading(doc As Word.Document, query As String) As Boolean
            Try
                ' 提取标题级别（如"标题1"、"标题 2"）
                Dim levelMatch As Match = Regex.Match(query, "标题\s*(\d)")
                Dim targetLevel As Integer = If(levelMatch.Success, Integer.Parse(levelMatch.Groups(1).Value), 1)

                ' 查找标题样式
                Dim targetStyle As String = $"标题 {targetLevel}"

                For Each para As Word.Paragraph In doc.Paragraphs
                    Try
                        If para.Style.NameLocal.Contains(targetStyle) OrElse
                           para.Style.NameLocal.Contains($"Heading {targetLevel}") Then
                            para.Range.Select()
                            Debug.WriteLine($"[ContentLocator] 找到并选中: {targetStyle}")
                            Return True
                        End If
                    Catch
                        ' 跳过无样式的段落
                    End Try
                Next

                ' 如果没找到精确匹配，找任意标题
                If Not levelMatch.Success Then
                    For Each para As Word.Paragraph In doc.Paragraphs
                        Try
                            Dim styleName As String = para.Style.NameLocal
                            If styleName.Contains("标题") OrElse styleName.Contains("Heading") Then
                                para.Range.Select()
                                Debug.WriteLine($"[ContentLocator] 找到并选中标题: {styleName}")
                                Return True
                            End If
                        Catch
                        End Try
                    Next
                End If

                Debug.WriteLine($"[ContentLocator] 未找到标题")
                Return False

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] FindHeading 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 选中指定段落
        ''' </summary>
        Private Function SelectParagraph(doc As Word.Document, paraNumber As Integer) As Boolean
            Try
                If paraNumber < 1 OrElse paraNumber > doc.Paragraphs.Count Then
                    Debug.WriteLine($"[ContentLocator] 段落索引超出范围: {paraNumber}")
                    Return False
                End If

                doc.Paragraphs(paraNumber).Range.Select()
                Debug.WriteLine($"[ContentLocator] 选中第 {paraNumber} 段")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] SelectParagraph 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 选中下一段
        ''' </summary>
        Private Function SelectNextParagraph(doc As Word.Document) As Boolean
            Try
                Dim sel As Word.Selection = _app.Selection
                If sel Is Nothing OrElse sel.Paragraphs.Count = 0 Then
                    Return False
                End If

                Dim currentPara As Word.Paragraph = sel.Paragraphs(1)

                ' 查找当前段落索引
                Dim currentIndex As Integer = 1
                For Each para As Word.Paragraph In doc.Paragraphs
                    If para.Range.Start = currentPara.Range.Start Then
                        Exit For
                    End If
                    currentIndex += 1
                Next

                ' 选中下一段
                If currentIndex < doc.Paragraphs.Count Then
                    doc.Paragraphs(currentIndex + 1).Range.Select()
                    Debug.WriteLine($"[ContentLocator] 选中下一段: {currentIndex + 1}")
                    Return True
                End If

                Return False

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] SelectNextParagraph 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 选中上一段
        ''' </summary>
        Private Function SelectPreviousParagraph(doc As Word.Document) As Boolean
            Try
                Dim sel As Word.Selection = _app.Selection
                If sel Is Nothing OrElse sel.Paragraphs.Count = 0 Then
                    Return False
                End If

                Dim currentPara As Word.Paragraph = sel.Paragraphs(1)

                ' 查找当前段落索引
                Dim currentIndex As Integer = 1
                For Each para As Word.Paragraph In doc.Paragraphs
                    If para.Range.Start = currentPara.Range.Start Then
                        Exit For
                    End If
                    currentIndex += 1
                Next

                ' 选中上一段
                If currentIndex > 1 Then
                    doc.Paragraphs(currentIndex - 1).Range.Select()
                    Debug.WriteLine($"[ContentLocator] 选中上一段: {currentIndex - 1}")
                    Return True
                End If

                Return False

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] SelectPreviousParagraph 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 查找包含指定文本的段落
        ''' </summary>
        Private Function FindParagraphContaining(doc As Word.Document, searchText As String) As Boolean
            Try
                For Each para As Word.Paragraph In doc.Paragraphs
                    If para.Range.Text.Contains(searchText) Then
                        para.Range.Select()
                        Debug.WriteLine($"[ContentLocator] 找到包含'{searchText}'的段落")
                        Return True
                    End If
                Next

                Debug.WriteLine($"[ContentLocator] 未找到包含'{searchText}'的段落")
                Return False

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] FindParagraphContaining 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 查找文本
        ''' </summary>
        Private Function FindText(doc As Word.Document, searchText As String) As Boolean
            Try
                Dim findRange As Word.Range = doc.Content
                findRange.Find.ClearFormatting()
                findRange.Find.Text = searchText
                findRange.Find.Forward = True
                findRange.Find.Wrap = Word.WdFindWrap.wdFindStop

                If findRange.Find.Execute() Then
                    findRange.Select()
                    Debug.WriteLine($"[ContentLocator] 找到文本: {searchText}")
                    Return True
                End If

                Debug.WriteLine($"[ContentLocator] 未找到文本: {searchText}")
                Return False

            Catch ex As Exception
                Debug.WriteLine($"[ContentLocator] FindText 失败: {ex.Message}")
                Return False
            End Try
        End Function

    End Class

End Namespace
