Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports Office = Microsoft.Office.Core
Imports System.Text.RegularExpressions

Namespace Handlers

    ''' <summary>
    ''' PowerPoint 生成处理器 - 根据文本大纲自动生成幻灯片
    ''' </summary>
    Public Class PptGenerationHandler

        Private ReadOnly _app As PowerPoint.Application

        Public Sub New(app As PowerPoint.Application)
            _app = app
        End Sub

        ''' <summary>
        ''' 从 AI 响应生成幻灯片
        ''' </summary>
        Public Function GenerateSlidesFromAI(aiResponse As String, Optional insertAfterCurrent As Boolean = True) As Integer
            Try
                ' 解析 AI 响应，提取幻灯片结构
                Dim slides As List(Of SlideContent) = ParseSlideContent(aiResponse)

                If slides.Count = 0 Then
                    Return 0
                End If

                ' 获取当前演示文稿
                Dim presentation As PowerPoint.Presentation = _app.ActivePresentation
                If presentation Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("[PptGenerationHandler] 没有活动的演示文稿")
                    Return 0
                End If

                ' 确定插入位置
                Dim insertIndex As Integer = 1
                If insertAfterCurrent AndAlso _app.ActiveWindow IsNot Nothing Then
                    Try
                        insertIndex = _app.ActiveWindow.Selection.SlideRange(1).SlideIndex + 1
                    Catch
                        insertIndex = presentation.Slides.Count + 1
                    End Try
                Else
                    insertIndex = presentation.Slides.Count + 1
                End If

                ' 生成幻灯片
                Dim createdCount As Integer = 0
                For Each slideContent As SlideContent In slides
                    If CreateSlide(presentation, slideContent, insertIndex + createdCount) Then
                        createdCount = createdCount + 1
                    End If
                Next

                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 成功生成 {0} 张幻灯片", createdCount))
                Return createdCount

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 生成幻灯片失败: {0}", ex.Message))
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' 幻灯片内容结构
        ''' </summary>
        Public Class SlideContent
            Public Property Title As String
            Public Property BulletPoints As List(Of String)
            Public Property Notes As String
            Public Property LayoutType As SlideLayoutType

            Public Sub New()
                BulletPoints = New List(Of String)()
                LayoutType = SlideLayoutType.TitleAndContent
            End Sub
        End Class

        ''' <summary>
        ''' 幻灯片布局类型
        ''' </summary>
        Public Enum SlideLayoutType
            TitleSlide = 1          ' 标题幻灯片
            TitleAndContent = 2     ' 标题和内容
            SectionHeader = 3       ' 章节标题
            TwoContent = 4          ' 两栏内容
            Blank = 12              ' 空白
        End Enum

        ''' <summary>
        ''' 解析 AI 响应，提取幻灯片内容
        ''' </summary>
        Private Function ParseSlideContent(aiResponse As String) As List(Of SlideContent)
            Dim slides As New List(Of SlideContent)()

            Try
                ' 按行分割
                Dim lines As String() = aiResponse.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)

                Dim currentSlide As SlideContent = Nothing

                For Each line As String In lines
                    Dim trimmedLine As String = line.Trim()

                    ' 跳过空行
                    If String.IsNullOrEmpty(trimmedLine) Then
                        Continue For
                    End If

                    ' 检测新幻灯片标题（# 开头或 "幻灯片" 关键字）
                    If trimmedLine.StartsWith("#") OrElse trimmedLine.Contains("幻灯片") Then
                        ' 保存上一张幻灯片
                        If currentSlide IsNot Nothing AndAlso Not String.IsNullOrEmpty(currentSlide.Title) Then
                            slides.Add(currentSlide)
                        End If

                        ' 创建新幻灯片
                        currentSlide = New SlideContent()
                        currentSlide.Title = trimmedLine.TrimStart("#"c).Trim()

                        ' 移除 "幻灯片 X:" 前缀
                        currentSlide.Title = Regex.Replace(currentSlide.Title, "^幻灯片\s*\d+[:：]?\s*", "")

                        ' 判断是否为章节标题（短标题且没有冒号）
                        If currentSlide.Title.Length <= 10 AndAlso Not currentSlide.Title.Contains(":") Then
                            currentSlide.LayoutType = SlideLayoutType.SectionHeader
                        End If

                    ' 检测要点（- 或 * 开头）
                    ElseIf currentSlide IsNot Nothing AndAlso (trimmedLine.StartsWith("-") OrElse trimmedLine.StartsWith("*")) Then
                        Dim bulletText As String = trimmedLine.Substring(1).Trim()
                        If Not String.IsNullOrEmpty(bulletText) Then
                            currentSlide.BulletPoints.Add(bulletText)
                        End If

                    ' 检测编号要点（1. 2. 等）
                    ElseIf currentSlide IsNot Nothing AndAlso Regex.IsMatch(trimmedLine, "^\d+\.") Then
                        Dim bulletText As String = Regex.Replace(trimmedLine, "^\d+\.\s*", "")
                        If Not String.IsNullOrEmpty(bulletText) Then
                            currentSlide.BulletPoints.Add(bulletText)
                        End If

                    ' 其他内容作为要点
                    ElseIf currentSlide IsNot Nothing AndAlso Not trimmedLine.StartsWith("```") Then
                        ' 跳过代码块标记
                        currentSlide.BulletPoints.Add(trimmedLine)
                    End If
                Next

                ' 添加最后一张幻灯片
                If currentSlide IsNot Nothing AndAlso Not String.IsNullOrEmpty(currentSlide.Title) Then
                    slides.Add(currentSlide)
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 解析幻灯片内容失败: {0}", ex.Message))
            End Try

            Return slides
        End Function

        ''' <summary>
        ''' 创建单张幻灯片
        ''' </summary>
        Private Function CreateSlide(presentation As PowerPoint.Presentation, content As SlideContent, index As Integer) As Boolean
            Try
                ' 选择布局
                Dim layout As PowerPoint.CustomLayout = presentation.SlideMaster.CustomLayouts(content.LayoutType)

                ' 创建幻灯片
                Dim slide As PowerPoint.Slide = presentation.Slides.AddSlide(index, layout)

                ' 设置标题
                If slide.Shapes.HasTitle Then
                    slide.Shapes.Title.TextFrame.TextRange.Text = content.Title
                End If

                ' 添加内容
                If content.BulletPoints.Count > 0 Then
                    ' 查找内容占位符
                    For Each shape As PowerPoint.Shape In slide.Shapes
                        If shape.Type = Office.MsoShapeType.msoPlaceholder Then
                            If shape.PlaceholderFormat.Type = PowerPoint.PpPlaceholderType.ppPlaceholderBody OrElse
                               shape.PlaceholderFormat.Type = PowerPoint.PpPlaceholderType.ppPlaceholderObject Then

                                ' 清空占位符
                                shape.TextFrame.TextRange.Text = ""

                                ' 添加要点
                                For i As Integer = 0 To content.BulletPoints.Count - 1
                                    If i = 0 Then
                                        shape.TextFrame.TextRange.Text = content.BulletPoints(i)
                                    Else
                                        shape.TextFrame.TextRange.InsertAfter(vbCrLf & content.BulletPoints(i))
                                    End If
                                Next

                                Exit For
                            End If
                        End If
                    Next
                End If

                ' 添加备注
                If Not String.IsNullOrEmpty(content.Notes) Then
                    slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = content.Notes
                End If

                Return True

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 创建幻灯片失败: {0}", ex.Message))
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 应用设计主题
        ''' </summary>
        Public Sub ApplyTheme(themeName As String)
            Try
                If _app.ActivePresentation Is Nothing Then
                    Return
                End If

                ' PowerPoint 主题路径通常在 Office 安装目录
                ' 这里提供基础实现，实际可以扩展支持更多主题
                Dim presentation As PowerPoint.Presentation = _app.ActivePresentation

                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 应用主题: {0}", themeName))

                ' 主题应用需要主题文件路径
                ' presentation.ApplyTheme("C:\Program Files\Microsoft Office\root\Document Themes 16\" & themeName & ".thmx")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 应用主题失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 批量设置幻灯片字体
        ''' </summary>
        Public Sub SetFontForAllSlides(fontName As String, Optional fontSize As Single = 0)
            Try
                If _app.ActivePresentation Is Nothing Then
                    Return
                End If

                Dim presentation As PowerPoint.Presentation = _app.ActivePresentation

                For Each slide As PowerPoint.Slide In presentation.Slides
                    For Each shape As PowerPoint.Shape In slide.Shapes
                        If shape.HasTextFrame Then
                            If fontName <> "" Then
                                shape.TextFrame.TextRange.Font.Name = fontName
                            End If
                            If fontSize > 0 Then
                                shape.TextFrame.TextRange.Font.Size = fontSize
                            End If
                        End If
                    Next
                Next

                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 设置字体: {0}", fontName))

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[PptGenerationHandler] 设置字体失败: {0}", ex.Message))
            End Try
        End Sub

    End Class

End Namespace
