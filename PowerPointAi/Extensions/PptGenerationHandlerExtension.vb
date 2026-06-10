Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports System.Windows.Forms
Imports PowerPointAi.Handlers

Namespace Extensions

    ''' <summary>
    ''' PptGenerationHandler 集成扩展
    ''' 为 ChatControl 提供幻灯片大纲检测和生成功能
    ''' </summary>
    Public Module PptGenerationHandlerExtension

        ''' <summary>
        ''' 检测 AI 响应是否包含幻灯片大纲
        ''' </summary>
        Public Function DetectSlideOutline(aiResponse As String) As Boolean
            If String.IsNullOrWhiteSpace(aiResponse) Then
                Return False
            End If

            ' 检测 Markdown 标题（# 开头）
            Dim hasHeading As Boolean = aiResponse.Contains("#")

            ' 检测要点标记（- 或 * 或数字编号）
            Dim hasBullet As Boolean = False
            Dim lines As String() = aiResponse.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                Dim trimmed = line.Trim()
                If trimmed.StartsWith("-") OrElse
                   trimmed.StartsWith("*") OrElse
                   System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d+\.") Then
                    hasBullet = True
                    Exit For
                End If
            Next

            ' 检测"幻灯片"关键字
            Dim hasSlideKeyword As Boolean = aiResponse.Contains("幻灯片") OrElse
                                              aiResponse.ToLower().Contains("slide")

            ' 至少满足以下条件之一：
            ' 1. 有标题 + 有要点
            ' 2. 有"幻灯片"关键字 + 有标题
            Return (hasHeading AndAlso hasBullet) OrElse (hasSlideKeyword AndAlso hasHeading)
        End Function

        ''' <summary>
        ''' 尝试从 AI 响应中生成幻灯片
        ''' </summary>
        Public Function TryGenerateSlides(aiResponse As String, pptApp As PowerPoint.Application) As Boolean
            Try
                ' 检测是否包含幻灯片大纲
                If Not DetectSlideOutline(aiResponse) Then
                    Return False
                End If

                ' 询问用户确认
                Dim result As DialogResult = MessageBox.Show(
                    "检测到幻灯片大纲，是否自动生成幻灯片？" & vbCrLf & vbCrLf &
                    "将在当前幻灯片后插入新幻灯片。",
                    "生成幻灯片",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If result = DialogResult.Yes Then
                    ' 创建 PptGenerationHandler
                    Dim handler As New PptGenerationHandler(pptApp)

                    ' 生成幻灯片
                    Dim count As Integer = handler.GenerateSlidesFromAI(aiResponse, insertAfterCurrent:=True)

                    If count > 0 Then
                        MessageBox.Show(
                            String.Format("成功生成 {0} 张幻灯片！", count),
                            "生成完成",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                        System.Diagnostics.Debug.WriteLine("[PptGenerationHandlerExtension] 成功生成 " & count & " 张幻灯片")
                        Return True
                    Else
                        MessageBox.Show(
                            "未能生成幻灯片，请检查大纲格式。",
                            "生成失败",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
                        Return False
                    End If
                End If

                Return False

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[PptGenerationHandlerExtension] 错误: " & ex.Message)
                MessageBox.Show(
                    "生成幻灯片时出错: " & ex.Message,
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示幻灯片生成帮助
        ''' </summary>
        Public Sub ShowSlideGenerationHelp(pptApp As PowerPoint.Application)
            Try
                Dim helpText As String =
                    "PowerPoint 幻灯片自动生成助手" & vbCrLf & vbCrLf &
                    "您可以用自然语言描述演示文稿大纲，AI 会自动生成幻灯片。" & vbCrLf & vbCrLf &
                    "示例格式：" & vbCrLf &
                    "# 幻灯片 1: 项目介绍" & vbCrLf &
                    "- 项目背景" & vbCrLf &
                    "- 项目目标" & vbCrLf &
                    "- 预期成果" & vbCrLf & vbCrLf &
                    "# 幻灯片 2: 技术架构" & vbCrLf &
                    "- 前端技术" & vbCrLf &
                    "- 后端技术" & vbCrLf &
                    "- 数据库" & vbCrLf & vbCrLf &
                    "提示：" & vbCrLf &
                    "1. 使用 # 开头表示新幻灯片标题" & vbCrLf &
                    "2. 使用 - 或 * 或 1. 表示要点" & vbCrLf &
                    "3. AI 会自动生成并插入幻灯片"

                MessageBox.Show(helpText, "幻灯片生成帮助", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[PptGenerationHandlerExtension] 显示帮助失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 生成示例大纲文本
        ''' </summary>
        Public Function GetSampleOutline() As String
            Return "# 项目介绍" & vbCrLf &
                   "- 项目背景和意义" & vbCrLf &
                   "- 主要目标" & vbCrLf &
                   "- 预期成果" & vbCrLf & vbCrLf &
                   "# 技术方案" & vbCrLf &
                   "- 前端技术栈" & vbCrLf &
                   "- 后端架构" & vbCrLf &
                   "- 数据存储方案" & vbCrLf & vbCrLf &
                   "# 项目进度" & vbCrLf &
                   "- 第一阶段：需求分析" & vbCrLf &
                   "- 第二阶段：系统设计" & vbCrLf &
                   "- 第三阶段：开发实施" & vbCrLf &
                   "- 第四阶段：测试上线"
        End Function

        ''' <summary>
        ''' 验证大纲格式是否正确
        ''' </summary>
        Public Function ValidateOutlineFormat(outline As String, ByRef errorMessage As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(outline) Then
                    errorMessage = "大纲内容为空"
                    Return False
                End If

                ' 检查是否有标题
                If Not outline.Contains("#") AndAlso Not outline.Contains("幻灯片") Then
                    errorMessage = "未找到幻灯片标题（需要 # 开头或包含'幻灯片'关键字）"
                    Return False
                End If

                ' 检查是否有内容
                Dim lines = outline.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                Dim hasContent As Boolean = False
                For Each line In lines
                    Dim trimmed = line.Trim()
                    If trimmed.StartsWith("-") OrElse trimmed.StartsWith("*") OrElse
                       System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d+\.") Then
                        hasContent = True
                        Exit For
                    End If
                Next

                If Not hasContent Then
                    errorMessage = "未找到幻灯片内容（需要使用 - 或 * 或 1. 开头的要点）"
                    Return False
                End If

                errorMessage = ""
                Return True

            Catch ex As Exception
                errorMessage = "验证失败: " & ex.Message
                Return False
            End Try
        End Function

    End Module

End Namespace
