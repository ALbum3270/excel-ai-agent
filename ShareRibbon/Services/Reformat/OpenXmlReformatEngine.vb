' ShareRibbon\Services\Reformat\OpenXmlReformatEngine.vb
' OpenXML排版引擎 - 使用OpenXML SDK直接操作.docx文件，无需Word COM
' 在临时文档上执行DSL排版指令，安全隔离原文档

Imports System.Diagnostics
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Newtonsoft.Json.Linq

''' <summary>
''' OpenXML排版引擎 - 直接操作docx文件执行排版指令
''' 优势：不依赖Word COM、不污染原文档Undo栈、可预览后确认
''' </summary>
Public Class OpenXmlReformatEngine

    ''' <summary>
    ''' 在临时文档上执行完整的排版指令集
    ''' </summary>
    ''' <param name="tempDocPath">临时文档路径</param>
    ''' <param name="instructions">DSL指令列表</param>
    ''' <returns>执行结果统计</returns>
    Public Shared Function ExecuteInstructions(
        tempDocPath As String,
        instructions As List(Of Instruction)) As ExecutionResult

        Dim result As New ExecutionResult()

        Try
            Using wordDoc = WordprocessingDocument.Open(tempDocPath, True)
                For Each instruction In instructions
                    Try
                        Dim success = ExecuteSingleInstruction(wordDoc, instruction)
                        If success Then
                            result.SuccessCount += 1
                        Else
                            result.FailureCount += 1
                            result.Errors.Add(New InstructionError(ErrorLevel.Error, $"指令 {instruction.Operation} 执行返回False", instruction.Operation))
                        End If
                    Catch ex As Exception
                        result.FailureCount += 1
                        result.Errors.Add(New InstructionError(ErrorLevel.Error, $"指令 {instruction.Operation} 异常: {ex.Message}", instruction.Operation))
                        Debug.WriteLine($"[OpenXmlReformatEngine] 指令执行异常: {ex}")
                    End Try
                Next

                ' 保存修改（Using结束会自动Dispose释放文件句柄）
                wordDoc.MainDocumentPart.Document.Save()
            End Using
        Catch ex As Exception
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, $"OpenXML文档操作失败: {ex.Message}", ""))
            Debug.WriteLine($"[OpenXmlReformatEngine] 文档操作失败: {ex}")
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 执行单条指令
    ''' </summary>
    Private Shared Function ExecuteSingleInstruction(
        wordDoc As WordprocessingDocument,
        instruction As Instruction) As Boolean

        Select Case instruction.Operation.ToLower()
            Case "setparagraphstyle"
                Return ExecuteSetParagraphStyle(wordDoc, instruction)
            Case "setcharacterformat"
                Return ExecuteSetCharacterFormat(wordDoc, instruction)
            Case "setpagesetup"
                Return ExecuteSetPageSetup(wordDoc, instruction)
            Case Else
                Debug.WriteLine($"[OpenXmlReformatEngine] 未支持的指令: {instruction.Operation}")
                Return False
        End Select
    End Function

#Region "段落样式指令"

    Private Shared Function ExecuteSetParagraphStyle(
        wordDoc As WordprocessingDocument,
        instruction As Instruction) As Boolean

        Dim body = wordDoc.MainDocumentPart?.Document?.Body
        If body Is Nothing Then Return False

        ' 获取目标段落索引
        Dim paraIndex = GetParaIndex(instruction)
        If paraIndex < 0 Then Return False

        Dim paragraphs = body.Elements(Of Paragraph)().ToList()
        If paraIndex >= paragraphs.Count Then Return False

        Dim para = paragraphs(paraIndex)
        Dim pPr = para.GetFirstChild(Of ParagraphProperties)()
        If pPr Is Nothing Then
            pPr = New ParagraphProperties()
            para.PrependChild(pPr)
        End If

        Dim params = instruction.Params
        If params Is Nothing Then Return False

        ' 对齐方式
        Dim alignment = GetStringParam(params, "alignment")
        If Not String.IsNullOrEmpty(alignment) Then
            Dim jc = pPr.GetFirstChild(Of Justification)()
            If jc IsNot Nothing Then jc.Remove()
            pPr.Append(New Justification() With {.Val = GetJustificationValue(alignment)})
        End If

        ' 首行缩进（字符数）
        Dim firstLineIndent = GetNumberParam(params, "firstLineIndent")
        If firstLineIndent > 0 Then
            Dim indent = pPr.GetFirstChild(Of Indentation)()
            If indent Is Nothing Then
                indent = New Indentation()
                pPr.Append(indent)
            End If
            indent.FirstLineChars = CInt(firstLineIndent * 100).ToString()
        End If

        ' 左缩进
        Dim leftIndent = GetNumberParam(params, "leftIndent")
        If leftIndent > 0 Then
            Dim indent = pPr.GetFirstChild(Of Indentation)()
            If indent Is Nothing Then
                indent = New Indentation()
                pPr.Append(indent)
            End If
            indent.Left = CInt(leftIndent * 567).ToString() ' cm → twips
        End If

        ' 行距
        Dim lineSpacing = GetNumberParam(params, "lineSpacing")
        If lineSpacing > 0 Then
            Dim spacing = pPr.GetFirstChild(Of SpacingBetweenLines)()
            If spacing IsNot Nothing Then spacing.Remove()
            spacing = New SpacingBetweenLines()

            If Math.Abs(lineSpacing - 1.0) < 0.01 Then
                spacing.Line = "240"
                spacing.LineRule = LineSpacingRuleValues.Auto
            ElseIf Math.Abs(lineSpacing - 1.5) < 0.01 Then
                spacing.Line = "360"
                spacing.LineRule = LineSpacingRuleValues.Auto
            ElseIf Math.Abs(lineSpacing - 2.0) < 0.01 Then
                spacing.Line = "480"
                spacing.LineRule = LineSpacingRuleValues.Auto
            Else
                spacing.Line = CInt(lineSpacing * 240).ToString()
                spacing.LineRule = LineSpacingRuleValues.Auto
            End If
            pPr.Append(spacing)
        End If

        ' 段前间距
        Dim spaceBefore = GetNumberParam(params, "spaceBefore")
        If spaceBefore > 0 Then
            Dim spacing = pPr.GetFirstChild(Of SpacingBetweenLines)()
            If spacing Is Nothing Then
                spacing = New SpacingBetweenLines()
                pPr.Append(spacing)
            End If
            spacing.Before = CInt(spaceBefore * 240).ToString()
        End If

        ' 段后间距
        Dim spaceAfter = GetNumberParam(params, "spaceAfter")
        If spaceAfter > 0 Then
            Dim spacing = pPr.GetFirstChild(Of SpacingBetweenLines)()
            If spacing Is Nothing Then
                spacing = New SpacingBetweenLines()
                pPr.Append(spacing)
            End If
            spacing.After = CInt(spaceAfter * 240).ToString()
        End If

        ' 与下段同页（孤行控制）
        Dim keepWithNext = GetBoolParam(params, "keepWithNext")
        If keepWithNext Then
            Dim keep = pPr.GetFirstChild(Of KeepNext)()
            If keep IsNot Nothing Then keep.Remove()
            pPr.Append(New KeepNext())
        End If

        ' 段前分页
        Dim pageBreakBefore = GetBoolParam(params, "pageBreakBefore")
        If pageBreakBefore Then
            Dim pb = pPr.GetFirstChild(Of PageBreakBefore)()
            If pb IsNot Nothing Then pb.Remove()
            pPr.Append(New PageBreakBefore())
        End If

        ' 段中不分页
        Dim keepLinesTogether = GetBoolParam(params, "keepLinesTogether")
        If keepLinesTogether Then
            Dim kl = pPr.GetFirstChild(Of KeepLines)()
            If kl IsNot Nothing Then kl.Remove()
            pPr.Append(New KeepLines())
        End If

        Return True
    End Function

#End Region

#Region "字符格式指令"

    Private Shared Function ExecuteSetCharacterFormat(
        wordDoc As WordprocessingDocument,
        instruction As Instruction) As Boolean

        Dim body = wordDoc.MainDocumentPart?.Document?.Body
        If body Is Nothing Then Return False

        Dim paraIndex = GetParaIndex(instruction)
        If paraIndex < 0 Then Return False

        Dim paragraphs = body.Elements(Of Paragraph)().ToList()
        If paraIndex >= paragraphs.Count Then Return False

        Dim para = paragraphs(paraIndex)
        Dim params = instruction.Params
        If params Is Nothing Then Return False

        ' 获取或创建段落中所有Run的RunProperties
        For Each run In para.Elements(Of Run)()
            Dim rPr = run.GetFirstChild(Of RunProperties)()
            If rPr Is Nothing Then
                rPr = New RunProperties()
                run.PrependChild(rPr)
            End If

            ' 中文字体
            Dim fontNameCN = GetStringParam(params, "fontNameCN")
            If Not String.IsNullOrEmpty(fontNameCN) Then
                Dim rf = rPr.GetFirstChild(Of RunFonts)()
                If rf Is Nothing Then
                    rf = New RunFonts()
                    rPr.PrependChild(rf)
                End If
                rf.EastAsia = fontNameCN
            End If

            ' 英文字体
            Dim fontNameEN = GetStringParam(params, "fontNameEN")
            If Not String.IsNullOrEmpty(fontNameEN) Then
                Dim rf = rPr.GetFirstChild(Of RunFonts)()
                If rf Is Nothing Then
                    rf = New RunFonts()
                    rPr.PrependChild(rf)
                End If
                rf.Ascii = fontNameEN
                rf.HighAnsi = fontNameEN
            End If

            ' 字号（pt → half-point）
            Dim fontSize = GetNumberParam(params, "fontSize")
            If fontSize > 0 Then
                Dim fs = rPr.GetFirstChild(Of FontSize)()
                If fs IsNot Nothing Then fs.Remove()
                rPr.Append(New FontSize() With {.Val = CInt(fontSize * 2).ToString()})

                Dim fsCs = rPr.GetFirstChild(Of FontSizeComplexScript)()
                If fsCs IsNot Nothing Then fsCs.Remove()
                rPr.Append(New FontSizeComplexScript() With {.Val = CInt(fontSize * 2).ToString()})
            End If

            ' 加粗
            If params.ContainsKey("bold") Then
                Dim bold = GetBoolParam(params, "bold")
                Dim b = rPr.GetFirstChild(Of Bold)()
                If b IsNot Nothing Then b.Remove()
                If bold Then
                    rPr.Append(New Bold())
                End If
            End If

            ' 斜体
            If params.ContainsKey("italic") Then
                Dim italic = GetBoolParam(params, "italic")
                Dim i = rPr.GetFirstChild(Of Italic)()
                If i IsNot Nothing Then i.Remove()
                If italic Then
                    rPr.Append(New Italic())
                End If
            End If

            ' 下划线
            If params.ContainsKey("underline") Then
                Dim underline = GetBoolParam(params, "underline")
                Dim u = rPr.GetFirstChild(Of Underline)()
                If u IsNot Nothing Then u.Remove()
                If underline Then
                    rPr.Append(New Underline() With {.Val = UnderlineValues.Single})
                End If
            End If

            ' 字体颜色
            Dim fontColor = GetStringParam(params, "fontColor")
            If Not String.IsNullOrEmpty(fontColor) Then
                Dim color = rPr.GetFirstChild(Of DocumentFormat.OpenXml.Wordprocessing.Color)()
                If color IsNot Nothing Then color.Remove()
                rPr.Append(New DocumentFormat.OpenXml.Wordprocessing.Color() With {
                    .Val = fontColor.TrimStart("#"c)
                })
            End If
        Next

        Return True
    End Function

#End Region

#Region "页面设置指令"

    Private Shared Function ExecuteSetPageSetup(
        wordDoc As WordprocessingDocument,
        instruction As Instruction) As Boolean

        Dim body = wordDoc.MainDocumentPart?.Document?.Body
        If body Is Nothing Then Return False

        Dim params = instruction.Params
        If params Is Nothing Then Return False

        ' 查找或创建SectionProperties（在body末尾）
        Dim sectPr = body.Elements(Of SectionProperties)().FirstOrDefault()
        If sectPr Is Nothing Then
            sectPr = New SectionProperties()
            body.Append(sectPr)
        End If

        ' 页边距
        Dim margins = params("margins")
        If margins IsNot Nothing Then
            Dim pgMar = sectPr.GetFirstChild(Of PageMargin)()
            If pgMar Is Nothing Then
                pgMar = New PageMargin()
                sectPr.PrependChild(pgMar)
            End If

            Dim top = GetNumberParam(margins, "top")
            If top > 0 Then pgMar.Top = CUInt(top * 567) ' cm → twips

            Dim bottom = GetNumberParam(margins, "bottom")
            If bottom > 0 Then pgMar.Bottom = CUInt(bottom * 567)

            Dim left = GetNumberParam(margins, "left")
            If left > 0 Then pgMar.Left = CUInt(left * 567)

            Dim right = GetNumberParam(margins, "right")
            If right > 0 Then pgMar.Right = CUInt(right * 567)
        End If

        ' 纸张方向
        Dim orientation = GetStringParam(params, "orientation")
        If Not String.IsNullOrEmpty(orientation) Then
            Dim pgSz = sectPr.GetFirstChild(Of PageSize)()
            If pgSz Is Nothing Then
                pgSz = New PageSize()
                sectPr.PrependChild(pgSz)
            End If

            If orientation.ToLower() = "landscape" Then
                pgSz.Orient = PageOrientationValues.Landscape
                ' 交换宽高
                If pgSz.Code Is Nothing Then
                    Dim tmp = pgSz.Width
                    pgSz.Width = pgSz.Height
                    pgSz.Height = tmp
                End If
            Else
                pgSz.Orient = PageOrientationValues.Portrait
            End If
        End If

        Return True
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 从指令中获取段落索引
    ''' </summary>
    Private Shared Function GetParaIndex(instruction As Instruction) As Integer
        If instruction.Target IsNot Nothing Then
            Dim idx = instruction.Target("index")
            If idx IsNot Nothing Then
                Return CInt(idx)
            End If
        End If
        Return -1
    End Function

    Private Shared Function GetStringParam(token As JToken, key As String) As String
        If token Is Nothing Then Return ""
        Dim val = token(key)
        If val IsNot Nothing Then Return val.ToString()
        Return ""
    End Function

    Private Shared Function GetNumberParam(token As JToken, key As String) As Double
        If token Is Nothing Then Return 0
        Dim val = token(key)
        If val Is Nothing Then Return 0
        Try
            Return CDbl(val)
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function GetBoolParam(token As JToken, key As String) As Boolean
        If token Is Nothing Then Return False
        Dim val = token(key)
        If val Is Nothing Then Return False
        Try
            Return CBool(val)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 将对齐字符串转为OpenXML值
    ''' </summary>
    Private Shared Function GetJustificationValue(alignment As String) As JustificationValues
        Select Case alignment.ToLower()
            Case "center"
                Return JustificationValues.Center
            Case "right"
                Return JustificationValues.Right
            Case "justify"
                Return JustificationValues.Both
            Case Else
                Return JustificationValues.Left
        End Select
    End Function

#End Region

End Class
