' ShareRibbon\Services\Reformat\LegacyTemplateConverter.vb
' 旧模板 → SemanticStyleMapping 转换器

''' <summary>
''' 旧模板转换器 - 将现有ReformatTemplate转换为SemanticStyleMapping
''' 提供向后兼容能力
''' </summary>
Public Class LegacyTemplateConverter

    ''' <summary>
    ''' 将ReformatTemplate转换为SemanticStyleMapping
    ''' </summary>
    Public Shared Function Convert(template As ReformatTemplate) As SemanticStyleMapping
        If template Is Nothing Then Return Nothing

        ' 先检查缓存
        Dim cached = SemanticMappingManager.Instance.GetMappingBySourceId(template.Id)
        If cached IsNot Nothing Then Return cached

        Dim mapping As New SemanticStyleMapping()
        mapping.Name = template.Name
        mapping.SourceType = SemanticMappingSourceType.FromLegacy
        mapping.SourceId = template.Id

        ' 转换正文样式规则
        ConvertBodyStyles(template.BodyStyles, mapping)

        ' 转换版式骨架
        ConvertLayout(template.Layout, mapping)

        ' 转换页面设置（直接复用）
        If template.PageSettings IsNot Nothing Then
            mapping.PageConfig = template.PageSettings
        End If

        ' 确保基础标签
        EnsureBasicTags(mapping)

        ' 公文场景：添加公文特有的语义标签
        If template.Category = "公文" OrElse template.Name.Contains("公文") Then
            AddOfficialDocumentTags(mapping)
        End If

        ' 缓存转换结果
        SemanticMappingManager.Instance.AddMapping(mapping)

        Return mapping
    End Function

    ''' <summary>转换正文样式规则为语义标签</summary>
    Private Shared Sub ConvertBodyStyles(styles As List(Of StyleRule), mapping As SemanticStyleMapping)
        If styles Is Nothing Then Return

        For Each rule In styles
            Dim tag = MapRuleToTag(rule)
            If tag IsNot Nothing Then
                mapping.SemanticTags.Add(tag)
            End If
        Next
    End Sub

    ''' <summary>将StyleRule映射到SemanticTag</summary>
    Private Shared Function MapRuleToTag(rule As StyleRule) As SemanticTag
        If rule Is Nothing Then Return Nothing

        Dim tagId As String = ""
        Dim displayName As String = rule.RuleName
        Dim parentId As String = ""
        Dim matchHint As String = rule.MatchCondition

        Dim name = If(rule.RuleName, "").ToLower()

        Select Case True
            Case name.Contains("一级") OrElse name.Contains("大标题") OrElse name.Contains("章标题")
                tagId = SemanticTagRegistry.TAG_TITLE_1
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("二级") OrElse name.Contains("节标题")
                tagId = SemanticTagRegistry.TAG_TITLE_2
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("三级") OrElse name.Contains("小标题")
                tagId = SemanticTagRegistry.TAG_TITLE_3
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("正文") OrElse name.Contains("body")
                tagId = SemanticTagRegistry.TAG_BODY_NORMAL
                parentId = SemanticTagRegistry.TAG_BODY

            Case name.Contains("强调") OrElse name.Contains("emphasis")
                tagId = SemanticTagRegistry.TAG_BODY_EMPHASIS
                parentId = SemanticTagRegistry.TAG_BODY

            Case name.Contains("列表") OrElse name.Contains("list")
                tagId = SemanticTagRegistry.TAG_LIST_ORDERED
                parentId = SemanticTagRegistry.TAG_LIST

            Case name.Contains("引用") OrElse name.Contains("quote")
                tagId = SemanticTagRegistry.TAG_QUOTE
                parentId = ""

            Case name.Contains("题注") OrElse name.Contains("caption")
                tagId = SemanticTagRegistry.TAG_CAPTION
                parentId = ""

            Case Else
                ' 无法识别的规则作为正文处理
                tagId = SemanticTagRegistry.TAG_BODY_NORMAL
                parentId = SemanticTagRegistry.TAG_BODY
                matchHint = $"来自旧规则: {rule.RuleName}"
        End Select

        ' 避免重复
        Dim tag As New SemanticTag(tagId, displayName, parentId, SemanticTagRegistry.GetTagLevel(tagId), matchHint)

        ' 复制格式配置
        If rule.Font IsNot Nothing Then tag.Font = rule.Font
        If rule.Paragraph IsNot Nothing Then tag.Paragraph = rule.Paragraph
        If rule.Color IsNot Nothing Then tag.Color = rule.Color

        Return tag
    End Function

    ''' <summary>转换版式骨架</summary>
    Private Shared Sub ConvertLayout(layout As LayoutConfig, mapping As SemanticStyleMapping)
        If layout Is Nothing OrElse layout.Elements Is Nothing Then Return
        mapping.LayoutSkeleton = layout
    End Sub

    ''' <summary>确保基础标签存在</summary>
    Private Shared Sub EnsureBasicTags(mapping As SemanticStyleMapping)
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_BODY_NORMAL) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_BODY_NORMAL, "正文",
                SemanticTagRegistry.TAG_BODY, 2, "普通正文段落"))
        End If

        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_TITLE_1) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_TITLE_1, "一级标题",
                SemanticTagRegistry.TAG_TITLE, 2, "主要章节标题"))
        End If
    End Sub

    ''' <summary>添加公文特有的语义标签</summary>
    Private Shared Sub AddOfficialDocumentTags(mapping As SemanticStyleMapping)
        ' 发文机关标志（红色大字）
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "header.org") Then
            Dim tag = New SemanticTag("header.org", "发文机关标志", "header", 2,
                "发文机关全称+文件字样，如'XX市人民政府文件'，通常居中红色大号字体")
            tag.Font = New FontConfig("方正小标宋简体", "Arial", 22, True)
            tag.Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceBefore = 2}
            tag.Color = New ColorConfig("#C00000")
            mapping.SemanticTags.Add(tag)
        End If

        ' 发文字号
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "header.refno") Then
            Dim tag = New SemanticTag("header.refno", "发文字号", "header", 2,
                "公文发文字号，格式如'×政发〔2024〕15号'，通常位于红色分隔线下方")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("center", 0, 1.0)
            mapping.SemanticTags.Add(tag)
        End If

        ' 签发人（上行文）
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "header.signer") Then
            Dim tag = New SemanticTag("header.signer", "签发人", "header", 2,
                "上行文才有的签发人信息，格式'签发人：×××'，通常右对齐")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("right", 0, 1.0) With {.SpaceBefore = 0.5}
            mapping.SemanticTags.Add(tag)
        End If

        ' 文件标题
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "title.main") Then
            Dim tag = New SemanticTag("title.main", "文件标题", "title", 2,
                "公文的主标题，如'关于加强安全生产工作的通知'，通常居中")
            tag.Font = New FontConfig("方正小标宋简体", "Arial", 22, True)
            tag.Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceBefore = 1.5, .SpaceAfter = 1.5}
            mapping.SemanticTags.Add(tag)
        End If

        ' 主送机关
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "title.recipient") Then
            Dim tag = New SemanticTag("title.recipient", "主送机关", "title", 2,
                "公文的主送机关，如'各区县人民政府：'，通常顶格左对齐")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("left", 0, 1.0)
            mapping.SemanticTags.Add(tag)
        End If

        ' 附件说明
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "body.attachment") Then
            Dim tag = New SemanticTag("body.attachment", "附件说明", "body", 2,
                "附件说明段落，格式'附件：1. ××××'，通常首行缩进")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("left", 2, 1.875) With {.SpaceBefore = 1}
            mapping.SemanticTags.Add(tag)
        End If

        ' 发文机关署名
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "footer.signature") Then
            Dim tag = New SemanticTag("footer.signature", "发文机关署名", "footer", 2,
                "公文末尾的发文机关署名，如'XX市人民政府'，通常右对齐")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("right", 2, 1.0) With {.SpaceBefore = 2}
            mapping.SemanticTags.Add(tag)
        End If

        ' 成文日期
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "footer.date") Then
            Dim tag = New SemanticTag("footer.date", "成文日期", "footer", 2,
                "公文成文日期，格式'2024年1月15日'，通常右对齐")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("right", 0, 1.0)
            mapping.SemanticTags.Add(tag)
        End If

        ' 抄送
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "footer.cc") Then
            Dim tag = New SemanticTag("footer.cc", "抄送机关", "footer", 2,
                "版记中的抄送信息，格式'抄送：×××，×××。'")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 14)
            tag.Paragraph = New ParagraphConfig("left", 2, 1.0)
            mapping.SemanticTags.Add(tag)
        End If
    End Sub
End Class
