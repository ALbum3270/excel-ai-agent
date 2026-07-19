Imports System.Globalization
Imports Newtonsoft.Json.Linq

Namespace Design

    Public Class DeckDesignSpec
        Public Property DesignSystem As String = "executive-light"
        Public Property DesignTokens As JObject
        Public Property Slides As New List(Of SlideDesignSpec)()

        Public Shared Function Parse(params As JObject) As DeckDesignSpec
            Dim result As New DeckDesignSpec()
            If params Is Nothing Then Return result
            Dim designToken = params("designSystem")
            If designToken IsNot Nothing Then
                If designToken.Type = JTokenType.String Then
                    If Not String.IsNullOrWhiteSpace(designToken.ToString()) Then result.DesignSystem = designToken.ToString()
                ElseIf designToken.Type = JTokenType.Object Then
                    Dim designName = If(designToken("name")?.ToString(), "").Trim()
                    If Not String.IsNullOrWhiteSpace(designName) Then result.DesignSystem = designName
                End If
            End If
            Dim tokenOverrides = TryCast(params("designTokens"), JObject)
            If tokenOverrides Is Nothing AndAlso designToken IsNot Nothing AndAlso designToken.Type = JTokenType.Object Then
                tokenOverrides = TryCast(designToken, JObject)
            End If
            If tokenOverrides IsNot Nothing Then result.DesignTokens = DirectCast(tokenOverrides.DeepClone(), JObject)

            Dim slides = TryCast(params("slides"), JArray)
            If slides Is Nothing Then Return result
            For index = 0 To slides.Count - 1
                Dim raw = TryCast(slides(index), JObject)
                If raw IsNot Nothing Then result.Slides.Add(SlideDesignSpec.Parse(raw, index))
            Next
            Return result
        End Function
    End Class

    Public Class SlideDesignSpec
        Public Property Id As String
        Public Property SlideType As String = "content"
        Public Property RequestedSlideType As String
        Public Property SlideTypeRecognized As Boolean = True
        Public Property Title As String
        Public Property Subtitle As String
        Public Property Eyebrow As String
        Public Property KeyMessage As String
        Public Property Body As String
        Public Property Items As New List(Of DesignItem)()
        Public Property Metrics As New List(Of DesignMetric)()
        Public Property ImagePath As String
        Public Property Notes As String
        Public Property SectionNumber As String
        Public Property Cta As String
        Public Property Source As String
        Public Property LayoutVariant As String
        Public Property ColumnHeaders As New List(Of String)()
        Public Property Chart As DesignChart
        Public Property Table As DesignTable
        Public Property XAxisLabel As String
        Public Property YAxisLabel As String

        Public Shared Function Parse(raw As JObject, index As Integer) As SlideDesignSpec
            Dim scene = TryCast(raw("scene"), JObject)
            Dim source = DirectCast(raw.DeepClone(), JObject)
            If scene IsNot Nothing Then
                source.Merge(scene, New JsonMergeSettings With {
                    .MergeArrayHandling = MergeArrayHandling.Replace
                })
            End If
            Dim requestedSlideType = FirstString(source, "slideType", "type", "layout")
            Dim slideTypeRecognized As Boolean = True
            Dim normalizedSlideType = NormalizeSlideType(requestedSlideType, slideTypeRecognized)
            Dim result As New SlideDesignSpec With {
                .Id = FirstString(source, "id"),
                .SlideType = normalizedSlideType,
                .RequestedSlideType = requestedSlideType,
                .SlideTypeRecognized = slideTypeRecognized,
                .Title = FirstString(source, "title", "headline"),
                .Subtitle = FirstString(source, "subtitle", "subheading"),
                .Eyebrow = FirstString(source, "eyebrow", "kicker"),
                .KeyMessage = FirstString(source, "keyMessage", "conclusion", "insight"),
                .Body = FirstString(source, "body", "content", "description"),
                .ImagePath = FirstString(source, "imagePath", "image"),
                .Notes = FirstString(source, "notes", "speakerNotes"),
                .SectionNumber = FirstString(source, "sectionNumber", "number"),
                .Cta = FirstString(source, "cta", "callToAction", "action"),
                .Source = FirstString(source, "source", "citation", "footnote"),
                .LayoutVariant = FirstString(source, "variant", "layoutVariant", "composition"),
                .XAxisLabel = FirstString(source, "xAxisLabel", "horizontalAxis"),
                .YAxisLabel = FirstString(source, "yAxisLabel", "verticalAxis")
            }
            result.Items = ReadItems(source)
            result.Metrics = ReadMetrics(source)
            result.ColumnHeaders = ReadStringArray(source, "columnHeaders", "headers")
            result.Chart = ReadChart(source)
            result.Table = ReadTable(source)
            If String.IsNullOrWhiteSpace(result.Id) Then result.Id = $"slide-{index + 1}"
            If result.SlideType = "content" AndAlso result.Items.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(result.Body) Then
                For Each line In SplitContent(result.Body)
                    result.Items.Add(New DesignItem With {.Title = line})
                Next
            End If
            Return result
        End Function

        Private Shared Function ReadItems(source As JObject) As List(Of DesignItem)
            Dim result As New List(Of DesignItem)()
            For Each name In {"items", "bullets", "steps", "columns", "layers", "quadrants"}
                Dim values = TryCast(source(name), JArray)
                If values Is Nothing Then Continue For
                For index = 0 To values.Count - 1
                    If values(index).Type = JTokenType.String Then
                        result.Add(ParseStringItem(values(index).ToString(), index + 1))
                    ElseIf values(index).Type = JTokenType.Object Then
                        Dim item = DirectCast(values(index), JObject)
                        Dim parsed As New DesignItem With {
                            .Index = index + 1,
                            .Title = FirstString(item, "title", "label", "name", "text"),
                            .Body = FirstString(item, "body", "description", "content"),
                            .Value = FirstString(item, "value", "metric"),
                            .Emphasis = If(item("emphasis")?.Value(Of Boolean)(), False)
                        }
                        Dim features = TryCast(item("features"), JArray)
                        If features IsNot Nothing Then
                            parsed.Features = features.Select(Function(feature) feature.ToString()).ToList()
                        End If
                        result.Add(parsed)
                    End If
                Next
                If result.Count > 0 Then Exit For
            Next
            Return result
        End Function

        Private Shared Function ParseStringItem(value As String, index As Integer) As DesignItem
            Dim text = If(value, "").Trim()
            Dim separator = text.IndexOf("："c)
            If separator < 0 Then separator = text.IndexOf(":"c)
            If separator > 0 AndAlso separator < text.Length - 1 Then
                Return New DesignItem With {
                    .Index = index,
                    .Title = text.Substring(0, separator).Trim(),
                    .Body = text.Substring(separator + 1).Trim()
                }
            End If
            Return New DesignItem With {.Index = index, .Title = text}
        End Function

        Private Shared Function ReadMetrics(source As JObject) As List(Of DesignMetric)
            Dim result As New List(Of DesignMetric)()
            Dim values = TryCast(source("metrics"), JArray)
            If values Is Nothing Then Return result
            For Each token In values
                If token.Type <> JTokenType.Object Then Continue For
                Dim item = DirectCast(token, JObject)
                result.Add(New DesignMetric With {
                    .Value = FirstString(item, "value", "number"),
                    .Label = FirstString(item, "label", "title", "name"),
                    .Delta = FirstString(item, "delta", "change"),
                    .Description = FirstString(item, "description", "body"),
                    .Source = FirstString(item, "source", "citation")
                })
            Next
            Return result
        End Function

        Private Shared Function ReadStringArray(source As JObject, ParamArray names As String()) As List(Of String)
            Dim result As New List(Of String)()
            If source Is Nothing Then Return result
            For Each name In names
                Dim values = TryCast(source.GetValue(name, StringComparison.OrdinalIgnoreCase), JArray)
                If values Is Nothing Then Continue For
                For Each value In values
                    Dim text = If(value?.ToString(), "").Trim()
                    If Not String.IsNullOrWhiteSpace(text) Then result.Add(text)
                Next
                If result.Count > 0 Then Exit For
            Next
            Return result
        End Function

        Private Shared Function ReadChart(source As JObject) As DesignChart
            Dim raw = TryCast(source?.GetValue("chart", StringComparison.OrdinalIgnoreCase), JObject)
            If raw Is Nothing Then Return Nothing
            Dim result As New DesignChart With {
                .ChartType = FirstString(raw, "chartType", "type"),
                .Title = FirstString(raw, "title", "label"),
                .Source = FirstString(raw, "source", "citation"),
                .ValueSuffix = FirstString(raw, "valueSuffix", "unit", "suffix"),
                .Categories = ReadStringArray(raw, "categories", "labels")
            }
            If String.IsNullOrWhiteSpace(result.ChartType) Then result.ChartType = "column"
            Dim series = TryCast(raw.GetValue("series", StringComparison.OrdinalIgnoreCase), JArray)
            If series IsNot Nothing Then
                For Each token In series
                    Dim item = TryCast(token, JObject)
                    If item Is Nothing Then Continue For
                    result.Series.Add(New DesignChartSeries With {
                        .Name = FirstString(item, "name", "label"),
                        .Color = FirstString(item, "color"),
                        .Values = ReadDoubleArray(item, "values", "data")
                    })
                Next
            End If
            Return result
        End Function

        Private Shared Function ReadDoubleArray(source As JObject, ParamArray names As String()) As List(Of Double)
            Dim result As New List(Of Double)()
            If source Is Nothing Then Return result
            For Each name In names
                Dim values = TryCast(source.GetValue(name, StringComparison.OrdinalIgnoreCase), JArray)
                If values Is Nothing Then Continue For
                For Each token In values
                    Dim number As Double
                    If Double.TryParse(token.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, number) Then
                        result.Add(number)
                    Else
                        result.Add(Double.NaN)
                    End If
                Next
                Exit For
            Next
            Return result
        End Function

        Private Shared Function ReadTable(source As JObject) As DesignTable
            Dim raw = TryCast(source?.GetValue("table", StringComparison.OrdinalIgnoreCase), JObject)
            If raw Is Nothing Then Return Nothing
            Dim result As New DesignTable With {
                .Title = FirstString(raw, "title", "label"),
                .Source = FirstString(raw, "source", "citation"),
                .Headers = ReadStringArray(raw, "headers", "columns")
            }
            Dim highlightToken = raw.GetValue("highlightColumn", StringComparison.OrdinalIgnoreCase)
            If highlightToken IsNot Nothing Then
                Dim highlightColumn As Integer
                If Integer.TryParse(highlightToken.ToString(), highlightColumn) Then result.HighlightColumn = highlightColumn
            End If
            Dim rows = TryCast(raw.GetValue("rows", StringComparison.OrdinalIgnoreCase), JArray)
            If rows IsNot Nothing Then
                For Each rowToken In rows
                    Dim row = TryCast(rowToken, JArray)
                    If row Is Nothing Then Continue For
                    result.Rows.Add(row.Select(Function(cell) If(cell?.ToString(), "").Trim()).ToList())
                Next
            End If
            Return result
        End Function

        Private Shared Function SplitContent(value As String) As IEnumerable(Of String)
            Return If(value, "").Split({vbCrLf, vbLf, "•", "；"}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(item) item.Trim().TrimStart("-"c, "*"c)).
                Where(Function(item) item.Length > 0).
                Take(6)
        End Function

        Private Shared Function FirstString(source As JObject, ParamArray names As String()) As String
            If source Is Nothing Then Return ""
            For Each name In names
                Dim value = ReadString(source, name)
                If Not String.IsNullOrWhiteSpace(value) Then Return value
            Next
            Return ""
        End Function

        Private Shared Function ReadString(source As JObject, name As String) As String
            Return If(source?.GetValue(name, StringComparison.OrdinalIgnoreCase)?.ToString(), "").Trim()
        End Function

        Private Shared Function NormalizeSlideType(value As String, ByRef recognized As Boolean) As String
            recognized = True
            Dim normalized = If(value, "content").Trim().ToLowerInvariant().Replace("_", "-")
            Select Case normalized
                Case "", "content", "title-content", "titleandcontent", "title-and-content", "body", "bullets"
                    Return "content"
                Case "title", "titleonly", "cover", "封面"
                    Return "cover"
                Case "section", "sectionheader", "章节"
                    Return "section"
                Case "statement", "key-message", "big-number", "观点"
                    Return "statement"
                Case "two-column", "twocontent", "twotext", "双栏"
                    Return "two-column"
                Case "comparison", "compare", "对比"
                    Return "comparison"
                Case "kpi", "metrics", "dashboard", "数据"
                    Return "kpi"
                Case "process", "flow", "timeline", "roadmap", "流程", "时间线"
                    Return "process"
                Case "architecture", "system", "架构"
                    Return "architecture"
                Case "matrix", "quadrant", "矩阵"
                    Return "matrix"
                Case "quote", "引用"
                    Return "quote"
                Case "closing", "summary", "结束"
                    Return "closing"
                Case Else
                    recognized = False
                    Return normalized
            End Select
        End Function
    End Class

    Public Class DesignItem
        Public Property Index As Integer
        Public Property Title As String
        Public Property Body As String
        Public Property Value As String
        Public Property Emphasis As Boolean
        Public Property Features As New List(Of String)()
    End Class

    Public Class DesignMetric
        Public Property Value As String
        Public Property Label As String
        Public Property Delta As String
        Public Property Description As String
        Public Property Source As String
    End Class

    Public Class DesignChart
        Public Property ChartType As String = "column"
        Public Property Title As String
        Public Property Categories As New List(Of String)()
        Public Property Series As New List(Of DesignChartSeries)()
        Public Property ValueSuffix As String
        Public Property Source As String
    End Class

    Public Class DesignChartSeries
        Public Property Name As String
        Public Property Values As New List(Of Double)()
        Public Property Color As String
    End Class

    Public Class DesignTable
        Public Property Title As String
        Public Property Headers As New List(Of String)()
        Public Property Rows As New List(Of List(Of String))()
        Public Property HighlightColumn As Integer = -1
        Public Property Source As String
    End Class

    Public Class DesignTokens
        Public Property Name As String
        Public Property Background As String
        Public Property Surface As String
        Public Property SurfaceAlt As String
        Public Property Primary As String
        Public Property Secondary As String
        Public Property TextPrimary As String
        Public Property TextSecondary As String
        Public Property Divider As String
        Public Property Positive As String
        Public Property Negative As String
        Public Property FontFamily As String
        Public Property DisplaySize As Single
        Public Property TitleSize As Single
        Public Property BodySize As Single
        Public Property CaptionSize As Single
        Public Property Dark As Boolean
    End Class

    Public Class SceneRect
        Public Property X As Single
        Public Property Y As Single
        Public Property Width As Single
        Public Property Height As Single

        Public Sub New()
        End Sub

        Public Sub New(x As Single, y As Single, width As Single, height As Single)
            Me.X = x : Me.Y = y : Me.Width = width : Me.Height = height
        End Sub
    End Class

    Public Class SceneNode
        Public Property Id As String
        Public Property Kind As String
        Public Property Bounds As SceneRect
        Public Property Text As String
        Public Property StyleRole As String
        Public Property FillColor As String
        Public Property LineColor As String
        Public Property TextColor As String
        Public Property FontSize As Single
        Public Property Bold As Boolean
        Public Property Alignment As String = "left"
        Public Property FillTransparency As Single
        Public Property LineWeight As Single = 1
        Public Property CornerRadius As Single = 10
        Public Property Collision As Boolean = True
        Public Property Shadow As Boolean
        Public Property ImagePath As String
    End Class

    Public Class SlideRenderPlan
        Public Property SlideType As String
        Public Property Background As String
        Public Property Nodes As New List(Of SceneNode)()
        Public Property Notes As String
    End Class

    Public Class VisualIssue
        Public Property Code As String
        Public Property NodeId As String
        Public Property Message As String
        Public Property Severity As String
        Public Property Repaired As Boolean
    End Class

    Public Class VisualVerificationReport
        Public Property Issues As New List(Of VisualIssue)()
        Public Property RepairCount As Integer
        Public Property AestheticScore As Integer = 100
        Public Property Metrics As New Dictionary(Of String, Double)()
        Public ReadOnly Property Passed As Boolean
            Get
                Return Not Issues.Any(Function(item) item.Severity = "error" AndAlso Not item.Repaired)
            End Get
        End Property
    End Class

End Namespace
