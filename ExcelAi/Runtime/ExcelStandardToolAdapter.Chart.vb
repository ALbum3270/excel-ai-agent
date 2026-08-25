Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent
Imports ExcelAgent.Core.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend NotInheritable Partial Class ExcelStandardToolAdapter
        Private Shared Function ExecuteCreateChart(application As Object, params As JObject) As ToolResult
            Const toolId As String = "CreateChart"
            Dim sourceSpec = FirstText(params, "dataRange")
            If String.IsNullOrWhiteSpace(sourceSpec) Then Return Invalid(toolId, "CreateChart requires dataRange")
            Dim source = ResolveRange(application, sourceSpec, captureValues:=False)
            Dim chartType = ParseChartType(FirstText(params, "type"))
            Dim position = ResolveChartPosition(application, source, FirstText(params, "position"))
            Dim chartObjectsRef = BuildWorksheetRef(source.SheetName) & "/chartobjects"
            Dim addBatch = NewBatch()
            AddOperation(addBatch, "add-chart", chartObjectsRef, "create", "ChartObjects", "Add", "method",
                         New JObject From {{"Left", position.Left}, {"Top", position.Top}, {"Width", 450}, {"Height", 320}},
                         New JObject From {{"exists", True}})
            Dim addResult = ExecuteBatch(application, toolId, addBatch, "Excel 已创建待配置图表")
            If Not addResult.Success Then Return addResult
            Dim chartObjectRef = GetOperationResultRef(addResult, 0)
            If String.IsNullOrWhiteSpace(chartObjectRef) Then Return SemanticFailure(toolId, addResult, "ChartObjects.Add returned no canonical object ref")
            Dim chartRef = chartObjectRef & "/chart"
            Dim configureBatch = NewBatch()
            AddOperation(configureBatch, "set-chart-type", chartRef, "set", "Chart", "ChartType", "property",
                         New JObject From {{"value", chartType}}, New JObject From {{"ChartType", chartType}})
            Dim setSourceArguments As New JObject From {{"Source", New JObject From {{"ref", source.RangeRef}}}}
            Dim plotBy = FirstText(params, "plotBy")
            If Not String.IsNullOrWhiteSpace(plotBy) Then setSourceArguments("PlotBy") = If(String.Equals(plotBy, "row", StringComparison.OrdinalIgnoreCase), CInt(XlRowCol.xlRows), CInt(XlRowCol.xlColumns))
            AddOperation(configureBatch, "set-chart-source", chartRef, "invoke", "Chart", "SetSourceData", "method", setSourceArguments, New JObject())
            Dim title = FirstText(params, "title")
            If Not String.IsNullOrWhiteSpace(title) Then
                AddOperation(configureBatch, "enable-title", chartRef, "set", "Chart", "HasTitle", "property",
                             New JObject From {{"value", True}}, New JObject From {{"HasTitle", True}})
                AddOperation(configureBatch, "set-title", chartRef & "/charttitle", "set", "ChartTitle", "Text", "property",
                             New JObject From {{"value", title}}, New JObject From {{"Text", title}})
            End If
            AddOperation(configureBatch, "enable-legend", chartRef, "set", "Chart", "HasLegend", "property",
                         New JObject From {{"value", True}}, New JObject From {{"HasLegend", True}})
            Dim legendPosition = FirstText(params, "legendPosition")
            If Not String.IsNullOrWhiteSpace(legendPosition) Then
                Dim parsedLegend = ParseLegendPosition(legendPosition)
                AddOperation(configureBatch, "set-legend", chartRef & "/legend", "set", "Legend", "Position", "property",
                             New JObject From {{"value", parsedLegend}}, New JObject From {{"Position", parsedLegend}})
            End If
            configureBatch.SuccessCriteria.Add(New OperationCriterion With {
                .Id = "chart-has-series", .TargetRef = chartRef, .PropertyName = "SeriesCount",
                .Operator = "gte", .ExpectedValue = New JValue(1), .Required = True
            })
            Dim configured = ExecuteBatch(application, toolId, configureBatch, "已验证图表类型、数据源和基础外观")
            If Not configured.Success Then Return MarkCompositeMutationFailure(configured, chartObjectRef)
            configured = AnnotateVerifiedRequestProjection(
                configured,
                New JValue(If(String.IsNullOrWhiteSpace(FirstText(params, "type")), "column", FirstText(params, "type"))),
                requestProperty:="type",
                verifiedPropertyName:="ChartType")
            If Not String.IsNullOrWhiteSpace(title) Then
                configured = AnnotateVerifiedRequestProjection(
                    configured,
                    New JValue(title),
                    requestProperty:="title",
                    verifiedPropertyName:="Text")
            End If
            If Not String.IsNullOrWhiteSpace(legendPosition) Then
                configured = AnnotateVerifiedRequestProjection(
                    configured,
                    New JValue(legendPosition),
                    requestProperty:="legendPosition",
                    verifiedPropertyName:="Position")
            End If
            Dim finalResult = ConfigureChartSeries(application, params, chartRef, configured)
            If Not finalResult.Success Then Return MarkCompositeMutationFailure(finalResult, chartObjectRef)
            Return AnnotateArtifactAnchor(finalResult, position.TargetRef, chartObjectRef)
        End Function

        Private Shared Function ConfigureChartSeries(application As Object,
                                                       params As JObject,
                                                       chartRef As String,
                                                       currentResult As ToolResult) As ToolResult
            Dim seriesNames = params("seriesNames")
            Dim categoryAxis = params("categoryAxis")
            If seriesNames Is Nothing AndAlso categoryAxis Is Nothing Then Return currentResult
            Dim batch = NewBatch()
            If seriesNames IsNot Nothing Then
                Dim names = ResolveSequenceValues(application, seriesNames)
                For index = 0 To names.Count - 1
                    AddOperation(batch,
                                 "series-name-" & (index + 1).ToString(CultureInfo.InvariantCulture),
                                 chartRef & "/series/" & (index + 1).ToString(CultureInfo.InvariantCulture),
                                 "set", "Series", "Name", "property",
                                 New JObject From {{"value", names(index).DeepClone()}},
                                 New JObject From {{"Name", names(index).DeepClone()}})
                Next
            End If
            If categoryAxis IsNot Nothing Then
                Dim argument As JToken
                Dim expectedHash As String
                If categoryAxis.Type = JTokenType.String Then
                    Dim axisRange = ResolveRange(application, categoryAxis.ToString(), captureValues:=True)
                    argument = New JObject From {{"ref", axisRange.RangeRef}}
                    expectedHash = ExcelOperationObserver.ComputeSequenceHash(axisRange.Values)
                Else
                    argument = categoryAxis.DeepClone()
                    expectedHash = ExcelOperationObserver.ComputeSequenceHash(categoryAxis)
                End If
                AddOperation(batch, "set-category-axis", chartRef & "/series/1", "set", "Series", "XValues", "property",
                             New JObject From {{"value", argument}}, New JObject From {{"XValuesHash", expectedHash}})
            End If
            If batch.Operations.Count = 0 Then Return currentResult
            Dim seriesResult = ExecuteBatch(application, "CreateChart", batch, "已验证图表系列名称和分类轴")
            If Not seriesResult.Success Then Return seriesResult
            Return MergeCompositeSuccess(currentResult, seriesResult)
        End Function

        Private Shared Function ResolveChartPosition(application As Object,
                                                     source As RangeDescriptor,
                                                     positionSpec As String) As ChartPosition
            Dim targetSpec = positionSpec
            If String.IsNullOrWhiteSpace(targetSpec) Then targetSpec = QuoteSheet(source.SheetName) & "!" & source.Address
            If targetSpec.IndexOf("!"c) < 0 Then targetSpec = QuoteSheet(source.SheetName) & "!" & targetSpec
            Dim descriptor = ResolveRange(application, targetSpec, captureValues:=False)
            Using resolved = ExcelObjectResolver.Resolve(application, descriptor.RangeRef)
                Dim left = CDbl(resolved.Value.Left)
                If String.IsNullOrWhiteSpace(positionSpec) Then left += CDbl(resolved.Value.Width) + 20.0R
                Return New ChartPosition With {
                    .Left = left,
                    .Top = CDbl(resolved.Value.Top),
                    .TargetRef = descriptor.RangeRef
                }
            End Using
        End Function

        Private Shared Function ResolveSequenceValues(application As Object, token As JToken) As JArray
            If token Is Nothing Then Return New JArray()
            If token.Type = JTokenType.Array Then Return New JArray(DirectCast(token, JArray).Select(Function(item) item.DeepClone()))
            Dim descriptor = ResolveRange(application, token.ToString(), captureValues:=True)
            Dim result As New JArray()
            For Each item In FlattenValues(descriptor.Values)
                result.Add(If(item Is Nothing, JValue.CreateNull(), JToken.FromObject(item)))
            Next
            Return result
        End Function

        Private Shared Function ParseChartType(value As String) As Integer
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "line" : Return CInt(XlChartType.xlLine)
                Case "pie" : Return CInt(XlChartType.xlPie)
                Case "bar" : Return CInt(XlChartType.xlBarClustered)
                Case "scatter" : Return CInt(XlChartType.xlXYScatter)
                Case "area" : Return CInt(XlChartType.xlArea)
                Case Else : Return CInt(XlChartType.xlColumnClustered)
            End Select
        End Function

        Private Shared Function ParseLegendPosition(value As String) As Integer
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "left" : Return CInt(XlLegendPosition.xlLegendPositionLeft)
                Case "top" : Return CInt(XlLegendPosition.xlLegendPositionTop)
                Case "bottom" : Return CInt(XlLegendPosition.xlLegendPositionBottom)
                Case "corner" : Return CInt(XlLegendPosition.xlLegendPositionCorner)
                Case Else : Return CInt(XlLegendPosition.xlLegendPositionRight)
            End Select
        End Function

        Private Class ChartPosition
            Public Property Left As Double
            Public Property Top As Double
            Public Property TargetRef As String
        End Class
    End Class

End Namespace
