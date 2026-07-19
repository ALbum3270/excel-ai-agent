' ShareRibbon\Controls\Services\ExecutionPlanRenderer.vb
' 执行计划渲染服务：将JSON命令转换为用户友好的执行步骤

Imports Newtonsoft.Json.Linq

''' <summary>
''' 执行计划渲染服务
''' 将大模型返回的JSON命令转换为用户可理解的执行步骤
''' </summary>
Public Class ExecutionPlanRenderer

#Region "命令描述映射"

    ' 操作类型中文描述
    Private Shared ReadOnly OperationDescriptions As New Dictionary(Of String, String) From {
        {"removeDuplicates", "删除重复项"},
        {"fillEmpty", "填充空值"},
        {"trim", "去除空格"},
        {"replace", "替换内容"},
        {"transpose", "转置数据"},
        {"split", "拆分列"},
        {"merge", "合并列"},
        {"summary", "生成摘要"},
        {"pivot", "创建透视表"},
        {"groupby", "分组汇总"},
        {"ranking", "排名分析"}
    }

    ' 图表类型中文描述
    Private Shared ReadOnly ChartTypeDescriptions As New Dictionary(Of String, String) From {
        {"Column", "柱状图"},
        {"Line", "折线图"},
        {"Pie", "饼图"},
        {"Bar", "条形图"},
        {"Scatter", "散点图"},
        {"Area", "面积图"}
    }

#End Region

#Region "公共方法"

    ''' <summary>
    ''' 将JSON命令解析为执行计划
    ''' </summary>
    Public Function ParseJsonToExecutionPlan(jsonCommand As String) As List(Of ExecutionStep)
        Dim plan As New List(Of ExecutionStep)()

        Try
            Dim json = JObject.Parse(jsonCommand)
            Dim command = json("command")?.ToString()
            Dim params = json("params")

            If String.IsNullOrEmpty(command) Then
                Return plan
            End If

            ' 根据命令类型生成步骤
            Select Case command.ToLower()
                Case "applyformula", "formula", "calculate"
                    plan.AddRange(GenerateFormulaSteps(params))
                Case "writedata", "write", "setvalue"
                    plan.AddRange(GenerateWriteDataSteps(params))
                Case "formatrange", "format", "style"
                    plan.AddRange(GenerateFormatSteps(params))
                Case "createchart", "chart"
                    plan.AddRange(GenerateChartSteps(params))
                Case "cleandata", "clean"
                    plan.AddRange(GenerateCleanDataSteps(params))
                Case "dataanalysis", "analyze"
                    plan.AddRange(GenerateAnalysisSteps(params))
                Case "transformdata", "transform"
                    plan.AddRange(GenerateTransformSteps(params))
                Case "generatereport", "report"
                    plan.AddRange(GenerateReportSteps(params))
                Case Else
                    ' 通用处理
                    plan.Add(New ExecutionStep(1, $"执行 {command} 命令", "default"))
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ParseJsonToExecutionPlan 出错: {ex.Message}")
            plan.Add(New ExecutionStep(1, "解析命令失败", "default"))
        End Try

        Return plan
    End Function

#End Region

#Region "步骤生成方法"

    ''' <summary>
    ''' 生成公式应用步骤
    ''' </summary>
    Private Function GenerateFormulaSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim targetRange = If(params?("targetRange")?.ToString(), "目标区域")
        Dim formula = If(params?("formula")?.ToString(), "")
        Dim fillDown = If(params?("fillDown")?.Value(Of Boolean)(), False)

        steps.Add(New ExecutionStep(1, $"在 {targetRange} 应用公式", "formula") With {
            .WillModify = targetRange,
            .EstimatedTime = "1秒"
        })

        If Not String.IsNullOrEmpty(formula) Then
            Dim formulaDesc = GetFormulaDescription(formula)
            steps.Add(New ExecutionStep(2, $"公式内容: {formulaDesc}", "formula"))
        End If

        If fillDown Then
            steps.Add(New ExecutionStep(3, "自动向下填充公式", "formula"))
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据写入步骤
    ''' </summary>
    Private Function GenerateWriteDataSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim targetRange = If(params?("targetRange")?.ToString(), "目标区域")
        
        steps.Add(New ExecutionStep(1, $"向 {targetRange} 写入数据", "data") With {
            .WillModify = targetRange,
            .EstimatedTime = "1秒"
        })

        Return steps
    End Function

    ''' <summary>
    ''' 生成格式化步骤
    ''' </summary>
    Private Function GenerateFormatSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim range = If(params?("range")?.ToString(), If(params?("targetRange")?.ToString(), "目标区域"))
        Dim style = If(params?("style")?.ToString(), "")
        
        steps.Add(New ExecutionStep(1, $"选择 {range} 区域", "search") With {
            .EstimatedTime = "1秒"
        })

        Dim formatDesc = "应用格式设置"
        If Not String.IsNullOrEmpty(style) Then
            formatDesc = $"应用 {style} 样式"
        End If

        Dim formatDetails As New List(Of String)()
        If params?("bold")?.Value(Of Boolean)() = True Then formatDetails.Add("加粗")
        If params?("italic")?.Value(Of Boolean)() = True Then formatDetails.Add("斜体")
        If params?("borders")?.Value(Of Boolean)() = True Then formatDetails.Add("边框")
        
        If formatDetails.Count > 0 Then
            formatDesc &= $" ({String.Join(", ", formatDetails)})"
        End If

        steps.Add(New ExecutionStep(2, formatDesc, "format") With {
            .WillModify = range,
            .EstimatedTime = "1秒"
        })

        Return steps
    End Function

    ''' <summary>
    ''' 生成图表创建步骤
    ''' </summary>
    Private Function GenerateChartSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim chartType = If(params?("type")?.ToString(), "Column")
        Dim dataRange = If(params?("dataRange")?.ToString(), "数据区域")
        Dim title = If(params?("title")?.ToString(), "")
        Dim position = If(params?("position")?.ToString(), "")

        Dim chartTypeName = If(ChartTypeDescriptions.ContainsKey(chartType), ChartTypeDescriptions(chartType), chartType)

        steps.Add(New ExecutionStep(1, $"读取 {dataRange} 作为图表数据源", "search") With {
            .EstimatedTime = "1秒"
        })

        steps.Add(New ExecutionStep(2, $"创建 {chartTypeName}", "chart") With {
            .EstimatedTime = "2秒"
        })

        If Not String.IsNullOrEmpty(title) Then
            steps.Add(New ExecutionStep(3, $"设置图表标题: {title}", "chart"))
        End If

        If Not String.IsNullOrEmpty(position) Then
            steps.Add(New ExecutionStep(4, $"将图表放置在 {position}", "chart") With {
                .WillModify = position
            })
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据清洗步骤
    ''' </summary>
    Private Function GenerateCleanDataSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim operation = If(params?("operation")?.ToString(), "clean")
        Dim range = If(params?("range")?.ToString(), "数据区域")

        Dim operationDesc = If(OperationDescriptions.ContainsKey(operation), OperationDescriptions(operation), operation)

        steps.Add(New ExecutionStep(1, $"扫描 {range} 区域", "search") With {
            .EstimatedTime = "1秒"
        })

        steps.Add(New ExecutionStep(2, $"执行清洗操作: {operationDesc}", "clean") With {
            .WillModify = range,
            .EstimatedTime = "2秒"
        })

        steps.Add(New ExecutionStep(3, "验证清洗结果", "data"))

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据分析步骤
    ''' </summary>
    Private Function GenerateAnalysisSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim analysisType = If(params?("type")?.ToString(), "summary")
        Dim sourceRange = If(params?("sourceRange")?.ToString(), "数据区域")
        Dim targetRange = If(params?("targetRange")?.ToString(), "")

        Dim analysisDesc = If(OperationDescriptions.ContainsKey(analysisType), OperationDescriptions(analysisType), analysisType)

        steps.Add(New ExecutionStep(1, $"读取 {sourceRange} 数据", "search") With {
            .EstimatedTime = "1秒"
        })

        steps.Add(New ExecutionStep(2, $"执行分析: {analysisDesc}", "data") With {
            .EstimatedTime = "3秒"
        })

        If Not String.IsNullOrEmpty(targetRange) Then
            steps.Add(New ExecutionStep(3, $"输出结果到 {targetRange}", "data") With {
                .WillModify = targetRange
            })
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据转换步骤
    ''' </summary>
    Private Function GenerateTransformSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim operation = If(params?("operation")?.ToString(), "transform")
        Dim sourceRange = If(params?("sourceRange")?.ToString(), "源区域")
        Dim targetRange = If(params?("targetRange")?.ToString(), "")

        Dim operationDesc = If(OperationDescriptions.ContainsKey(operation), OperationDescriptions(operation), operation)

        steps.Add(New ExecutionStep(1, $"读取 {sourceRange} 数据", "search"))
        steps.Add(New ExecutionStep(2, $"执行转换: {operationDesc}", "data") With {
            .EstimatedTime = "2秒"
        })

        If Not String.IsNullOrEmpty(targetRange) Then
            steps.Add(New ExecutionStep(3, $"输出到 {targetRange}", "data") With {
                .WillModify = targetRange
            })
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成报表生成步骤
    ''' </summary>
    Private Function GenerateReportSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim sourceRange = If(params?("sourceRange")?.ToString(), "数据区域")
        Dim targetSheet = If(params?("targetSheet")?.ToString(), "新工作表")
        Dim title = If(params?("title")?.ToString(), "报表")
        Dim includeChart = If(params?("includeChart")?.Value(Of Boolean)(), False)

        steps.Add(New ExecutionStep(1, $"收集 {sourceRange} 数据", "search"))
        steps.Add(New ExecutionStep(2, $"创建报表工作表: {targetSheet}", "data") With {
            .EstimatedTime = "1秒"
        })
        steps.Add(New ExecutionStep(3, $"填充数据并设置标题: {title}", "data"))
        steps.Add(New ExecutionStep(4, "应用报表格式", "format") With {
            .EstimatedTime = "2秒"
        })

        If includeChart Then
            steps.Add(New ExecutionStep(5, "添加数据图表", "chart") With {
                .EstimatedTime = "2秒"
            })
        End If

        Return steps
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 获取公式的友好描述
    ''' </summary>
    Private Function GetFormulaDescription(formula As String) As String
        If String.IsNullOrEmpty(formula) Then Return ""

        ' 移除开头的=
        formula = formula.TrimStart("="c)

        ' 识别常见公式
        Dim upperFormula = formula.ToUpper()
        
        If upperFormula.StartsWith("SUM(") Then
            Return "求和"
        ElseIf upperFormula.StartsWith("AVERAGE(") Then
            Return "计算平均值"
        ElseIf upperFormula.StartsWith("COUNT(") Then
            Return "计数"
        ElseIf upperFormula.StartsWith("MAX(") Then
            Return "取最大值"
        ElseIf upperFormula.StartsWith("MIN(") Then
            Return "取最小值"
        ElseIf upperFormula.StartsWith("VLOOKUP(") Then
            Return "垂直查找"
        ElseIf upperFormula.StartsWith("IF(") Then
            Return "条件判断"
        ElseIf upperFormula.StartsWith("SUMIF(") Then
            Return "条件求和"
        ElseIf upperFormula.StartsWith("COUNTIF(") Then
            Return "条件计数"
        ElseIf upperFormula.Contains("+") Then
            Return "加法运算"
        ElseIf upperFormula.Contains("-") Then
            Return "减法运算"
        ElseIf upperFormula.Contains("*") Then
            Return "乘法运算"
        ElseIf upperFormula.Contains("/") Then
            Return "除法运算"
        Else
            ' 截断过长的公式
            If formula.Length > 30 Then
                Return formula.Substring(0, 27) & "..."
            End If
            Return formula
        End If
    End Function

#End Region

End Class
