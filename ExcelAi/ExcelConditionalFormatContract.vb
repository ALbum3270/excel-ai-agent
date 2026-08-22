Imports System.Globalization
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq

''' <summary>
''' Canonical contract shared by the ConditionalFormat executor and observer.
''' It converts the compact model-facing condition/color syntax into Excel values,
''' then verifies the post-state against the same normalized representation.
''' </summary>
Friend NotInheritable Class ExcelConditionalFormatContract

    Private Shared ReadOnly NamedColors As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
        {"black", ExcelRgb(0, 0, 0)},
        {"white", ExcelRgb(255, 255, 255)},
        {"red", ExcelRgb(255, 0, 0)},
        {"green", ExcelRgb(0, 128, 0)},
        {"lime", ExcelRgb(0, 255, 0)},
        {"lightgreen", ExcelRgb(144, 238, 144)},
        {"darkgreen", ExcelRgb(0, 100, 0)},
        {"blue", ExcelRgb(0, 0, 255)},
        {"lightblue", ExcelRgb(173, 216, 230)},
        {"darkblue", ExcelRgb(0, 0, 139)},
        {"yellow", ExcelRgb(255, 255, 0)},
        {"orange", ExcelRgb(255, 165, 0)},
        {"purple", ExcelRgb(128, 0, 128)},
        {"pink", ExcelRgb(255, 192, 203)},
        {"gray", ExcelRgb(128, 128, 128)},
        {"grey", ExcelRgb(128, 128, 128)},
        {"lightgray", ExcelRgb(211, 211, 211)},
        {"lightgrey", ExcelRgb(211, 211, 211)},
        {"红色", ExcelRgb(255, 0, 0)},
        {"绿色", ExcelRgb(0, 128, 0)},
        {"浅绿色", ExcelRgb(144, 238, 144)},
        {"蓝色", ExcelRgb(0, 0, 255)},
        {"浅蓝色", ExcelRgb(173, 216, 230)},
        {"黄色", ExcelRgb(255, 255, 0)},
        {"橙色", ExcelRgb(255, 165, 0)},
        {"紫色", ExcelRgb(128, 0, 128)},
        {"粉色", ExcelRgb(255, 192, 203)},
        {"灰色", ExcelRgb(128, 128, 128)},
        {"浅灰色", ExcelRgb(211, 211, 211)},
        {"白色", ExcelRgb(255, 255, 255)},
        {"黑色", ExcelRgb(0, 0, 0)}
    }

    Private Sub New()
    End Sub

    Friend Shared Function ParseHighlightCondition(rawCondition As String) As ExcelHighlightCondition
        Dim normalized = If(rawCondition, "").Trim()
        If String.IsNullOrWhiteSpace(normalized) Then normalized = ">0"

        normalized = Regex.Replace(normalized, "^大于等于\s*", ">=")
        normalized = Regex.Replace(normalized, "^小于等于\s*", "<=")
        normalized = Regex.Replace(normalized, "^不等于\s*", "<>")
        normalized = Regex.Replace(normalized, "^大于\s*", ">")
        normalized = Regex.Replace(normalized, "^小于\s*", "<")
        normalized = Regex.Replace(normalized, "^等于\s*", "=")
        normalized = Regex.Replace(normalized, "^greater\s+than\s+or\s+equal\s+to\s*", ">=", RegexOptions.IgnoreCase)
        normalized = Regex.Replace(normalized, "^less\s+than\s+or\s+equal\s+to\s*", "<=", RegexOptions.IgnoreCase)
        normalized = Regex.Replace(normalized, "^greater\s+than\s*", ">", RegexOptions.IgnoreCase)
        normalized = Regex.Replace(normalized, "^less\s+than\s*", "<", RegexOptions.IgnoreCase)
        normalized = Regex.Replace(normalized, "^not\s+equal\s+to\s*", "<>", RegexOptions.IgnoreCase)
        normalized = Regex.Replace(normalized, "^equal\s+to\s*", "=", RegexOptions.IgnoreCase)

        Dim match = Regex.Match(normalized, "^(>=|<=|<>|!=|==|=|>|<)\s*(.+)$")
        Dim operatorText As String
        Dim operand As String
        If match.Success Then
            operatorText = match.Groups(1).Value
            operand = match.Groups(2).Value.Trim()
        Else
            ' Backward compatibility: a bare value means greater-than, matching the
            ' behavior of the original tool implementation.
            operatorText = ">"
            operand = normalized
        End If

        If operand.StartsWith("=", StringComparison.Ordinal) Then operand = operand.Substring(1).Trim()
        If String.IsNullOrWhiteSpace(operand) Then Throw New FormatException("条件格式缺少比较值")

        Dim numericValue As Decimal
        If Decimal.TryParse(operand, NumberStyles.Number, CultureInfo.InvariantCulture, numericValue) OrElse
           Decimal.TryParse(operand, NumberStyles.Number, CultureInfo.CurrentCulture, numericValue) Then
            operand = numericValue.ToString(CultureInfo.InvariantCulture)
        End If

        Dim excelOperator As XlFormatConditionOperator
        Select Case operatorText
            Case ">"
                excelOperator = XlFormatConditionOperator.xlGreater
            Case ">="
                excelOperator = XlFormatConditionOperator.xlGreaterEqual
            Case "<"
                excelOperator = XlFormatConditionOperator.xlLess
            Case "<="
                excelOperator = XlFormatConditionOperator.xlLessEqual
            Case "=", "=="
                excelOperator = XlFormatConditionOperator.xlEqual
            Case "<>", "!="
                excelOperator = XlFormatConditionOperator.xlNotEqual
            Case Else
                Throw New FormatException($"不支持的条件运算符: {operatorText}")
        End Select

        Return New ExcelHighlightCondition With {
            .OperatorText = operatorText,
            .ExcelOperator = excelOperator,
            .FormulaOperand = operand
        }
    End Function

    Friend Shared Function ParseColor(colorText As String) As Integer
        Dim normalized = If(colorText, "").Trim()
        If String.IsNullOrWhiteSpace(normalized) Then Throw New FormatException("颜色不能为空")

        Dim compactName = Regex.Replace(normalized, "[\s_-]+", "")
        Dim namedValue As Integer
        If NamedColors.TryGetValue(compactName, namedValue) Then Return namedValue

        Dim hex = normalized
        If hex.StartsWith("#", StringComparison.Ordinal) Then hex = hex.Substring(1)
        If Regex.IsMatch(hex, "^[0-9a-fA-F]{3}$") Then
            hex = String.Concat(hex(0), hex(0), hex(1), hex(1), hex(2), hex(2))
        End If
        If Regex.IsMatch(hex, "^[0-9a-fA-F]{6}$") Then
            Return ExcelRgb(Convert.ToInt32(hex.Substring(0, 2), 16),
                            Convert.ToInt32(hex.Substring(2, 2), 16),
                            Convert.ToInt32(hex.Substring(4, 2), 16))
        End If

        Dim rgbMatch = Regex.Match(normalized, "^rgb\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$", RegexOptions.IgnoreCase)
        If rgbMatch.Success Then
            Dim red = Integer.Parse(rgbMatch.Groups(1).Value, CultureInfo.InvariantCulture)
            Dim green = Integer.Parse(rgbMatch.Groups(2).Value, CultureInfo.InvariantCulture)
            Dim blue = Integer.Parse(rgbMatch.Groups(3).Value, CultureInfo.InvariantCulture)
            If red <= 255 AndAlso green <= 255 AndAlso blue <= 255 Then Return ExcelRgb(red, green, blue)
        End If

        Throw New FormatException($"不支持的颜色: {colorText}；请使用 #RRGGBB、rgb(r,g,b) 或受支持的颜色名称")
    End Function

    Friend Shared Function EvaluatePostState(ruleText As String,
                                             conditionText As String,
                                             colorText As String,
                                             actualRules As JArray) As JObject
        Dim normalizedRule = If(ruleText, "").Trim().ToLowerInvariant()
        Dim rules = If(actualRules, New JArray())
        Dim expected As New JObject From {{"rule", normalizedRule}}

        Try
            Select Case normalizedRule
                Case "highlight"
                    Dim condition = ParseHighlightCondition(conditionText)
                    Dim expectedColor = ParseColor(If(String.IsNullOrWhiteSpace(colorText), "#FFC7CE", colorText))
                    expected("type") = CInt(XlFormatConditionType.xlCellValue)
                    expected("operator") = CInt(condition.ExcelOperator)
                    expected("formula1") = condition.FormulaOperand
                    expected("interiorColor") = expectedColor

                    For Each rule As JObject In rules.OfType(Of JObject)()
                        If TokenInteger(rule, "type") = CInt(XlFormatConditionType.xlCellValue) AndAlso
                           TokenInteger(rule, "operator") = CInt(condition.ExcelOperator) AndAlso
                           String.Equals(NormalizeObservedFormula(rule("formula1")?.ToString()),
                                         condition.FormulaOperand,
                                         StringComparison.OrdinalIgnoreCase) AndAlso
                           TokenInteger(rule, "interiorColor", Integer.MinValue) = expectedColor Then
                            Return VerificationResult(True, "matched", expected, rules.Count)
                        End If
                    Next
                    Return VerificationResult(False, "no_matching_highlight_rule", expected, rules.Count)

                Case "databar", "colorscale", "iconset"
                    Dim expectedType As Integer
                    Select Case normalizedRule
                        Case "databar"
                            expectedType = CInt(XlFormatConditionType.xlDataBar)
                        Case "colorscale"
                            expectedType = CInt(XlFormatConditionType.xlColorScale)
                        Case Else
                            expectedType = CInt(XlFormatConditionType.xlIconSets)
                    End Select
                    expected("type") = expectedType
                    Dim matched = rules.OfType(Of JObject)().Any(Function(rule) TokenInteger(rule, "type") = expectedType)
                    Return VerificationResult(matched, If(matched, "matched", "no_matching_rule_type"), expected, rules.Count)

                Case Else
                    Return VerificationResult(False, "unsupported_rule", expected, rules.Count)
            End Select
        Catch ex As Exception
            expected("contractError") = ex.Message
            Return VerificationResult(False, "invalid_expected_contract", expected, rules.Count)
        End Try
    End Function

    Private Shared Function NormalizeObservedFormula(formula As String) As String
        Dim normalized = If(formula, "").Trim()
        If normalized.StartsWith("=", StringComparison.Ordinal) Then normalized = normalized.Substring(1).Trim()
        If normalized.Length >= 2 AndAlso normalized.StartsWith("""", StringComparison.Ordinal) AndAlso normalized.EndsWith("""", StringComparison.Ordinal) Then
            normalized = normalized.Substring(1, normalized.Length - 2).Replace("""""", """")
        End If
        Return normalized.Trim()
    End Function

    Private Shared Function VerificationResult(satisfied As Boolean,
                                               reason As String,
                                               expected As JObject,
                                               actualRuleCount As Integer) As JObject
        Return New JObject From {
            {"satisfied", satisfied},
            {"reason", reason},
            {"expected", expected},
            {"actualRuleCount", actualRuleCount}
        }
    End Function

    Private Shared Function TokenInteger(token As JObject, name As String, Optional fallback As Integer = 0) As Integer
        If token Is Nothing OrElse token(name) Is Nothing Then Return fallback
        Dim value As Integer
        If Integer.TryParse(token(name).ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, value) Then Return value
        Return fallback
    End Function

    Private Shared Function ExcelRgb(red As Integer, green As Integer, blue As Integer) As Integer
        Return red Or (green << 8) Or (blue << 16)
    End Function
End Class

Friend NotInheritable Class ExcelHighlightCondition
    Friend Property OperatorText As String
    Friend Property ExcelOperator As XlFormatConditionOperator
    Friend Property FormulaOperand As String
End Class
