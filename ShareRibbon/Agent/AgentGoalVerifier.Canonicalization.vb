Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public NotInheritable Partial Class AgentGoalVerifier

        Private Shared Function CanonicalObjectIdentity(value As String) As String
            Dim decoded = Uri.UnescapeDataString(If(value, "")).Trim().Replace("\", "/")
            decoded = Regex.Replace(decoded, "/+", "/").TrimEnd("/"c)

            Dim workbookMatch = Regex.Match(
                decoded,
                "(?:^|Excel:|/)workbooks/(?<workbook>[^/]+)(?:/|$)",
                RegexOptions.IgnoreCase)
            Dim workbook = If(workbookMatch.Success, workbookMatch.Groups("workbook").Value, "active")

            Dim worksheetMatch = Regex.Match(decoded, "(?:^|/)worksheets/(?<sheet>[^/!]+)", RegexOptions.IgnoreCase)
            If worksheetMatch.Success Then
                Dim worksheetIdentity = BuildWorksheetIdentity(workbook, worksheetMatch.Groups("sheet").Value)
                Dim tailStart = worksheetMatch.Index + worksheetMatch.Length
                Dim tail = decoded.Substring(tailStart).Trim("/"c)
                If String.IsNullOrWhiteSpace(tail) Then Return worksheetIdentity
                Return worksheetIdentity & "/" & tail.ToLowerInvariant()
            End If

            Dim excelSheetMatch = Regex.Match(decoded, "^Excel:(?<sheet>[^!/:]+)(?:!|$)", RegexOptions.IgnoreCase)
            If excelSheetMatch.Success Then
                Return BuildWorksheetIdentity("active", excelSheetMatch.Groups("sheet").Value)
            End If

            If decoded.IndexOf("/"c) < 0 AndAlso decoded.IndexOf("!"c) < 0 AndAlso
               decoded.IndexOf(":"c) < 0 Then
                Return BuildWorksheetIdentity("active", decoded)
            End If
            If workbookMatch.Success Then Return "workbook:" & NormalizeWorkbookName(workbook)
            Return NormalizeTarget(decoded)
        End Function

        Private Shared Function BuildWorksheetIdentity(workbook As String, sheet As String) As String
            Return "workbook:" & NormalizeWorkbookName(workbook) &
                "/worksheet:" & NormalizeWorksheetName(sheet)
        End Function

        Friend Shared Function CanonicalContractTargetReference(value As String) As String
            Dim rangeRef As ExcelRangeRef = Nothing
            If TryParseExcelRange(value, rangeRef) Then
                Return String.Join(":", {
                    "range",
                    NormalizeWorkbookName(rangeRef.Workbook),
                    NormalizeWorksheetName(rangeRef.Sheet),
                    NormalizeChildPath(rangeRef.ChildPath),
                    rangeRef.StartColumn.ToString(CultureInfo.InvariantCulture),
                    rangeRef.StartRow.ToString(CultureInfo.InvariantCulture),
                    rangeRef.EndColumn.ToString(CultureInfo.InvariantCulture),
                    rangeRef.EndRow.ToString(CultureInfo.InvariantCulture)
                })
            End If
            Return CanonicalObjectIdentity(value)
        End Function

        Friend Shared Function BindActiveWorkbookReference(value As String,
                                                            workbookName As String) As String
            Dim source = If(value, "").Trim()
            Dim workbook = If(workbookName, "").Trim()
            If String.IsNullOrWhiteSpace(source) OrElse String.IsNullOrWhiteSpace(workbook) Then Return source
            Dim encodedWorkbook = Uri.EscapeDataString(workbook)

            If Regex.IsMatch(source, "(?:^|Excel:|/)workbooks/active(?:/|$)", RegexOptions.IgnoreCase) Then
                Return Regex.Replace(
                    source,
                    "(?<=workbooks/)active(?=/|$)",
                    encodedWorkbook,
                    RegexOptions.IgnoreCase)
            End If
            If Regex.IsMatch(source, "(?:^|Excel:|/)workbooks/[^/]+(?:/|$)", RegexOptions.IgnoreCase) Then Return source

            Dim rangeMatch = Regex.Match(
                source,
                "^(?:Excel:)?(?<sheet>[^!/:]+)!(?<address>[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?)$",
                RegexOptions.IgnoreCase)
            If rangeMatch.Success Then
                Dim sheet = rangeMatch.Groups("sheet").Value.Trim("'"c)
                Return $"Excel:workbooks/{encodedWorkbook}/worksheets/{Uri.EscapeDataString(sheet)}/ranges/{rangeMatch.Groups("address").Value}"
            End If

            Dim sheetMatch = Regex.Match(source, "^(?:Excel:)?(?<sheet>[^!/:/]+)$", RegexOptions.IgnoreCase)
            If sheetMatch.Success Then
                Dim sheet = sheetMatch.Groups("sheet").Value.Trim("'"c)
                If Not String.Equals(sheet, "ActiveSheet", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(sheet, "ActiveWorkbook", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(sheet, "Selection", StringComparison.OrdinalIgnoreCase) Then
                    Return $"Excel:workbooks/{encodedWorkbook}/worksheets/{Uri.EscapeDataString(sheet)}"
                End If
            End If
            Return source
        End Function

        Friend Shared Function ResolveContextWorkbookName(contextPack As Context.ContextPack) As String
            If contextPack Is Nothing Then Return ""
            Dim hostWorkbook = contextPack.Host?("workbook")?.ToString()
            If Not String.IsNullOrWhiteSpace(hostWorkbook) Then Return hostWorkbook.Trim()
            Return If(contextPack.Document?.Name, "").Trim()
        End Function

        ''' <summary>
        ''' Canonicalizes an expected contract value with the same equivalence rules used by
        ''' ExpectedMatches. This prevents two success criteria from disguising the same host
        ''' assertion with harmless JSON, casing, whitespace, or numeric representation changes.
        ''' </summary>
        Friend Shared Function CanonicalContractExpectedValue(value As JToken,
                                                               comparisonOperator As String,
                                                               Optional propertyName As String = "") As String
            Dim normalizedOperator = If(comparisonOperator, "equals").Trim().ToLowerInvariant()
            If normalizedOperator = "exists" Then Return "exists"
            Dim recursiveSemanticNormalization =
                normalizedOperator = "contains" OrElse normalizedOperator = "covers"
            Return CanonicalExpectedToken(
                If(value, JValue.CreateNull()),
                recursiveSemanticNormalization,
                True,
                propertyName)
        End Function

        Private Shared Function CanonicalExpectedToken(value As JToken,
                                                       recursiveSemanticNormalization As Boolean,
                                                       isRoot As Boolean,
                                                       Optional propertyName As String = "") As String
            If value Is Nothing OrElse value.Type = JTokenType.Null OrElse
               value.Type = JTokenType.Undefined Then Return PackCanonical("null", "")

            If IsNumericToken(value) AndAlso (recursiveSemanticNormalization OrElse isRoot) Then
                Dim number As Decimal
                If Decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, number) Then
                    Return PackCanonical("number", number.ToString("G29", CultureInfo.InvariantCulture))
                End If
            End If

            If value.Type = JTokenType.String AndAlso (recursiveSemanticNormalization OrElse isRoot) Then
                Dim normalized = NormalizeSemanticString(value.Value(Of String)(), propertyName)
                If IsCaseInsensitiveSemanticProperty(propertyName) Then normalized = normalized.ToLowerInvariant()
                Return PackCanonical("string", normalized)
            End If

            If value.Type = JTokenType.Object Then
                Dim comparer As StringComparer = If(
                    recursiveSemanticNormalization,
                    StringComparer.OrdinalIgnoreCase,
                    StringComparer.Ordinal)
                Dim properties = DirectCast(value, JObject).Properties().
                    OrderBy(Function(prop) prop.Name, comparer).
                    Select(
                        Function(prop)
                            Dim name = If(recursiveSemanticNormalization,
                                          prop.Name.ToLowerInvariant(),
                                          prop.Name)
                            Return PackCanonical("property", name) &
                                CanonicalExpectedToken(prop.Value, recursiveSemanticNormalization, False, prop.Name)
                        End Function)
                Return PackCanonical("object", String.Concat(properties))
            End If

            If value.Type = JTokenType.Array Then
                Dim items = DirectCast(value, JArray).
                    Select(Function(item) CanonicalExpectedToken(item, recursiveSemanticNormalization, False, propertyName))
                Return PackCanonical("array", String.Concat(items))
            End If

            Return PackCanonical(
                value.Type.ToString().ToLowerInvariant(),
                value.ToString(Newtonsoft.Json.Formatting.None))
        End Function

        Private Shared Function PackCanonical(kind As String, payload As String) As String
            Dim content = If(payload, "")
            Return If(kind, "") & content.Length.ToString(CultureInfo.InvariantCulture) & ":" & content
        End Function

    End Class

End Namespace
