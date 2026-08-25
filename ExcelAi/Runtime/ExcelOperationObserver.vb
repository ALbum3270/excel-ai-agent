Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent.OfficeOperations

Namespace OfficeRuntime

    ''' <summary>
    ''' Captures and verifies the same properties declared by an operation batch.
    ''' ExpectedEffects is therefore executable specification, not explanatory text.
    ''' </summary>
    Friend NotInheritable Class ExcelOperationObserver
        Private Sub New()
        End Sub

        Public Shared Function CaptureState(application As Object,
                                            batch As OfficeOperationBatch,
                                            Optional additionalTargetRefs As IEnumerable(Of String) = Nothing) As JObject
            Dim state As New JObject()
            Dim refs = CollectTargetRefs(batch, additionalTargetRefs)
            Dim propertyMap = BuildObservedPropertyMap(batch)
            Dim workbook As Object = Nothing
            Dim worksheets As Object = Nothing
            Try
                workbook = application?.ActiveWorkbook
                If workbook IsNot Nothing Then
                    state("workbook") = CStr(If(workbook.Name, ""))
                    worksheets = workbook.Worksheets
                    state("sheetCount") = CInt(worksheets.Count)
                End If
            Catch ex As Exception
                state("captureError") = AppLogger.Redact(ex.Message)
            Finally
                ReleaseCom(worksheets)
                ReleaseCom(workbook)
            End Try

            Dim targets As New JObject()
            For Each targetRef In refs.OrderBy(Function(item) item, StringComparer.OrdinalIgnoreCase)
                Try
                    Using resolved = ExcelObjectResolver.Resolve(application, targetRef)
                        Dim properties As HashSet(Of String) = Nothing
                        propertyMap.TryGetValue(targetRef, properties)
                        targets(targetRef) = CaptureResolvedValue(resolved.Value,
                                                                  resolved.ObjectKind,
                                                                  If(properties, New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)))
                    End Using
                Catch ex As Exception
                    targets(targetRef) = New JObject From {
                        {"exists", False},
                        {"error", AppLogger.Redact(ex.Message)}
                    }
                End Try
            Next
            state("targets") = targets
            Return state
        End Function

        Public Shared Function BuildObservation(batch As OfficeOperationBatch,
                                                operationResults As JArray,
                                                beforeState As JObject,
                                                afterState As JObject,
                                                targetRefs As IEnumerable(Of String),
                                                warnings As IEnumerable(Of String)) As JObject
            Dim changed = Not JToken.DeepEquals(If(beforeState, New JObject()), If(afterState, New JObject()))
            Dim targetArray = New JArray()
            If targetRefs IsNot Nothing Then
                For Each targetRef In targetRefs.Where(Function(item) Not String.IsNullOrWhiteSpace(item)).Distinct(StringComparer.OrdinalIgnoreCase)
                    targetArray.Add(targetRef)
                Next
            End If

            Dim warningArray = New JArray()
            If warnings IsNot Nothing Then
                For Each warning In warnings.Where(Function(item) Not String.IsNullOrWhiteSpace(item))
                    warningArray.Add(warning)
                Next
            End If

            Dim succeededCount = If(operationResults, New JArray()).OfType(Of JObject)().Count(
                Function(item) String.Equals(item("status")?.ToString(), "succeeded", StringComparison.OrdinalIgnoreCase))
            Dim totalCount = If(operationResults?.Count, 0)
            Dim effectType = InferMutationEffect(batch)
            Return New JObject From {
                {"kind", "office_operation_batch"},
                {"summary", $"Excel 声明式操作完成 {succeededCount}/{totalCount} 项"},
                {"changed", changed},
                {"writeExpected", HasMutatingOperations(batch)},
                {"appType", "Excel"},
                {"effectType", effectType},
                {"invalidationRefs", BuildStructuralInvalidationRefs(batch)},
                {"atomic", If(batch?.Atomic, True)},
                {"targetRefs", targetArray},
                {"operations", If(operationResults, New JArray())},
                {"before", If(beforeState, New JObject())},
                {"after", If(afterState, New JObject())},
                {"diff", BuildDiff(beforeState, afterState)},
                {"warnings", warningArray}
            }
        End Function

        Public Shared Function VerifyExpectedEffects(batch As OfficeOperationBatch,
                                                     operationResults As JArray,
                                                     afterState As JObject) As JArray
            Dim verification As New JArray()
            If batch Is Nothing OrElse afterState Is Nothing Then Return verification
            Dim targets = TryCast(afterState("targets"), JObject)
            If targets Is Nothing Then Return verification

            For Each operation In If(batch.Operations, New List(Of OfficeOperation)())
                If operation Is Nothing OrElse operation.ExpectedEffects Is Nothing OrElse Not operation.ExpectedEffects.HasValues Then Continue For
                Dim operationResult = operationResults?.OfType(Of JObject)().FirstOrDefault(
                    Function(item) String.Equals(item("id")?.ToString(), operation.Id, StringComparison.OrdinalIgnoreCase))
                Dim resultRef = operationResult?("resultRef")?.ToString()
                Dim verifyRef = If(String.IsNullOrWhiteSpace(resultRef), operation.TargetRef, resultRef)
                Dim snapshot = TryCast(targets(verifyRef), JObject)
                For Each expectedProperty In operation.ExpectedEffects.Properties()
                    Dim actual = If(snapshot Is Nothing, Nothing, snapshot(expectedProperty.Name))
                    Dim passed = TokensEquivalent(expectedProperty.Name, actual, expectedProperty.Value)
                    verification.Add(New JObject From {
                        {"id", $"{operation.Id}:{expectedProperty.Name}"},
                        {"required", True},
                        {"targetRef", verifyRef},
                        {"effectType", InferOperationEffect(operation, If(operation.Action, "").Trim().ToLowerInvariant())},
                        {"property", expectedProperty.Name},
                        {"status", If(passed, "passed", "failed")},
                        {"expected", expectedProperty.Value.DeepClone()},
                        {"actual", If(actual?.DeepClone(), JValue.CreateNull())}
                    })
                Next
            Next

            For index = 0 To If(batch.SuccessCriteria?.Count, 0) - 1
                Dim criterion = batch.SuccessCriteria(index)
                If criterion Is Nothing OrElse String.IsNullOrWhiteSpace(criterion.TargetRef) Then Continue For
                Dim snapshot = TryCast(targets(criterion.TargetRef), JObject)
                Dim actual = If(snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(criterion.PropertyName),
                                Nothing,
                                snapshot(criterion.PropertyName))
                Dim passed = EvaluateCriterion(actual, criterion)
                Dim criterionEffect = InferCriterionEffect(batch, operationResults, criterion)
                verification.Add(New JObject From {
                    {"id", If(criterion.Id, $"criterion-{index + 1}")},
                    {"required", criterion.Required},
                    {"targetRef", criterion.TargetRef},
                    {"effectType", criterionEffect},
                    {"property", If(criterion.PropertyName, "")},
                    {"operator", If(criterion.Operator, "equals")},
                    {"status", If(passed, "passed", "failed")},
                    {"expected", If(criterion.ExpectedValue?.DeepClone(), JValue.CreateNull())},
                    {"actual", If(actual?.DeepClone(), JValue.CreateNull())}
                })
            Next
            Return verification
        End Function

        Private Shared Function InferCriterionEffect(batch As OfficeOperationBatch,
                                                      operationResults As JArray,
                                                      criterion As OperationCriterion) As String
            If criterion Is Nothing Then Return "property_state"
            If String.Equals(criterion.PropertyName, "exists", StringComparison.OrdinalIgnoreCase) Then
                Dim expectedExists As Boolean
                If criterion.ExpectedValue IsNot Nothing AndAlso
                   Boolean.TryParse(criterion.ExpectedValue.ToString(), expectedExists) Then
                    Return If(expectedExists, "object_exists", "object_absent")
                End If
            End If

            Dim effects As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each operation In If(batch?.Operations, New List(Of OfficeOperation)())
                If operation Is Nothing Then Continue For
                Dim operationResult = operationResults?.OfType(Of JObject)().FirstOrDefault(
                    Function(item) String.Equals(item("id")?.ToString(), operation.Id, StringComparison.OrdinalIgnoreCase))
                Dim resultRef = If(operationResult?("resultRef")?.ToString(), "")
                If Not String.Equals(operation.TargetRef, criterion.TargetRef, StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(resultRef, criterion.TargetRef, StringComparison.OrdinalIgnoreCase) Then Continue For
                effects.Add(InferOperationEffect(operation, If(operation.Action, "").Trim().ToLowerInvariant()))
            Next
            If effects.Count = 1 Then Return effects.First()
            Return "property_state"
        End Function

        Public Shared Function HasRequiredVerificationFailure(verification As JArray) As Boolean
            Return verification IsNot Nothing AndAlso verification.OfType(Of JObject)().Any(
                Function(item) If(item("required")?.Value(Of Boolean)(), True) AndAlso
                    String.Equals(item("status")?.ToString(), "failed", StringComparison.OrdinalIgnoreCase))
        End Function

        Public Shared Function HasMutatingOperations(batch As OfficeOperationBatch) As Boolean
            Return batch?.Operations IsNot Nothing AndAlso batch.Operations.Any(
                Function(operation) operation IsNot Nothing AndAlso
                    Not String.Equals(operation.Action, "get", StringComparison.OrdinalIgnoreCase) AndAlso
                    Not String.Equals(operation.Action, "collection_item", StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' Maps a declarative COM batch to the canonical state family it can prove. Mixed or
        ''' unknown batches are invalidation-only; callers must split them to obtain positive
        ''' completion evidence for a specific outcome.
        ''' </summary>
        Private Shared Function InferMutationEffect(batch As OfficeOperationBatch) As String
            Dim effects As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each operation In If(batch?.Operations, New List(Of OfficeOperation)())
                If operation Is Nothing Then Continue For
                Dim action = If(operation.Action, "").Trim().ToLowerInvariant()
                If action = "get" OrElse action = "collection_item" Then Continue For
                effects.Add(InferOperationEffect(operation, action))
            Next
            If effects.Count = 0 Then Return "read_coverage"
            If effects.Count = 1 Then Return effects.First()
            Return "unclassified_mutation"
        End Function

        Private Shared Function InferOperationEffect(operation As OfficeOperation,
                                                     action As String) As String
            Dim memberId = If(operation?.MemberId, "").ToLowerInvariant()
            If action = "set" Then
                If Regex.IsMatch(memberId, "\.property\.(formula|formular1c1)(?:\(|$)", RegexOptions.IgnoreCase) Then
                    Return "formula_state"
                End If
                If Regex.IsMatch(memberId, "\.property\.(value|value2)(?:\(|$)", RegexOptions.IgnoreCase) Then
                    Return "data_state"
                End If
                Return "property_state"
            End If

            If action = "create" Then
                If memberId.Contains("chart") Then Return "artifact"
                Return "object_exists"
            End If

            If action = "delete" Then
                If memberId.Contains("excel.range.") OrElse memberId.Contains("excel.rows.") OrElse
                   memberId.Contains("excel.columns.") Then Return "data_state"
                Return "object_absent"
            End If

            If action = "invoke" Then
                If memberId.Contains(".method.sort") Then Return "order_state"
                If memberId.Contains(".method.autofilter") OrElse memberId.Contains(".method.filter") Then Return "filter_state"
                If memberId.Contains(".method.clear") OrElse memberId.Contains(".method.clearcontents") OrElse
                   (memberId.Contains("excel.range.") AndAlso memberId.Contains(".method.delete")) Then Return "data_state"
                If memberId.Contains(".method.delete") Then Return "object_absent"
                If memberId.Contains("chartobjects.method.add") OrElse memberId.Contains("charts.method.add") Then Return "artifact"
                If memberId.Contains("worksheets.method.add") OrElse memberId.Contains("sheets.method.add") OrElse
                   memberId.Contains("worksheet.method.copy") Then Return "object_exists"
                If memberId.Contains("excel.range.") AndAlso memberId.Contains(".method.copy") Then Return "data_state"
            End If

            Return "unclassified_mutation"
        End Function

        ''' <summary>
        ''' Insert/Delete changes the coordinate identity of every downstream range. A
        ''' tombstone at the worksheet root is deliberately conservative and prevents old
        ''' cell/read evidence from being re-used after rows or columns shift.
        ''' </summary>
        Private Shared Function BuildStructuralInvalidationRefs(batch As OfficeOperationBatch) As JArray
            Dim result As New JArray()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each operation In If(batch?.Operations, New List(Of OfficeOperation)())
                If operation Is Nothing Then Continue For
                Dim memberId = If(operation.MemberId, "")
                Dim structural = Regex.IsMatch(
                    memberId,
                    "Excel\.(?:Range|Rows|Columns)\.method\.(?:Insert|Delete)(?:\(\)|$)",
                    RegexOptions.IgnoreCase)
                If Not structural Then Continue For
                Dim worksheetRef = GetWorksheetRootReference(operation.TargetRef)
                If Not String.IsNullOrWhiteSpace(worksheetRef) AndAlso seen.Add(worksheetRef) Then
                    result.Add(worksheetRef)
                End If
            Next
            Return result
        End Function

        Private Shared Function GetWorksheetRootReference(targetRef As String) As String
            Dim source = If(targetRef, "").Trim()
            Dim canonicalMatch = Regex.Match(
                source,
                "^(?<root>Excel:workbooks/[^/]+/worksheets/[^/]+)(?:/|$)",
                RegexOptions.IgnoreCase)
            If canonicalMatch.Success Then Return canonicalMatch.Groups("root").Value

            Dim shortMatch = Regex.Match(source, "^(?:Excel:)?(?<sheet>[^!/:]+)!", RegexOptions.IgnoreCase)
            If shortMatch.Success Then
                Return $"Excel:workbooks/active/worksheets/{Uri.EscapeDataString(shortMatch.Groups("sheet").Value.Trim("'"c))}"
            End If
            Return ""
        End Function

        Private Shared Function CollectTargetRefs(batch As OfficeOperationBatch,
                                                  additionalTargetRefs As IEnumerable(Of String)) As HashSet(Of String)
            Dim refs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If batch?.Operations IsNot Nothing Then
                For Each operation In batch.Operations
                    If operation IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(operation.TargetRef) Then refs.Add(operation.TargetRef)
                Next
            End If
            If batch?.SuccessCriteria IsNot Nothing Then
                For Each criterion In batch.SuccessCriteria
                    If criterion IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(criterion.TargetRef) Then refs.Add(criterion.TargetRef)
                Next
            End If
            If additionalTargetRefs IsNot Nothing Then
                For Each targetRef In additionalTargetRefs
                    If Not String.IsNullOrWhiteSpace(targetRef) Then refs.Add(targetRef)
                Next
            End If
            Return refs
        End Function

        Private Shared Function BuildObservedPropertyMap(batch As OfficeOperationBatch) As Dictionary(Of String, HashSet(Of String))
            Dim result As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            If batch?.Operations IsNot Nothing Then
                For Each operation In batch.Operations
                    If operation Is Nothing OrElse String.IsNullOrWhiteSpace(operation.TargetRef) Then Continue For
                    Dim properties As HashSet(Of String) = Nothing
                    If Not result.TryGetValue(operation.TargetRef, properties) Then
                        properties = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                        result(operation.TargetRef) = properties
                    End If
                    If operation.ExpectedEffects IsNot Nothing Then
                        For Each expectedProperty In operation.ExpectedEffects.Properties()
                            properties.Add(expectedProperty.Name)
                        Next
                    End If
                    Dim capability As OfficeCapabilityMember = Nothing
                    Dim member As MemberInfo = Nothing
                    If ExcelApiCatalogProvider.TryGetMemberBinding(operation.MemberId, capability, member) AndAlso
                       TypeOf member Is PropertyInfo Then properties.Add(member.Name)
                Next
            End If
            If batch?.SuccessCriteria IsNot Nothing Then
                For Each criterion In batch.SuccessCriteria
                    If criterion Is Nothing OrElse String.IsNullOrWhiteSpace(criterion.TargetRef) OrElse String.IsNullOrWhiteSpace(criterion.PropertyName) Then Continue For
                    Dim properties As HashSet(Of String) = Nothing
                    If Not result.TryGetValue(criterion.TargetRef, properties) Then
                        properties = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                        result(criterion.TargetRef) = properties
                    End If
                    properties.Add(criterion.PropertyName)
                Next
            End If
            Return result
        End Function

        Private Shared Function CaptureResolvedValue(value As Object,
                                                     objectKind As String,
                                                     properties As HashSet(Of String)) As JObject
            Dim snapshot As New JObject From {
                {"exists", value IsNot Nothing},
                {"objectKind", If(objectKind, "")}
            }
            If value Is Nothing Then Return snapshot

            Try
                Select Case If(objectKind, "").ToLowerInvariant()
                    Case "workbook"
                        snapshot("Name") = CStr(value.Name)
                        Dim worksheets As Object = Nothing
                        Try
                            worksheets = value.Worksheets
                            snapshot("WorksheetCount") = CInt(worksheets.Count)
                        Finally
                            ReleaseCom(worksheets)
                        End Try
                    Case "worksheets", "chartobjects", "listobjects", "pivottables", "borders", "seriescollection"
                        snapshot("Count") = CInt(value.Count)
                    Case "worksheet"
                        snapshot("Name") = CStr(value.Name)
                        CaptureWorksheetShape(value, snapshot)
                    Case "range"
                        snapshot("Address") = CStr(value.Address(False, False))
                        CaptureRangeShape(value, snapshot)
                    Case "chartobject", "listobject", "pivottable"
                        Try
                            snapshot("Name") = CStr(value.Name)
                        Catch
                        End Try
                    Case "chart"
                        CaptureChartShape(value, snapshot)
                    Case "charttitle"
                        snapshot("Text") = CStr(value.Text)
                    Case "legend"
                        snapshot("Position") = CInt(value.Position)
                    Case "series"
                        snapshot("Name") = CStr(value.Name)
                        ' Excel may expose the same category sequence as a one-dimensional
                        ' SAFEARRAY while a Range snapshot is represented as a two-dimensional
                        ' matrix. Verify the semantic scalar sequence, not the COM array shape.
                        snapshot("XValuesHash") = ComputeSequenceHash(value.XValues)
                End Select
            Catch ex As Exception
                snapshot("identityError") = AppLogger.Redact(ex.Message)
            End Try

            For Each propertyName In properties
                If snapshot(propertyName) IsNot Nothing Then Continue For
                Try
                    Dim propertyValue = Interaction.CallByName(value, propertyName, CallType.Get)
                    snapshot(propertyName) = ConvertObservedValue(propertyValue)
                    ReleaseComIfNeeded(propertyValue)
                Catch ex As Exception
                    snapshot(propertyName & "Error") = AppLogger.Redact(ex.Message)
                End Try
            Next
            Return snapshot
        End Function

        Private Shared Sub CaptureWorksheetShape(worksheet As Object, snapshot As JObject)
            Dim usedRange As Object = Nothing
            Dim rows As Object = Nothing
            Dim columns As Object = Nothing
            Dim chartObjects As Object = Nothing
            Try
                usedRange = worksheet.UsedRange
                rows = usedRange.Rows
                columns = usedRange.Columns
                snapshot("UsedRows") = CInt(rows.Count)
                snapshot("UsedColumns") = CInt(columns.Count)
                snapshot("UsedValueHash") = ComputeValueHash(usedRange.Value2)
                snapshot("UsedFormulaHash") = ComputeValueHash(usedRange.Formula)
                snapshot("AutoFilterMode") = CBool(worksheet.AutoFilterMode)
                snapshot("FilterMode") = CBool(worksheet.FilterMode)
                snapshot("ProtectContents") = CBool(worksheet.ProtectContents)
                chartObjects = worksheet.ChartObjects()
                snapshot("ChartCount") = CInt(chartObjects.Count)
            Finally
                ReleaseCom(chartObjects)
                ReleaseCom(columns)
                ReleaseCom(rows)
                ReleaseCom(usedRange)
            End Try
        End Sub

        Private Shared Sub CaptureRangeShape(target As Object, snapshot As JObject)
            Dim rows As Object = Nothing
            Dim columns As Object = Nothing
            Dim cells As Object = Nothing
            Dim topLeft As Object = Nothing
            Try
                rows = target.Rows
                columns = target.Columns
                cells = target.Cells
                topLeft = cells.Item(1, 1)
                snapshot("RowCount") = CInt(rows.Count)
                snapshot("ColumnCount") = CInt(columns.Count)
                snapshot("ValueHash") = ComputeValueHash(target.Value2)
                snapshot("FormulaHash") = ComputeValueHash(target.Formula)
                snapshot("RowSetHash") = ComputeRowSetHash(target.Value2)
                snapshot("TopLeftValue") = ConvertObservedValue(topLeft.Value2)
                snapshot("TopLeftFormula") = ConvertObservedValue(topLeft.Formula)
                snapshot("NonEmptyFormulaCount") = CountNonEmptyFormulas(target)
                CaptureOptionalProperty(target, snapshot, "MergeCells")
                CaptureOptionalProperty(target, snapshot, "Hidden")
                CaptureOptionalProperty(target, snapshot, "ColumnWidth")
                CaptureOptionalProperty(target, snapshot, "RowHeight")
            Finally
                ReleaseCom(topLeft)
                ReleaseCom(cells)
                ReleaseCom(columns)
                ReleaseCom(rows)
            End Try
        End Sub

        Private Shared Sub CaptureOptionalProperty(target As Object, snapshot As JObject, propertyName As String)
            Try
                Dim propertyValue = Interaction.CallByName(target, propertyName, CallType.Get)
                snapshot(propertyName) = ConvertObservedValue(propertyValue)
                ReleaseComIfNeeded(propertyValue)
            Catch
            End Try
        End Sub

        Private Shared Sub CaptureChartShape(chart As Object, snapshot As JObject)
            Dim seriesCollection As Object = Nothing
            Try
                snapshot("ChartType") = CInt(chart.ChartType)
                snapshot("HasTitle") = CBool(chart.HasTitle)
                snapshot("HasLegend") = CBool(chart.HasLegend)
                seriesCollection = chart.SeriesCollection()
                snapshot("SeriesCount") = CInt(seriesCollection.Count)
            Finally
                ReleaseCom(seriesCollection)
            End Try
        End Sub

        Private Shared Function CountNonEmptyFormulas(target As Object) As Integer
            Dim cells As Object = Nothing
            Try
                cells = target.Cells
                Dim count = Math.Min(CInt(cells.Count), 10000)
                Dim populated As Integer = 0
                For index = 1 To count
                    Dim cell As Object = Nothing
                    Try
                        cell = cells.Item(index)
                        Dim formula = cell.Formula
                        If formula IsNot Nothing AndAlso formula.ToString().StartsWith("=", StringComparison.Ordinal) Then populated += 1
                    Finally
                        ReleaseCom(cell)
                    End Try
                Next
                Return populated
            Finally
                ReleaseCom(cells)
            End Try
        End Function

        Private Shared Function ConvertObservedValue(value As Object) As JToken
            If value Is Nothing Then Return JValue.CreateNull()
            If Marshal.IsComObject(value) Then Return New JValue(value.GetType().FullName)
            If TypeOf value Is Array Then Return New JValue(ComputeHash(SerializeValue(value)))
            Try
                Return JToken.FromObject(value)
            Catch
                Return New JValue(value.ToString())
            End Try
        End Function

        Private Shared Function EvaluateCriterion(actual As JToken, criterion As OperationCriterion) As Boolean
            Dim op = If(criterion?.Operator, "equals").Trim().ToLowerInvariant()
            Select Case op
                Case "exists"
                    Return actual IsNot Nothing AndAlso actual.Type <> JTokenType.Null
                Case "not_equals"
                    Return Not TokensEquivalent(criterion.PropertyName, actual, criterion.ExpectedValue)
                Case "contains"
                    Return actual IsNot Nothing AndAlso criterion.ExpectedValue IsNot Nothing AndAlso
                           actual.ToString().IndexOf(criterion.ExpectedValue.ToString(), StringComparison.OrdinalIgnoreCase) >= 0
                Case "gte", "lte"
                    Dim actualNumber As Decimal
                    Dim expectedNumber As Decimal
                    If Not Decimal.TryParse(If(actual?.ToString(), ""), NumberStyles.Any, CultureInfo.InvariantCulture, actualNumber) OrElse
                       Not Decimal.TryParse(If(criterion.ExpectedValue?.ToString(), ""), NumberStyles.Any, CultureInfo.InvariantCulture, expectedNumber) Then Return False
                    Return If(op = "gte", actualNumber >= expectedNumber, actualNumber <= expectedNumber)
                Case Else
                    Return TokensEquivalent(criterion.PropertyName, actual, criterion.ExpectedValue)
            End Select
        End Function

        Private Shared Function TokensEquivalent(propertyName As String, actual As JToken, expected As JToken) As Boolean
            If actual Is Nothing OrElse expected Is Nothing Then Return actual Is Nothing AndAlso expected Is Nothing
            If actual.Type = JTokenType.Null OrElse expected.Type = JTokenType.Null Then Return actual.Type = expected.Type

            If String.Equals(propertyName, "NumberFormat", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(propertyName, "NumberFormatLocal", StringComparison.OrdinalIgnoreCase) Then
                Return String.Equals(NormalizeNumberFormat(actual.ToString()),
                                     NormalizeNumberFormat(expected.ToString()),
                                     StringComparison.OrdinalIgnoreCase)
            End If

            Dim actualNumber As Decimal
            Dim expectedNumber As Decimal
            If Decimal.TryParse(actual.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, actualNumber) AndAlso
               Decimal.TryParse(expected.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, expectedNumber) Then
                Return actualNumber = expectedNumber
            End If
            Return JToken.DeepEquals(actual, expected) OrElse
                   String.Equals(actual.ToString(), expected.ToString(), StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function NormalizeNumberFormat(value As String) As String
            Return If(value, "").Trim().Replace(" ", "")
        End Function

        Private Shared Function BuildDiff(beforeState As JObject, afterState As JObject) As JObject
            Return New JObject From {
                {"sheetCountDelta", GetInteger(afterState, "sheetCount") - GetInteger(beforeState, "sheetCount")},
                {"targetStateChanged", Not JToken.DeepEquals(beforeState?("targets"), afterState?("targets"))}
            }
        End Function

        Private Shared Function GetInteger(source As JObject, name As String) As Integer
            Return If(source?(name)?.Value(Of Integer)(), 0)
        End Function

        Private Shared Function SerializeValue(value As Object) As String
            If value Is Nothing Then Return ""
            Try
                Return JsonConvert.SerializeObject(value, Formatting.None)
            Catch
                Return value.ToString()
            End Try
        End Function

        Friend Shared Function ComputeValueHash(value As Object) As String
            Dim token As JToken = Nothing
            Try
                token = If(value Is Nothing, JValue.CreateNull(), JToken.FromObject(value))
            Catch
                token = New JValue(If(value?.ToString(), ""))
            End Try
            Return ComputeHash(CanonicalValueString(token))
        End Function

        Friend Shared Function ComputeRowSetHash(value As Object) As String
            Dim token As JToken = Nothing
            Try
                token = If(value Is Nothing, JValue.CreateNull(), JToken.FromObject(value))
            Catch
                token = New JValue(If(value?.ToString(), ""))
            End Try
            Dim rows = TryCast(token, JArray)
            If rows Is Nothing OrElse rows.Count = 0 OrElse rows.Any(Function(item) item.Type <> JTokenType.Array) Then
                Return ComputeValueHash(value)
            End If
            Dim normalizedRows = rows.Select(AddressOf CanonicalValueString).
                                      OrderBy(Function(item) item, StringComparer.Ordinal).
                                      ToArray()
            Return ComputeHash(String.Join(vbLf, normalizedRows))
        End Function

        Friend Shared Function ComputeSequenceHash(value As Object) As String
            Dim flattened As New List(Of String)()
            Dim token As JToken = Nothing
            Try
                token = If(value Is Nothing, JValue.CreateNull(), JToken.FromObject(value))
            Catch
                token = New JValue(If(value?.ToString(), ""))
            End Try
            Dim scalarValues As New List(Of JValue)()
            CollectScalarValues(token, scalarValues)
            For Each scalar In scalarValues
                flattened.Add(CanonicalValueString(scalar))
            Next
            Return ComputeHash(String.Join(vbLf, flattened))
        End Function

        Private Shared Sub CollectScalarValues(token As JToken, values As List(Of JValue))
            If token Is Nothing Then Return
            Dim scalar = TryCast(token, JValue)
            If scalar IsNot Nothing Then
                values.Add(scalar)
                Return
            End If
            For Each child In token.Children()
                CollectScalarValues(child, values)
            Next
        End Sub

        Private Shared Function CanonicalValueString(token As JToken) As String
            If token Is Nothing OrElse token.Type = JTokenType.Null OrElse token.Type = JTokenType.Undefined Then Return "null"
            If token.Type = JTokenType.Array Then
                Return "[" & String.Join(",", DirectCast(token, JArray).Select(AddressOf CanonicalValueString)) & "]"
            End If
            If token.Type = JTokenType.Object Then
                Return "{" & String.Join(",", DirectCast(token, JObject).Properties().
                    OrderBy(Function(item) item.Name, StringComparer.Ordinal).
                    Select(Function(item) JsonConvert.ToString(item.Name) & ":" & CanonicalValueString(item.Value))) & "}"
            End If
            If token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float Then
                Dim number As Decimal
                If Decimal.TryParse(token.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, number) Then
                    Return "n:" & number.ToString("G29", CultureInfo.InvariantCulture)
                End If
                Return "n:" & token.ToString()
            End If
            If token.Type = JTokenType.Boolean Then Return If(token.Value(Of Boolean)(), "true", "false")
            Return "s:" & JsonConvert.ToString(token.ToString())
        End Function

        Private Shared Function ComputeHash(value As String) As String
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(If(value, "")))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Sub ReleaseComIfNeeded(value As Object)
            If value IsNot Nothing AndAlso Marshal.IsComObject(value) Then ReleaseCom(value)
        End Sub

        Private Shared Sub ReleaseCom(value As Object)
            If value Is Nothing Then Return
            Try
                If Marshal.IsComObject(value) Then Marshal.ReleaseComObject(value)
            Catch
            End Try
        End Sub
    End Class

End Namespace
