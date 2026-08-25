Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend Class ExcelOperationException
        Inherits Exception

        Public ReadOnly Property ErrorCode As String
        Public ReadOnly Property Recoverable As Boolean

        Public Sub New(errorCode As String, message As String, Optional recoverable As Boolean = True)
            MyBase.New(message)
            Me.ErrorCode = errorCode
            Me.Recoverable = recoverable
        End Sub
    End Class

    Friend NotInheritable Class ResolvedExcelObject
        Implements IDisposable

        Public Property Value As Object
        Public Property CanonicalRef As String
        Public Property ObjectKind As String

        Private ReadOnly _ownedObjects As New List(Of Object)()
        Private _disposed As Boolean

        Public Sub Track(value As Object)
            If value Is Nothing OrElse Not Marshal.IsComObject(value) Then Return
            If _ownedObjects.Any(Function(item) Object.ReferenceEquals(item, value)) Then Return
            _ownedObjects.Add(value)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            For index = _ownedObjects.Count - 1 To 0 Step -1
                Try
                    If Marshal.IsComObject(_ownedObjects(index)) Then Marshal.ReleaseComObject(_ownedObjects(index))
                Catch
                End Try
            Next
            _ownedObjects.Clear()
            Value = Nothing
        End Sub
    End Class

    ''' <summary>
    ''' Resolves canonical Excel refs into short-lived COM objects. No COM object is
    ''' stored in an Agent plan, observation, memory item, or operation result.
    ''' </summary>
    Friend NotInheritable Class ExcelObjectResolver
        Private Sub New()
        End Sub

        Public Shared Function Resolve(application As Object, targetRef As String) As ResolvedExcelObject
            Dim parsed As OfficeObjectRef = Nothing
            Dim parseError As String = ""
            If Not OfficeObjectRef.TryParse(targetRef, parsed, parseError) Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeObjectRefInvalid, parseError)
            End If
            If Not String.Equals(parsed.AppType, "Excel", StringComparison.OrdinalIgnoreCase) Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeHostUnsupported,
                                                  $"Target ref belongs to {parsed.AppType}",
                                                  recoverable:=False)
            End If
            If application Is Nothing Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeDocMissing,
                                                  "Excel application is unavailable",
                                                  recoverable:=False)
            End If

            Dim workbook As Object = Nothing
            Try
                workbook = application.ActiveWorkbook
            Catch
            End Try
            If workbook Is Nothing Then
                Throw New ExcelOperationException(ExceptionClassifier.CodeDocMissing,
                                                  "No active Excel workbook",
                                                  recoverable:=False)
            End If

            Dim workbookName = CStr(If(workbook.Name, ""))
            Dim requestedWorkbook = DecodeSegment(parsed.DocumentRef)
            If Not String.Equals(requestedWorkbook, "active", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(requestedWorkbook, workbookName, StringComparison.OrdinalIgnoreCase) Then
                ReleaseCom(workbook)
                Throw New ExcelOperationException(ExceptionClassifier.CodeObjectNotFound,
                                                  $"Active workbook does not match '{requestedWorkbook}'")
            End If

            Dim resolved As New ResolvedExcelObject With {
                .Value = workbook,
                .CanonicalRef = parsed.ToCanonicalString(),
                .ObjectKind = "Workbook"
            }
            resolved.Track(workbook)
            Try
                ResolvePath(parsed.Path, workbook, resolved)
                Return resolved
            Catch
                resolved.Dispose()
                Throw
            End Try
        End Function

        Public Shared Function BuildRangeRef(sheetName As String, address As String) As String
            Return "Excel:workbooks/active/worksheets/" & EncodeSegment(sheetName) &
                   "/ranges/" & EncodeSegment(address)
        End Function

        Public Shared Function EncodeSegment(value As String) As String
            Return Uri.EscapeDataString(If(value, ""))
        End Function

        Private Shared Function DecodeSegment(value As String) As String
            Return Uri.UnescapeDataString(If(value, ""))
        End Function

        Private Shared Sub ResolvePath(path As String, workbook As Object, resolved As ResolvedExcelObject)
            If String.IsNullOrWhiteSpace(path) Then Return
            Dim segments = path.Split("/"c)
            Dim current As Object = workbook
            Dim index As Integer = 0

            While index < segments.Length
                Dim segment = DecodeSegment(segments(index)).Trim().ToLowerInvariant()
                Select Case segment
                    Case "worksheets", "sheets"
                        EnsureKind(current, "Workbook", resolved.ObjectKind)
                        Dim worksheets As Object = current.Worksheets
                        resolved.Track(worksheets)
                        current = worksheets
                        resolved.ObjectKind = "Worksheets"
                        If index + 1 < segments.Length Then
                            index += 1
                            Dim key = DecodeSegment(segments(index))
                            Dim worksheet As Object = ResolveCollectionItem(worksheets, key, "worksheet")
                            resolved.Track(worksheet)
                            current = worksheet
                            resolved.ObjectKind = "Worksheet"
                        End If

                    Case "ranges", "range"
                        EnsureKind(current, "Worksheet", resolved.ObjectKind)
                        If index + 1 >= segments.Length Then ThrowInvalid("Range address is missing")
                        index += 1
                        Dim address = DecodeSegment(segments(index))
                        Dim target As Object = current.Range(address)
                        If target Is Nothing Then ThrowNotFound("range", address)
                        resolved.Track(target)
                        current = target
                        resolved.ObjectKind = "Range"

                    Case "rows"
                        If resolved.ObjectKind <> "Range" AndAlso resolved.ObjectKind <> "Worksheet" Then ThrowTypeMismatch("Range or Worksheet", resolved.ObjectKind)
                        Dim rows As Object = current.Rows
                        resolved.Track(rows)
                        current = rows
                        resolved.ObjectKind = "Range"
                        If HasNextIndex(segments, index + 1) Then
                            index += 1
                            Dim row = rows.Item(ParseIndex(DecodeSegment(segments(index)), "row"))
                            resolved.Track(row)
                            current = row
                        End If

                    Case "columns"
                        If resolved.ObjectKind <> "Range" AndAlso resolved.ObjectKind <> "Worksheet" Then ThrowTypeMismatch("Range or Worksheet", resolved.ObjectKind)
                        Dim columns As Object = current.Columns
                        resolved.Track(columns)
                        current = columns
                        resolved.ObjectKind = "Range"
                        If HasNextIndex(segments, index + 1) Then
                            index += 1
                            Dim column = columns.Item(ParseIndex(DecodeSegment(segments(index)), "column"))
                            resolved.Track(column)
                            current = column
                        End If

                    Case "cells"
                        If resolved.ObjectKind <> "Range" AndAlso resolved.ObjectKind <> "Worksheet" Then ThrowTypeMismatch("Range or Worksheet", resolved.ObjectKind)
                        Dim cells As Object = current.Cells
                        resolved.Track(cells)
                        current = cells
                        resolved.ObjectKind = "Range"

                    Case "entirerow"
                        EnsureKind(current, "Range", resolved.ObjectKind)
                        Dim entireRow As Object = current.EntireRow
                        resolved.Track(entireRow)
                        current = entireRow
                        resolved.ObjectKind = "Range"

                    Case "entirecolumn"
                        EnsureKind(current, "Range", resolved.ObjectKind)
                        Dim entireColumn As Object = current.EntireColumn
                        resolved.Track(entireColumn)
                        current = entireColumn
                        resolved.ObjectKind = "Range"

                    Case "font"
                        EnsureKind(current, "Range", resolved.ObjectKind)
                        Dim font As Object = current.Font
                        resolved.Track(font)
                        current = font
                        resolved.ObjectKind = "Font"

                    Case "interior"
                        EnsureKind(current, "Range", resolved.ObjectKind)
                        Dim interior As Object = current.Interior
                        resolved.Track(interior)
                        current = interior
                        resolved.ObjectKind = "Interior"

                    Case "borders"
                        EnsureKind(current, "Range", resolved.ObjectKind)
                        Dim borders As Object = current.Borders
                        resolved.Track(borders)
                        current = borders
                        resolved.ObjectKind = "Borders"
                        If HasNextIndex(segments, index + 1) Then
                            index += 1
                            Dim border = borders.Item(ParseIndex(DecodeSegment(segments(index)), "border"))
                            resolved.Track(border)
                            current = border
                            resolved.ObjectKind = "Border"
                        End If

                    Case "chartobjects"
                        EnsureKind(current, "Worksheet", resolved.ObjectKind)
                        Dim chartObjects As Object = current.ChartObjects()
                        resolved.Track(chartObjects)
                        current = chartObjects
                        resolved.ObjectKind = "ChartObjects"
                        If index + 1 < segments.Length Then
                            index += 1
                            Dim chartObject = ResolveCollectionItem(chartObjects, DecodeSegment(segments(index)), "chart object")
                            resolved.Track(chartObject)
                            current = chartObject
                            resolved.ObjectKind = "ChartObject"
                        End If

                    Case "chart"
                        EnsureKind(current, "ChartObject", resolved.ObjectKind)
                        Dim chart As Object = current.Chart
                        resolved.Track(chart)
                        current = chart
                        resolved.ObjectKind = "Chart"

                    Case "charttitle"
                        EnsureKind(current, "Chart", resolved.ObjectKind)
                        Dim chartTitle As Object = current.ChartTitle
                        resolved.Track(chartTitle)
                        current = chartTitle
                        resolved.ObjectKind = "ChartTitle"

                    Case "legend"
                        EnsureKind(current, "Chart", resolved.ObjectKind)
                        Dim legend As Object = current.Legend
                        resolved.Track(legend)
                        current = legend
                        resolved.ObjectKind = "Legend"

                    Case "series", "seriescollection"
                        EnsureKind(current, "Chart", resolved.ObjectKind)
                        Dim seriesCollection As Object = current.SeriesCollection()
                        resolved.Track(seriesCollection)
                        current = seriesCollection
                        resolved.ObjectKind = "SeriesCollection"
                        If index + 1 < segments.Length Then
                            index += 1
                            Dim series = seriesCollection.Item(ParseIndex(DecodeSegment(segments(index)), "series"))
                            resolved.Track(series)
                            current = series
                            resolved.ObjectKind = "Series"
                        End If

                    Case "listobjects"
                        EnsureKind(current, "Worksheet", resolved.ObjectKind)
                        Dim listObjects As Object = current.ListObjects
                        resolved.Track(listObjects)
                        current = listObjects
                        resolved.ObjectKind = "ListObjects"
                        If index + 1 < segments.Length Then
                            index += 1
                            Dim listObject = ResolveCollectionItem(listObjects, DecodeSegment(segments(index)), "list object")
                            resolved.Track(listObject)
                            current = listObject
                            resolved.ObjectKind = "ListObject"
                        End If

                    Case "pivottables"
                        EnsureKind(current, "Worksheet", resolved.ObjectKind)
                        Dim pivotTables As Object = current.PivotTables()
                        resolved.Track(pivotTables)
                        current = pivotTables
                        resolved.ObjectKind = "PivotTables"
                        If index + 1 < segments.Length Then
                            index += 1
                            Dim pivotTable = ResolveCollectionItem(pivotTables, DecodeSegment(segments(index)), "pivot table")
                            resolved.Track(pivotTable)
                            current = pivotTable
                            resolved.ObjectKind = "PivotTable"
                        End If

                    Case Else
                        ThrowInvalid($"Unsupported Excel object path segment '{DecodeSegment(segments(index))}'")
                End Select
                index += 1
            End While
            resolved.Value = current
        End Sub

        Private Shared Function ResolveCollectionItem(collection As Object, key As String, kind As String) As Object
            Try
                If key.StartsWith("name:", StringComparison.OrdinalIgnoreCase) Then Return collection.Item(key.Substring(5))
                Dim numericIndex As Integer
                If Integer.TryParse(key, numericIndex) AndAlso numericIndex > 0 Then Return collection.Item(numericIndex)
                Return collection.Item(key)
            Catch ex As Exception
                Throw New ExcelOperationException(ExceptionClassifier.CodeObjectNotFound,
                                                  $"Excel {kind} '{key}' was not found")
            End Try
        End Function

        Private Shared Function HasNextIndex(segments As String(), index As Integer) As Boolean
            If segments Is Nothing OrElse index >= segments.Length Then Return False
            Dim ignored As Integer
            Return Integer.TryParse(DecodeSegment(segments(index)), ignored)
        End Function

        Private Shared Function ParseIndex(value As String, kind As String) As Integer
            Dim parsed As Integer
            If Not Integer.TryParse(value, parsed) OrElse parsed < 1 Then ThrowInvalid($"Invalid {kind} index '{value}'")
            Return parsed
        End Function

        Private Shared Sub EnsureKind(value As Object, expectedKind As String, actualKind As String)
            If value Is Nothing OrElse Not String.Equals(expectedKind, actualKind, StringComparison.OrdinalIgnoreCase) Then
                ThrowTypeMismatch(expectedKind, actualKind)
            End If
        End Sub

        Private Shared Sub ThrowInvalid(message As String)
            Throw New ExcelOperationException(ExceptionClassifier.CodeObjectRefInvalid, message)
        End Sub

        Private Shared Sub ThrowNotFound(kind As String, identity As String)
            Throw New ExcelOperationException(ExceptionClassifier.CodeObjectNotFound,
                                              $"Excel {kind} '{identity}' was not found")
        End Sub

        Private Shared Sub ThrowTypeMismatch(expected As String, actual As String)
            Throw New ExcelOperationException(ExceptionClassifier.CodeObjectTypeMismatch,
                                              $"Expected {expected}, actual {If(actual, "Nothing")}")
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
