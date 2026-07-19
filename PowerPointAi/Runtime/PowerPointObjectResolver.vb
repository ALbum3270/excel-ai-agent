Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend Class PowerPointOperationException
        Inherits Exception

        Public ReadOnly Property ErrorCode As String
        Public ReadOnly Property Recoverable As Boolean

        Public Sub New(errorCode As String, message As String, Optional recoverable As Boolean = True)
            MyBase.New(message)
            Me.ErrorCode = errorCode
            Me.Recoverable = recoverable
        End Sub
    End Class

    Friend Class ResolvedOfficeObject
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
    ''' Resolves canonical PowerPoint refs into short-lived COM objects. Resolved RCWs
    ''' are owned by the returned scope and must not cross an operation boundary.
    ''' </summary>
    Friend NotInheritable Class PowerPointObjectResolver
        Private Sub New()
        End Sub

        Public Shared Function Resolve(targetRef As String) As ResolvedOfficeObject
            Dim parsed As OfficeObjectRef = Nothing
            Dim parseError As String = ""
            If Not OfficeObjectRef.TryParse(targetRef, parsed, parseError) Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectRefInvalid, parseError)
            End If
            If Not String.Equals(parsed.AppType, "PowerPoint", StringComparison.OrdinalIgnoreCase) Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeHostUnsupported,
                                                       $"Target ref belongs to {parsed.AppType}",
                                                       recoverable:=False)
            End If

            Dim application = Globals.ThisAddIn.Application
            If application Is Nothing Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeDocMissing,
                                                       "PowerPoint application is unavailable",
                                                       recoverable:=False)
            End If

            Dim presentation As PowerPoint.Presentation = Nothing
            Try
                presentation = application.ActivePresentation
            Catch
            End Try
            If presentation Is Nothing Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeDocMissing,
                                                       "No active PowerPoint presentation",
                                                       recoverable:=False)
            End If
            If Not String.Equals(parsed.DocumentRef, "active", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(parsed.DocumentRef, presentation.Name, StringComparison.OrdinalIgnoreCase) Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectNotFound,
                                                       $"Active presentation does not match '{parsed.DocumentRef}'")
            End If

            Dim resolved As New ResolvedOfficeObject With {
                .Value = presentation,
                .CanonicalRef = parsed.ToCanonicalString(),
                .ObjectKind = "Presentation"
            }
            resolved.Track(presentation)
            Try
                ResolvePath(parsed.Path, presentation, resolved)
                Return resolved
            Catch
                resolved.Dispose()
                Throw
            End Try
        End Function

        Private Shared Sub ResolvePath(path As String,
                                       presentation As PowerPoint.Presentation,
                                       resolved As ResolvedOfficeObject)
            If String.IsNullOrWhiteSpace(path) Then Return
            Dim segments = path.Split("/"c)
            Dim current As Object = presentation
            Dim index As Integer = 0

            While index < segments.Length
                Dim segment = segments(index).Trim().ToLowerInvariant()
                Select Case segment
                    Case "slides"
                        If Not TypeOf current Is PowerPoint.Presentation Then ThrowTypeMismatch("Slides", current)
                        Dim slides = DirectCast(current, PowerPoint.Presentation).Slides
                        resolved.Track(slides)
                        current = slides
                        resolved.ObjectKind = "Slides"
                        If HasIndex(segments, index + 1) Then
                            index += 1
                            Dim slideIndex = ParseIndex(segments(index), "slide")
                            If slideIndex < 1 OrElse slideIndex > slides.Count Then ThrowNotFound("slide", slideIndex)
                            Dim slide = slides(slideIndex)
                            resolved.Track(slide)
                            current = slide
                            resolved.ObjectKind = "Slide"
                        End If

                    Case "shapes"
                        If Not TypeOf current Is PowerPoint.Slide Then ThrowTypeMismatch("Shapes", current)
                        Dim shapes = DirectCast(current, PowerPoint.Slide).Shapes
                        resolved.Track(shapes)
                        current = shapes
                        resolved.ObjectKind = "Shapes"
                        If HasIndex(segments, index + 1) Then
                            index += 1
                            Dim shapeIndex = ParseIndex(segments(index), "shape")
                            If shapeIndex < 1 OrElse shapeIndex > shapes.Count Then ThrowNotFound("shape", shapeIndex)
                            Dim shape = shapes(shapeIndex)
                            resolved.Track(shape)
                            current = shape
                            resolved.ObjectKind = "Shape"
                        End If

                    Case "smartart"
                        If Not TypeOf current Is PowerPoint.Shape Then ThrowTypeMismatch("SmartArt", current)
                        Dim shape = DirectCast(current, PowerPoint.Shape)
                        Dim smartArt As SmartArt = Nothing
                        Try
                            smartArt = shape.SmartArt
                        Catch
                        End Try
                        If smartArt Is Nothing Then ThrowNotFound("SmartArt", 0)
                        resolved.Track(smartArt)
                        current = smartArt
                        resolved.ObjectKind = "SmartArt"

                    Case "nodes"
                        Dim nodes As SmartArtNodes = Nothing
                        If TypeOf current Is SmartArt Then
                            nodes = DirectCast(current, SmartArt).Nodes
                        ElseIf TypeOf current Is SmartArtNode Then
                            nodes = DirectCast(current, SmartArtNode).Nodes
                        Else
                            ThrowTypeMismatch("SmartArtNodes", current)
                        End If
                        If nodes Is Nothing Then ThrowNotFound("SmartArtNodes", 0)
                        resolved.Track(nodes)
                        current = nodes
                        resolved.ObjectKind = "SmartArtNodes"
                        If HasIndex(segments, index + 1) Then
                            index += 1
                            Dim nodeIndex = ParseIndex(segments(index), "SmartArt node")
                            If nodeIndex < 1 OrElse nodeIndex > nodes.Count Then ThrowNotFound("SmartArt node", nodeIndex)
                            Dim node = nodes.Item(nodeIndex)
                            resolved.Track(node)
                            current = node
                            resolved.ObjectKind = "SmartArtNode"
                        End If

                    Case "textframe"
                        If Not TypeOf current Is PowerPoint.Shape Then ThrowTypeMismatch("TextFrame", current)
                        Dim textFrame = DirectCast(current, PowerPoint.Shape).TextFrame
                        resolved.Track(textFrame)
                        current = textFrame
                        resolved.ObjectKind = "TextFrame"

                    Case "textframe2"
                        Dim textFrame2 As TextFrame2 = Nothing
                        If TypeOf current Is PowerPoint.Shape Then
                            textFrame2 = DirectCast(current, PowerPoint.Shape).TextFrame2
                        ElseIf TypeOf current Is SmartArtNode Then
                            textFrame2 = DirectCast(current, SmartArtNode).TextFrame2
                        Else
                            ThrowTypeMismatch("TextFrame2", current)
                        End If
                        resolved.Track(textFrame2)
                        current = textFrame2
                        resolved.ObjectKind = "TextFrame2"

                    Case "textrange"
                        Dim textRange As Object = Nothing
                        If TypeOf current Is PowerPoint.TextFrame Then
                            textRange = DirectCast(current, PowerPoint.TextFrame).TextRange
                            resolved.ObjectKind = "TextRange"
                        ElseIf TypeOf current Is TextFrame2 Then
                            textRange = DirectCast(current, TextFrame2).TextRange
                            resolved.ObjectKind = "TextRange2"
                        Else
                            ThrowTypeMismatch("TextRange", current)
                        End If
                        resolved.Track(textRange)
                        current = textRange

                    Case Else
                        Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectRefInvalid,
                                                               $"Unsupported PowerPoint object path segment '{segments(index)}'")
                End Select
                index += 1
            End While

            resolved.Value = current
        End Sub

        Private Shared Function HasIndex(segments As String(), index As Integer) As Boolean
            If segments Is Nothing OrElse index >= segments.Length Then Return False
            Dim ignored As Integer
            Return Integer.TryParse(segments(index), ignored)
        End Function

        Private Shared Function ParseIndex(value As String, kind As String) As Integer
            Dim parsed As Integer
            If Not Integer.TryParse(value, parsed) OrElse parsed < 1 Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectRefInvalid,
                                                       $"Invalid {kind} index '{value}'")
            End If
            Return parsed
        End Function

        Private Shared Sub ThrowNotFound(kind As String, index As Integer)
            Dim suffix = If(index > 0, $" at index {index}", "")
            Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectNotFound,
                                                   $"PowerPoint {kind} was not found{suffix}")
        End Sub

        Private Shared Sub ThrowTypeMismatch(expected As String, actual As Object)
            Dim actualName = If(actual Is Nothing, "Nothing", actual.GetType().FullName)
            Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectTypeMismatch,
                                                   $"Expected {expected}, actual {actualName}")
        End Sub
    End Class

End Namespace
