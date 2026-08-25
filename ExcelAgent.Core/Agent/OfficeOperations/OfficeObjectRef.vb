Imports System.Linq

Namespace Agent.OfficeOperations

    ''' <summary>
    ''' Serializable reference to an Excel object. COM RCWs must never be stored in
    ''' Agent memory, plans, observations, or run traces.
    ''' </summary>
    Public Class OfficeObjectRef
        Public Property AppType As String
        Public Property DocumentRef As String
        Public Property Path As String

        Public Function ToCanonicalString() As String
            Dim normalizedApp = NormalizeAppType(AppType)
            Dim root = GetDocumentRoot(normalizedApp)
            Dim document = TrimPath(DocumentRef)
            Dim objectPath = TrimPath(Path)

            If String.IsNullOrWhiteSpace(normalizedApp) OrElse
               String.IsNullOrWhiteSpace(root) OrElse
               String.IsNullOrWhiteSpace(document) Then
                Return ""
            End If

            If document.StartsWith(root & "/", StringComparison.OrdinalIgnoreCase) Then
                document = document.Substring(root.Length + 1)
            End If

            Dim canonical = normalizedApp & ":" & root & "/" & document
            If Not String.IsNullOrWhiteSpace(objectPath) Then canonical &= "/" & objectPath
            Return canonical
        End Function

        Public Overrides Function ToString() As String
            Return ToCanonicalString()
        End Function

        Public Shared Function TryParse(value As String,
                                        ByRef objectRef As OfficeObjectRef,
                                        Optional ByRef errorMessage As String = Nothing) As Boolean
            objectRef = Nothing
            errorMessage = ""
            If String.IsNullOrWhiteSpace(value) Then
                errorMessage = "Office object ref is empty"
                Return False
            End If

            Dim separatorIndex = value.IndexOf(":"c)
            If separatorIndex <= 0 OrElse separatorIndex >= value.Length - 1 Then
                errorMessage = "Office object ref must use '<AppType>:<path>' format"
                Return False
            End If

            Dim appType = NormalizeAppType(value.Substring(0, separatorIndex))
            Dim expectedRoot = GetDocumentRoot(appType)
            If String.IsNullOrWhiteSpace(expectedRoot) Then
                errorMessage = "Unsupported Office app type"
                Return False
            End If

            Dim rawPath = TrimPath(value.Substring(separatorIndex + 1))
            Dim segments = rawPath.Split("/"c)
            If segments.Length < 2 OrElse segments.Any(Function(item) String.IsNullOrWhiteSpace(item)) Then
                errorMessage = "Office object ref must include document collection and document identity"
                Return False
            End If
            If Not String.Equals(segments(0), expectedRoot, StringComparison.OrdinalIgnoreCase) Then
                errorMessage = $"Office object ref root '{segments(0)}' does not match app type {appType}"
                Return False
            End If

            objectRef = New OfficeObjectRef With {
                .AppType = appType,
                .DocumentRef = segments(1),
                .Path = String.Join("/", segments.Skip(2))
            }
            Return True
        End Function

        Public Shared Function NormalizeAppType(value As String) As String
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "xls", "xlsx", "excel"
                    Return "Excel"
                Case Else
                    Return ""
            End Select
        End Function

        Private Shared Function GetDocumentRoot(appType As String) As String
            Select Case NormalizeAppType(appType)
                Case "Excel"
                    Return "workbooks"
                Case Else
                    Return ""
            End Select
        End Function

        Private Shared Function TrimPath(value As String) As String
            Return If(value, "").Trim().Trim("/"c)
        End Function
    End Class

End Namespace
