Namespace Agent.OfficeOperations

    Public Class OfficeCapabilitySearchRequest
        Public Property Query As String
        Public Property TargetType As String
        Public Property IncludeReadOnly As Boolean = True
        Public Property MaxResults As Integer = 12
    End Class

    Public Class OfficeCapabilitySearchResult
        Public Property AppType As String
        Public Property Query As String
        Public Property Members As New List(Of OfficeCapabilityMember)()
        Public Property CatalogFingerprint As String
        Public Property Truncated As Boolean
        Public Property Warnings As New List(Of String)()
    End Class

    Public Class OfficeCapabilityMember
        Public Property MemberId As String
        Public Property DeclaringType As String
        Public Property MemberName As String
        Public Property MemberKind As String
        Public Property Parameters As New List(Of OfficeCapabilityParameter)()
        Public Property ReturnType As String
        Public Property RiskLevel As String
        Public Property Executable As Boolean
        Public Property UnsupportedReason As String
        Public Property Aliases As New List(Of String)()
    End Class

    Public Class OfficeCapabilityParameter
        Public Property Name As String
        Public Property ParameterType As String
        Public Property Required As Boolean
        Public Property DefaultValue As Object
        Public Property Description As String
    End Class

End Namespace
