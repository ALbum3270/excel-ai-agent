Imports Newtonsoft.Json.Linq

Namespace Agent.OfficeOperations

    Public Class OfficeOperationBatch
        Public Property SchemaVersion As String = "1.0"
        Public Property AppType As String
        Public Property Operations As New List(Of OfficeOperation)()
        Public Property Atomic As Boolean = True
        Public Property SuccessCriteria As New List(Of OperationCriterion)()
    End Class

    Public Class OfficeOperation
        Public Property Id As String
        Public Property TargetRef As String
        Public Property Action As String
        Public Property MemberId As String
        Public Property Arguments As New JObject()
        Public Property ExpectedEffects As New JObject()
    End Class

    Public Class OperationCriterion
        Public Property Id As String
        Public Property Description As String
        Public Property TargetRef As String
        Public Property PropertyName As String
        Public Property [Operator] As String = "equals"
        Public Property ExpectedValue As JToken
        Public Property Required As Boolean = True
    End Class

End Namespace
