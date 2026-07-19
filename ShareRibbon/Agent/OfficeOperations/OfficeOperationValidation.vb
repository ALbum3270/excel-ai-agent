Namespace Agent.OfficeOperations

    Public Class OfficeOperationValidationResult
        Public Property Errors As New List(Of String)()

        Public ReadOnly Property IsValid As Boolean
            Get
                Return Errors.Count = 0
            End Get
        End Property

        Public Function ToErrorMessage() As String
            Return String.Join("; ", Errors)
        End Function
    End Class

    Public NotInheritable Class OfficeOperationValidation
        Public Const CurrentSchemaVersion As String = "1.0"
        Public Const MaxOperationsPerBatch As Integer = 50
        Public Const MaxCapabilityResults As Integer = 20

        Private Shared ReadOnly AllowedActions As New HashSet(Of String)(
            {"get", "set", "invoke", "create", "delete", "collection_item"},
            StringComparer.OrdinalIgnoreCase)

        Private Sub New()
        End Sub

        Public Shared Function ValidateCapabilitySearch(request As OfficeCapabilitySearchRequest) As OfficeOperationValidationResult
            Dim result As New OfficeOperationValidationResult()
            If request Is Nothing Then
                result.Errors.Add("OPERATION_SCHEMA_INVALID: capability search request is required")
                Return result
            End If
            If String.IsNullOrWhiteSpace(request.Query) Then
                result.Errors.Add("OPERATION_SCHEMA_INVALID: capability search query is required")
            End If
            If request.MaxResults < 1 OrElse request.MaxResults > MaxCapabilityResults Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: maxResults must be between 1 and {MaxCapabilityResults}")
            End If
            Return result
        End Function

        Public Shared Function ValidateBatch(batch As OfficeOperationBatch) As OfficeOperationValidationResult
            Dim result As New OfficeOperationValidationResult()
            If batch Is Nothing Then
                result.Errors.Add("OPERATION_SCHEMA_INVALID: batch is required")
                Return result
            End If

            If Not String.Equals(batch.SchemaVersion, CurrentSchemaVersion, StringComparison.Ordinal) Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: unsupported schemaVersion '{If(batch.SchemaVersion, "")}'")
            End If

            Dim normalizedApp = OfficeObjectRef.NormalizeAppType(batch.AppType)
            If String.IsNullOrWhiteSpace(normalizedApp) Then
                result.Errors.Add($"HOST_UNSUPPORTED: unsupported appType '{If(batch.AppType, "")}'")
            End If

            If batch.Operations Is Nothing OrElse batch.Operations.Count = 0 Then
                result.Errors.Add("OPERATION_SCHEMA_INVALID: operations must contain at least one item")
                Return result
            End If
            If batch.Operations.Count > MaxOperationsPerBatch Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: operations exceeds limit {MaxOperationsPerBatch}")
            End If

            Dim operationIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For index = 0 To batch.Operations.Count - 1
                ValidateOperation(batch.Operations(index), index, normalizedApp, operationIds, result)
            Next

            If batch.SuccessCriteria IsNot Nothing Then
                For index = 0 To batch.SuccessCriteria.Count - 1
                    Dim criterion = batch.SuccessCriteria(index)
                    If criterion Is Nothing Then
                        result.Errors.Add($"OPERATION_SCHEMA_INVALID: successCriteria[{index}] is null")
                        Continue For
                    End If
                    If Not String.IsNullOrWhiteSpace(criterion.TargetRef) Then
                        ValidateTargetRef(criterion.TargetRef,
                                          normalizedApp,
                                          $"successCriteria[{index}].targetRef",
                                          result)
                    End If
                Next
            End If

            Return result
        End Function

        Public Shared Function ValidateCanonicalRef(value As String,
                                                    expectedAppType As String,
                                                    ByRef errorMessage As String) As Boolean
            Dim parsed As OfficeObjectRef = Nothing
            If Not OfficeObjectRef.TryParse(value, parsed, errorMessage) Then Return False

            Dim normalizedExpected = OfficeObjectRef.NormalizeAppType(expectedAppType)
            If Not String.IsNullOrWhiteSpace(normalizedExpected) AndAlso
               Not String.Equals(parsed.AppType, normalizedExpected, StringComparison.OrdinalIgnoreCase) Then
                errorMessage = $"HOST_UNSUPPORTED: target ref appType {parsed.AppType} does not match {normalizedExpected}"
                Return False
            End If
            Return True
        End Function

        Private Shared Sub ValidateOperation(operation As OfficeOperation,
                                             index As Integer,
                                             normalizedApp As String,
                                             operationIds As HashSet(Of String),
                                             result As OfficeOperationValidationResult)
            Dim prefix = $"operations[{index}]"
            If operation Is Nothing Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: {prefix} is null")
                Return
            End If

            If String.IsNullOrWhiteSpace(operation.Id) Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: {prefix}.id is required")
            ElseIf Not operationIds.Add(operation.Id.Trim()) Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: duplicate operation id '{operation.Id}'")
            End If

            If String.IsNullOrWhiteSpace(operation.Action) OrElse Not AllowedActions.Contains(operation.Action.Trim()) Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: {prefix}.action must be get/set/invoke/create/delete/collection_item")
            End If
            If String.IsNullOrWhiteSpace(operation.MemberId) Then
                result.Errors.Add($"OPERATION_SCHEMA_INVALID: {prefix}.memberId is required")
            End If

            ValidateTargetRef(operation.TargetRef, normalizedApp, prefix & ".targetRef", result)
        End Sub

        Private Shared Sub ValidateTargetRef(value As String,
                                             normalizedApp As String,
                                             fieldName As String,
                                             result As OfficeOperationValidationResult)
            Dim errorMessage As String = ""
            If Not ValidateCanonicalRef(value, normalizedApp, errorMessage) Then
                result.Errors.Add($"OBJECT_REF_INVALID: {fieldName}: {errorMessage}")
            End If
        End Sub
    End Class

End Namespace
