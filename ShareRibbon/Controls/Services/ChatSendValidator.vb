Public Enum ChatSendValidationFailure
    None = 0
    MissingApiKey
    MissingApiUrl
    MissingQuestion
End Enum

Public Class ChatSendValidationResult
    Public Property IsValid As Boolean
    Public Property Failure As ChatSendValidationFailure

    Public Shared Function Success() As ChatSendValidationResult
        Return New ChatSendValidationResult With {
            .IsValid = True,
            .Failure = ChatSendValidationFailure.None
        }
    End Function

    Public Shared Function Fail(failure As ChatSendValidationFailure) As ChatSendValidationResult
        Return New ChatSendValidationResult With {
            .IsValid = False,
            .Failure = failure
        }
    End Function
End Class

Public Class ChatSendValidator
    Public Function Validate(apiUrl As String, apiKey As String, question As String) As ChatSendValidationResult
        If String.IsNullOrWhiteSpace(apiKey) Then
            Return ChatSendValidationResult.Fail(ChatSendValidationFailure.MissingApiKey)
        End If

        If String.IsNullOrWhiteSpace(apiUrl) Then
            Return ChatSendValidationResult.Fail(ChatSendValidationFailure.MissingApiUrl)
        End If

        If String.IsNullOrWhiteSpace(question) Then
            Return ChatSendValidationResult.Fail(ChatSendValidationFailure.MissingQuestion)
        End If

        Return ChatSendValidationResult.Success()
    End Function
End Class
