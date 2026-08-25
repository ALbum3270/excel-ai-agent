Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Produces stable content identities for tool payloads. A hash records lineage only; it
    ''' never substitutes for host verification of the user's requested result.
    ''' </summary>
    Friend NotInheritable Class OutcomeEvidenceValueHasher
        Private Sub New()
        End Sub

        Friend Shared Function Compute(value As Object) As String
            Dim token As JToken = Nothing
            Try
                token = TryCast(value, JToken)
                If token Is Nothing AndAlso value IsNot Nothing Then token = JToken.FromObject(value)
            Catch
                token = Nothing
            End Try
            If token Is Nothing Then Return ""

            Dim canonical = token.ToString(Formatting.None)
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(canonical))).
                    Replace("-", "").ToLowerInvariant()
            End Using
        End Function
    End Class

End Namespace
