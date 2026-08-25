Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    Partial Public Class LoopEngine

        Private Shared Function BuildActionSignature(toolCall As ToolCall) As String
            If toolCall Is Nothing Then Return ""
            Dim parameters = If(toolCall.Parameters Is Nothing,
                                 "{}",
                                 CanonicalizeJsonToken(toolCall.Parameters).
                                     ToString(Newtonsoft.Json.Formatting.None))
            Return If(toolCall.ToolId, "").Trim().ToLowerInvariant() & "|" & parameters
        End Function

        Private Shared Function CanonicalizeJsonToken(token As JToken) As JToken
            If token Is Nothing Then Return JValue.CreateNull()
            If token.Type = JTokenType.Object Then
                Dim canonicalObject As New JObject()
                For Each prop In DirectCast(token, JObject).
                    Properties().
                    OrderBy(Function(item) item.Name, StringComparer.Ordinal)
                    canonicalObject.Add(prop.Name, CanonicalizeJsonToken(prop.Value))
                Next
                Return canonicalObject
            End If
            If token.Type = JTokenType.Array Then
                Dim canonicalArray As New JArray()
                For Each item In DirectCast(token, JArray)
                    canonicalArray.Add(CanonicalizeJsonToken(item))
                Next
                Return canonicalArray
            End If
            Return token.DeepClone()
        End Function

        Private Shared Function IsActionBlocked(blocked As Dictionary(Of String, BlockedActionState),
                                                signature As String,
                                                worldRevision As Long) As Boolean
            If blocked Is Nothing OrElse String.IsNullOrWhiteSpace(signature) Then Return False
            Dim state As BlockedActionState = Nothing
            If Not blocked.TryGetValue(signature, state) OrElse state Is Nothing Then Return False
            If state.Permanent OrElse state.WorldRevision = worldRevision Then Return True

            ' A deterministic no-change failure belongs to the world snapshot in which it
            ' occurred. Once another verified action changes the host, the same call may be valid.
            blocked.Remove(signature)
            Return False
        End Function

        Private Shared Function WasSuccessfulMutationAlreadyApplied(
            successful As Dictionary(Of String, Long),
            signature As String,
            worldRevision As Long) As Boolean
            If successful Is Nothing OrElse String.IsNullOrWhiteSpace(signature) Then Return False
            Dim appliedRevision As Long = -1
            If Not successful.TryGetValue(signature, appliedRevision) Then Return False
            Return appliedRevision = worldRevision
        End Function

        Private Shared Function HasRequestBoundVerification(result As ToolResult) As Boolean
            If result Is Nothing OrElse Not result.Success OrElse result.Observation Is Nothing Then Return False
            Try
                Dim observation = TryCast(result.Observation, JToken)
                If observation Is Nothing Then observation = JToken.FromObject(result.Observation)
                Dim root = TryCast(observation, JObject)
                If root Is Nothing Then Return False
                If root("requestExpected") IsNot Nothing Then Return True

                Dim verification = root("verification")
                If TypeOf verification Is JObject Then
                    Return DirectCast(verification, JObject)("requestExpected") IsNot Nothing
                End If
                Dim items = TryCast(verification, JArray)
                If items Is Nothing Then Return False
                Return items.OfType(Of JObject)().Any(
                    Function(item) item("requestExpected") IsNot Nothing AndAlso
                        String.Equals(item("status")?.ToString(), "passed", StringComparison.OrdinalIgnoreCase))
            Catch
                Return False
            End Try
        End Function

        Private Shared Sub BlockAction(blocked As Dictionary(Of String, BlockedActionState),
                                       signature As String,
                                       worldRevision As Long,
                                       permanent As Boolean)
            If blocked Is Nothing OrElse String.IsNullOrWhiteSpace(signature) Then Return
            Dim existing As BlockedActionState = Nothing
            If blocked.TryGetValue(signature, existing) AndAlso existing IsNot Nothing Then
                If existing.Permanent Then Return
                existing.Permanent = permanent
                existing.WorldRevision = worldRevision
                Return
            End If
            blocked(signature) = New BlockedActionState With {
                .WorldRevision = worldRevision,
                .Permanent = permanent
            }
        End Sub

    End Class

End Namespace
