Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Safely snapshots observed values before they are placed in a verification contract.
    ''' </summary>
    Friend NotInheritable Class OutcomeProjectionValue
        Private Sub New()
        End Sub

        Friend Shared Function CloneToken(value As Object) As JToken
            If value Is Nothing Then Return Nothing
            Try
                Dim token = TryCast(value, JToken)
                If token Is Nothing Then token = JToken.FromObject(value)
                Return token.DeepClone()
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Returns the observed value for a named host property. Atomic verification keeps
        ''' the property name in an object wrapper (for example {"ValueHash":"..."}); a
        ''' generated requirement for ValueHash must compare the contained value itself.
        ''' </summary>
        Friend Shared Function ClonePropertyValue(value As Object, propertyName As String) As JToken
            Dim cloned = CloneToken(value)
            If cloned Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return cloned

            Dim observedObject = TryCast(cloned, JObject)
            If observedObject Is Nothing Then Return cloned

            Dim propertyValue = observedObject.GetValue(
                propertyName.Trim(),
                StringComparison.OrdinalIgnoreCase)
            Return If(propertyValue Is Nothing, cloned, propertyValue.DeepClone())
        End Function
    End Class

End Namespace
