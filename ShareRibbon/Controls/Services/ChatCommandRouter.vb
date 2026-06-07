Imports System.Diagnostics
Imports Newtonsoft.Json.Linq

Public Class ChatCommandRouter
    Private ReadOnly _handlers As New Dictionary(Of String, Action(Of JObject))(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _unknownHandler As Action(Of String)

    Public Sub New(Optional unknownHandler As Action(Of String) = Nothing)
        _unknownHandler = unknownHandler
    End Sub

    Public Sub Register(messageType As String, handler As Action(Of JObject))
        If String.IsNullOrWhiteSpace(messageType) Then Throw New ArgumentException("messageType is required", NameOf(messageType))
        If handler Is Nothing Then Throw New ArgumentNullException(NameOf(handler))
        _handlers(messageType.Trim()) = handler
    End Sub

    Public Function Dispatch(jsonDoc As JObject) As Boolean
        If jsonDoc Is Nothing Then Throw New ArgumentNullException(NameOf(jsonDoc))

        Dim typeToken = jsonDoc("type")
        If typeToken Is Nothing OrElse String.IsNullOrWhiteSpace(typeToken.ToString()) Then
            Debug.WriteLine("[ChatCommandRouter] message type is missing")
            Return False
        End If

        Dim messageType = typeToken.ToString()
        Dim handler As Action(Of JObject) = Nothing
        If _handlers.TryGetValue(messageType, handler) Then
            handler(jsonDoc)
            Return True
        End If

        If _unknownHandler IsNot Nothing Then
            _unknownHandler(messageType)
        Else
            Debug.WriteLine($"[ChatCommandRouter] unknown message type: {messageType}")
        End If

        Return False
    End Function
End Class
