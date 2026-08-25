Public Class HistoryMessage
    Public Property role As String
    Public Property content As String
    Public Property Timestamp As DateTime = DateTime.Now
    Public Property Uuid As String = Guid.NewGuid().ToString("N")
End Class
