' ShareRibbon\Log\SimpleLogger.vb
' Compatibility facade — routes to AppLogger (P0-4).

''' <summary>
''' Legacy logger API. Prefer AppLogger with explicit module names.
''' Kept so existing callers (e.g. ExcelAi/DeepseekControl) continue to compile.
''' </summary>
Public Module SimpleLogger
    Public Sub LogError(msg As String, Optional ex As Exception = Nothing)
        AppLogger.Error("SimpleLogger", msg, ex)
    End Sub

End Module
