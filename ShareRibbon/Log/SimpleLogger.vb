' ShareRibbon\Log\SimpleLogger.vb
' Compatibility facade — routes to AppLogger (P0-4).

Imports System.IO

''' <summary>
''' Legacy logger API. Prefer AppLogger with explicit module names.
''' Kept so existing callers (e.g. ExcelAi/DeepseekControl) continue to compile.
''' </summary>
Public Module SimpleLogger
    ''' <summary>Historical path retained for callers that read it; new writes go to AppLogger path.</summary>
    Public ReadOnly Property LogFile As String
        Get
            Return AppLogger.GetLogFilePath()
        End Get
    End Property

    Public Sub LogInfo(msg As String)
        AppLogger.Info("SimpleLogger", msg)
    End Sub

    Public Sub LogError(msg As String, Optional ex As Exception = Nothing)
        AppLogger.Error("SimpleLogger", msg, ex)
    End Sub

    Public Sub LogWarn(msg As String)
        AppLogger.Warn("SimpleLogger", msg)
    End Sub

    Public Sub LogDebug(msg As String)
        AppLogger.Debug("SimpleLogger", msg)
    End Sub
End Module
