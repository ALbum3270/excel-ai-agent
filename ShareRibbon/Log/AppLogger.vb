' ShareRibbon\Log\AppLogger.vb
' Structured logger with level, module, correlation id, and secret redaction (P0-4).

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' Application-wide logger. Prefer AppLogger over Debug.WriteLine for Agent/Tool failures.
''' Never log raw API keys, Bearer tokens, or Authorization headers.
''' </summary>
Public NotInheritable Class AppLogger
    Private Sub New()
    End Sub

    Public Enum LogLevel
        Debug = 0
        Info = 1
        Warn = 2
        [Error] = 3
    End Enum

    Private Shared ReadOnly _lock As New Object()
    Private Shared _correlationId As String = ""
    Private Shared _minLevel As LogLevel = LogLevel.Debug

    Private Shared ReadOnly SecretPatterns As Regex() = {
        New Regex("(?i)(api[_-]?key|apikey|access[_-]?token|secret|password)\s*[:=]\s*[""']?([^\s""',}{]+)", RegexOptions.Compiled),
        New Regex("(?i)(Bearer)\s+([A-Za-z0-9\-._~+/]+=*)", RegexOptions.Compiled),
        New Regex("(?i)(Authorization\s*[:=]\s*)([^\s,;]+)", RegexOptions.Compiled),
        New Regex("(?i)(sk-[A-Za-z0-9]{8,})", RegexOptions.Compiled)
    }

    Public Shared Property MinimumLevel As LogLevel
        Get
            Return _minLevel
        End Get
        Set(value As LogLevel)
            _minLevel = value
        End Set
    End Property

    ''' <summary>Current request/session correlation id (AsyncLocal-like via thread-static + override).</summary>
    Public Shared Property CorrelationId As String
        Get
            If String.IsNullOrWhiteSpace(_correlationId) Then Return "-"
            Return _correlationId
        End Get
        Set(value As String)
            _correlationId = If(value, "")
        End Set
    End Property

    Public Shared Function BeginScope(Optional correlationId As String = Nothing) As String
        Dim id = If(String.IsNullOrWhiteSpace(correlationId), Guid.NewGuid().ToString("N").Substring(0, 12), correlationId.Trim())
        AppLogger.CorrelationId = id
        Return id
    End Function

    Public Shared Sub ClearScope()
        _correlationId = ""
    End Sub

    Public Shared Sub Debug(moduleName As String, message As String, Optional ex As Exception = Nothing)
        Write(LogLevel.Debug, moduleName, message, ex)
    End Sub

    Public Shared Sub Info(moduleName As String, message As String, Optional ex As Exception = Nothing)
        Write(LogLevel.Info, moduleName, message, ex)
    End Sub

    Public Shared Sub Warn(moduleName As String, message As String, Optional ex As Exception = Nothing)
        Write(LogLevel.Warn, moduleName, message, ex)
    End Sub

    Public Shared Sub [Error](moduleName As String, message As String, Optional ex As Exception = Nothing)
        Write(LogLevel.Error, moduleName, message, ex)
    End Sub

    Public Shared Sub Write(level As LogLevel, moduleName As String, message As String, Optional ex As Exception = Nothing)
        If level < _minLevel Then Return

        Dim modName = If(String.IsNullOrWhiteSpace(moduleName), "App", moduleName.Trim())
        Dim safeMessage = Redact(If(message, ""))
        Dim line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level.ToString().ToUpperInvariant()}] [cid={CorrelationId}] [{modName}] {safeMessage}"

        If ex IsNot Nothing Then
            line &= $" | ex={ex.GetType().Name}: {Redact(ex.Message)}"
        End If

        Try
            Diagnostics.Debug.WriteLine(line)
        Catch
        End Try

        Try
            AppendToFile(line, ex)
        Catch
            ' Never throw from logger
        End Try
    End Sub

    Public Shared Function Redact(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text
        Dim result = text
        For Each pattern In SecretPatterns
            result = pattern.Replace(result, AddressOf RedactMatch)
        Next
        Return result
    End Function

    Private Shared Function RedactMatch(m As Match) As String
        If m Is Nothing Then Return "***"
        If m.Groups.Count >= 3 AndAlso m.Groups(2).Success Then
            Return m.Groups(1).Value & ":***"
        End If
        Return "***"
    End Function

    Private Shared Sub AppendToFile(line As String, ex As Exception)
        Dim logPath = GetLogFilePath()
        Dim dir = IO.Path.GetDirectoryName(logPath)
        If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine(line)
        If ex IsNot Nothing AndAlso Not String.IsNullOrEmpty(ex.StackTrace) Then
            sb.AppendLine("  StackTrace: " & Redact(ex.StackTrace))
        End If

        SyncLock _lock
            File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8)
        End SyncLock
    End Sub

    Public Shared Function GetLogFilePath() As String
        Dim folderName = ConfigSettings.OfficeAiAppDataFolder
        Dim baseDir = IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            folderName,
            "logs")
        Return IO.Path.Combine(baseDir, $"office-ai-{DateTime.Now:yyyyMMdd}.log")
    End Function
End Class
