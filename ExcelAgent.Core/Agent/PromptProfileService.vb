Imports System.IO
Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Class PromptProfile
        Public Property PersonalPrompt As String = ""
        Public Property UserProfile As String = ""
        Public Property ExternalPrompt As String = ""
        Public Property SourceSummary As New List(Of String)()

        Public ReadOnly Property HasAny As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(PersonalPrompt) OrElse
                       Not String.IsNullOrWhiteSpace(ExternalPrompt)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Loads only explicit Excel Agent prompts. No database, memory profile, or legacy
    ''' application prompt can silently alter the current run.
    ''' </summary>
    Public NotInheritable Class PromptProfileService
        Private Const MaxPromptChars As Integer = 5000
        Private Const MaxFileChars As Integer = 2500

        Private Sub New()
        End Sub

        Public Shared Function Load(appType As String) As PromptProfile
            Dim profile As New PromptProfile()
            If Not String.Equals(If(appType, ""), "excel", StringComparison.OrdinalIgnoreCase) Then Return profile

            profile.PersonalPrompt = TrimToLimit(Global.ExcelAgent.Core.ConfigSettings.propmtContent, MaxPromptChars)
            If Not String.IsNullOrWhiteSpace(profile.PersonalPrompt) Then
                profile.SourceSummary.Add("selected:" & If(Global.ExcelAgent.Core.ConfigSettings.propmtName, "custom"))
            End If
            profile.ExternalPrompt = LoadExternalPrompts(profile.SourceSummary)
            Return profile
        End Function

        Public Shared Function GetExternalPromptDirectory() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Global.ExcelAgent.Core.ConfigSettings.OfficeAiAppDataFolder,
                "Prompts")
        End Function

        Private Shared Function LoadExternalPrompts(sources As List(Of String)) As String
            Dim baseDir = GetExternalPromptDirectory()
            If Not Directory.Exists(baseDir) Then Return ""

            Dim sb As New StringBuilder()
            For Each filePath In Directory.GetFiles(baseDir, "*.*", SearchOption.TopDirectoryOnly).
                Where(Function(candidate)
                          Dim extension = System.IO.Path.GetExtension(candidate).ToLowerInvariant()
                          Return extension = ".md" OrElse extension = ".txt" OrElse extension = ".json"
                      End Function).
                OrderBy(Function(candidate) candidate, StringComparer.OrdinalIgnoreCase)
                Try
                    Dim content = File.ReadAllText(filePath, Encoding.UTF8)
                    If Path.GetExtension(filePath).Equals(".json", StringComparison.OrdinalIgnoreCase) Then
                        content = ExtractJsonPrompt(content)
                    End If
                    content = TrimToLimit(content, MaxFileChars)
                    If String.IsNullOrWhiteSpace(content) Then Continue For
                    sb.AppendLine("# " & Path.GetFileName(filePath))
                    sb.AppendLine(content.Trim())
                    sb.AppendLine()
                    sources.Add("file:" & Path.GetFileName(filePath))
                Catch ex As Exception
                    AppLogger.Warn("PromptProfile", "Cannot read prompt file: " & ex.Message)
                End Try
            Next
            Return TrimToLimit(sb.ToString().Trim(), MaxPromptChars)
        End Function

        Private Shared Function ExtractJsonPrompt(json As String) As String
            Try
                Dim root = JObject.Parse(json)
                If root("enabled") IsNot Nothing AndAlso Not root("enabled").Value(Of Boolean)() Then Return ""
                Return If(root("content")?.ToString(), If(root("prompt")?.ToString(), ""))
            Catch
                Return json
            End Try
        End Function

        Private Shared Function TrimToLimit(value As String, maxChars As Integer) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            If value.Length <= maxChars Then Return value
            Return value.Substring(0, maxChars) & vbCrLf & "...(truncated)"
        End Function
    End Class
End Namespace
