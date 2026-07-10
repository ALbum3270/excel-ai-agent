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
                       Not String.IsNullOrWhiteSpace(UserProfile) OrElse
                       Not String.IsNullOrWhiteSpace(ExternalPrompt)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Agent 提示词画像服务：统一读取用户个人风格、用户画像和外接提示词。
    ''' 这些内容只能影响表达偏好和业务背景，不能覆盖系统/工具/安全协议。
    ''' </summary>
    Public Class PromptProfileService
        Private Const MaxPersonalPromptChars As Integer = 3000
        Private Const MaxUserProfileChars As Integer = 2500
        Private Const MaxExternalPromptChars As Integer = 5000
        Private Const MaxSingleExternalFileChars As Integer = 2000

        Public Shared Function Load(appType As String) As PromptProfile
            Dim profile As New PromptProfile()
            Dim scenario = NormalizeScenario(appType)

            profile.PersonalPrompt = BuildPersonalPrompt(scenario, profile.SourceSummary)
            profile.UserProfile = BuildUserProfile(profile.SourceSummary)
            profile.ExternalPrompt = BuildExternalPrompt(scenario, profile.SourceSummary)

            Return profile
        End Function

        Private Shared Function BuildPersonalPrompt(scenario As String, sources As List(Of String)) As String
            Dim sb As New StringBuilder()

            AppendNamedBlock(sb, "当前选中的聊天提示词", Global.ShareRibbon.ConfigSettings.propmtContent)
            If Not String.IsNullOrWhiteSpace(Global.ShareRibbon.ConfigSettings.propmtContent) Then
                sources.Add($"selected:{If(Global.ShareRibbon.ConfigSettings.propmtName, "unnamed")}")
            End If

            Try
                Dim commonPrompt = Global.ShareRibbon.PromptTemplateRepository.GetSystemPrompt("common")
                Dim scenarioPrompt = Global.ShareRibbon.PromptTemplateRepository.GetSystemPrompt(scenario)

                AppendNamedBlock(sb, "外接提示词(common)", commonPrompt)
                If Not String.IsNullOrWhiteSpace(commonPrompt) Then sources.Add("prompt_template:common")

                If Not String.Equals(scenario, "common", StringComparison.OrdinalIgnoreCase) Then
                    AppendNamedBlock(sb, $"外接提示词({scenario})", scenarioPrompt)
                    If Not String.IsNullOrWhiteSpace(scenarioPrompt) Then sources.Add($"prompt_template:{scenario}")
                End If
            Catch ex As Exception
                Debug.WriteLine($"[PromptProfileService] 读取 prompt_template 失败: {ex.Message}")
            End Try

            Return TrimToLimit(sb.ToString().Trim(), MaxPersonalPromptChars)
        End Function

        Private Shared Function BuildUserProfile(sources As List(Of String)) As String
            Try
                If Not Global.ShareRibbon.MemoryConfig.EnableUserProfile Then Return ""
                Dim content = Global.ShareRibbon.MemoryRepository.GetUserProfile()
                If Not String.IsNullOrWhiteSpace(content) Then sources.Add("memory:user_profile")
                Return TrimToLimit(content, MaxUserProfileChars)
            Catch ex As Exception
                Debug.WriteLine($"[PromptProfileService] 读取用户画像失败: {ex.Message}")
                Return ""
            End Try
        End Function

        Private Shared Function BuildExternalPrompt(scenario As String, sources As List(Of String)) As String
            Dim baseDir = GetExternalPromptDirectory()
            If Not Directory.Exists(baseDir) Then Return ""

            Dim sb As New StringBuilder()
            AppendExternalFiles(sb, baseDir, "common", sources)
            If Not String.Equals(scenario, "common", StringComparison.OrdinalIgnoreCase) Then
                AppendExternalFiles(sb, baseDir, scenario, sources)
            End If

            Return TrimToLimit(sb.ToString().Trim(), MaxExternalPromptChars)
        End Function

        Public Shared Function GetExternalPromptDirectory() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Global.ShareRibbon.ConfigSettings.OfficeAiAppDataFolder,
                "Prompts")
        End Function

        Private Shared Sub AppendExternalFiles(sb As StringBuilder, baseDir As String, scope As String, sources As List(Of String))
            Dim directNames = {
                Path.Combine(baseDir, scope & ".md"),
                Path.Combine(baseDir, scope & ".txt"),
                Path.Combine(baseDir, scope & ".json")
            }

            For Each file In directNames
                AppendExternalFile(sb, file, scope, sources)
            Next

            Dim scopedDir = Path.Combine(baseDir, scope)
            If Not Directory.Exists(scopedDir) Then Return

            For Each file In Directory.GetFiles(scopedDir, "*.*", SearchOption.TopDirectoryOnly).
                Where(Function(p)
                          Dim ext = Path.GetExtension(p).ToLowerInvariant()
                          Return ext = ".md" OrElse ext = ".txt" OrElse ext = ".json"
                      End Function).
                OrderBy(Function(p) p)
                AppendExternalFile(sb, file, scope, sources)
            Next
        End Sub

        Private Shared Sub AppendExternalFile(sb As StringBuilder, filePath As String, scope As String, sources As List(Of String))
            If Not File.Exists(filePath) Then Return

            Try
                Dim content = File.ReadAllText(filePath, Encoding.UTF8)
                If Path.GetExtension(filePath).Equals(".json", StringComparison.OrdinalIgnoreCase) Then
                    content = ExtractJsonPromptContent(content, scope)
                End If

                content = TrimToLimit(content, MaxSingleExternalFileChars)
                If String.IsNullOrWhiteSpace(content) Then Return

                sb.AppendLine($"# {scope}/{Path.GetFileName(filePath)}")
                sb.AppendLine(content.Trim())
                sb.AppendLine()
                sources.Add($"file:{Path.GetFileName(filePath)}")
            Catch ex As Exception
                Debug.WriteLine($"[PromptProfileService] 读取外接提示词失败 {filePath}: {ex.Message}")
            End Try
        End Sub

        Private Shared Function ExtractJsonPromptContent(json As String, scope As String) As String
            Try
                Dim jo = JObject.Parse(json)
                Dim enabledToken = jo("enabled")
                If enabledToken IsNot Nothing AndAlso Not enabledToken.Value(Of Boolean)() Then Return ""

                Dim app = If(jo("application")?.ToString(), jo("appType")?.ToString())
                If Not String.IsNullOrWhiteSpace(app) AndAlso
                   Not String.Equals(NormalizeScenario(app), scope, StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.Equals(app, "common", StringComparison.OrdinalIgnoreCase) Then
                    Return ""
                End If

                Return If(jo("content")?.ToString(), If(jo("prompt")?.ToString(), ""))
            Catch
                Return json
            End Try
        End Function

        Private Shared Sub AppendNamedBlock(sb As StringBuilder, title As String, content As String)
            If String.IsNullOrWhiteSpace(content) Then Return
            sb.AppendLine($"# {title}")
            sb.AppendLine(content.Trim())
            sb.AppendLine()
        End Sub

        Private Shared Function TrimToLimit(text As String, maxChars As Integer) As String
            If String.IsNullOrWhiteSpace(text) Then Return ""
            If text.Length <= maxChars Then Return text
            Return text.Substring(0, maxChars) & vbCrLf & "...(已截断，避免提示词过长)"
        End Function

        Private Shared Function NormalizeScenario(appType As String) As String
            Select Case If(appType, "").Trim().ToLowerInvariant()
                Case "word"
                    Return "word"
                Case "powerpoint", "ppt"
                    Return "ppt"
                Case "common"
                    Return "common"
                Case Else
                    Return "excel"
            End Select
        End Function
    End Class
End Namespace
