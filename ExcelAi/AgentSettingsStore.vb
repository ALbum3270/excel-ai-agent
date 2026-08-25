Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports ExcelAgent.Core

Public Class ExcelAgentSettings
    Public Property Platform As String = "OpenAI compatible"
    Public Property ApiUrl As String = ""
    Public Property ProtectedApiKey As String = ""
    Public Property ModelName As String = ""
    Public Property ReasoningMode As String = ReasoningRequestHelper.ReasoningDefault
    Public Property PromptName As String = ""
    Public Property PromptContent As String = ""
End Class

Public NotInheritable Class AgentSettingsStore
    Private Shared ReadOnly SettingsPath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        "settings.json")

    Private Sub New()
    End Sub

    Public Shared Function Load() As ExcelAgentSettings
        Dim settings As ExcelAgentSettings = Nothing
        Try
            If File.Exists(SettingsPath) Then
                settings = JsonConvert.DeserializeObject(Of ExcelAgentSettings)(File.ReadAllText(SettingsPath, Encoding.UTF8))
            End If
        Catch ex As Exception
            AppLogger.Warn("Settings", "Cannot load settings: " & ex.Message)
        End Try
        If settings Is Nothing Then settings = New ExcelAgentSettings()
        Apply(settings)
        Return settings
    End Function

    Public Shared Sub Save(settings As ExcelAgentSettings, plainApiKey As String)
        If settings Is Nothing Then Throw New ArgumentNullException(NameOf(settings))
        settings.ProtectedApiKey = Protect(plainApiKey)
        Dim directoryPath = Path.GetDirectoryName(SettingsPath)
        Directory.CreateDirectory(directoryPath)
        File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(settings, Formatting.Indented), New UTF8Encoding(False))
        Apply(settings)
    End Sub

    Public Shared Function GetApiKey(settings As ExcelAgentSettings) As String
        If settings Is Nothing Then Return ""
        Return Unprotect(settings.ProtectedApiKey)
    End Function

    Public Shared Function IsConfigured() As Boolean
        Return Not String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) AndAlso
               Not String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) AndAlso
               Not String.IsNullOrWhiteSpace(ConfigSettings.ModelName)
    End Function

    Private Shared Sub Apply(settings As ExcelAgentSettings)
        ConfigSettings.platform = settings.Platform
        ConfigSettings.ApiUrl = settings.ApiUrl
        ConfigSettings.ApiKey = GetApiKey(settings)
        ConfigSettings.ModelName = settings.ModelName
        ConfigSettings.ReasoningMode = ReasoningRequestHelper.NormalizeReasoningMode(settings.ReasoningMode)
        ConfigSettings.propmtName = settings.PromptName
        ConfigSettings.propmtContent = settings.PromptContent
    End Sub

    Private Shared Function Protect(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        Dim bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(value), Nothing, DataProtectionScope.CurrentUser)
        Return Convert.ToBase64String(bytes)
    End Function

    Private Shared Function Unprotect(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Try
            Dim bytes = ProtectedData.Unprotect(Convert.FromBase64String(value), Nothing, DataProtectionScope.CurrentUser)
            Return Encoding.UTF8.GetString(bytes)
        Catch ex As Exception
            AppLogger.Warn("Settings", "Cannot decrypt API key: " & ex.Message)
            Return ""
        End Try
    End Function
End Class
