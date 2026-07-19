Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Namespace Security

    ''' <summary>
    ''' 安全配置管理器 - 加密存储敏感信息（如 API Key）
    ''' </summary>
    Public Class SecureConfigManager

        Private Shared ReadOnly Entropy As Byte() = Encoding.UTF8.GetBytes("AiHelper-Office-2024")

        ''' <summary>
        ''' 加密字符串
        ''' </summary>
        Public Shared Function Encrypt(plainText As String) As String
            If String.IsNullOrEmpty(plainText) Then
                Return String.Empty
            End If

            Try
                Dim plainBytes As Byte() = Encoding.UTF8.GetBytes(plainText)
                Dim encryptedBytes As Byte() = ProtectedData.Protect(plainBytes, Entropy, DataProtectionScope.CurrentUser)
                Return Convert.ToBase64String(encryptedBytes)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigManager] 加密失败: {0}", ex.Message))
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' 解密字符串
        ''' </summary>
        Public Shared Function Decrypt(encryptedText As String) As String
            If String.IsNullOrEmpty(encryptedText) Then
                Return String.Empty
            End If

            Try
                Dim encryptedBytes As Byte() = Convert.FromBase64String(encryptedText)
                Dim plainBytes As Byte() = ProtectedData.Unprotect(encryptedBytes, Entropy, DataProtectionScope.CurrentUser)
                Return Encoding.UTF8.GetString(plainBytes)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigManager] 解密失败: {0}", ex.Message))
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' 保存加密的 API Key
        ''' </summary>
        Public Shared Function SaveApiKey(apiKey As String, provider As String) As Boolean
            Try
                Dim encrypted As String = Encrypt(apiKey)
                If String.IsNullOrEmpty(encrypted) Then
                    Return False
                End If

                ' 保存到配置文件
                Dim configPath As String = GetConfigFilePath()
                Dim config As New Dictionary(Of String, String)

                ' 读取现有配置
                If File.Exists(configPath) Then
                    Dim lines As String() = File.ReadAllLines(configPath)
                    For Each line In lines
                        If line.Contains("=") Then
                            Dim parts As String() = line.Split(New Char() {"="c}, 2)
                            If parts.Length = 2 Then
                                config(parts(0).Trim()) = parts(1).Trim()
                            End If
                        End If
                    Next
                End If

                ' 更新 API Key
                Dim key As String = String.Format("ApiKey_{0}", provider)
                config(key) = encrypted

                ' 保存配置
                Dim configLines As New List(Of String)
                For Each kvp In config
                    configLines.Add(String.Format("{0}={1}", kvp.Key, kvp.Value))
                Next

                File.WriteAllLines(configPath, configLines.ToArray())

                System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigManager] API Key 已保存: {0}", provider))
                Return True

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigManager] 保存 API Key 失败: {0}", ex.Message))
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 读取加密的 API Key
        ''' </summary>
        Public Shared Function LoadApiKey(provider As String) As String
            Try
                Dim configPath As String = GetConfigFilePath()
                If Not File.Exists(configPath) Then
                    Return String.Empty
                End If

                Dim key As String = String.Format("ApiKey_{0}", provider)
                Dim lines As String() = File.ReadAllLines(configPath)

                For Each line In lines
                    If line.StartsWith(key & "=") Then
                        Dim encrypted As String = line.Substring((key & "=").Length).Trim()
                        Return Decrypt(encrypted)
                    End If
                Next

                Return String.Empty

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigManager] 读取 API Key 失败: {0}", ex.Message))
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' 验证 API Key 格式
        ''' </summary>
        Public Shared Function ValidateApiKey(apiKey As String, provider As String) As Boolean
            If String.IsNullOrWhiteSpace(apiKey) Then
                Return False
            End If

            ' 根据不同提供商验证格式
            Select Case provider.ToLower()
                Case "claude", "anthropic"
                    ' Claude API Key 格式: sk-ant-...
                    Return apiKey.StartsWith("sk-ant-") AndAlso apiKey.Length > 20

                Case "openai"
                    ' OpenAI API Key 格式: sk-...
                    Return apiKey.StartsWith("sk-") AndAlso apiKey.Length > 20

                Case "deepseek"
                    ' DeepSeek API Key 格式
                    Return apiKey.Length > 20

                Case Else
                    ' 默认验证：至少 20 字符
                    Return apiKey.Length >= 20
            End Select
        End Function

        ''' <summary>
        ''' 获取配置文件路径
        ''' </summary>
        Private Shared Function GetConfigFilePath() As String
            Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim configDir As String = Path.Combine(appDataPath, "AiHelper")

            If Not Directory.Exists(configDir) Then
                Directory.CreateDirectory(configDir)
            End If

            Return Path.Combine(configDir, "secure.config")
        End Function

        ''' <summary>
        ''' 获取 API Key 的掩码显示（用于 UI）
        ''' </summary>
        Public Shared Function GetMaskedApiKey(provider As String) As String
            Dim apiKey As String = LoadApiKey(provider)
            If String.IsNullOrEmpty(apiKey) Then
                Return "(未配置)"
            End If

            If apiKey.Length <= 8 Then
                Return "****"
            End If

            ' 显示前4个和后4个字符
            Dim prefix As String = apiKey.Substring(0, 4)
            Dim suffix As String = apiKey.Substring(apiKey.Length - 4)
            Return String.Format("{0}...{1}", prefix, suffix)
        End Function

        ''' <summary>
        ''' 数据隐私设置
        ''' </summary>
        Public Class PrivacySettings

            ''' <summary>
            ''' 是否允许发送文档内容到 AI
            ''' </summary>
            Public Shared Property AllowSendDocumentContent As Boolean
                Get
                    Return GetSetting("AllowSendDocumentContent", "true") = "true"
                End Get
                Set(value As Boolean)
                    SaveSetting("AllowSendDocumentContent", If(value, "true", "false"))
                End Set
            End Property

            ''' <summary>
            ''' 是否允许记录对话历史
            ''' </summary>
            Public Shared Property AllowSaveConversationHistory As Boolean
                Get
                    Return GetSetting("AllowSaveConversationHistory", "true") = "true"
                End Get
                Set(value As Boolean)
                    SaveSetting("AllowSaveConversationHistory", If(value, "true", "false"))
                End Set
            End Property

            ''' <summary>
            ''' 是否允许收集使用统计
            ''' </summary>
            Public Shared Property AllowUsageStatistics As Boolean
                Get
                    Return GetSetting("AllowUsageStatistics", "true") = "true"
                End Get
                Set(value As Boolean)
                    SaveSetting("AllowUsageStatistics", If(value, "true", "false"))
                End Set
            End Property

            ''' <summary>
            ''' 用户是否已同意隐私政策
            ''' </summary>
            Public Shared Property PrivacyPolicyAccepted As Boolean
                Get
                    Return GetSetting("PrivacyPolicyAccepted", "false") = "true"
                End Get
                Set(value As Boolean)
                    SaveSetting("PrivacyPolicyAccepted", If(value, "true", "false"))
                End Set
            End Property

            Private Shared Function GetSetting(key As String, defaultValue As String) As String
                Try
                    Dim configPath As String = GetConfigFilePath()
                    If Not File.Exists(configPath) Then
                        Return defaultValue
                    End If

                    Dim lines As String() = File.ReadAllLines(configPath)
                    For Each line In lines
                        If line.StartsWith(key & "=") Then
                            Return line.Substring((key & "=").Length).Trim()
                        End If
                    Next

                    Return defaultValue

                Catch ex As Exception
                    Return defaultValue
                End Try
            End Function

            Private Shared Sub SaveSetting(key As String, value As String)
                Try
                    Dim configPath As String = GetConfigFilePath()
                    Dim config As New Dictionary(Of String, String)

                    ' 读取现有配置
                    If File.Exists(configPath) Then
                        Dim lines As String() = File.ReadAllLines(configPath)
                        For Each line In lines
                            If line.Contains("=") Then
                                Dim parts As String() = line.Split(New Char() {"="c}, 2)
                                If parts.Length = 2 Then
                                    config(parts(0).Trim()) = parts(1).Trim()
                                End If
                            End If
                        Next
                    End If

                    ' 更新设置
                    config(key) = value

                    ' 保存配置
                    Dim configLines As New List(Of String)
                    For Each kvp In config
                        configLines.Add(String.Format("{0}={1}", kvp.Key, kvp.Value))
                    Next

                    File.WriteAllLines(configPath, configLines.ToArray())

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine(String.Format("[PrivacySettings] 保存设置失败: {0}", ex.Message))
                End Try
            End Sub

        End Class

    End Class

End Namespace
