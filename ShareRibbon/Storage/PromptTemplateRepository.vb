' ShareRibbon\Storage\PromptTemplateRepository.vb
' prompt_template 表 CRUD 与按场景/Skills 加载

Imports System.Data.SQLite

''' <summary>
''' prompt_template 表访问
''' </summary>
Public Class PromptTemplateRepository

    ' 系统提示词内存缓存，避免每次发消息都查库
    Private Shared _systemPromptCache As New Dictionary(Of String, String)()
    Private Shared _cacheLock As New Object()

    ''' <summary>
    ''' 按 scenario 获取系统提示词（is_skill=0），带内存缓存
    ''' </summary>
    Public Shared Function GetSystemPrompt(scenario As String) As String
        Dim scenarioNorm = If(String.IsNullOrEmpty(scenario), "common", scenario.ToLowerInvariant())

        SyncLock _cacheLock
            If _systemPromptCache.ContainsKey(scenarioNorm) Then
                Return _systemPromptCache(scenarioNorm)
            End If
        End SyncLock

        OfficeAiDatabase.EnsureInitialized()
        Dim result As String = ""
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Dim sql = "SELECT content FROM prompt_template WHERE scenario=@s AND is_skill=0 ORDER BY sort, id LIMIT 1"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@s", scenarioNorm)
                Dim obj = cmd.ExecuteScalar()
                result = If(obj Is Nothing OrElse obj Is DBNull.Value, "", obj.ToString())
            End Using
        End Using

        SyncLock _cacheLock
            _systemPromptCache(scenarioNorm) = result
        End SyncLock

        Return result
    End Function

    ''' <summary>
    ''' 变量替换：{{变量名}} 替换为 vars 字典中的值
    ''' </summary>
    Public Shared Function ReplaceVariables(template As String, vars As Dictionary(Of String, String)) As String
        If String.IsNullOrEmpty(template) Then Return ""
        If vars Is Nothing OrElse vars.Count = 0 Then Return template

        Dim result = template
        For Each kv In vars
            result = result.Replace("{{" & kv.Key & "}}", If(kv.Value, ""))
        Next
        Return result
    End Function
End Class
