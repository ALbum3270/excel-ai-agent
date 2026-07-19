' ShareRibbon\Services\Memory\SkillsIndexService.vb
' Incrementally indexes filesystem Skills into SQLite for fast request-time recall.

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Threading.Tasks

Public Class SkillsIndexService
    Implements ISkillSelector

    Private Shared ReadOnly _indexLock As New Object()
    Private Shared _isIndexing As Boolean = False
    Private Shared _lastKickoff As DateTime = DateTime.MinValue
    Private Shared ReadOnly _kickoffCooldown As TimeSpan = TimeSpan.FromMinutes(2)

    Public Shared Sub KickoffIndex(Optional forceRefresh As Boolean = False)
        SyncLock _indexLock
            If _isIndexing Then Return
            If Not forceRefresh AndAlso (DateTime.Now - _lastKickoff) < _kickoffCooldown Then Return
            _isIndexing = True
            _lastKickoff = DateTime.Now
        End SyncLock

        Task.Run(Async Function()
                     Try
                         Await IndexInstalledSkillsAsync(forceRefresh)
                     Catch ex As Exception
                         Debug.WriteLine($"[SkillsIndexService] 后台索引失败: {ex.Message}")
                     Finally
                         SyncLock _indexLock
                             _isIndexing = False
                         End SyncLock
                     End Try
                 End Function)
    End Sub

    Public Shared Async Function IndexInstalledSkillsAsync(Optional forceRefresh As Boolean = True) As Task(Of Integer)
        SkillsDirectoryService.EnsureDirectoryExists()
        Dim skills = SkillsDirectoryService.GetSkillsCatalog(forceRefresh)
        Dim activeNames As New List(Of String)()
        Dim updated As Integer = 0

        For Each skill In skills
            If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then Continue For

            Dim record = BuildSkillRecord(skill)
            Dim existing = AgentMemoryRepository.GetSkillRegistryByName(record.SkillName)
            Dim contentChanged = HasSkillContentChanged(skill, existing)
            If Not NeedsReindex(record, existing, contentChanged) Then
                activeNames.Add(record.SkillName)
                Continue For
            End If

            If existing IsNot Nothing AndAlso Not contentChanged Then
                record.EmbeddingJson = existing.EmbeddingJson
            End If

            If EmbeddingService.IsEmbeddingAvailable() AndAlso (contentChanged OrElse String.IsNullOrWhiteSpace(record.EmbeddingJson)) Then
                Dim indexText = BuildSkillIndexText(skill)
                If indexText.Length > 4000 Then indexText = indexText.Substring(0, 4000)
                Dim embedding = Await EmbeddingService.GetEmbeddingAsync(indexText)
                If embedding IsNot Nothing AndAlso embedding.Length > 0 Then
                    record.EmbeddingJson = EmbeddingService.SerializeVector(embedding)
                End If
            End If

            AgentMemoryRepository.UpsertSkillRegistry(record)
            activeNames.Add(record.SkillName)
            updated += 1
        Next

        AgentMemoryRepository.DisableSkillsNotIn(activeNames)
        Debug.WriteLine($"[SkillsIndexService] Skills discovered={skills.Count}, active={activeNames.Count}, updated={updated}")
        Return activeNames.Count
    End Function

    Public Shared Function SelectSkillDefinitions(query As String, intentType As String, appType As String, Optional topN As Integer = 3) As List(Of SkillFileDefinition)
        KickoffIndex()

        Dim selected As New List(Of SkillFileDefinition)()
        Dim directMatches = SkillsService.MatchSkills(If(query, ""), Math.Max(topN * 3, 6))

        ' Filesystem metadata is the freshest authority. Strong semantic matches must be
        ' considered before a possibly stale database index, especially immediately after
        ' installing a new specialized Skill. This remains metadata-driven and app-scoped.
        For Each match In directMatches.Where(Function(item) item.MatchScore >= 10).
                                         OrderByDescending(Function(item) item.MatchScore)
            AddSelectedSkill(selected, SkillsDirectoryService.LoadSkillDetail(match.Skill), appType, topN)
            If selected.Count >= topN Then Exit For
        Next

        Dim records = AgentMemoryRepository.RetrieveSkillRegistry(query, intentType, appType, topN)

        If records IsNot Nothing Then
            For Each record In records
                Dim skill = SkillsDirectoryService.GetSkillByNameOrPath(record.SkillName, record.FilePath)
                AddSelectedSkill(selected, skill, appType, topN)
            Next
        End If

        ' 首次启动时数据库索引仍在后台构建；文件目录是权威来源，必须同步兜底。
        ' 同时补足数据库召回不足的结果，避免过期索引遮蔽新安装 Skill。
        If selected.Count < Math.Max(1, topN) Then
            For Each match In directMatches
                AddSelectedSkill(selected, SkillsDirectoryService.LoadSkillDetail(match.Skill), appType, topN)
                If selected.Count >= topN Then Exit For
            Next
        End If

        ' 开放需求可能没有任何可词法穷举的关键词。专用 Skill 未命中时，使用宿主显式
        ' 声明的 baseline Skill，让 Agent 在其工具边界内理解新任务。
        If selected.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(query) Then
            Dim baseline = SkillsDirectoryService.GetSkillsCatalog().
                FirstOrDefault(Function(skill) SupportsApplication(skill, appType) AndAlso IsDefaultForApplication(skill))
            AddSelectedSkill(selected, SkillsDirectoryService.LoadSkillDetail(baseline), appType, topN)
            If baseline IsNot Nothing Then
                Debug.WriteLine($"[SkillsIndexService] 使用宿主 baseline Skill: {baseline.Name}")
            End If
        End If

        Debug.WriteLine($"[SkillsIndexService] Skill matches={selected.Count}, app={appType}, query={If(query, "").Substring(0, Math.Min(If(query, "").Length, 80))}")

        Return selected
    End Function

    Private Shared Sub AddSelectedSkill(selected As List(Of SkillFileDefinition),
                                        skill As SkillFileDefinition,
                                        appType As String,
                                        topN As Integer)
        If selected Is Nothing OrElse skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then Return
        If selected.Count >= Math.Max(1, topN) Then Return
        If Not SupportsApplication(skill, appType) Then Return
        If selected.Any(Function(s) String.Equals(s.Name, skill.Name, StringComparison.OrdinalIgnoreCase)) Then Return
        selected.Add(skill)
    End Sub

    Private Shared Function SupportsApplication(skill As SkillFileDefinition, appType As String) As Boolean
        If skill Is Nothing OrElse String.IsNullOrWhiteSpace(appType) Then Return True
        Dim scope = If(skill.Application, "")
        If String.IsNullOrWhiteSpace(scope) Then scope = InferAppScope(skill)
        ' 完全没有宿主信号的 Skill 视为 common；能够推断宿主时必须隔离，避免把
        ' Excel 的“图表生成/数据分析”等旧 Skill 注入 PowerPoint。
        If String.IsNullOrWhiteSpace(scope) Then Return True
        Dim requested = NormalizeAppType(appType)
        Return scope.Split({","c, ";"c, "|"c}, StringSplitOptions.RemoveEmptyEntries).
            Any(Function(value) NormalizeAppType(value) = requested)
    End Function

    Private Shared Function NormalizeAppType(value As String) As String
        Dim normalized = If(value, "").Trim().ToLowerInvariant()
        If normalized = "ppt" OrElse normalized = "power point" Then Return "powerpoint"
        If normalized = "xls" OrElse normalized = "xlsx" Then Return "excel"
        If normalized = "doc" OrElse normalized = "docx" Then Return "word"
        Return normalized
    End Function

    Private Shared Function IsDefaultForApplication(skill As SkillFileDefinition) As Boolean
        If skill?.Metadata Is Nothing Then Return False
        Dim raw As Object = Nothing
        If Not skill.Metadata.TryGetValue("default_for_application", raw) Then Return False
        Dim enabled As Boolean
        Return Boolean.TryParse(If(raw, "").ToString(), enabled) AndAlso enabled
    End Function

    Public Function SelectSkills(query As String, intentType As String, appType As String, topN As Integer) As List(Of SkillRegistryRecord) Implements ISkillSelector.SelectSkills
        Return AgentMemoryRepository.RetrieveSkillRegistry(query, intentType, appType, topN)
    End Function

    Private Shared Function BuildSkillRecord(skill As SkillFileDefinition) As SkillRegistryRecord
        Return New SkillRegistryRecord With {
            .SkillName = skill.Name,
            .FilePath = If(skill.FilePath, ""),
            .AppScope = InferAppScope(skill),
            .IntentTypes = InferIntentTypes(skill),
            .TriggerKeywords = BuildTriggerKeywords(skill),
            .Description = If(String.IsNullOrWhiteSpace(skill.Description), FirstContentLine(skill.Content), skill.Description),
            .Enabled = True
        }
    End Function

    Private Shared Function BuildSkillIndexText(skill As SkillFileDefinition) As String
        Dim parts As New List(Of String)()
        parts.Add(skill.Name)
        parts.Add(skill.Description)
        parts.Add(skill.Application)
        parts.Add(skill.Compatibility)
        If skill.Tags IsNot Nothing Then parts.Add(String.Join(", ", skill.Tags))
        If skill.AllowedTools IsNot Nothing Then parts.Add(String.Join(", ", skill.AllowedTools))
        Return String.Join(vbCrLf, parts.Where(Function(p) Not String.IsNullOrWhiteSpace(p)))
    End Function

    Private Shared Function BuildTriggerKeywords(skill As SkillFileDefinition) As String
        Dim terms As New List(Of String)()
        AddTerm(terms, skill.Name)
        AddTerm(terms, skill.Application)
        AddTerm(terms, skill.Compatibility)

        If skill.Tags IsNot Nothing Then
            For Each tag In skill.Tags
                AddTerm(terms, tag)
            Next
        End If

        If skill.Metadata IsNot Nothing Then
            For Each key In New String() {"keywords", "triggers", "trigger_keywords", "intent", "intent_types"}
                If skill.Metadata.ContainsKey(key) Then AddTerm(terms, skill.Metadata(key)?.ToString())
            Next
        End If

        Dim desc = If(skill.Description, "")
        For Each token In desc.Split({" "c, ","c, "，"c, ";"c, "；"c, "/"c, "\"c, "|"c}, StringSplitOptions.RemoveEmptyEntries)
            If token.Trim().Length >= 2 AndAlso token.Trim().Length <= 32 Then AddTerm(terms, token.Trim())
        Next

        Return String.Join(", ", terms.Distinct(StringComparer.OrdinalIgnoreCase))
    End Function

    Private Shared Function InferAppScope(skill As SkillFileDefinition) As String
        Dim text = String.Join(" ", New String() {
            If(skill.Name, ""),
            If(skill.Description, ""),
            If(skill.Application, ""),
            If(skill.Compatibility, ""),
            If(skill.Context, ""),
            If(skill.Agent, ""),
            If(skill.FilePath, ""),
            If(skill.Tags Is Nothing, "", String.Join(" ", skill.Tags))
        }).ToLowerInvariant()

        Dim apps As New List(Of String)()
        If text.Contains("word") OrElse text.Contains("docx") OrElse text.Contains("文档") Then apps.Add("Word")
        If text.Contains("excel") OrElse text.Contains("xlsx") OrElse text.Contains("表格") Then apps.Add("Excel")
        If text.Contains("powerpoint") OrElse text.Contains("ppt") OrElse text.Contains("幻灯片") Then apps.Add("PowerPoint")
        Return String.Join(",", apps.Distinct(StringComparer.OrdinalIgnoreCase))
    End Function

    Private Shared Function InferIntentTypes(skill As SkillFileDefinition) As String
        Dim intents As New List(Of String)()
        If skill.Metadata IsNot Nothing Then
            For Each key In New String() {"intent", "intents", "intent_type", "intent_types"}
                If skill.Metadata.ContainsKey(key) Then AddTerm(intents, skill.Metadata(key)?.ToString())
            Next
        End If
        If Not String.IsNullOrWhiteSpace(skill.Context) Then AddTerm(intents, skill.Context)
        Return String.Join(", ", intents.Distinct(StringComparer.OrdinalIgnoreCase))
    End Function

    Private Shared Function NeedsReindex(record As SkillRegistryRecord, existing As SkillRegistryRecord, contentChanged As Boolean) As Boolean
        If existing Is Nothing Then Return True
        If Not existing.Enabled Then Return True
        If contentChanged Then Return True
        If Not String.Equals(NormalizePath(existing.FilePath), NormalizePath(record.FilePath), StringComparison.OrdinalIgnoreCase) Then Return True
        If Not String.Equals(If(existing.AppScope, ""), If(record.AppScope, ""), StringComparison.Ordinal) Then Return True
        If Not String.Equals(If(existing.IntentTypes, ""), If(record.IntentTypes, ""), StringComparison.Ordinal) Then Return True
        If Not String.Equals(If(existing.TriggerKeywords, ""), If(record.TriggerKeywords, ""), StringComparison.Ordinal) Then Return True
        If Not String.Equals(If(existing.Description, ""), If(record.Description, ""), StringComparison.Ordinal) Then Return True
        If EmbeddingService.IsEmbeddingAvailable() AndAlso String.IsNullOrWhiteSpace(existing.EmbeddingJson) Then Return True
        Return False
    End Function

    Private Shared Function HasSkillContentChanged(skill As SkillFileDefinition, existing As SkillRegistryRecord) As Boolean
        If existing Is Nothing OrElse String.IsNullOrWhiteSpace(existing.LastIndexedAt) Then Return True

        Dim indexedAt As DateTime
        If Not DateTime.TryParse(existing.LastIndexedAt, indexedAt) Then Return True

        Dim lastWrite = GetSkillLastWriteTime(skill)
        If Not lastWrite.HasValue Then Return False
        Return lastWrite.Value > indexedAt.AddSeconds(1)
    End Function

    Private Shared Function GetSkillLastWriteTime(skill As SkillFileDefinition) As DateTime?
        If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.FilePath) Then Return Nothing

        Try
            If Directory.Exists(skill.FilePath) Then
                Dim latest = Directory.GetLastWriteTime(skill.FilePath)
                For Each filePath In Directory.GetFiles(skill.FilePath, "*.*", SearchOption.AllDirectories)
                    Dim fileWrite = File.GetLastWriteTime(filePath)
                    If fileWrite > latest Then latest = fileWrite
                Next
                Return latest
            End If

            If File.Exists(skill.FilePath) Then
                Return File.GetLastWriteTime(skill.FilePath)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SkillsIndexService] 读取 Skill 更新时间失败: {ex.Message}")
        End Try

        Return Nothing
    End Function

    Private Shared Sub AddTerm(terms As List(Of String), value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        For Each part In value.Split({","c, "，"c, ";"c, "；"c, "|"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim term = part.Trim()
            If term.Length > 0 AndAlso term.Length <= 80 Then terms.Add(term)
        Next
    End Sub

    Private Shared Function FirstContentLine(content As String) As String
        If String.IsNullOrWhiteSpace(content) Then Return ""
        For Each line In content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim trimmed = line.Trim()
            If trimmed.Length > 0 AndAlso Not trimmed.StartsWith("#") Then Return trimmed
        Next
        Return ""
    End Function

    Private Shared Function NormalizePath(pathValue As String) As String
        If String.IsNullOrWhiteSpace(pathValue) Then Return ""
        Try
            Return Path.GetFullPath(pathValue.Trim())
        Catch
            Return pathValue.Trim()
        End Try
    End Function
End Class
