Imports System.IO
Imports System.Text

Public Class SkillFileDefinition
    Public Property Name As String
    Public Property Description As String
    Public Property Content As String
    Public Property AllowedTools As New List(Of String)()
    Public Property FilePath As String
    Public Property IsContentLoaded As Boolean
    Public Property Tags As New List(Of String)()
    Public Property Application As String = "Excel"
End Class

''' <summary>
''' Filesystem loader for the single standalone Excel Skill format. It reads front matter for
''' discovery and loads the body only after the Skill is selected.
''' </summary>
Public NotInheritable Class SkillsDirectoryService
    Private Sub New()
    End Sub

    Public Shared Function GetSkillsDirectory() As String
        Return ResolveSkillsDirectory()
    End Function

    Public Shared Function GetSkillsDirectories() As List(Of String)
        Dim directoryPath = ResolveSkillsDirectory()
        Return If(String.IsNullOrWhiteSpace(directoryPath), New List(Of String)(), New List(Of String) From {directoryPath})
    End Function

    Public Shared Function GetAllSkills(Optional forceRefresh As Boolean = False) As List(Of SkillFileDefinition)
        Return GetSkillsCatalog(forceRefresh).
            Select(Function(skill) LoadSkillDetail(skill)).
            Where(Function(skill) skill IsNot Nothing).ToList()
    End Function

    Public Shared Function GetSkillsCatalog(Optional forceRefresh As Boolean = False) As List(Of SkillFileDefinition)
        Dim result As New List(Of SkillFileDefinition)()
        Dim root = ResolveSkillsDirectory()
        If String.IsNullOrWhiteSpace(root) OrElse Not Directory.Exists(root) Then Return result

        For Each filePath In Directory.GetFiles(root, "SKILL.md", SearchOption.AllDirectories).
            OrderBy(Function(path) path, StringComparer.OrdinalIgnoreCase)
            Dim skill = ReadSkill(filePath, False)
            If skill IsNot Nothing AndAlso
               String.Equals(If(skill.Application, "Excel"), "Excel", StringComparison.OrdinalIgnoreCase) Then result.Add(skill)
        Next
        Return result
    End Function

    Public Shared Function LoadSkillDetail(skill As SkillFileDefinition) As SkillFileDefinition
        If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.FilePath) Then Return Nothing
        Return ReadSkill(skill.FilePath, True)
    End Function

    Public Shared Function GetSkillByNameOrPath(skillName As String, filePath As String) As SkillFileDefinition
        If Not String.IsNullOrWhiteSpace(filePath) AndAlso File.Exists(filePath) Then Return ReadSkill(filePath, True)
        Return GetSkillsCatalog(True).
            Where(Function(skill) String.Equals(skill.Name, skillName, StringComparison.OrdinalIgnoreCase)).
            Select(Function(skill) LoadSkillDetail(skill)).FirstOrDefault()
    End Function

    Private Shared Function ResolveSkillsDirectory() As String
        For Each startPath In {Path.GetDirectoryName(GetType(SkillsDirectoryService).Assembly.Location), AppDomain.CurrentDomain.BaseDirectory}
            Dim current = startPath
            While Not String.IsNullOrWhiteSpace(current)
                For Each relativePath In {"Skills", Path.Combine("ExcelAgent.Core", "Skills")}
                    Dim candidate = Path.Combine(current, relativePath)
                    If Directory.Exists(candidate) Then Return candidate
                Next
                current = Path.GetDirectoryName(current)
            End While
        Next
        Return ""
    End Function

    Private Shared Function ReadSkill(filePath As String, includeBody As Boolean) As SkillFileDefinition
        Try
            Dim text = File.ReadAllText(filePath, Encoding.UTF8)
            Dim metadata As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim body = text
            If text.StartsWith("---", StringComparison.Ordinal) Then
                Dim normalized = text.Replace(vbCrLf, vbLf)
                Dim closing = normalized.IndexOf(vbLf & "---" & vbLf, 3, StringComparison.Ordinal)
                If closing > 0 Then
                    For Each line In normalized.Substring(3, closing - 3).Split(ChrW(10))
                        Dim separator = line.IndexOf(":"c)
                        If separator <= 0 Then Continue For
                        metadata(line.Substring(0, separator).Trim()) = line.Substring(separator + 1).Trim().Trim(""""c, "'"c)
                    Next
                    body = normalized.Substring(closing + 5).Trim()
                End If
            End If

            Dim skill As New SkillFileDefinition With {
                .FilePath = filePath,
                .Name = GetMetadata(metadata, "name", Path.GetFileName(Path.GetDirectoryName(filePath))),
                .Description = GetMetadata(metadata, "description", "Excel table automation"),
                .Application = GetMetadata(metadata, "application", "Excel"),
                .IsContentLoaded = includeBody,
                .Content = If(includeBody, body, "")
            }
            skill.AllowedTools = SplitList(GetMetadata(metadata, "allowed-tools", ""))
            skill.Tags = SplitList(GetMetadata(metadata, "tags", ""))
            Return skill
        Catch ex As Exception
            AppLogger.Warn("Skills", "Cannot load Skill: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Function GetMetadata(metadata As Dictionary(Of String, String), key As String, fallback As String) As String
        Dim value = ""
        If metadata.TryGetValue(key, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then Return value
        Return fallback
    End Function

    Private Shared Function SplitList(value As String) As List(Of String)
        Return If(value, "").Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
            Select(Function(item) item.Trim()).Where(Function(item) item.Length > 0).
            Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function
End Class
