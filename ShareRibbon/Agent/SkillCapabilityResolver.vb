Imports System.Collections.Generic
Imports System.Linq

Namespace Agent

    ''' <summary>
    ''' Keeps one primary Skill as the execution boundary while preventing a specialised
    ''' workflow Skill from hiding tools that the already-built task contract requires.
    ''' When that happens, the application's declared baseline Skill becomes primary and
    ''' the specialised matches remain available as secondary trace metadata.
    ''' </summary>
    Public NotInheritable Class SkillCapabilityResolver
        Private Sub New()
        End Sub

        Public Shared Function ResolvePrimarySkill(selectedSkills As List(Of SkillFileDefinition),
                                                   requiredTools As IEnumerable(Of String),
                                                   appType As String) As List(Of SkillFileDefinition)
            Dim selected = If(selectedSkills, New List(Of SkillFileDefinition)()).
                Where(Function(skill) skill IsNot Nothing).
                ToList()
            Dim required = New HashSet(Of String)(
                If(requiredTools, Enumerable.Empty(Of String)()).
                    Where(Function(toolId) Not String.IsNullOrWhiteSpace(toolId)).
                    Select(Function(toolId) toolId.Trim()),
                StringComparer.OrdinalIgnoreCase)
            If required.Count = 0 OrElse (selected.Count > 0 AndAlso Covers(selected(0), required)) Then
                Return selected
            End If

            Dim baseline = SkillsDirectoryService.GetSkillsCatalog().
                Where(Function(skill) SupportsApplication(skill, appType) AndAlso IsDefaultForApplication(skill)).
                Select(Function(skill) SkillsDirectoryService.LoadSkillDetail(skill)).
                FirstOrDefault(Function(skill) Covers(skill, required))
            If baseline Is Nothing Then Return selected

            Dim resolved As New List(Of SkillFileDefinition) From {baseline}
            For Each skill In selected
                If Not String.Equals(skill.Name, baseline.Name, StringComparison.OrdinalIgnoreCase) Then
                    resolved.Add(skill)
                End If
            Next
            AppLogger.Info("SkillCapabilityResolver",
                           $"Primary Skill promoted to baseline name={baseline.Name} requiredTools={String.Join(",", required.OrderBy(Function(id) id))}")
            Return resolved
        End Function

        Private Shared Function Covers(skill As SkillFileDefinition,
                                       required As HashSet(Of String)) As Boolean
            If skill?.AllowedTools Is Nothing Then Return False
            Dim allowed = New HashSet(Of String)(skill.AllowedTools, StringComparer.OrdinalIgnoreCase)
            Return required.All(Function(toolId) allowed.Contains(toolId))
        End Function

        Private Shared Function IsDefaultForApplication(skill As SkillFileDefinition) As Boolean
            If skill?.Metadata Is Nothing Then Return False
            Dim raw As Object = Nothing
            If Not skill.Metadata.TryGetValue("default_for_application", raw) Then Return False
            Dim enabled As Boolean
            Return Boolean.TryParse(If(raw, "").ToString(), enabled) AndAlso enabled
        End Function

        Private Shared Function SupportsApplication(skill As SkillFileDefinition,
                                                     appType As String) As Boolean
            If skill Is Nothing OrElse String.IsNullOrWhiteSpace(appType) Then Return True
            Dim requested = NormalizeAppType(appType)
            Dim scope = If(skill.Application, "")
            If String.IsNullOrWhiteSpace(scope) Then Return True
            Return scope.Split({","c, ";"c, "|"c}, StringSplitOptions.RemoveEmptyEntries).
                Any(Function(value) NormalizeAppType(value) = requested)
        End Function

        Private Shared Function NormalizeAppType(appType As String) As String
            Dim value = If(appType, "").Trim().ToLowerInvariant()
            Select Case value
                Case "ppt", "powerpoint", "power point"
                    Return "powerpoint"
                Case "xls", "xlsx", "excel"
                    Return "excel"
                Case "doc", "docx", "word"
                    Return "word"
                Case Else
                    Return value
            End Select
        End Function
    End Class

End Namespace
