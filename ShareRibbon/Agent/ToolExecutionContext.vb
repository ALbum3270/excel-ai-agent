Imports System.Collections.Generic
Imports System.Linq

Namespace Agent

    ''' <summary>
    ''' Carries per-run tool visibility and execution policy.
    ''' H1: appType + selected Skill allowed-tools are enforced at execution time.
    ''' </summary>
    Public Class ToolExecutionContext
        Public Property AppType As String
        Public Property RunId As String
        Public Property CorrelationId As String
        Public Property PrimarySkillName As String
        Public Property AllowedTools As HashSet(Of String)
        Public Property EnforceAllowedTools As Boolean = True

        Public Shared Function FromSession(session As AgentSession, skill As AgentSkill) As ToolExecutionContext
            Dim context As New ToolExecutionContext With {
                .AppType = If(session?.AppType, ""),
                .RunId = If(session?.Id, ""),
                .CorrelationId = If(session?.Id, ""),
                .PrimarySkillName = If(skill?.Name, ""),
                .AllowedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .EnforceAllowedTools = False
            }

            If skill IsNot Nothing AndAlso skill.RequiredTools IsNot Nothing Then
                For Each toolId In skill.RequiredTools
                    If Not String.IsNullOrWhiteSpace(toolId) Then
                        context.AllowedTools.Add(toolId.Trim())
                    End If
                Next

                context.EnforceAllowedTools = context.AllowedTools.Count > 0
            End If

            Return context
        End Function

        Public Function HasPrimarySkillGate() As Boolean
            Return EnforceAllowedTools AndAlso
                   AllowedTools IsNot Nothing AndAlso
                   AllowedTools.Count > 0
        End Function

        Public Function IsToolAllowed(toolId As String) As Boolean
            If Not HasPrimarySkillGate() Then Return True
            If String.IsNullOrWhiteSpace(toolId) Then Return False
            Return AllowedTools.Contains(toolId.Trim())
        End Function

        Public Function AllowedToolsText() As String
            If AllowedTools Is Nothing OrElse AllowedTools.Count = 0 Then Return ""
            Return String.Join(", ", AllowedTools.OrderBy(Function(id) id))
        End Function
    End Class

End Namespace
