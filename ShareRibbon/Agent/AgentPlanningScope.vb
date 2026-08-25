Imports System.Collections.Generic
Imports System.Linq

Namespace Agent

    ''' <summary>
    ''' Builds a compact planner view without changing the execution authorization boundary.
    ''' RequiredTools scopes what the model needs to see for this plan; the primary Skill still
    ''' authorizes its complete tool set during execution and repair.
    ''' </summary>
    Public NotInheritable Class AgentPlanningScope
        Private Sub New()
        End Sub

        Public Shared Function SelectTools(visibleTools As IEnumerable(Of ToolDescriptor),
                                           spec As AgentTaskSpec) As List(Of ToolDescriptor)
            Dim visible = If(visibleTools, Enumerable.Empty(Of ToolDescriptor)()).
                Where(Function(tool) tool IsNot Nothing).
                ToList()
            If spec?.RequiredTools Is Nothing OrElse spec.RequiredTools.Count = 0 Then Return visible

            Dim required = New HashSet(Of String)(
                spec.RequiredTools.Where(Function(toolId) Not String.IsNullOrWhiteSpace(toolId)),
                StringComparer.OrdinalIgnoreCase)
            Dim scoped = visible.Where(Function(tool) required.Contains(tool.Id)).ToList()

            ' A stale task hint must never produce an empty capability view. Execution remains
            ' governed by the Skill either way; falling back here preserves open-ended tasks.
            Return If(scoped.Count > 0, scoped, visible)
        End Function
    End Class

End Namespace
