Imports System.Collections.Generic
Imports System.Linq

Namespace Agent

    ''' <summary>
    ''' Deterministic policy check for capabilities explicitly required by the user.
    ''' It does not prescribe a plan, tool sequence, or implementation path; those choices
    ''' remain inside the adaptive ReAct loop.
    ''' </summary>
    Public NotInheritable Class AgentExecutionContract
        Private Sub New()
        End Sub

        Public Shared Function ValidatePlan(spec As AgentTaskSpec,
                                            plan As ExecutionPlan) As String
            ' A high-level plan is soft UI guidance. Capability policy is enforced only
            ' against observed execution, never against a predicted workflow.
            Return ""
        End Function

        Public Shared Function ValidateOutcome(session As AgentSession) As String
            If session?.Spec Is Nothing Then Return ""

            Dim successful = New HashSet(Of String)(
                If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item?.Action IsNot Nothing AndAlso
                                      item.Explanation IsNot Nothing AndAlso
                                      item.Explanation.Success).
                Select(Function(item) item.Action.ToolId),
                StringComparer.OrdinalIgnoreCase)
            Dim missing = ResolveRequiredCapabilities(session.Spec).
                Where(Function(toolId) Not successful.Contains(toolId)).
                ToList()
            If missing.Count > 0 Then
                Return $"任务未完成：用户要求的能力尚无成功观察：{String.Join("、", missing)}。"
            End If
            Return ""
        End Function

        Public Shared Function IsRequiredCapability(spec As AgentTaskSpec,
                                                     toolId As String) As Boolean
            If String.IsNullOrWhiteSpace(toolId) Then Return False
            Return ResolveRequiredCapabilities(spec).
                Any(Function(required) String.Equals(required, toolId, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>Legacy name retained for binary/source compatibility.</summary>
        Public Shared Function IsMandatoryTool(spec As AgentTaskSpec,
                                               toolId As String) As Boolean
            Return IsRequiredCapability(spec, toolId)
        End Function

        ''' <summary>
        ''' Single runtime authority for user-required capabilities.  Once a GoalContract
        ''' exists, legacy TaskSpec projections must not add, remove or replace goal policy.
        ''' Legacy fields are consulted only for persisted sessions that predate GoalContract.
        ''' </summary>
        Friend Shared Function ResolveRequiredCapabilities(spec As AgentTaskSpec) As List(Of String)
            If spec Is Nothing Then Return New List(Of String)()

            Dim result As New List(Of String)()
            If spec.GoalContract IsNot Nothing Then
                result.AddRange(spec.GoalContract.RequiredCapabilities)
            Else
                ' Compatibility path for persisted TaskSpec payloads created before the Goal
                ' Boundary existed.  This path disappears as soon as a frozen goal is attached.
                If spec.RequiredCapabilities IsNot Nothing Then result.AddRange(spec.RequiredCapabilities)
                If spec.MandatoryTools IsNot Nothing Then result.AddRange(spec.MandatoryTools)
            End If
            Return result.
                Where(Function(toolId) Not String.IsNullOrWhiteSpace(toolId)).
                Select(Function(toolId) toolId.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function
    End Class

End Namespace
