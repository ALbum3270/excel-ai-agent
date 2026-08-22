Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Deterministic acceptance contract for plans and completed executions.
    ''' The language model may choose parameters and sequencing, but it cannot omit or
    ''' replace tools that the semantic task classifier marked as mandatory.
    ''' </summary>
    Public NotInheritable Class AgentExecutionContract
        Private Sub New()
        End Sub

        Public Shared Function ValidatePlan(spec As AgentTaskSpec,
                                            plan As ExecutionPlan) As String
            If spec Is Nothing OrElse plan Is Nothing Then Return ""

            Dim plannedSequence = GetPlannedToolSequence(plan)
            Dim planned = New HashSet(Of String)(plannedSequence, StringComparer.OrdinalIgnoreCase)
            Dim missing = GetMandatoryTools(spec).
                Where(Function(toolId) Not planned.Contains(toolId)).
                ToList()
            If missing.Count > 0 Then
                Return $"计划未覆盖任务合同要求的工具：{String.Join("、", missing)}。不得以其他工具或降级写入替代。"
            End If

            Return ValidateSequence(spec.MandatoryToolSequence, plannedSequence, "计划")
        End Function

        Public Shared Function ValidateOutcome(session As AgentSession) As String
            If session?.Spec Is Nothing Then Return ""

            Dim successfulSequence = If(session.Iterations, New List(Of ReActIteration)()).
                Where(Function(item) item?.Action IsNot Nothing AndAlso
                                      item.Explanation IsNot Nothing AndAlso
                                      item.Explanation.Success).
                Select(Function(item) item.Action.ToolId).
                ToList()
            Dim successful = New HashSet(Of String)(successfulSequence, StringComparer.OrdinalIgnoreCase)
            Dim missing = GetMandatoryTools(session.Spec).
                Where(Function(toolId) Not successful.Contains(toolId)).
                ToList()
            If missing.Count > 0 Then
                Return $"任务未完成：执行记录中缺少任务合同要求的成功工具调用：{String.Join("、", missing)}。"
            End If

            Return ValidateSequence(session.Spec.MandatoryToolSequence, successfulSequence, "执行记录")
        End Function

        Public Shared Function IsMandatoryTool(spec As AgentTaskSpec,
                                               toolId As String) As Boolean
            If String.IsNullOrWhiteSpace(toolId) Then Return False
            Return GetMandatoryTools(spec).
                Any(Function(required) String.Equals(required, toolId, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Shared Function GetMandatoryTools(spec As AgentTaskSpec) As List(Of String)
            If spec?.MandatoryTools Is Nothing Then Return New List(Of String)()
            Return spec.MandatoryTools.
                Where(Function(toolId) Not String.IsNullOrWhiteSpace(toolId)).
                Select(Function(toolId) toolId.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function GetPlannedToolSequence(plan As ExecutionPlan) As List(Of String)
            Dim result As New List(Of String)()
            If plan?.Steps Is Nothing Then Return result

            For Each planStep In plan.Steps
                If Not String.IsNullOrWhiteSpace(planStep.ToolHint) Then
                    AddToolId(result, planStep.ToolHint)
                    Continue For
                End If
                Try
                    Dim envelope = JObject.Parse(If(planStep.Code, ""))
                    AddToolId(result, envelope("command")?.ToString())

                    Dim action = TryCast(envelope("action"), JObject)
                    AddToolId(result, action?("tool")?.ToString())

                    Dim commands = TryCast(envelope("commands"), JArray)
                    If commands Is Nothing Then Continue For
                    For Each commandItem In commands.OfType(Of JObject)()
                        AddToolId(result, commandItem("command")?.ToString())
                        Dim nestedAction = TryCast(commandItem("action"), JObject)
                        AddToolId(result, nestedAction?("tool")?.ToString())
                    Next
                Catch
                    ' Normal tool-call parsing reports malformed plan JSON separately.
                End Try
            Next
            Return result
        End Function

        Private Shared Function ValidateSequence(required As IEnumerable(Of String),
                                                 actual As IList(Of String),
                                                 sourceName As String) As String
            Dim sequence = If(required, Enumerable.Empty(Of String)()).
                Where(Function(toolId) Not String.IsNullOrWhiteSpace(toolId)).
                ToList()
            If sequence.Count < 2 Then Return ""

            Dim cursor As Integer = -1
            For Each requiredTool In sequence
                Dim found As Integer = -1
                For index = cursor + 1 To actual.Count - 1
                    If String.Equals(actual(index), requiredTool, StringComparison.OrdinalIgnoreCase) Then
                        found = index
                        Exit For
                    End If
                Next
                If found < 0 Then
                    Return $"{sourceName}未按任务合同顺序执行工具：{String.Join(" → ", sequence)}。"
                End If
                cursor = found
            Next
            Return ""
        End Function

        Private Shared Sub AddToolId(target As List(Of String), toolId As String)
            If target IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(toolId) Then
                target.Add(toolId.Trim())
            End If
        End Sub
    End Class

End Namespace
