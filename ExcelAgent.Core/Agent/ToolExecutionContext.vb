Imports System.Collections.Generic
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Carries per-run tool visibility and execution policy.
    ''' Application host tools, read-only policy, and selected private extensions are
    ''' represented independently so Skill retrieval cannot accidentally remove capabilities.
    ''' </summary>
    Public Class ToolExecutionContext
        Public Property AppType As String
        Public Property RunId As String
        Public Property CorrelationId As String
        Public Property PrimarySkillName As String
        Public Property AllowedTools As HashSet(Of String)
        Public Property ApprovedTools As HashSet(Of String)
        Public Property ApprovedToolCalls As HashSet(Of String)
        Public Property EnforceAllowedTools As Boolean = True
        Public Property ReadOnlyTask As Boolean
        Public Property AllowApplicationHostTools As Boolean
        Private _officeCapabilityDiscovered As Boolean

        Public Shared Function FromSession(session As AgentSession, skill As AgentSkill) As ToolExecutionContext
            Dim context As New ToolExecutionContext With {
                .AppType = If(session?.AppType, ""),
                .RunId = If(session?.Id, ""),
                .CorrelationId = If(session?.Id, ""),
                .PrimarySkillName = If(skill?.Name, ""),
                .AllowedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .ApprovedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .ApprovedToolCalls = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .EnforceAllowedTools = True,
                .ReadOnlyTask = String.Equals(If(session?.Spec?.MutationPolicy, ""),
                                              "read_only",
                                              StringComparison.OrdinalIgnoreCase),
                .AllowApplicationHostTools = True
            }

            ' Explicit grants are independent of the trusted Office host baseline. Task
            ' semantics and the selected Skill may opt into private/external capabilities;
            ' neither list can remove host tools or self-promote a descriptor into the host.
            AddExplicitGrants(context.AllowedTools, session?.Spec?.RequiredTools)
            AddExplicitGrants(context.AllowedTools, skill?.RequiredTools)

            Return context
        End Function

        Private Shared Sub AddExplicitGrants(target As HashSet(Of String),
                                              toolIds As IEnumerable(Of String))
            If target Is Nothing OrElse toolIds Is Nothing Then Return
            For Each toolId In toolIds
                If Not String.IsNullOrWhiteSpace(toolId) Then target.Add(toolId.Trim())
            Next
        End Sub

        Public Function HasPrimarySkillGate() As Boolean
            Return EnforceAllowedTools
        End Function

        Public Function IsToolAllowed(toolId As String) As Boolean
            If Not HasPrimarySkillGate() Then Return True
            If String.IsNullOrWhiteSpace(toolId) Then Return False
            ' A string has no trustworthy provenance, owner, or access mode.  Runtime contexts
            ' therefore fail closed and require the descriptor overload.  Manually constructed
            ' strict contexts retain exact-id behavior for legacy harness callers.
            If AllowApplicationHostTools Then Return False
            Return AllowedTools IsNot Nothing AndAlso AllowedTools.Contains(toolId.Trim())
        End Function

        Public Function IsToolAllowed(tool As ToolDescriptor) As Boolean
            If Not HasPrimarySkillGate() Then Return True
            If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.Id) Then Return False
            If ReadOnlyTask AndAlso Not IsNonMutating(tool) Then Return False

            Dim explicitlyGranted = AllowedTools IsNot Nothing AndAlso
                AllowedTools.Contains(tool.Id.Trim())
            If Not AllowApplicationHostTools Then Return explicitlyGranted

            Select Case tool.RegistrationProvenance
                Case ToolRegistrationProvenance.HostManifest
                    Return True
                Case ToolRegistrationProvenance.AgentInternal,
                     ToolRegistrationProvenance.Custom
                    Return explicitlyGranted
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsNonMutating(tool As ToolDescriptor) As Boolean
            Return tool IsNot Nothing AndAlso
                (String.Equals(tool.AccessMode, "read", StringComparison.OrdinalIgnoreCase) OrElse
                 String.Equals(tool.AccessMode, "compute", StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' The generic object-operation interface is a second-stage capability. Keeping
        ''' discovery state in the per-run context prevents a model from bypassing a
        ''' registered high-level tool with invented COM-shaped commands.
        ''' </summary>
        Public Function IsOfficeObjectOperationReady() As Boolean
            Return _officeCapabilityDiscovered
        End Function

        Public Sub RecordSuccessfulTool(toolId As String)
            If String.Equals(If(toolId, "").Trim(),
                             "DiscoverOfficeCapability",
                             StringComparison.OrdinalIgnoreCase) Then
                _officeCapabilityDiscovered = True
            End If
        End Sub

        Public Function AllowedToolsText() As String
            If Not EnforceAllowedTools Then Return "unrestricted"
            If ReadOnlyTask Then
                Return "read-only Office host tools; explicit grants: " &
                    String.Join(", ", If(AllowedTools, New HashSet(Of String)()).OrderBy(Function(id) id))
            End If
            If Not AllowApplicationHostTools Then
                Return String.Join(", ", If(AllowedTools, New HashSet(Of String)()).OrderBy(Function(id) id))
            End If
            Dim explicitTools = If(AllowedTools, New HashSet(Of String)()).OrderBy(Function(id) id).ToList()
            Return "trusted application host tools" &
                If(explicitTools.Count = 0, "", "; explicit grants: " & String.Join(", ", explicitTools))
        End Function

        Public Sub ApproveTool(toolId As String, Optional params As JObject = Nothing)
            If String.IsNullOrWhiteSpace(toolId) Then Return
            If params IsNot Nothing Then
                If ApprovedToolCalls Is Nothing Then ApprovedToolCalls = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                ApprovedToolCalls.Add(BuildApprovalKey(toolId, params))
                Return
            End If
            If ApprovedTools Is Nothing Then ApprovedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            ApprovedTools.Add(toolId.Trim())
        End Sub

        Public Function IsToolApproved(toolId As String, Optional params As JObject = Nothing) As Boolean
            If params IsNot Nothing AndAlso ApprovedToolCalls IsNot Nothing Then
                Return ApprovedToolCalls.Contains(BuildApprovalKey(toolId, params))
            End If
            Return ApprovedTools IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(toolId) AndAlso
                   ApprovedTools.Contains(toolId.Trim())
        End Function

        Public Function ConsumeToolApproval(toolId As String, Optional params As JObject = Nothing) As Boolean
            If params IsNot Nothing AndAlso ApprovedToolCalls IsNot Nothing Then
                Dim approvalKey = BuildApprovalKey(toolId, params)
                Return ApprovedToolCalls.Remove(approvalKey)
            End If
            If Not IsToolApproved(toolId, Nothing) Then Return False
            Return ApprovedTools.Remove(toolId.Trim())
        End Function

        Public Shared Function BuildApprovalKey(toolId As String, params As JObject) As String
            Dim normalizedToolId = If(toolId, "").Trim().ToLowerInvariant()
            Dim canonicalParams = CanonicalizeToken(If(params, New JObject())).ToString(Formatting.None)
            Dim bytes = Encoding.UTF8.GetBytes(canonicalParams)
            Using sha = SHA256.Create()
                Dim hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
                Return normalizedToolId & ":" & hash
            End Using
        End Function

        Private Shared Function CanonicalizeToken(token As JToken) As JToken
            If token Is Nothing Then Return JValue.CreateNull()

            If token.Type = JTokenType.Object Then
                Dim source = DirectCast(token, JObject)
                Dim result As New JObject()
                For Each prop In source.Properties().OrderBy(Function(item) item.Name, StringComparer.Ordinal)
                    result.Add(prop.Name, CanonicalizeToken(prop.Value))
                Next
                Return result
            End If

            If token.Type = JTokenType.Array Then
                Dim result As New JArray()
                For Each item In DirectCast(token, JArray)
                    result.Add(CanonicalizeToken(item))
                Next
                Return result
            End If

            Return token.DeepClone()
        End Function
    End Class

End Namespace
