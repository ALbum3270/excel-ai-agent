Namespace Agent

    ''' <summary>
    ''' Runtime projection of a selected directory-based Excel skill.
    ''' </summary>
    Public Class AgentSkill
        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property TriggerPatterns As New List(Of String)()
        Public Property RequiredTools As New List(Of String)()
        Public Property PromptTemplate As String
        Public Property MaxSteps As Integer = 8
        Public Property AutoApprove As Boolean = False
    End Class

End Namespace
