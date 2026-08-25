Namespace Agent.Harness

    Public Interface IRunTraceStore
        Sub StartRun(runId As String, turn As UserTurn, startedAt As DateTime)
        Sub AppendStep(runId As String, seq As Integer, toolId As String, status As String, message As String, errorCode As String, observation As Object, startedAt As DateTime, finishedAt As DateTime)
        Sub SetRunStatus(runId As String, status As String, message As String, errorCode As String)
        Sub CompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime)
    End Interface

    ''' <summary>
    ''' Run evidence remains in the live Agent session and workbook observations are not
    ''' retained after the run.
    ''' </summary>
    Public NotInheritable Class NoopRunTraceStore
        Implements IRunTraceStore

        Public Sub StartRun(runId As String, turn As UserTurn, startedAt As DateTime) Implements IRunTraceStore.StartRun
        End Sub

        Public Sub AppendStep(runId As String, seq As Integer, toolId As String, status As String, message As String, errorCode As String, observation As Object, startedAt As DateTime, finishedAt As DateTime) Implements IRunTraceStore.AppendStep
        End Sub

        Public Sub SetRunStatus(runId As String, status As String, message As String, errorCode As String) Implements IRunTraceStore.SetRunStatus
        End Sub

        Public Sub CompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime) Implements IRunTraceStore.CompleteRun
        End Sub
    End Class
End Namespace
