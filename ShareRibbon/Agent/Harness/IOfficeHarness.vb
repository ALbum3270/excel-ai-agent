Imports System.Threading
Imports System.Threading.Tasks

Namespace Agent.Harness

    Public Interface IOfficeHarness
        Event PhaseChanged As EventHandler(Of HarnessPhaseChangedEventArgs)
        Event StepChanged As EventHandler(Of HarnessStepChangedEventArgs)
        Event ContextReady As EventHandler(Of HarnessContextEventArgs)

        Function RunAsync(turn As UserTurn, cancellationToken As CancellationToken) As Task(Of HarnessRunResult)
    End Interface

End Namespace
