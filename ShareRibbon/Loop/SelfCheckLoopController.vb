' ShareRibbon\Loop\SelfCheckLoopController.vb
' 当前仅承载 Chat 路由仍在使用的发送前检查与响应格式校验。

Imports System.Threading.Tasks

''' <summary>
''' Chat 路由自检适配器。
''' </summary>
Public Class SelfCheckLoopController

    Private ReadOnly _contextChecker As IContextChecker
    Private ReadOnly _instructionValidator As IInstructionValidator

    Public Sub New(
        contextChecker As IContextChecker,
        instructionValidator As IInstructionValidator)

        If contextChecker Is Nothing Then Throw New ArgumentNullException(NameOf(contextChecker))
        If instructionValidator Is Nothing Then Throw New ArgumentNullException(NameOf(instructionValidator))

        _contextChecker = contextChecker
        _instructionValidator = instructionValidator
    End Sub

    Public Function PreSendCheckAsync(context As ExecutionContext) As Task(Of ContextCheckResult)
        Return _contextChecker.CheckAsync(context)
    End Function

    Public Function PostFlushValidateAsync(
        aiResponse As String,
        expectedFormat As InstructionFormat,
        context As ExecutionContext) As Task(Of ValidationResult)

        Return _instructionValidator.ValidateAsync(aiResponse, expectedFormat, context)
    End Function

End Class
