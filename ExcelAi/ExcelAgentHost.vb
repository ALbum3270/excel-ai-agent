Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAgent.Core
Imports ExcelAgent.Core.Agent
Imports ExcelAgent.Core.Agent.Context

Public Class ExcelContextSnapshot
    Public Property OfficeContext As OfficeContext
    Public Property ContextPack As ContextPack
    Public Property PromptText As String
End Class

''' <summary>
''' The only bridge from the shared Agent runtime to Excel COM. Every tool is dispatched to
''' an explicit adapter on Excel's UI thread and must return a structured ToolResult.
''' </summary>
Public NotInheritable Class ExcelAgentHost
    Private ReadOnly _application As Excel.Application
    Private ReadOnly _uiContext As SynchronizationContext

    Public Sub New(application As Excel.Application, uiContext As SynchronizationContext)
        _application = application
        _uiContext = If(uiContext, New Windows.Forms.WindowsFormsSynchronizationContext())
    End Sub

    Public Function CaptureSnapshot() As ExcelContextSnapshot
        Return InvokeOnUi(Function()
                              Dim officeContext = New Context.ExcelContextProvider(_application).GetContext()
                              Dim promptText = If(officeContext?.DocStructure?.Summary, "")
                              Return New ExcelContextSnapshot With {
                                  .OfficeContext = officeContext,
                                  .ContextPack = ContextPack.FromOfficeContext(officeContext, promptText),
                                  .PromptText = promptText
                              }
                          End Function)
    End Function

    Public Function CaptureContextPack() As ContextPack
        Return CaptureSnapshot().ContextPack
    End Function

    Public Function ExecuteTool(commandJson As String, language As String, preview As Boolean) As ToolResult
        Return InvokeOnUi(Function() ExecuteToolOnUi(commandJson))
    End Function

    Private Function ExecuteToolOnUi(commandJson As String) As ToolResult
        Try
            Dim command = JObject.Parse(commandJson)
            Dim toolId = If(command("command")?.ToString(), "").Trim()
            Dim parameters = TryCast(command("params"), JObject)
            If parameters Is Nothing Then parameters = New JObject()
            If String.IsNullOrWhiteSpace(toolId) Then
                Return ToolResult.Failed("", "Excel 工具命令缺少 command", errorCode:=ExceptionClassifier.CodeArgument)
            End If

            Select Case toolId.ToLowerInvariant()
                Case "readrange"
                    Return OfficeRuntime.ExcelReadRangeAdapter.Execute(_application, parameters)
                Case "discoverofficecapability"
                    Return OfficeRuntime.ExcelApiCatalogProvider.SearchAsToolResult(parameters)
                Case "officeobjectoperation"
                    Return OfficeRuntime.ExcelOperationExecutor.Execute(_application, parameters)
                Case "formatrange"
                    Return OfficeRuntime.ExcelFormatRangeAdapter.Execute(_application, parameters)
            End Select

            Dim result As ToolResult = Nothing
            If OfficeRuntime.ExcelStandardToolAdapter.TryExecute(_application, toolId, parameters, result) Then Return result
            Return ToolResult.Failed(toolId,
                                     "未注册的 Excel 工具: " & toolId,
                                     errorCode:=ExceptionClassifier.CodeNotFound,
                                     recoverable:=False)
        Catch ex As Exception
            Return ToolResult.FromException("ExcelHost", ex)
        End Try
    End Function

    Private Function InvokeOnUi(Of T)(callback As Func(Of T)) As T
        If SynchronizationContext.Current Is _uiContext Then Return callback()

        Dim value As T = Nothing
        Dim capturedError As Exception = Nothing
        _uiContext.Send(
            Sub(state)
                Try
                    value = callback()
                Catch ex As Exception
                    capturedError = ex
                End Try
            End Sub,
            Nothing)
        If capturedError IsNot Nothing Then Throw capturedError
        Return value
    End Function
End Class
