' ShareRibbon\Services\Reformat\TempDocumentService.vb
' 临时文档服务 - 管理排版/校对过程中的临时文件

Imports System.IO

''' <summary>
''' 临时文档服务 - 原文档 ↔ 临时文档的复制与清理
''' 所有排版/校对操作在临时文档上进行，确认后才合并回原文档
''' </summary>
Public Class TempDocumentService

    Private Shared ReadOnly _tempDir As String = Path.Combine(
        Path.GetTempPath(), "OfficeAI_Reformat")

    Private Shared _lock As New Object()

    ''' <summary>
    ''' 创建原文档的临时副本
    ''' </summary>
    ''' <param name="sourceDocPath">原文档完整路径</param>
    ''' <returns>临时文档路径</returns>
    Public Shared Function CreateTempDocument(sourceDocPath As String) As String
        EnsureTempDirectory()

        Dim docName = Path.GetFileNameWithoutExtension(sourceDocPath)
        Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim guidStr As String = Guid.NewGuid().ToString("N").Substring(0, 8)
        Dim tempFileName = $"{docName}_{timestamp}_{guidStr}.docx"
        Dim tempPath = Path.Combine(_tempDir, tempFileName)

        File.Copy(sourceDocPath, tempPath, overwrite:=True)

        Return tempPath
    End Function

    ''' <summary>
    ''' 从Word Document对象创建临时副本（获取路径后复制）
    ''' </summary>
    Public Shared Function CreateTempDocument(sourceDoc As Object) As String
        Dim docPath As String = ""
        Try
            docPath = sourceDoc.FullName
        Catch
            ' 未保存文档，先保存到临时目录
            docPath = Path.Combine(_tempDir, $"UnsavedDoc_{Guid.NewGuid().ToString("N")}.docx")
            Try
                sourceDoc.SaveAs2(docPath)
            Catch ex As Exception
                Throw New InvalidOperationException($"无法保存未保存的文档到临时路径: {ex.Message}", ex)
            End Try
        End Try

        Return CreateTempDocument(docPath)
    End Function

    ''' <summary>
    ''' 安全删除临时文件
    ''' </summary>
    Public Shared Sub Cleanup(tempPath As String)
        If String.IsNullOrEmpty(tempPath) Then Return

        Try
            If File.Exists(tempPath) Then
                File.Delete(tempPath)
            End If
        Catch ex As Exception
            Debug.WriteLine($"TempDocumentService.Cleanup 删除失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 清理所有过期的临时文件（超过24小时的）
    ''' </summary>
    Public Shared Sub CleanupExpired()
        Try
            If Not Directory.Exists(_tempDir) Then Return

            Dim cutoff = DateTime.Now.AddHours(-24)
            For Each tempFile In Directory.GetFiles(_tempDir, "*.docx")
                Try
                    Dim fi = New FileInfo(tempFile)
                    If fi.LastWriteTime < cutoff Then
                        System.IO.File.Delete(tempFile)
                    End If
                Catch
                End Try
            Next
        Catch ex As Exception
            Debug.WriteLine($"TempDocumentService.CleanupExpired 失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取临时目录路径
    ''' </summary>
    Public Shared Function GetTempDirectory() As String
        EnsureTempDirectory()
        Return _tempDir
    End Function

    ''' <summary>
    ''' 确保临时目录存在
    ''' </summary>
    Private Shared Sub EnsureTempDirectory()
        SyncLock _lock
            If Not Directory.Exists(_tempDir) Then
                Directory.CreateDirectory(_tempDir)
            End If
        End SyncLock
    End Sub

End Class
