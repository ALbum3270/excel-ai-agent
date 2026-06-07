Imports System.Collections.Generic

''' <summary>
''' 一次智能排版任务的固定上下文。
''' 生成预览时捕获，后续微调、切换和应用都应复用该对象，避免重新读取当前选区。
''' </summary>
Public Class ReformatJob
    Public Property JobId As String = Guid.NewGuid().ToString("N")
    Public Property CreatedAt As DateTime = DateTime.Now
    Public Property ScopeKind As ReformatScopeKind = ReformatScopeKind.Selection
    Public Property SourceDocumentName As String = ""
    Public Property ScopeStart As Integer = -1
    Public Property ScopeEnd As Integer = -1

    Public Property WordParagraphs As New List(Of Object)()
    Public Property ParagraphTexts As New List(Of String)()
    Public Property ParagraphStyles As New List(Of String)()
    Public Property ParagraphTypes As New List(Of String)()
    Public Property ParagraphFontSizes As New List(Of Single)()
    Public Property ParagraphIsBold As New List(Of Boolean)()

    Public Property PreviewPlan As ReformatPreviewPlan = Nothing

    Public ReadOnly Property ParagraphCount As Integer
        Get
            Return If(WordParagraphs IsNot Nothing, WordParagraphs.Count, 0)
        End Get
    End Property

    Public ReadOnly Property TextParagraphCount As Integer
        Get
            If ParagraphTypes Is Nothing Then Return ParagraphCount
            Dim count As Integer = 0
            For Each item In ParagraphTypes
                If String.Equals(item, "text", StringComparison.OrdinalIgnoreCase) Then
                    count += 1
                End If
            Next
            Return count
        End Get
    End Property

    Public Function HasUsableParagraphs() As Boolean
        Return WordParagraphs IsNot Nothing AndAlso WordParagraphs.Count > 0
    End Function

    Public Function GetScopeSummary() As String
        Dim scopeName As String
        Select Case ScopeKind
            Case ReformatScopeKind.WholeDocument
                scopeName = "全文"
            Case ReformatScopeKind.Section
                scopeName = "章节"
            Case Else
                scopeName = "选区"
        End Select

        Dim docPart = If(String.IsNullOrWhiteSpace(SourceDocumentName), "当前文档", SourceDocumentName)
        Return $"{docPart} / {scopeName} / {ParagraphCount} 段"
    End Function
End Class

Public Enum ReformatScopeKind
    Selection
    WholeDocument
    Section
End Enum
