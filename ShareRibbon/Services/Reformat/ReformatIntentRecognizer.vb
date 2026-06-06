Imports System.Text
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' 智能排版意图识别结果。
''' LLM 可用时优先使用结构化 JSON；不可用或解析失败时使用规则结果。
''' </summary>
Public Class ReformatIntentRecognitionResult
    Public Property Intent As FormatIntent = New FormatIntent()
    Public Property Confidence As Double = 0.0
    Public Property Source As String = "rules"
    Public Property ScopeHint As ReformatScopeKind = ReformatScopeKind.Selection
    Public Property RawResponse As String = ""
    Public Property Notes As New List(Of String)()
End Class

Public Class ReformatIntentRecognizer
    Private ReadOnly _textAnalyzer As Func(Of String, String, Task(Of String))
    Private ReadOnly _ruleParser As Func(Of String, DocumentAnalysisResult, FormatIntent)
    Private ReadOnly _knowledgeEngine As FormattingKnowledgeEngine
    Private ReadOnly _standardRegistry As FormattingStandardRegistry

    Public Sub New(textAnalyzer As Func(Of String, String, Task(Of String)),
                   ruleParser As Func(Of String, DocumentAnalysisResult, FormatIntent),
                   Optional knowledgeEngine As FormattingKnowledgeEngine = Nothing)
        _textAnalyzer = textAnalyzer
        _ruleParser = ruleParser
        _knowledgeEngine = If(knowledgeEngine, New FormattingKnowledgeEngine())
        _standardRegistry = New FormattingStandardRegistry(_knowledgeEngine)
    End Sub

    Public Async Function RecognizeAsync(userMessage As String,
                                         analysis As DocumentAnalysisResult,
                                         paragraphs As List(Of String)) As Task(Of ReformatIntentRecognitionResult)
        Dim fallback = BuildRuleFallback(userMessage, analysis)

        If _textAnalyzer Is Nothing OrElse String.IsNullOrWhiteSpace(userMessage) Then
            Return fallback
        End If

        Try
            Dim prompt = BuildPrompt(userMessage, analysis, paragraphs)
            Dim response = Await _textAnalyzer("reformat_intent", prompt)
            If String.IsNullOrWhiteSpace(response) Then
                fallback.Notes.Add("LLM intent response is empty.")
                Return fallback
            End If

            Dim parsed = ParseLlmResponse(response, fallback.Intent)
            If parsed Is Nothing Then
                fallback.RawResponse = response
                fallback.Notes.Add("LLM intent response could not be parsed.")
                Return fallback
            End If

            parsed.RawResponse = response
            parsed.Source = "llm"
            If parsed.Confidence <= 0 Then parsed.Confidence = 0.75
            Return parsed
        Catch ex As Exception
            fallback.Notes.Add("LLM intent recognition failed: " & ex.Message)
            Return fallback
        End Try
    End Function

    Private Function BuildRuleFallback(userMessage As String, analysis As DocumentAnalysisResult) As ReformatIntentRecognitionResult
        Dim result As New ReformatIntentRecognitionResult()
        If _ruleParser IsNot Nothing Then
            result.Intent = _ruleParser(userMessage, analysis)
        Else
            result.Intent = New FormatIntent()
        End If
        result.Confidence = 0.45
        result.Source = "rules"
        Return result
    End Function

    Private Function BuildPrompt(userMessage As String,
                                 analysis As DocumentAnalysisResult,
                                 paragraphs As List(Of String)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("你是 Word 文档智能排版的意图识别器。")
        sb.AppendLine("请把用户需求解析为严格 JSON，不要输出 markdown，不要解释。")
        sb.AppendLine()
        sb.AppendLine("输出 JSON 字段：")
        sb.AppendLine("{")
        sb.AppendLine("  ""intentType"": ""AutoFormat|StandardFormat|StyleClone|SpecificTweak|FormatCleanup"",")
        sb.AppendLine("  ""targetDocumentType"": ""OfficialDocument|AcademicPaper|BusinessReport|Contract|Resume|GeneralDocument|Unknown"",")
        sb.AppendLine("  ""targetStandardName"": ""标准名称或空字符串"",")
        sb.AppendLine("  ""specificRequests"": [""具体格式要求""],")
        sb.AppendLine("  ""scope"": ""selection|wholeDocument|section"",")
        sb.AppendLine("  ""confidence"": 0.0")
        sb.AppendLine("}")
        sb.AppendLine()
        sb.AppendLine("规则：")
        sb.AppendLine("1. 用户明确说公文、国标、GB/T 9704 时，优先 StandardFormat 和 GB/T 9704-2012。")
        sb.AppendLine("2. 用户说标题更醒目、正文紧一点、字号调大等局部变化时，使用 SpecificTweak。")
        sb.AppendLine("3. 用户说清理格式、去掉混乱格式时，使用 FormatCleanup。")
        sb.AppendLine("4. 用户说参考范文、照这个、格式克隆时，使用 StyleClone。")
        sb.AppendLine("5. specificRequests 保留用户的真实约束，例如页码位置、三线表、标题字号、行距、字体、颜色。")
        sb.AppendLine()

        sb.AppendLine("可用标准：")
        For Each candidate In _standardRegistry.GetAllCandidates().Take(30)
            If candidate.Standard Is Nothing Then Continue For
            sb.AppendLine("- " & candidate.Standard.Name & " [" & candidate.SourceType.ToString() & "]: " & candidate.Standard.Description)
        Next
        sb.AppendLine()

        If analysis IsNot Nothing Then
            sb.AppendLine("规则分析上下文：")
            sb.AppendLine("- documentType: " & analysis.DocumentType.ToString())
            sb.AppendLine("- confidence: " & analysis.Confidence.ToString("0.00"))
            sb.AppendLine("- paragraphCount: " & analysis.ParagraphCount.ToString())
            sb.AppendLine()
        End If

        sb.AppendLine("文档片段：")
        sb.AppendLine(BuildParagraphSample(paragraphs))
        sb.AppendLine()
        sb.AppendLine("用户需求：")
        sb.AppendLine(userMessage)

        Return sb.ToString()
    End Function

    Private Shared Function BuildParagraphSample(paragraphs As List(Of String)) As String
        If paragraphs Is Nothing OrElse paragraphs.Count = 0 Then Return "(empty)"

        Dim sb As New StringBuilder()
        Dim maxCount = Math.Min(paragraphs.Count, 12)
        For i = 0 To maxCount - 1
            Dim text = If(paragraphs(i), "").Trim()
            If text.Length > 120 Then text = text.Substring(0, 120)
            sb.AppendLine($"[{i}] {text}")
        Next
        Return sb.ToString()
    End Function

    Private Function ParseLlmResponse(response As String, fallbackIntent As FormatIntent) As ReformatIntentRecognitionResult
        Dim jsonText = ExtractJsonObject(response)
        If String.IsNullOrWhiteSpace(jsonText) Then Return Nothing

        Dim obj = JObject.Parse(jsonText)
        Dim result As New ReformatIntentRecognitionResult()
        Dim intent As New FormatIntent()

        intent.IntentType = ParseIntentType(obj("intentType")?.ToString(), fallbackIntent.IntentType)
        intent.TargetDocumentType = ParseDocumentType(obj("targetDocumentType")?.ToString(), fallbackIntent.TargetDocumentType)
        intent.TargetStandardName = If(obj("targetStandardName")?.ToString(), fallbackIntent.TargetStandardName)

        Dim requests = TryCast(obj("specificRequests"), JArray)
        If requests IsNot Nothing Then
            For Each item In requests
                Dim value = item?.ToString()
                If Not String.IsNullOrWhiteSpace(value) Then intent.SpecificRequests.Add(value)
            Next
        End If
        If intent.SpecificRequests.Count = 0 AndAlso fallbackIntent.SpecificRequests IsNot Nothing Then
            For Each item In fallbackIntent.SpecificRequests
                intent.SpecificRequests.Add(item)
            Next
        End If

        result.Intent = intent
        result.ScopeHint = ParseScope(obj("scope")?.ToString())
        Dim confidence As Double = 0
        If Double.TryParse(obj("confidence")?.ToString(), confidence) Then
            result.Confidence = Math.Max(0, Math.Min(1, confidence))
        End If

        Return result
    End Function

    Private Shared Function ExtractJsonObject(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return Nothing
        Dim clean = text.Trim()
        If clean.StartsWith("```") Then
            Dim firstBrace = clean.IndexOf("{"c)
            Dim lastBrace = clean.LastIndexOf("}"c)
            If firstBrace >= 0 AndAlso lastBrace > firstBrace Then
                Return clean.Substring(firstBrace, lastBrace - firstBrace + 1)
            End If
        End If

        If clean.StartsWith("{") Then Return clean

        Dim startIndex = clean.IndexOf("{"c)
        Dim endIndex = clean.LastIndexOf("}"c)
        If startIndex >= 0 AndAlso endIndex > startIndex Then
            Return clean.Substring(startIndex, endIndex - startIndex + 1)
        End If
        Return Nothing
    End Function

    Private Shared Function ParseIntentType(value As String, fallback As IntentType) As IntentType
        If String.IsNullOrWhiteSpace(value) Then Return fallback
        Select Case value.Trim().ToLowerInvariant()
            Case "autoformat", "auto", "format"
                Return IntentType.AutoFormat
            Case "standardformat", "standard"
                Return IntentType.StandardFormat
            Case "styleclone", "clone", "mirror"
                Return IntentType.StyleClone
            Case "specifictweak", "tweak", "refine"
                Return IntentType.SpecificTweak
            Case "formatcleanup", "cleanup"
                Return IntentType.FormatCleanup
            Case Else
                Return fallback
        End Select
    End Function

    Private Shared Function ParseDocumentType(value As String, fallback As DocumentType) As DocumentType
        If String.IsNullOrWhiteSpace(value) Then Return fallback
        Dim normalized = value.Trim()
        Select Case normalized.ToLowerInvariant()
            Case "official", "officialdocument", "公文"
                Return DocumentType.OfficialDocument
            Case "academic", "academicpaper", "paper", "论文"
                Return DocumentType.AcademicPaper
            Case "business", "businessreport", "report", "报告"
                Return DocumentType.BusinessReport
            Case "contract", "合同"
                Return DocumentType.Contract
            Case "resume", "简历"
                Return DocumentType.[Resume]
            Case "general", "generaldocument", "通用"
                Return DocumentType.GeneralDocument
            Case Else
                Dim parsed As DocumentType
                If [Enum].TryParse(Of DocumentType)(normalized, True, parsed) Then
                    Return parsed
                End If
                Return fallback
        End Select
    End Function

    Private Shared Function ParseScope(value As String) As ReformatScopeKind
        If String.IsNullOrWhiteSpace(value) Then Return ReformatScopeKind.Selection
        Select Case value.Trim().ToLowerInvariant()
            Case "wholedocument", "whole", "document", "全文", "整篇"
                Return ReformatScopeKind.WholeDocument
            Case "section", "章节"
                Return ReformatScopeKind.Section
            Case Else
                Return ReformatScopeKind.Selection
        End Select
    End Function
End Class
