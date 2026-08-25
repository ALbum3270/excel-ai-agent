Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace Agent.Goals

    ''' <summary>
    ''' Deterministic Adapter from exact user language to method capabilities whose use is
    ''' itself part of the requested outcome.  Model declarations are never evidence.  New
    ''' method capabilities are added here with their registry aliases and invocation grammar.
    ''' </summary>
    Friend NotInheritable Class GoalCapabilityEvidenceResolver
        Private NotInheritable Class CapabilityGrammar
            Friend ReadOnly CapabilityId As String
            Friend ReadOnly EnginePattern As String
            Friend ReadOnly OperationPattern As String

            Friend Sub New(capabilityId As String,
                           enginePattern As String,
                           operationPattern As String)
                Me.CapabilityId = capabilityId
                Me.EnginePattern = enginePattern
                Me.OperationPattern = operationPattern
            End Sub
        End Class

        Private Structure ClauseSpan
            Friend Start As Integer
            Friend Text As String
        End Structure

        Private Shared ReadOnly Grammars As New List(Of CapabilityGrammar) From {
            New CapabilityGrammar(
                "PythonCompute",
                "(?:PythonCompute|Python)",
                "(?:calculate|analyze|process|summarize|aggregate|compute|计算|分析|处理|汇总|统计|运行|执行)")
        }

        Private Const InvocationPrefix As String =
            "(?:(?:use|using|via|run|execute|with)\s+|(?:使用|用|通过|调用|运行|执行|基于)\s*)"
        Private Const EnglishNegativePrefix As String =
            "(?:do\s+not|don['’]?t|never|without|avoid|must\s+not|mustn['’]?t|should\s+not|shouldn['’]?t|cannot|can['’]?t|no\s+need\s+to)"
        Private Const ChineseNegativePrefix As String =
            "(?:不要|别|无需|无须|不必|不需要|禁止|不得|避免|不能|不可|莫|勿)"
        Private Const EnglishOptionalPrefix As String =
            "(?:if\s+possible|if\s+available|when\s+convenient|consider|optionally|maybe|might|could|preferably|try\s+to)"
        Private Const ChineseOptionalPrefix As String =
            "(?:如果可以|如有可能|可以的话|条件允许时|尽量|最好|建议|考虑|可选|视情况|酌情)"
        Private Const EnglishMetaPrefix As String =
            "(?:why|how|whether|can|could|would|should|is|are|do|does|did|what)"
        Private Const ChineseMetaPrefix As String =
            "(?:为什么|为何|怎么|如何|是否|能否|可否|请问)"

        Private Sub New()
        End Sub

        Friend Shared Function Resolve(rawUserRequest As String) As List(Of String)
            If String.IsNullOrWhiteSpace(rawUserRequest) Then Return New List(Of String)()
            Return Grammars.
                Where(Function(grammar) HasAffirmativeInvocation(rawUserRequest, grammar)).
                Select(Function(grammar) grammar.CapabilityId).
                ToList()
        End Function

        Friend Shared Function ClauseSupports(clauseText As String, capabilityId As String) As Boolean
            If String.Equals(If(capabilityId, "").Trim(), "PythonCompute", StringComparison.OrdinalIgnoreCase) Then
                Return IsExplicitPythonComputeRequest(clauseText)
            End If
            Return False
        End Function

        Friend Shared Function IsExplicitPythonComputeRequest(input As String) As Boolean
            Return HasAffirmativeInvocation(
                input,
                Grammars.First(Function(grammar) String.Equals(
                    grammar.CapabilityId,
                    "PythonCompute",
                    StringComparison.OrdinalIgnoreCase)))
        End Function

        ''' <summary>
        ''' A capability is policy only when an invocation is asserted by the user.  Engine
        ''' words inside negation, optional language, questions or quoted literals remain goal
        ''' text, but cannot silently become a hard execution-method constraint.
        ''' </summary>
        Private Shared Function HasAffirmativeInvocation(input As String,
                                                         grammar As CapabilityGrammar) As Boolean
            If String.IsNullOrWhiteSpace(input) OrElse grammar Is Nothing Then Return False

            Dim engine = grammar.EnginePattern & "(?![A-Za-z0-9_])"
            Dim explicitInvocation = InvocationPrefix & engine
            Dim explicitOperation = engine &
                "\s*(?:(?:to|for)\s+|(?:来|进行|用于)?\s*)" & grammar.OperationPattern
            Dim matches = Regex.Matches(
                input,
                $"(?:{explicitInvocation})|(?:{explicitOperation})",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            For Each invocation As Match In matches
                If IsInsideQuotedLiteral(input, invocation.Index) Then Continue For
                Dim clause = GetClauseSpan(input, invocation.Index)
                Dim relativeStart = Math.Max(0, invocation.Index - clause.Start)
                Dim prefix = clause.Text.Substring(0, Math.Min(relativeStart, clause.Text.Length))
                If HasNonAssertiveScope(prefix) Then Continue For
                Return True
            Next
            Return False
        End Function

        Private Shared Function HasNonAssertiveScope(prefix As String) As Boolean
            Dim text = If(prefix, "")
            If Regex.IsMatch(
                text,
                $"(?:{ChineseNegativePrefix}|{EnglishNegativePrefix})(?:\s|再|去|直接|实际|really|actually)*$",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then Return True
            If Regex.IsMatch(
                text,
                $"(?:{ChineseOptionalPrefix}|{EnglishOptionalPrefix})(?:\s|，|,)*$",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then Return True
            If Regex.IsMatch(
                text,
                $"^\s*(?:{ChineseMetaPrefix}|{EnglishMetaPrefix})(?:\s|，|,)*",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then Return True
            Return False
        End Function

        Private Shared Function GetClauseSpan(input As String, index As Integer) As ClauseSpan
            Dim start = Math.Max(0, Math.Min(index, input.Length))
            While start > 0 AndAlso Not IsClauseSeparator(input(start - 1))
                start -= 1
            End While
            Dim finish = Math.Max(start, Math.Min(index, input.Length))
            While finish < input.Length AndAlso Not IsClauseSeparator(input(finish))
                finish += 1
            End While
            Return New ClauseSpan With {
                .Start = start,
                .Text = input.Substring(start, finish - start)
            }
        End Function

        Private Shared Function IsClauseSeparator(value As Char) As Boolean
            Return value = "。"c OrElse value = "！"c OrElse value = "？"c OrElse
                value = "；"c OrElse value = ";"c OrElse value = "!"c OrElse
                value = "?"c OrElse value = ChrW(13) OrElse value = ChrW(10) OrElse
                value = "，"c OrElse value = ","c
        End Function

        Private Shared Function IsInsideQuotedLiteral(input As String, index As Integer) As Boolean
            If String.IsNullOrEmpty(input) OrElse index <= 0 Then Return False
            Return IsInsideQuotePair(input, index, ChrW(&H201C), ChrW(&H201D)) OrElse
                IsInsideQuotePair(input, index, ChrW(&H2018), ChrW(&H2019)) OrElse
                IsInsideQuotePair(input, index, ChrW(&H300C), ChrW(&H300D)) OrElse
                IsInsideQuotePair(input, index, ChrW(&H300E), ChrW(&H300F)) OrElse
                IsInsideSymmetricQuote(input, index, """"c)
        End Function

        Private Shared Function IsInsideQuotePair(input As String,
                                                  index As Integer,
                                                  opening As Char,
                                                  closing As Char) As Boolean
            Dim openIndex = input.LastIndexOf(opening, Math.Min(index - 1, input.Length - 1))
            If openIndex < 0 Then Return False
            Dim closeBefore = input.LastIndexOf(closing, Math.Min(index - 1, input.Length - 1))
            If closeBefore > openIndex Then Return False
            Return input.IndexOf(closing, index) >= index
        End Function

        Private Shared Function IsInsideSymmetricQuote(input As String,
                                                       index As Integer,
                                                       quote As Char) As Boolean
            Dim count As Integer = 0
            For position = 0 To Math.Min(index - 1, input.Length - 1)
                If input(position) = quote AndAlso (position = 0 OrElse input(position - 1) <> ChrW(92)) Then
                    count += 1
                End If
            Next
            Return (count Mod 2) = 1 AndAlso input.IndexOf(quote, index) >= index
        End Function
    End Class

End Namespace
