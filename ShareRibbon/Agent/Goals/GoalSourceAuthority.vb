Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.Goals

    ''' <summary>
    ''' Deterministic authority check for model-proposed source spans. An exact substring is
    ''' not sufficient: its boundaries must retain the surrounding polarity and modality of
    ''' the user clause. This prevents extracting "delete X" from "do not delete X" and
    ''' promoting the inverted fragment into the frozen goal.
    ''' </summary>
    Friend NotInheritable Class GoalSourceAuthority
        Private Sub New()
        End Sub

        Friend Shared Function IsCompleteSemanticSpan(rawUserRequest As String,
                                                       sourceStart As Integer,
                                                       sourceText As String) As Boolean
            Dim raw = If(rawUserRequest, "")
            Dim text = If(sourceText, "")
            If text.Length = 0 OrElse sourceStart < 0 OrElse
               sourceStart > raw.Length - text.Length Then Return False
            If Not String.Equals(raw.Substring(sourceStart, text.Length), text, StringComparison.Ordinal) Then Return False
            If sourceStart = 0 AndAlso text.Length = raw.Length Then Return True

            Return HasLeadingSemanticBoundary(raw, sourceStart) AndAlso
                HasTrailingSemanticBoundary(raw, sourceStart + text.Length)
        End Function

        Friend Shared Function CoversAuthoritativeText(rawUserRequest As String,
                                                       clauses As IEnumerable(Of CandidateGoalSourceClause)) As Boolean
            Dim raw = If(rawUserRequest, "")
            If raw.Length = 0 Then Return False
            Dim covered(raw.Length - 1) As Boolean
            For Each clause In If(clauses, Enumerable.Empty(Of CandidateGoalSourceClause)())
                If clause Is Nothing OrElse Not IsCompleteSemanticSpan(raw, clause.SourceStart, clause.Text) Then Continue For
                For index = clause.SourceStart To clause.SourceStart + clause.Text.Length - 1
                    covered(index) = True
                Next
            Next

            Dim gapStart As Integer = -1
            For index = 0 To raw.Length
                Dim isCovered = index < raw.Length AndAlso covered(index)
                If Not isCovered AndAlso gapStart < 0 Then gapStart = index
                If isCovered AndAlso gapStart >= 0 Then
                    If Not IsStructuralGap(raw.Substring(gapStart, index - gapStart)) Then Return False
                    gapStart = -1
                End If
            Next
            If gapStart >= 0 AndAlso Not IsStructuralGap(raw.Substring(gapStart)) Then Return False
            Return True
        End Function

        Private Shared Function IsStructuralGap(value As String) As Boolean
            Dim structural = Regex.Replace(
                If(value, ""),
                "(?:并且|以及|同时|然后|并|再|\b(?:and|then|also)\b)",
                "",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            For Each character In structural
                If Not Char.IsWhiteSpace(character) AndAlso Not IsHardSeparator(character) Then Return False
            Next
            Return True
        End Function

        Private Shared Function HasLeadingSemanticBoundary(raw As String, start As Integer) As Boolean
            Dim boundary = start
            While boundary > 0 AndAlso Char.IsWhiteSpace(raw(boundary - 1))
                boundary -= 1
            End While
            If boundary = 0 Then Return True
            If IsHardSeparator(raw(boundary - 1)) Then Return True

            Dim prefix = raw.Substring(0, boundary)
            Return Regex.IsMatch(
                prefix,
                "(?:并且|并|以及|同时|然后|再|\b(?:and|then|also)\b)\s*$",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
        End Function

        Private Shared Function HasTrailingSemanticBoundary(raw As String, finish As Integer) As Boolean
            Dim boundary = finish
            While boundary < raw.Length AndAlso Char.IsWhiteSpace(raw(boundary))
                boundary += 1
            End While
            If boundary = raw.Length Then Return True
            If IsHardSeparator(raw(boundary)) Then Return True

            Dim suffix = raw.Substring(boundary)
            Return Regex.IsMatch(
                suffix,
                "^\s*(?:并且|并|以及|同时|然后|再|\b(?:and|then|also)\b)",
                RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
        End Function

        Private Shared Function IsHardSeparator(value As Char) As Boolean
            Return value = "。"c OrElse value = "！"c OrElse value = "？"c OrElse
                value = "；"c OrElse value = "，"c OrElse value = "、"c OrElse
                value = "."c OrElse value = "!"c OrElse value = "?"c OrElse
                value = ";"c OrElse value = ","c OrElse value = ":"c OrElse
                value = "："c OrElse value = "|"c OrElse value = ChrW(13) OrElse
                value = ChrW(10)
        End Function
    End Class

End Namespace
