' ShareRibbon\Services\UnifiedMemoryService.vb
' 统一记忆管理服务 - 整合原子记忆与Ralph Loop记忆

''' <summary>
''' 统一记忆管理服务
''' </summary>
Public Class UnifiedMemoryService

    ''' <summary>
    ''' 计算记忆重要性（基于内容、用户反馈、访问频率）
    ''' </summary>
    Public Shared Function CalculateImportance(
        content As String,
        memoryType As String,
        metadata As Dictionary(Of String, Object)) As Double

        Dim baseScore As Double = 0.5

        ' 1. 根据类型调整基础分
        Select Case memoryType?.ToLowerInvariant()
            Case "user_explicit_intent"
                baseScore = 0.9
            Case "assistant_solution", "task_result"
                baseScore = 0.8
            Case "user_feedback"
                baseScore = 0.85
            Case "knowledge"
                baseScore = 0.75
            Case "skill_feedback"
                baseScore = 0.7
        End Select

        ' 2. 内容特征分析
        If Not String.IsNullOrWhiteSpace(content) Then
            If content.Length > 100 Then baseScore += 0.1
            If content.Contains("?") OrElse content.Contains("如何") OrElse content.Contains("怎么") Then baseScore += 0.05
            If content.Contains("!") OrElse content.Contains("重要") OrElse content.Contains("关键") Then baseScore += 0.08
        End If

        ' 3. 元数据调整
        If metadata IsNot Nothing Then
            If metadata.ContainsKey("user_explicit_save") AndAlso CBool(metadata("user_explicit_save")) Then
                baseScore += 0.15
            End If
            If metadata.ContainsKey("user_rating") Then
                Dim rating = Convert.ToInt32(metadata("user_rating"))
                baseScore += (rating - 3) * 0.1  ' 5星+0.2，3星0，1星-0.2
            End If
        End If

        Return Math.Min(1.0, Math.Max(0.1, baseScore))
    End Function

End Class
