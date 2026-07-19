Namespace Design

    Public NotInheritable Class SlideLayoutEngine
        Private Sub New()
        End Sub

        Public Shared Function Compile(spec As SlideDesignSpec,
                                       tokens As DesignTokens,
                                       slideWidth As Single,
                                       slideHeight As Single,
                                       slideIndex As Integer,
                                       totalSlides As Integer) As SlideRenderPlan
            Dim plan = SlideComponentLibrary.Build(spec, tokens, slideWidth, slideHeight, slideIndex, totalSlides)
            NormalizeBounds(plan, slideWidth, slideHeight)
            Return plan
        End Function

        Private Shared Sub NormalizeBounds(plan As SlideRenderPlan, slideWidth As Single, slideHeight As Single)
            If plan Is Nothing Then Return
            For Each node In plan.Nodes
                If node.Bounds Is Nothing Then node.Bounds = New SceneRect()
                node.Bounds.X = Math.Max(0, Math.Min(slideWidth, node.Bounds.X))
                node.Bounds.Y = Math.Max(0, Math.Min(slideHeight, node.Bounds.Y))
                node.Bounds.Width = Math.Max(0, Math.Min(slideWidth - node.Bounds.X, node.Bounds.Width))
                node.Bounds.Height = Math.Max(0, Math.Min(slideHeight - node.Bounds.Y, node.Bounds.Height))
            Next
        End Sub
    End Class

End Namespace
