Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Drawing.Text
Imports System.IO

Namespace Design

    ''' <summary>
    ''' Raster preview for the same SlideRenderPlan consumed by the COM renderer.
    ''' This enables fast design review and future vision-model scoring without
    ''' requiring a live PowerPoint process.
    ''' </summary>
    Public NotInheritable Class ScenePlanPreviewRenderer
        Private Sub New()
        End Sub

        Public Shared Sub RenderToPng(plan As SlideRenderPlan,
                                      tokens As DesignTokens,
                                      outputPath As String,
                                      Optional sourceWidth As Single = 960,
                                      Optional sourceHeight As Single = 540,
                                      Optional scale As Single = 2)
            If plan Is Nothing Then Throw New ArgumentNullException(NameOf(plan))
            If String.IsNullOrWhiteSpace(outputPath) Then Throw New ArgumentNullException(NameOf(outputPath))
            Dim outputDirectory = Path.GetDirectoryName(outputPath)
            If Not String.IsNullOrWhiteSpace(outputDirectory) Then System.IO.Directory.CreateDirectory(outputDirectory)

            Dim pixelWidth = Math.Max(1, CInt(sourceWidth * scale))
            Dim pixelHeight = Math.Max(1, CInt(sourceHeight * scale))
            Using bitmap As New Bitmap(pixelWidth, pixelHeight, PixelFormat.Format32bppArgb)
                bitmap.SetResolution(144, 144)
                Using drawingGraphics As Graphics = System.Drawing.Graphics.FromImage(bitmap)
                    drawingGraphics.SmoothingMode = SmoothingMode.AntiAlias
                    drawingGraphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit
                    drawingGraphics.CompositingQuality = CompositingQuality.HighQuality
                    drawingGraphics.Clear(ParseColor(plan.Background, tokens.Background))
                    drawingGraphics.ScaleTransform(scale, scale)
                    For Each node In plan.Nodes
                        RenderNode(drawingGraphics, node, tokens)
                    Next
                End Using
                bitmap.Save(outputPath, ImageFormat.Png)
            End Using
        End Sub

        Private Shared Sub RenderNode(graphics As Graphics, node As SceneNode, tokens As DesignTokens)
            If node Is Nothing OrElse node.Bounds Is Nothing Then Return
            Dim rect As New RectangleF(node.Bounds.X, node.Bounds.Y, node.Bounds.Width, node.Bounds.Height)
            Select Case node.Kind
                Case "round-rect"
                    If node.Shadow Then
                        Dim shadowRect As New RectangleF(rect.X, rect.Y + 2, rect.Width, rect.Height)
                        Using shadowPath = RoundedRectangle(shadowRect, Math.Min(node.CornerRadius, Math.Min(rect.Width, rect.Height) * 0.2F))
                            Using shadowBrush As New SolidBrush(Color.FromArgb(38, 0, 0, 0))
                                graphics.FillPath(shadowBrush, shadowPath)
                            End Using
                        End Using
                    End If
                    Using path = RoundedRectangle(rect, Math.Min(node.CornerRadius, Math.Min(rect.Width, rect.Height) * 0.2F))
                        Using brush As New SolidBrush(WithTransparency(ParseColor(node.FillColor, tokens.Surface), node.FillTransparency))
                            graphics.FillPath(brush, path)
                        End Using
                        Using pen As New Pen(WithTransparency(ParseColor(node.LineColor, node.FillColor), 0.25F), Math.Max(0.5F, node.LineWeight))
                            graphics.DrawPath(pen, path)
                        End Using
                    End Using
                Case "rect"
                    Using brush As New SolidBrush(WithTransparency(ParseColor(node.FillColor, tokens.Surface), node.FillTransparency))
                        graphics.FillRectangle(brush, rect)
                    End Using
                Case "circle"
                    Using brush As New SolidBrush(WithTransparency(ParseColor(node.FillColor, tokens.Primary), node.FillTransparency))
                        graphics.FillEllipse(brush, rect)
                    End Using
                Case "line"
                    Using pen As New Pen(ParseColor(node.LineColor, tokens.Divider), Math.Max(0.5F, node.LineWeight))
                        graphics.DrawLine(pen, rect.X, rect.Y, rect.X + rect.Width, rect.Y + rect.Height)
                    End Using
                Case "image"
                    If Not String.IsNullOrWhiteSpace(node.ImagePath) AndAlso File.Exists(node.ImagePath) Then
                        Using image = System.Drawing.Image.FromFile(node.ImagePath)
                            DrawImageCropFill(graphics, image, rect)
                        End Using
                    End If
                Case Else
                    DrawText(graphics, node, tokens, rect)
            End Select
        End Sub

        Private Shared Sub DrawImageCropFill(graphics As Graphics,
                                             image As System.Drawing.Image,
                                             target As RectangleF)
            If image Is Nothing OrElse image.Width <= 0 OrElse image.Height <= 0 OrElse
               target.Width <= 0 OrElse target.Height <= 0 Then Return
            Dim source As New RectangleF(0, 0, image.Width, image.Height)
            Dim sourceAspect = image.Width / CDbl(image.Height)
            Dim targetAspect = target.Width / CDbl(target.Height)
            If sourceAspect > targetAspect Then
                source.Width = CSng(image.Height * targetAspect)
                source.X = (image.Width - source.Width) / 2.0F
            ElseIf sourceAspect < targetAspect Then
                source.Height = CSng(image.Width / targetAspect)
                source.Y = (image.Height - source.Height) / 2.0F
            End If
            graphics.DrawImage(image, target, source, GraphicsUnit.Pixel)
        End Sub

        Private Shared Sub DrawText(graphics As Graphics, node As SceneNode, tokens As DesignTokens, rect As RectangleF)
            If String.IsNullOrWhiteSpace(node.Text) Then Return
            Dim selectedFontStyle As FontStyle = If(node.Bold, FontStyle.Bold, FontStyle.Regular)
            Using font As New Font(tokens.FontFamily, Math.Max(8, node.FontSize), selectedFontStyle, GraphicsUnit.Pixel)
                Using brush As New SolidBrush(ParseColor(node.TextColor, tokens.TextPrimary))
                    Using format As New StringFormat(StringFormatFlags.LineLimit)
                        format.Trimming = StringTrimming.EllipsisCharacter
                        format.LineAlignment = StringAlignment.Near
                        Select Case If(node.Alignment, "left").ToLowerInvariant()
                            Case "center" : format.Alignment = StringAlignment.Center
                            Case "right" : format.Alignment = StringAlignment.Far
                            Case Else : format.Alignment = StringAlignment.Near
                        End Select
                        graphics.DrawString(node.Text, font, brush, rect, format)
                    End Using
                End Using
            End Using
        End Sub

        Private Shared Function RoundedRectangle(rect As RectangleF, radius As Single) As GraphicsPath
            Dim diameter = Math.Max(1, radius * 2)
            Dim path As New GraphicsPath()
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90)
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90)
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90)
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90)
            path.CloseFigure()
            Return path
        End Function

        Private Shared Function ParseColor(value As String, fallback As String) As Color
            Try
                Return ColorTranslator.FromHtml(If(String.IsNullOrWhiteSpace(value), fallback, value))
            Catch
                Return ColorTranslator.FromHtml(If(String.IsNullOrWhiteSpace(fallback), "#FFFFFF", fallback))
            End Try
        End Function

        Private Shared Function WithTransparency(color As Color, transparency As Single) As Color
            Dim alpha = CInt(255 * (1 - Math.Max(0, Math.Min(1, transparency))))
            Return Color.FromArgb(alpha, color.R, color.G, color.B)
        End Function
    End Class

End Namespace
