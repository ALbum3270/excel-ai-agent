Imports System.Drawing
Imports System.Globalization
Imports Newtonsoft.Json.Linq

Namespace Design

    Public NotInheritable Class DesignSystemCatalog
        Private Sub New()
        End Sub

        Public Shared Function Resolve(name As String, Optional tokenOverrides As JObject = Nothing) As DesignTokens
            Dim result As DesignTokens
            Dim normalizedName = If(name, "").Trim().ToLowerInvariant()
            Select Case normalizedName
                Case "modern-tech", "tech", "科技"
                    result = ModernTech()
                Case "executive-light", "consulting", "咨询", "商务"
                    result = ExecutiveLight()
                Case "editorial-warm", "editorial", "人文", "温暖"
                    result = EditorialWarm()
                Case "executive-dark", "dark", "深色"
                    result = ExecutiveDark()
                Case Else
                    result = ExecutiveLight()
            End Select
            ApplyOverrides(result, tokenOverrides)
            If tokenOverrides IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(name) AndAlso Not IsSupported(name) Then
                result.Name = name.Trim()
            End If
            EnsureReadableContrast(result)
            Return result
        End Function

        Public Shared Function IsSupported(name As String) As Boolean
            Select Case If(name, "").Trim().ToLowerInvariant()
                Case "", "modern-tech", "tech", "科技",
                     "executive-light", "consulting", "咨询", "商务",
                     "editorial-warm", "editorial", "人文", "温暖",
                     "executive-dark", "dark", "深色"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Sub ApplyOverrides(tokens As DesignTokens, tokenOverrides As JObject)
            If tokens Is Nothing OrElse tokenOverrides Is Nothing Then Return
            Dim colors = TryCast(tokenOverrides("colors"), JObject)
            If colors Is Nothing Then colors = tokenOverrides
            tokens.Background = ReadColor(colors, "background", tokens.Background)
            tokens.Surface = ReadColor(colors, "surface", tokens.Surface)
            tokens.SurfaceAlt = ReadColor(colors, "surfaceAlt", tokens.SurfaceAlt)
            tokens.Primary = ReadColor(colors, "primary", tokens.Primary)
            tokens.Secondary = ReadColor(colors, "secondary", tokens.Secondary)
            tokens.TextPrimary = ReadColor(colors, "textPrimary", tokens.TextPrimary)
            tokens.TextSecondary = ReadColor(colors, "textSecondary", tokens.TextSecondary)
            tokens.Divider = ReadColor(colors, "divider", tokens.Divider)
            tokens.Positive = ReadColor(colors, "positive", tokens.Positive)
            tokens.Negative = ReadColor(colors, "negative", tokens.Negative)

            Dim typography = TryCast(tokenOverrides("typography"), JObject)
            If typography Is Nothing Then typography = tokenOverrides
            tokens.FontFamily = ReadString(typography, "fontFamily", tokens.FontFamily)
            tokens.DisplaySize = ReadSingle(typography, "displaySize", tokens.DisplaySize, 28, 60)
            tokens.TitleSize = ReadSingle(typography, "titleSize", tokens.TitleSize, 20, 42)
            tokens.BodySize = ReadSingle(typography, "bodySize", tokens.BodySize, 12, 24)
            tokens.CaptionSize = ReadSingle(typography, "captionSize", tokens.CaptionSize, 9.5F, 16)
        End Sub

        Private Shared Function ReadString(source As JObject, name As String, fallback As String) As String
            Dim value = source?.GetValue(name, StringComparison.OrdinalIgnoreCase)?.ToString()
            Return If(String.IsNullOrWhiteSpace(value), fallback, value.Trim())
        End Function

        Private Shared Function ReadColor(source As JObject, name As String, fallback As String) As String
            Dim value = ReadString(source, name, fallback)
            Try
                ColorTranslator.FromHtml(value)
                Return value
            Catch
                Return fallback
            End Try
        End Function

        Private Shared Function ReadSingle(source As JObject,
                                           name As String,
                                           fallback As Single,
                                           minimum As Single,
                                           maximum As Single) As Single
            Dim token = source?.GetValue(name, StringComparison.OrdinalIgnoreCase)
            If token Is Nothing Then Return fallback
            Dim value As Single
            If Single.TryParse(token.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, value) Then
                Return Math.Max(minimum, Math.Min(maximum, value))
            End If
            Return fallback
        End Function

        Private Shared Sub EnsureReadableContrast(tokens As DesignTokens)
            If tokens Is Nothing Then Return
            If MinimumContrast(tokens.TextPrimary, tokens.Background, tokens.Surface) < 4.5R Then
                tokens.TextPrimary = ChooseReadableText(tokens.Background, tokens.Surface)
            End If
            If MinimumContrast(tokens.TextSecondary, tokens.Background, tokens.Surface) < 3.0R Then
                tokens.TextSecondary = tokens.TextPrimary
            End If
        End Sub

        Private Shared Function ChooseReadableText(ParamArray backgrounds As String()) As String
            Dim light = MinimumContrast("#FFFFFF", backgrounds)
            Dim dark = MinimumContrast("#111827", backgrounds)
            Return If(light >= dark, "#FFFFFF", "#111827")
        End Function

        Private Shared Function MinimumContrast(foreground As String, ParamArray backgrounds As String()) As Double
            Dim result As Double = Double.MaxValue
            For Each background In backgrounds
                result = Math.Min(result, ContrastRatio(foreground, background))
            Next
            Return If(result = Double.MaxValue, 0, result)
        End Function

        Private Shared Function ContrastRatio(left As String, right As String) As Double
            Try
                Dim leftLuminance = RelativeLuminance(ColorTranslator.FromHtml(left))
                Dim rightLuminance = RelativeLuminance(ColorTranslator.FromHtml(right))
                Dim lighter = Math.Max(leftLuminance, rightLuminance)
                Dim darker = Math.Min(leftLuminance, rightLuminance)
                Return (lighter + 0.05R) / (darker + 0.05R)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function RelativeLuminance(color As Color) As Double
            Dim red = LinearColor(color.R / 255.0R)
            Dim green = LinearColor(color.G / 255.0R)
            Dim blue = LinearColor(color.B / 255.0R)
            Return 0.2126R * red + 0.7152R * green + 0.0722R * blue
        End Function

        Private Shared Function LinearColor(value As Double) As Double
            If value <= 0.03928R Then Return value / 12.92R
            Return Math.Pow((value + 0.055R) / 1.055R, 2.4R)
        End Function

        Private Shared Function ModernTech() As DesignTokens
            Return New DesignTokens With {
                .Name = "modern-tech", .Background = "#071521", .Surface = "#102A3D", .SurfaceAlt = "#163C56",
                .Primary = "#2DD4BF", .Secondary = "#38BDF8", .TextPrimary = "#F8FAFC", .TextSecondary = "#A9C0D2",
                .Divider = "#28506B", .Positive = "#34D399", .Negative = "#FB7185", .FontFamily = "Microsoft YaHei",
                .DisplaySize = 40, .TitleSize = 27, .BodySize = 16, .CaptionSize = 10.5F, .Dark = True
            }
        End Function

        Private Shared Function ExecutiveLight() As DesignTokens
            Return New DesignTokens With {
                .Name = "executive-light", .Background = "#F7F8FA", .Surface = "#FFFFFF", .SurfaceAlt = "#EDF1F5",
                .Primary = "#123B5D", .Secondary = "#D94F3D", .TextPrimary = "#17212B", .TextSecondary = "#667482",
                .Divider = "#D7DEE5", .Positive = "#16856B", .Negative = "#C64242", .FontFamily = "Microsoft YaHei",
                .DisplaySize = 38, .TitleSize = 26, .BodySize = 15.5F, .CaptionSize = 10.5F, .Dark = False
            }
        End Function

        Private Shared Function EditorialWarm() As DesignTokens
            Return New DesignTokens With {
                .Name = "editorial-warm", .Background = "#F5F0E7", .Surface = "#FFFCF6", .SurfaceAlt = "#E9DED0",
                .Primary = "#7B3F2C", .Secondary = "#C38A4A", .TextPrimary = "#2D2924", .TextSecondary = "#746A5E",
                .Divider = "#D7C9B9", .Positive = "#557A5A", .Negative = "#A84D45", .FontFamily = "Microsoft YaHei",
                .DisplaySize = 39, .TitleSize = 27, .BodySize = 16, .CaptionSize = 10.5F, .Dark = False
            }
        End Function

        Private Shared Function ExecutiveDark() As DesignTokens
            Return New DesignTokens With {
                .Name = "executive-dark", .Background = "#111318", .Surface = "#1D2028", .SurfaceAlt = "#292E39",
                .Primary = "#E6B85C", .Secondary = "#7EA6D8", .TextPrimary = "#F5F3EE", .TextSecondary = "#B7BBC5",
                .Divider = "#3A404D", .Positive = "#67C59B", .Negative = "#EE7B7B", .FontFamily = "Microsoft YaHei",
                .DisplaySize = 40, .TitleSize = 27, .BodySize = 16, .CaptionSize = 10.5F, .Dark = True
            }
        End Function
    End Class

End Namespace
