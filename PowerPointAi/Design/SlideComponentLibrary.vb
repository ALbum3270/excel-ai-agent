Imports System.Drawing

Namespace Design

    Public NotInheritable Partial Class SlideComponentLibrary
        Private Sub New()
        End Sub

        Public Shared Function Build(spec As SlideDesignSpec,
                                     tokens As DesignTokens,
                                     slideWidth As Single,
                                     slideHeight As Single,
                                     slideIndex As Integer,
                                     totalSlides As Integer) As SlideRenderPlan
            Dim plan As New SlideRenderPlan With {
                .SlideType = spec.SlideType,
                .Background = tokens.Background,
                .Notes = spec.Notes
            }
            Select Case spec.SlideType
                Case "cover" : BuildCover(plan, spec, tokens, slideWidth, slideHeight)
                Case "section" : BuildSection(plan, spec, tokens, slideWidth, slideHeight)
                Case "statement" : BuildStatement(plan, spec, tokens, slideWidth, slideHeight)
                Case "two-column" : BuildTwoColumn(plan, spec, tokens, slideWidth, slideHeight)
                Case "comparison" : BuildComparison(plan, spec, tokens, slideWidth, slideHeight)
                Case "kpi" : BuildKpi(plan, spec, tokens, slideWidth, slideHeight)
                Case "process" : BuildProcess(plan, spec, tokens, slideWidth, slideHeight)
                Case "architecture" : BuildArchitecture(plan, spec, tokens, slideWidth, slideHeight)
                Case "matrix" : BuildMatrix(plan, spec, tokens, slideWidth, slideHeight)
                Case "quote" : BuildQuote(plan, spec, tokens, slideWidth, slideHeight)
                Case "closing" : BuildClosing(plan, spec, tokens, slideWidth, slideHeight)
                Case Else : BuildContent(plan, spec, tokens, slideWidth, slideHeight)
            End Select
            If spec.SlideType <> "cover" AndAlso spec.SlideType <> "closing" Then
                AddFooter(plan, spec, tokens, slideWidth, slideHeight, slideIndex, totalSlides)
            End If
            Return plan
        End Function

        Private Shared Sub BuildCover(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddRect(plan, "cover-panel", w * 0.64F, 0, w * 0.36F, h, t.Surface, collision:=False)
            AddRect(plan, "cover-accent", w * 0.64F, 0, 8, h, t.Primary, collision:=False)
            If Not String.IsNullOrWhiteSpace(spec.ImagePath) Then
                AddImage(plan, "cover-image", spec.ImagePath, w * 0.67F, 42, w * 0.3F, h - 84)
            Else
                AddRect(plan, "cover-module-1", w * 0.71F, 84, w * 0.21F, 72, t.SurfaceAlt, t.Divider, collision:=False)
                AddRect(plan, "cover-module-2", w * 0.75F, 190, w * 0.18F, 72, t.Primary, t.Primary, collision:=False)
                AddRect(plan, "cover-module-3", w * 0.69F, 296, w * 0.24F, 72, t.Secondary, t.Secondary, collision:=False)
                AddLine(plan, "cover-flow-1", w * 0.81F, 156, w * 0.81F, 190, t.Divider, 2)
                AddLine(plan, "cover-flow-2", w * 0.81F, 262, w * 0.81F, 296, t.Divider, 2)
            End If
            AddText(plan, "eyebrow", spec.Eyebrow, 54, 58, w * 0.51F, 22, t.CaptionSize, t.Primary, True)
            AddText(plan, "title", spec.Title, 54, 105, w * 0.56F, 160, t.DisplaySize, t.TextPrimary, True)
            AddText(plan, "subtitle", FirstNonEmpty(spec.Subtitle, spec.KeyMessage), 56, 285, w * 0.51F, 85, t.BodySize + 1, t.TextSecondary, False)
            AddLine(plan, "cover-rule", 56, h - 92, w * 0.34F, h - 92, t.Divider, 1.4F)
            AddText(plan, "cover-note", spec.Body, 56, h - 74, w * 0.45F, 20, t.CaptionSize, t.TextSecondary, False)
        End Sub

        Private Shared Sub BuildSection(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddText(plan, "section-number", spec.SectionNumber, 52, 70, 220, 110, 72, t.SurfaceAlt, True)
            AddRect(plan, "section-rule", 58, 205, 90, 6, t.Primary, collision:=False)
            AddText(plan, "section-title", spec.Title, 58, 235, w * 0.75F, 92, t.DisplaySize, t.TextPrimary, True)
            AddText(plan, "section-subtitle", FirstNonEmpty(spec.Subtitle, spec.KeyMessage, spec.Body), 60, 342, w * 0.64F, 70, t.BodySize + 1, t.TextSecondary, False)
        End Sub

        Private Shared Sub BuildStatement(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddText(plan, "eyebrow", spec.Eyebrow, 58, 50, 300, 20, t.CaptionSize, t.Primary, True)
            AddText(plan, "statement", FirstNonEmpty(spec.KeyMessage, spec.Title), 58, 105, w * 0.78F, 170, t.DisplaySize + 2, t.TextPrimary, True)
            AddRect(plan, "statement-accent", 58, 306, 7, 95, t.Primary, collision:=False)
            AddText(plan, "statement-body", FirstNonEmpty(spec.Body, spec.Subtitle), 84, 312, w * 0.68F, 92, t.BodySize + 1, t.TextSecondary, False)
            AddRect(plan, "statement-rail", w - 74, 96, 8, 176, t.Secondary, collision:=False)
            Dim evidence = spec.Items.Take(3).ToList()
            If evidence.Count > 0 Then
                Dim margin As Single = 84, gap As Single = 14, top As Single = 410
                Dim itemWidth = (w - margin * 2 - gap * (evidence.Count - 1)) / evidence.Count
                For index = 0 To evidence.Count - 1
                    Dim x = margin + index * (itemWidth + gap)
                    AddRect(plan, $"evidence-{index + 1}", x, top, itemWidth, 70, t.Surface, t.Divider)
                    AddCircle(plan, $"evidence-dot-{index + 1}", x + 16, top + 17, 12, If(index = 0, t.Primary, t.Secondary), 0)
                    AddText(plan, $"evidence-text-{index + 1}", ComposeItemText(evidence(index)), x + 38, top + 13, itemWidth - 52, 46, t.CaptionSize, t.TextSecondary, False)
                Next
            End If
        End Sub

        Private Shared Sub BuildContent(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            Dim items = spec.Items.Take(6).ToList()
            If spec.Chart IsNot Nothing Then
                BuildChartContent(plan, spec, items, t, w, h)
                Return
            End If
            If spec.Table IsNot Nothing Then
                BuildTableContent(plan, spec, items, t, w, h)
                Return
            End If
            If Not String.IsNullOrWhiteSpace(spec.ImagePath) Then
                BuildImageContent(plan, spec.ImagePath, items, t, w, h)
                Return
            End If
            If items.Count >= 3 AndAlso
               (String.Equals(spec.LayoutVariant, "feature-left", StringComparison.OrdinalIgnoreCase) OrElse
                items.Any(Function(item) item.Emphasis)) Then
                BuildFeatureContent(plan, items, t, w, h)
                Return
            End If
            If items.All(Function(item) String.IsNullOrWhiteSpace(item.Body)) Then
                BuildEditorialListContent(plan, items, t, w, h)
                Return
            End If
            Dim columns = If(items.Count <= 3,
                             Math.Max(1, items.Count),
                             If(items.Count = 4, 2, 3))
            Dim rows = CInt(Math.Ceiling(items.Count / CDbl(columns)))
            Dim margin As Single = 58, gap As Single = 16, top As Single = 162, bottom As Single = 54
            Dim cardW = (w - margin * 2 - gap * (columns - 1)) / columns
            Dim cardH = (h - top - bottom - gap * Math.Max(0, rows - 1)) / Math.Max(1, rows)
            For index = 0 To items.Count - 1
                Dim col = index Mod columns, row = index \ columns
                AddCard(plan, $"card-{index + 1}", items(index), margin + col * (cardW + gap), top + row * (cardH + gap), cardW, cardH, t, index)
            Next
        End Sub

        Private Shared Sub BuildChartContent(plan As SlideRenderPlan,
                                             spec As SlideDesignSpec,
                                             items As List(Of DesignItem),
                                             t As DesignTokens,
                                             w As Single,
                                             h As Single)
            Dim chart = spec.Chart
            Dim left As Single = 58, top As Single = 166, bottom As Single = 58, gap As Single = 22
            Dim chartWidth = (w - left * 2 - gap) * 0.69F
            Dim chartHeight = h - top - bottom - 22
            AddRect(plan, "chart-frame", left, top, chartWidth, chartHeight, t.Surface, t.Divider, collision:=False)
            AddText(plan, "chart-title", chart.Title, left + 22, top + 10, chartWidth * 0.46F, 38,
                    t.CaptionSize + 1, t.TextPrimary, True)

            Dim legendX = left + chartWidth * 0.5F
            Dim legendWidth = (chartWidth * 0.46F) / Math.Max(1, chart.Series.Count)
            For seriesIndex = 0 To chart.Series.Count - 1
                If String.IsNullOrWhiteSpace(chart.Series(seriesIndex).Name) Then Continue For
                Dim color = ResolveSeriesColor(chart.Series(seriesIndex), seriesIndex, t)
                Dim x = legendX + seriesIndex * legendWidth
                AddPlainRect(plan, $"chart-legend-swatch-{seriesIndex + 1}", x, top + 16, 10, 10, color)
                AddText(plan, $"chart-legend-{seriesIndex + 1}", chart.Series(seriesIndex).Name,
                        x + 16, top + 9, legendWidth - 20, 38, t.CaptionSize, t.TextSecondary, False)
            Next

            Dim plotX = left + 52
            Dim plotY = top + 68
            Dim plotWidth = chartWidth - 72
            Dim plotHeight = chartHeight - 122
            Dim baseline = plotY + plotHeight
            Dim values = chart.Series.SelectMany(Function(series) series.Values).ToList()
            Dim minimum = Math.Min(0.0R, values.Min())
            Dim maximum = Math.Max(0.0R, values.Max())
            If Math.Abs(maximum - minimum) < 0.000001R Then maximum = minimum + 1.0R
            Dim zeroY = ChartValueToY(0, plotY, plotHeight, minimum, maximum)
            For gridIndex = 0 To 4
                Dim y = plotY + plotHeight * gridIndex / 4.0F
                AddLine(plan, $"chart-grid-{gridIndex}", plotX, y, plotX + plotWidth, y, t.Divider, 0.6F)
            Next
            AddLine(plan, "chart-zero-axis", plotX, zeroY, plotX + plotWidth, zeroY, t.TextSecondary, 1.1F)
            If maximum > 0 Then
                AddText(plan, "chart-scale-max", FormatChartValue(maximum, chart.ValueSuffix),
                        left + 4, plotY - 6, 44, 16, 9.5F, t.TextSecondary, False, "right")
            End If
            If minimum < 0 Then
                AddText(plan, "chart-scale-min", FormatChartValue(minimum, chart.ValueSuffix),
                        left + 4, baseline - 8, 44, 16, 9.5F, t.TextSecondary, False, "right")
            End If
            Dim zeroLabelHasClearance = minimum >= 0 OrElse maximum <= 0 OrElse
                                        (zeroY > plotY + 18 AndAlso zeroY < baseline - 18)
            If zeroLabelHasClearance Then
                AddText(plan, "chart-scale-zero", FormatChartValue(0, chart.ValueSuffix),
                        left + 4, zeroY - 8, 44, 16, 9.5F, t.TextSecondary, False, "right")
            End If

            If String.Equals(chart.ChartType, "line", StringComparison.OrdinalIgnoreCase) Then
                BuildLineChart(plan, chart, t, plotX, plotY, plotWidth, plotHeight, minimum, maximum)
            Else
                BuildColumnChart(plan, chart, t, plotX, plotY, plotWidth, plotHeight, minimum, maximum, zeroY)
            End If

            AddText(plan, "chart-source", chart.Source, left + 22, top + chartHeight + 5,
                    chartWidth - 44, 14, 9.5F, t.TextSecondary, False)
            BuildSideInsights(plan, items.Take(4).ToList(), t,
                              left + chartWidth + gap, top, w - left - (left + chartWidth + gap), chartHeight)
        End Sub

        Private Shared Sub BuildColumnChart(plan As SlideRenderPlan,
                                            chart As DesignChart,
                                            t As DesignTokens,
                                            plotX As Single,
                                             plotY As Single,
                                             plotWidth As Single,
                                             plotHeight As Single,
                                             minimum As Double,
                                             maximum As Double,
                                             zeroY As Single)
            Dim groupWidth = plotWidth / chart.Categories.Count
            Dim barRegionWidth = groupWidth * 0.68F
            Dim barWidth = barRegionWidth / chart.Series.Count
            Dim showValues = chart.Series.Count = 1 AndAlso chart.Categories.Count <= 6
            For categoryIndex = 0 To chart.Categories.Count - 1
                Dim groupX = plotX + categoryIndex * groupWidth
                For seriesIndex = 0 To chart.Series.Count - 1
                    Dim value = chart.Series(seriesIndex).Values(categoryIndex)
                    Dim valueY = ChartValueToY(value, plotY, plotHeight, minimum, maximum)
                    Dim barHeight = Math.Abs(valueY - zeroY)
                    Dim x = groupX + (groupWidth - barRegionWidth) / 2.0F + seriesIndex * barWidth
                    Dim y = Math.Min(valueY, zeroY)
                    If barHeight > 0.5F Then
                        AddPlainRect(plan, $"chart-bar-{categoryIndex + 1}-{seriesIndex + 1}",
                                     x + 1, y, Math.Max(3, barWidth - 2), barHeight,
                                     ResolveSeriesColor(chart.Series(seriesIndex), seriesIndex, t))
                    End If
                    If showValues Then
                        Dim labelY = If(value >= 0,
                                        Math.Max(plotY - 2, valueY - 18),
                                        Math.Min(plotY + plotHeight - 16, valueY + 2))
                        AddText(plan, $"chart-value-{categoryIndex + 1}-{seriesIndex + 1}",
                                FormatChartValue(value, chart.ValueSuffix),
                                x - 8, labelY, barWidth + 16, 16,
                                9.5F, t.TextSecondary, True, "center")
                    End If
                Next
                AddText(plan, $"chart-category-{categoryIndex + 1}", chart.Categories(categoryIndex),
                        groupX + 2, plotY + plotHeight + 8, groupWidth - 4, 30,
                        9.5F, t.TextSecondary, False, "center")
            Next
        End Sub

        Private Shared Sub BuildLineChart(plan As SlideRenderPlan,
                                          chart As DesignChart,
                                          t As DesignTokens,
                                          plotX As Single,
                                           plotY As Single,
                                           plotWidth As Single,
                                           plotHeight As Single,
                                           minimum As Double,
                                           maximum As Double)
            Dim stepWidth = plotWidth / Math.Max(1, chart.Categories.Count - 1)
            Dim labelWidth = Math.Min(78.0F, plotWidth / chart.Categories.Count)
            Dim showValues = chart.Series.Count = 1 AndAlso chart.Categories.Count <= 6
            For seriesIndex = 0 To chart.Series.Count - 1
                Dim color = ResolveSeriesColor(chart.Series(seriesIndex), seriesIndex, t)
                Dim previousX As Single = 0
                Dim previousY As Single = 0
                For categoryIndex = 0 To chart.Categories.Count - 1
                    Dim value = chart.Series(seriesIndex).Values(categoryIndex)
                    Dim x = plotX + categoryIndex * stepWidth
                    Dim y = ChartValueToY(value, plotY, plotHeight, minimum, maximum)
                    If categoryIndex > 0 Then
                        AddLine(plan, $"chart-line-{seriesIndex + 1}-{categoryIndex}",
                                previousX, previousY, x, y, color, 2.4F)
                    End If
                    AddCircle(plan, $"chart-point-{seriesIndex + 1}-{categoryIndex + 1}",
                              x - 5, y - 5, 10, color, 0)
                    If showValues Then
                        Dim labelY = If(value >= 0,
                                        Math.Max(plotY - 2, y - 22),
                                        Math.Min(plotY + plotHeight - 16, y + 4))
                        AddText(plan, $"chart-value-{seriesIndex + 1}-{categoryIndex + 1}",
                                FormatChartValue(value, chart.ValueSuffix),
                                x - 32, labelY, 64, 16,
                                9.5F, t.TextSecondary, True, "center")
                    End If
                    previousX = x
                    previousY = y
                Next
            Next
            For categoryIndex = 0 To chart.Categories.Count - 1
                Dim x = plotX + categoryIndex * stepWidth
                AddText(plan, $"chart-category-{categoryIndex + 1}", chart.Categories(categoryIndex),
                        x - labelWidth / 2.0F, plotY + plotHeight + 8, labelWidth, 30,
                        9.5F, t.TextSecondary, False, "center")
            Next
        End Sub

        Private Shared Function ChartValueToY(value As Double,
                                              plotY As Single,
                                              plotHeight As Single,
                                              minimum As Double,
                                              maximum As Double) As Single
            Dim range = maximum - minimum
            If range <= 0 Then Return plotY + plotHeight
            Return plotY + CSng((maximum - value) / range * plotHeight)
        End Function

        Private Shared Sub BuildTableContent(plan As SlideRenderPlan,
                                             spec As SlideDesignSpec,
                                             items As List(Of DesignItem),
                                             t As DesignTokens,
                                             w As Single,
                                             h As Single)
            Dim table = spec.Table
            Dim left As Single = 58, top As Single = 166, bottom As Single = 58, gap As Single = 22
            Dim tableWidth = (w - left * 2 - gap) * 0.69F
            Dim tableHeight = h - top - bottom - 22
            Dim titleHeight As Single = 42
            Dim headerHeight As Single = 40
            Dim gridTop = top + titleHeight
            Dim gridHeight = tableHeight - titleHeight
            Dim columnWidths = ResolveTableColumnWidths(table, tableWidth)
            Dim columnLefts As New List(Of Single)()
            Dim columnCursor = left
            For Each width In columnWidths
                columnLefts.Add(columnCursor)
                columnCursor += width
            Next
            Dim rowHeight = (gridHeight - headerHeight) / table.Rows.Count
            AddText(plan, "table-title", table.Title, left, top + 8, tableWidth, 26,
                    t.BodySize, t.TextPrimary, True)

            For columnIndex = 0 To table.Headers.Count - 1
                Dim x = columnLefts(columnIndex)
                Dim columnWidth = columnWidths(columnIndex)
                Dim highlighted = columnIndex = table.HighlightColumn
                AddPlainRect(plan, $"table-header-fill-{columnIndex + 1}", x, gridTop,
                             columnWidth, headerHeight, If(highlighted, t.Primary, t.SurfaceAlt))
                 AddText(plan, $"table-header-{columnIndex + 1}-title", table.Headers(columnIndex),
                         x + 10, gridTop + 9, columnWidth - 20, 22,
                         t.CaptionSize + 1, If(highlighted, t.Background, t.TextPrimary), True,
                         If(columnIndex = 0, "left", "center"))
            Next

            For rowIndex = 0 To table.Rows.Count - 1
                Dim y = gridTop + headerHeight + rowIndex * rowHeight
                For columnIndex = 0 To table.Headers.Count - 1
                    Dim x = columnLefts(columnIndex)
                    Dim columnWidth = columnWidths(columnIndex)
                    Dim highlighted = columnIndex = table.HighlightColumn
                    Dim fill = If(highlighted, t.SurfaceAlt, If(rowIndex Mod 2 = 0, t.Surface, t.Background))
                    AddPlainRect(plan, $"table-cell-fill-{rowIndex + 1}-{columnIndex + 1}",
                                 x, y, columnWidth, rowHeight, fill)
                    AddText(plan, $"table-cell-{rowIndex + 1}-{columnIndex + 1}",
                            table.Rows(rowIndex)(columnIndex), x + 10, y + 8,
                             columnWidth - 20, rowHeight - 14, t.CaptionSize + 1,
                             If(highlighted, t.TextPrimary, If(columnIndex = 0, t.TextPrimary, t.TextSecondary)),
                             columnIndex = 0 OrElse highlighted, If(columnIndex = 0, "left", "center"))
                Next
            Next

            For columnIndex = 0 To table.Headers.Count
                Dim x = If(columnIndex = table.Headers.Count,
                           left + tableWidth,
                           columnLefts(columnIndex))
                AddLine(plan, $"table-grid-v-{columnIndex}", x, gridTop, x, gridTop + gridHeight, t.Divider, 0.7F)
            Next
            AddLine(plan, "table-grid-header", left, gridTop + headerHeight,
                    left + tableWidth, gridTop + headerHeight, t.Divider, 0.9F)
            For rowIndex = 1 To table.Rows.Count
                Dim y = gridTop + headerHeight + rowIndex * rowHeight
                AddLine(plan, $"table-grid-h-{rowIndex}", left, y, left + tableWidth, y, t.Divider, 0.7F)
            Next
            AddLine(plan, "table-grid-top", left, gridTop, left + tableWidth, gridTop, t.Divider, 0.7F)
            AddText(plan, "table-source", table.Source, left, top + tableHeight + 5,
                    tableWidth, 14, 9.5F, t.TextSecondary, False)
            BuildSideInsights(plan, items.Take(4).ToList(), t,
                               left + tableWidth + gap, top, w - left - (left + tableWidth + gap), tableHeight)
        End Sub

        Private Shared Function ResolveTableColumnWidths(table As DesignTable,
                                                         totalWidth As Single) As List(Of Single)
            Dim weights As New List(Of Double)()
            For columnIndex = 0 To table.Headers.Count - 1
                Dim maximumUnits = EstimateLayoutTextUnits(table.Headers(columnIndex))
                For Each row In table.Rows
                    maximumUnits = Math.Max(maximumUnits, EstimateLayoutTextUnits(row(columnIndex)))
                Next
                Dim weight = Math.Max(0.8R, Math.Min(1.8R, 0.65R + maximumUnits / 14.0R))
                If columnIndex = 0 Then weight = Math.Max(1.15R, weight)
                weights.Add(weight)
            Next
            Dim totalWeight = Math.Max(0.1R, weights.Sum())
            Return weights.Select(Function(weight) CSng(totalWidth * weight / totalWeight)).ToList()
        End Function

        Private Shared Function EstimateLayoutTextUnits(value As String) As Double
            Dim units As Double = 0
            For Each character In If(value, "")
                If Char.IsWhiteSpace(character) Then
                    units += 0.32R
                ElseIf AscW(character) >= &H2E80 Then
                    units += 1.0R
                ElseIf Char.IsUpper(character) Then
                    units += 0.68R
                Else
                    units += 0.55R
                End If
            Next
            Return units
        End Function

        Private Shared Sub BuildSideInsights(plan As SlideRenderPlan,
                                             items As List(Of DesignItem),
                                             t As DesignTokens,
                                             x As Single,
                                             y As Single,
                                             width As Single,
                                             height As Single)
            Dim rowHeight = height / Math.Max(1, items.Count)
            For index = 0 To items.Count - 1
                Dim rowY = y + index * rowHeight
                AddText(plan, $"side-insight-index-{index + 1}", (index + 1).ToString("00"),
                        x, rowY + 10, 34, 20, t.CaptionSize, t.Primary, True)
                If items.Count >= 4 Then
                    AddText(plan, $"side-insight-{index + 1}-compact", ComposeItemText(items(index)),
                            x + 44, rowY + 8, width - 44, rowHeight - 16,
                            t.CaptionSize + 1, t.TextPrimary, index = 0)
                Else
                    AddText(plan, $"side-insight-{index + 1}-title", items(index).Title,
                            x + 44, rowY + 8, width - 44, 32, t.BodySize, t.TextPrimary, True)
                    AddText(plan, $"side-insight-{index + 1}-body", items(index).Body,
                            x + 44, rowY + 43, width - 44, rowHeight - 54,
                            t.CaptionSize, t.TextSecondary, False)
                End If
                If index < items.Count - 1 Then
                    AddLine(plan, $"side-insight-divider-{index + 1}", x + 44, rowY + rowHeight - 2,
                            x + width, rowY + rowHeight - 2, t.Divider, 0.7F)
                End If
            Next
        End Sub

        Private Shared Function ResolveSeriesColor(series As DesignChartSeries,
                                                   index As Integer,
                                                   t As DesignTokens) As String
            If series IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(series.Color) Then
                Try
                    ColorTranslator.FromHtml(series.Color)
                    Return series.Color
                Catch
                End Try
            End If
            Select Case index Mod 4
                Case 0 : Return t.Primary
                Case 1 : Return t.Secondary
                Case 2 : Return t.Positive
                Case Else : Return t.Negative
            End Select
        End Function

        Private Shared Function FormatChartValue(value As Double, suffix As String) As String
            Dim magnitude = Math.Abs(value)
            Dim formatted As String
            If magnitude > 0 AndAlso (magnitude < 0.001R OrElse magnitude >= 1000000.0R) Then
                formatted = value.ToString("G4", Globalization.CultureInfo.InvariantCulture)
            Else
                formatted = value.ToString("0.####", Globalization.CultureInfo.InvariantCulture)
            End If
            Return formatted & If(suffix, "")
        End Function



        Private Shared Sub AddHeader(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single)
            AddText(plan, "eyebrow", spec.Eyebrow, 58, 30, 360, 18, t.CaptionSize, t.Primary, True)
            AddText(plan, "title", spec.Title, 58, 58, w - 116, 55, t.TitleSize, t.TextPrimary, True)
            AddText(plan, "key-message", FirstNonEmpty(spec.KeyMessage, spec.Subtitle), 60, 117, w - 120, 34, t.BodySize, t.TextSecondary, False)
        End Sub

        Private Shared Sub AddFooter(plan As SlideRenderPlan,
                                     spec As SlideDesignSpec,
                                     t As DesignTokens,
                                     w As Single,
                                     h As Single,
                                     index As Integer,
                                     total As Integer)
            AddLine(plan, "footer-rule", 58, h - 31, w - 58, h - 31, t.Divider, 0.8F)
            AddText(plan, "footer-source", spec.Source, 58, h - 24, w * 0.62F, 14, 9.5F, t.TextSecondary, False)
            AddText(plan, "footer-page", (index + 1).ToString("00") & " / " & total.ToString("00"), w - 138, h - 24, 80, 14, 9.5F, t.TextSecondary, True, "right")
        End Sub

        Private Shared Sub AddCard(plan As SlideRenderPlan, id As String, item As DesignItem, x As Single, y As Single, w As Single, h As Single, t As DesignTokens, index As Integer)
            AddRect(plan, id, x, y, w, h, If(item.Emphasis, t.SurfaceAlt, t.Surface),
                    If(item.Emphasis, t.Primary, t.Divider), shadow:=item.Emphasis)
            AddText(plan, id & "-number", (index + 1).ToString("00"), x + 18, y + 15, 38, 20, t.CaptionSize, t.Primary, True)
            AddText(plan, id & "-title", item.Title, x + 18, y + 48, w - 36, 42, t.BodySize + 1, t.TextPrimary, True)
            AddText(plan, id & "-body", item.Body, x + 18, y + 98, w - 36, h - 116, t.CaptionSize + 1, t.TextSecondary, False)
        End Sub

        Private Shared Sub AddText(plan As SlideRenderPlan, id As String, text As String, x As Single, y As Single, w As Single, h As Single, size As Single, color As String, bold As Boolean, Optional alignment As String = "left")
            If String.IsNullOrWhiteSpace(text) Then Return
            plan.Nodes.Add(New SceneNode With {.Id = id, .Kind = "text", .Text = text, .Bounds = New SceneRect(x, y, w, h), .FontSize = size, .TextColor = color, .Bold = bold, .Alignment = alignment, .Collision = False})
        End Sub

        Private Shared Sub AddRect(plan As SlideRenderPlan,
                                   id As String,
                                   x As Single,
                                   y As Single,
                                   w As Single,
                                   h As Single,
                                   fill As String,
                                   Optional line As String = Nothing,
                                   Optional collision As Boolean = True,
                                   Optional shadow As Boolean = False)
            plan.Nodes.Add(New SceneNode With {
                .Id = id,
                .Kind = "round-rect",
                .Bounds = New SceneRect(x, y, w, h),
                .FillColor = fill,
                .LineColor = If(line, fill),
                .Collision = collision,
                .Shadow = shadow
            })
        End Sub

        Private Shared Sub AddPlainRect(plan As SlideRenderPlan,
                                        id As String,
                                        x As Single,
                                        y As Single,
                                        w As Single,
                                        h As Single,
                                        fill As String)
            plan.Nodes.Add(New SceneNode With {
                .Id = id,
                .Kind = "rect",
                .Bounds = New SceneRect(x, y, w, h),
                .FillColor = fill,
                .LineColor = "",
                .Collision = False
            })
        End Sub

        Private Shared Sub AddCircle(plan As SlideRenderPlan, id As String, x As Single, y As Single, diameter As Single, fill As String, transparency As Single)
            plan.Nodes.Add(New SceneNode With {.Id = id, .Kind = "circle", .Bounds = New SceneRect(x, y, diameter, diameter), .FillColor = fill, .FillTransparency = transparency, .LineColor = fill, .Collision = False})
        End Sub

        Private Shared Sub AddLine(plan As SlideRenderPlan, id As String, x1 As Single, y1 As Single, x2 As Single, y2 As Single, color As String, weight As Single)
            plan.Nodes.Add(New SceneNode With {.Id = id, .Kind = "line", .Bounds = New SceneRect(x1, y1, x2 - x1, y2 - y1), .LineColor = color, .LineWeight = weight, .Collision = False})
        End Sub

        Private Shared Sub AddImage(plan As SlideRenderPlan,
                                    id As String,
                                    imagePath As String,
                                    x As Single,
                                    y As Single,
                                    width As Single,
                                    height As Single)
            plan.Nodes.Add(New SceneNode With {
                .Id = id,
                .Kind = "image",
                .ImagePath = imagePath,
                .Bounds = New SceneRect(x, y, width, height),
                .Collision = True
            })
        End Sub

        Private Shared Function FirstNonEmpty(ParamArray values As String()) As String
            If values Is Nothing Then Return ""
            For Each value In values
                If Not String.IsNullOrWhiteSpace(value) Then Return value
            Next
            Return ""
        End Function

        Private Shared Function ComposeItemText(item As DesignItem) As String
            If item Is Nothing Then Return ""
            If String.IsNullOrWhiteSpace(item.Body) Then Return item.Title
            If String.IsNullOrWhiteSpace(item.Title) Then Return item.Body
            Return item.Title & "：" & item.Body
        End Function
    End Class

End Namespace
