---
name: powerpoint-professional-design-agent
description: Use when PowerPoint must produce a polished business, consulting, product, strategy, technology, pitch, or executive presentation rather than a basic title-and-bullets deck. 适用于精美、高端、复杂、专业商业汇报和咨询风格 PPT。
application: PowerPoint
keywords: ppt, 精美, 精致, 好看, 漂亮, 美观, 高级, 高端, 专业设计, 咨询风格, 麦肯锡, 路演, 发布会, polished, beautiful, premium, high-end
tags: powerpoint, professional-design, pitch-deck, consulting, executive, architecture, visual-storytelling, premium, 专业, 专业设计, 精美, 精美PPT, 高端, 高端PPT, 复杂, 复杂PPT, 商业汇报, 咨询风格
allowed-tools: CreateSlides, InsertTable, DiscoverOfficeCapability, OfficeObjectOperation
intent_types: slide_generation, professional_design, visual_storytelling, pitch_deck, executive_presentation
---

# PowerPoint Professional Design Agent

Create presentation-ready slides through narrative planning, page archetype selection, one coherent design system, deterministic layout, and visual verification.

## Design Process

1. Read the current presentation size, theme, slide count, selected slide, existing colors, and available assets.
2. Convert the user's goal into a narrative arc. Every slide must have one conclusion, not merely a topic label.
3. Select one page archetype per slide: `cover`, `section`, `statement`, `content`, `two-column`, `comparison`, `kpi`, `process`, `architecture`, `matrix`, `quote`, or `closing`.
4. Choose one design system for the entire deck: `modern-tech`, `executive-light`, `executive-dark`, or `editorial-warm`.
   If the current deck or user provides brand colors and fonts, pass them through `designTokens` instead of forcing a preset palette.
5. Call `CreateSlides` once with a coherent deck-level Scene specification. Do not generate one basic title-and-content command per page.
6. Put a real accessible `imagePath` in the `CreateSlides` Scene when an image is required. Never create a fake image placeholder.
7. Use `DiscoverOfficeCapability` and `OfficeObjectOperation` only for long-tail Office objects that the Scene compiler or high-level tools do not cover.
   Every mutating object operation must declare observable `expectedEffects` or batch `successCriteria`; a COM call returning without an exception is not completion evidence.
8. Observe `visualVerification`, `slideResults`, and warnings. Repair overflow, high density, missing artifacts, or failed pages before reporting completion.
9. Never invent market statistics, customer numbers, benchmarks, or ROI figures. Every external metric must include `source`; if no reliable source is available, omit the number or explicitly label it as an assumption/illustrative estimate.
10. Do not add generic AI branding, template watermarks, decorative circles, or filler labels unless the user or active brand template requires them.
11. Never copy titles, numbers, labels, or claims from this Skill's examples. Every visible string must be grounded in the user's request, the active presentation, or an explicitly identified source.
12. Do not pad `items`, `steps`, `layers`, `metrics`, or other collections to reach a preferred layout count. Select a composition that fits the real information quantity.
13. For comparison tables, provide `columnHeaders` as `[dimension, left option, right option]`; do not let generic labels imply which alternative is current, traditional, recommended, or AI-driven.
14. If a visual unit cannot fit at its semantic minimum font size, shorten it, split the slide, or select another composition. Do not solve density by shrinking body text into caption sizes.
15. For a simple verified column or line chart, put a `chart` object directly on a `content` Scene so it is laid out, rendered, verified, and rolled back with the deck. Use 1-4 `items` for the conclusions the chart proves.
16. Chart categories must contain 2-8 labels, series must contain 1-3 finite numeric arrays of equal length, and external data must provide `chart.source`. Signed values are supported for change, variance, profit/loss, and other zero-baseline comparisons. Do not convert qualitative claims into invented numbers.
17. For a compact evidence table, put `table:{title,headers,rows,highlightColumn,source}` on a `content` Scene. Use 2-5 columns, 1-6 rows, a zero-based optional highlight column, and 1-4 conclusion items.
18. A content Scene may use only one main visual among `imagePath`, `chart`, and `table`. Split the slide when more than one main visual is required.
19. Every `matrix` Scene must provide concise semantic `xAxisLabel` and `yAxisLabel`; never assume the axes are effort and impact.
20. Use `architecture.variant: hub-spoke` only when the first item is the real core platform or capability and the remaining 2-4 items are its surrounding capabilities. Use the default architecture stack for actual layers.

## Slide Scene Contract

```json
{
  "designSystem": "executive-light",
  "slides": [
    {
      "slideType": "cover",
      "eyebrow": "PRODUCT STRATEGY",
      "title": "AI Native Office Agent",
      "subtitle": "From chat assistant to autonomous document execution"
    },
    {
      "slideType": "comparison",
      "title": "The two operating models optimize for different outcomes",
      "columnHeaders": ["Dimension", "Current model", "Target model"],
      "items": [
        {"label": "Decision flow", "features": ["User coordinates each step", "Agent plans and verifies the workflow"]},
        {"label": "Recovery", "features": ["Manual diagnosis", "Observe and repair loop"]},
        {"label": "Governance", "features": ["Process convention", "Safety-gated execution"]}
      ]
    },
    {
      "slideType": "process",
      "title": "The Agent closes the execution loop",
      "items": [
        {"title": "Context", "body": "Read the active document"},
        {"title": "Plan", "body": "Create a verifiable plan"},
        {"title": "Act", "body": "Execute native Office tools"},
        {"title": "Observe", "body": "Inspect the real result"},
        {"title": "Repair", "body": "Correct deviations automatically"}
      ]
    }
  ]
}
```

## Visual Standards

- Use conclusion-style titles and 3–5 visual units per slide.
- Keep one dominant focal point and a clear reading order.
- Use whitespace deliberately; do not fill every region.
- Avoid paragraphs longer than roughly 80 Chinese characters or 45 English words per visual unit.
- Use KPI for metrics, comparison for alternatives, process for sequences, architecture for layered systems, and matrix for two-dimensional prioritization.
- Use a `content` Scene with `chart.chartType: column|line` for verified quantitative comparisons or trends. Prefer editable Scene charts over screenshots of charts.
- Use a `content` Scene with `table` for compact evidence grids that need more than the three-column `comparison` contract. Keep cells concise enough to remain readable.
- Maintain the same colors, typography hierarchy, spacing rhythm, footer, and shape language across the deck.
- Use `variant: feature-left` or mark one item with `emphasis: true` when a content page needs a dominant focal point instead of an equal-card grid.
- Mark exactly one comparison item with `emphasis: true` only when the evidence supports a recommended or dominant option; the composition will allocate it stronger visual weight without inventing a verdict.
- Use `variant: hero-left` for KPI pages when one verified metric is the main conclusion and the remaining metrics are supporting evidence.
- Use `variant: vertical` for process pages with longer step descriptions; use the default horizontal timeline only when every step is concise.
- Use `variant: hub-spoke` for a core-and-capabilities architecture; keep the first item as the semantic core rather than selecting it for visual convenience.
- Vary composition intentionally across adjacent slides. Do not repeat the same card grid, header rhythm, or decorative motif throughout the deck.
- A successful tool call is not sufficient: the rendered slide must pass visual verification.
- Prefer real images, charts, logos, and cited data when accessible. If no real asset is available, use a structured information graphic and report the missing asset instead of drawing a fake placeholder.
