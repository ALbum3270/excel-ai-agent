---
name: powerpoint-deck-agent
description: Use for PowerPoint tasks that need slide generation, deck structure, layout cleanup, speaker notes, visual consistency, translation, review, or chart/table insertion.
application: PowerPoint
tags: powerpoint, ppt, slide, deck, presentation, layout, theme, chart, notes, review
allowed-tools: CreateSlides, InsertSlide, FormatSlide, FormatText, InsertText, InsertTable, InsertChart, InsertImage, ApplyTheme, ApplyTransition, AddSpeakerNotes, BeautifySlides, ExecuteVBA
intent_types: slide_generation, formatting, review, translation, presentation
---

# PowerPoint Deck Agent

Use this skill when the user asks PowerPoint to create or improve slides.

## Operating Rules

1. Read current slide index, total slide count, selected shapes, text boxes, slide titles, and visible content before planning.
2. Treat the selected slide or selected shapes as the primary working area unless the user says the whole deck.
3. Prefer slide tools over VBA. Use `ExecuteVBA` only for gaps in registered tools.
4. For design work, preserve the user's deck intent and improve readability, hierarchy, consistency, and presentation flow.
5. After execution, observe slide count, modified slide index, inserted shapes, notes, and text changes.
6. If the result is not correct, repair the plan using observed slide state.

## Common Tasks

- Generate slides from an outline
- Beautify the current slide or whole deck
- Rewrite slide text for business reporting
- Add speaker notes
- Insert charts or tables
- Apply consistent theme, spacing, alignment, and transitions
