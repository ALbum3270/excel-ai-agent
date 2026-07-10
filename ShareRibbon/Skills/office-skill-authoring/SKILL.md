---
name: office-skill-authoring
description: Use when creating or reviewing Office AI Skills. Enforces SKILL.md directory structure, concise YAML metadata, progressive disclosure, references, scripts, app scope, and allowed tool boundaries.
application: Excel,Word,PowerPoint
tags: skill, skills, authoring, SKILL.md, progressive-disclosure, harness, agent-loop
allowed-tools: memory.search, memory.list_recent
intent_types: skill_authoring, architecture
---

# Office Skill Authoring

Use this skill when adding or reviewing a Skill for Office AI Agent.

## Required Structure

Each directory-based Skill must use:

```text
skill-name/
├── SKILL.md
├── references/        optional detailed docs, loaded only after the skill is selected
├── scripts/           optional executable helpers
└── assets/            optional templates or examples
```

`SKILL.md` must start with YAML front matter. Required fields:

```yaml
---
name: skill-name
description: One concise sentence that explains when this skill should be used.
---
```

Recommended fields for this project:

```yaml
application: Excel
tags: excel, formula, chart
allowed-tools: ApplyFormula, CreateChart
intent_types: data_analysis, formula
```

## Rules

1. Keep the description action-oriented; it is used for first-pass skill selection.
2. Put long examples and domain rules in `references/` so the first prompt remains small.
3. Restrict `allowed-tools` to the smallest useful set.
4. State app scope clearly: Excel, Word, PowerPoint, or common.
5. A Skill should teach the agent how to decide and execute; it should not be a pile of keyword routes.
6. Scripts must live in `scripts/` and declare arguments in nearby docs when needed.
