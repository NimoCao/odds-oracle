---
name: football-betting-review
description: Generic football betting review and ledger skill. Use when the user asks to record betting screenshots or text tickets, extract bet details, append or update a review workbook, settle results, summarize ROI, analyze strategy performance, or review football betting discipline.
---

# Football Betting Review

## Overview

Use this skill as the single user-facing entry point for football betting records and review. It handles screenshots, text tickets, settlement updates, workbook entry, and strategy performance summaries.

Do not expose legacy review skills to the user. This skill owns the review flow.

## Default Workbook

Prefer the user's named workbook only when they explicitly ask for a custom file. Otherwise use the user-facing output review workbook:

`/Users/gangzi/Documents/交易大师/outputs/football_betting/足球复盘表.xlsx`

Blank templates live in `/Users/gangzi/Documents/交易大师/templates/` and must not be edited during normal use.

## Workflow

1. Identify the source:
   - Screenshot/image: visually extract all visible bet details.
   - Text ticket: parse each distinct bet.
   - Settlement/result update: update an existing row only when the target row is unambiguous.
2. Read `references/screenshot_extraction_rules.md` when extracting from screenshots.
3. Normalize records with the schema in `references/workbook_schema.md`.
4. Ask one concise follow-up only if required fields cannot be inferred safely.
5. Use `/Users/gangzi/Documents/交易大师/scripts/football_tables_tool.mjs` to append records to the 复盘表. If extraction confidence is low, summarize the parsed records and ask before writing.
6. Verify appended rows, formula columns, result normalization, and backup creation.
7. Summarize:
   - Number of rows appended or updated.
   - Row numbers and bet summaries.
   - Backup path when created.
   - Fields inferred or left blank.

## Required Fields

To append a normal record, require:

- `match` or `fixture`
- `bet_kind`
- `selection`
- `stake`
- `odds`

Default `result` to `未结算` only when the ticket is clearly unsettled. Use the settled result shown in the ticket when visible.

## Table Tool Usage

```bash
/Users/gangzi/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node \
  /Users/gangzi/Documents/交易大师/scripts/football_tables_tool.mjs append \
  --table review \
  --records-json '[{"fixture":"巴西 vs 摩洛哥","market":"亚盘","bet_kind":"赛前","selection":"摩洛哥 +0.5/1","strategy_tag":"S02-反热门受让小球","stake":100,"odds":2.04,"result":"赢"}]'
```

## Review Discipline

When summarizing a batch, separate:

- Mainbook bets.
- Live correction bets.
- Lottery/high-odds bets.
- Chasing or rule-breaking bets.

When the user only asks for a summary and says not to record, do not write to the workbook.
