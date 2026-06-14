---
name: football-betting-recommender
description: Generic football betting recommendation router for pre-match and live/in-play analysis. Use when the user asks for 足球下注建议, 赛前推荐, 临场/滚球建议, 亚盘/让球, 大小球, 胜平负, 串关, 波胆, 比分, 对冲, 止盈, 补仓, or asks to analyze a match using the user's betting strategy.
---

# Football Betting Recommender

## Overview

Use this skill as the single user-facing entry point for football betting recommendations. Route every request through one method file and one strategy file before producing advice.

Do not expose legacy recommendation skills to the user. This skill owns the recommendation flow.

## Routing

First classify the scenario:

- Read `references/methods/pre_match.md` for 赛前, 早盘, 初盘, 赛前推荐, 盘口分析, 胜平负, 亚盘, 让球, 大小球, 串关 legs, or any match that has not started.
- Read `references/methods/live_betting.md` for 滚球, 走地, 临场, 比赛中, 补仓, 对冲, 止盈, next goal, live handicap, live totals, or any match already in progress.
- If the user does not specify the scenario, infer from kickoff time and current match state. If the state is unclear and the answer depends on it, check current sources before deciding.

Then classify the user strategy:

- Read `references/strategies/S01_conservative-mainbook.md` for 稳健, 主仓, 默认, 保守, 单关, 亚盘, 受让保护, 大小球, or when the user does not name a strategy.
- Read `references/strategies/S02_value-dog-under.md` for 反热门, 受让, 弱队不败, 小球, 强队热, 不被打穿, or underdog value.
- Read `references/strategies/S03_live-correction.md` for 滚球纠偏, 错杀, 补仓, 比分误导, 落后方压制, or live market overreaction.
- Read `references/strategies/S04_small-stake-lottery.md` for 波胆, 比分, 高赔, 小博大, 娱乐仓, 彩票仓, or speculative parlays.

If several strategies match, choose the narrowest one for the requested action. For live requests, prefer `S03_live-correction.md` unless the user explicitly asks for 波胆 or 高赔.

## Data Discipline

Use current match data and cite sources when facts can change. Browse for latest news, lineups, odds, injuries, suspensions, weather, and live stats unless the user explicitly provides all needed facts and says not to browse.

For live betting, include the data timestamp and treat the recommendation as expiring quickly.

Do not present any bet as certain. Tie every recommendation to the current line, odds, and a maximum stake.

## Table Contract

The user-facing betting system has exactly three output workbooks:

- 数据表: `/Users/gangzi/Documents/交易大师/outputs/football_betting/足球数据表.xlsx`
- 推荐表: `/Users/gangzi/Documents/交易大师/outputs/football_betting/足球推荐表.xlsx`
- 复盘表: `/Users/gangzi/Documents/交易大师/outputs/football_betting/足球复盘表.xlsx`

Blank templates live in `/Users/gangzi/Documents/交易大师/templates/` and must not be edited during normal use.

Use this repository tool for all table writes:

```bash
/Users/gangzi/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node \
  /Users/gangzi/Documents/交易大师/scripts/football_tables_tool.mjs append \
  --table data|recommendation|review \
  --records-json '[{...}]'
```

For formal recommendations, append:

1. The source snapshot and key evidence to `--table data`.
2. The final actionable recommendation or no-bet decision to `--table recommendation`.

Do not update the 复盘表 from this skill unless the user is recording an actual placed bet; that belongs to `football-betting-review`.

If the user explicitly says not to record, only answer in chat and do not write any workbook.

## Output Contract

Lead with the action:

```text
使用策略：
当前结论：
可执行下注：
可接受盘口/赔率：
建议仓位：
不下注条件：
主要风险：
复盘标签：
```

For live betting, also include:

```text
数据时间：
当前分钟/比分：
实时数据是否支持：
建议有效期：
```

Use `观望/不下注` when the evidence is thin, the price is gone, the strategy's forbidden conditions are triggered, or the market conflicts with the data.
