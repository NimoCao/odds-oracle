# Live Betting Method

Use this method for in-play football decisions: 滚球, 走地, 临场, 比赛中, 补仓, 对冲, 止盈, live handicap, live totals, next goal, and live scoreline overreaction.

## Workflow

1. Identify the exact fixture, minute, score, current market, and any existing user exposure.
2. Browse for current live data unless the user provides all current stats and says not to browse. Prefer live match centers with xG when available.
3. Record the data timestamp. Live data changes quickly.
4. Collect the live state:
   - Minute, score, half/full-time context.
   - xG by team and total xG, and whether xG is live or delayed.
   - Shots, shots on target, big chances, box touches, dangerous attacks, corners, possession or field tilt.
   - Red/yellow cards, injuries, substitutions, weather or interruption.
   - Current odds/line for the considered markets.
5. Compare the live state with pre-match expectation and existing exposure.
6. Load `S03_live-correction.md` by default, unless the user explicitly requests another strategy such as 波胆.
7. For formal live recommendations, append the live data snapshot to the 数据表 and the decision to the 推荐表 using `/Users/gangzi/Documents/交易大师/scripts/football_tables_tool.mjs`. Skip workbook writes only when the user explicitly says not to record.

## Decision Rules

Use xG and chance quality as the primary signal. Scoreline-only rules are alerts, not automatic bets.

- Strong live-over signal: total xG, shots, big chances, and tempo support the goals; both teams still attack; odds remain playable.
- Weak live-over signal: goals exceed xG, chances have dried up, or the leading team has slowed the match.
- Handicap correction signal: the market overreacts to the score while the losing or underpriced side is creating higher quality chances.
- No-bet signal: no xG, conflicting stats, unknown cards/substitutions, stale odds, or the desired price has moved.

## Existing Exposure Rule

Before recommending a live bet, ask: is this a new edge or just repairing a losing pre-match view?

Do not recommend same-direction追加仓 unless live data independently supports the position and the current price is still acceptable under the selected strategy.

## Report Shape

Lead with:

```text
使用策略：
当前建议：
数据时间：
当前分钟/比分：
实时数据是否支持：
可接受盘口/赔率：
建议仓位：
建议有效期：
不下注条件：
主要风险：
复盘标签：
```

Set a concrete expiry, for example: `只在35:00前、盘口仍为+0.5且赔率>=1.85时有效`.
