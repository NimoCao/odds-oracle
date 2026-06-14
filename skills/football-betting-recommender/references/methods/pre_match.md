# Pre-Match Method

Use this method for matches that have not started, including 赛前推荐, 盘口分析, 胜平负, 亚盘/让球, 大小球, and pre-match串关 leg selection.

## Workflow

1. Identify the fixture, kickoff time, competition, neutral/home context, and requested market scope.
2. Gather current information unless the user provided it and explicitly said not to browse:
   - Recent official-match form, preferably each team's last 10 meaningful official matches.
   - Lineup validity in that sample, manager changes, injuries, suspensions, expected XI, rotation and motivation.
   - Tactical matchup: tempo, chance creation route, defensive weakness, set pieces, transition exposure, pressing, finishing volatility.
   - Current odds and line movement for 胜平负, 亚盘/让球, and 大小球.
3. Build a data-led baseline before reading the market. If samples are weak or news is uncertain, lower confidence.
4. Analyze the three market groups separately:
   - `胜平负`: implied probability, draw risk, public heat, and price value.
   - `亚盘/让球`: whether the line protects the stronger side, invites public money, or creates value on the underdog.
   - `大小球`: expected pace, shot quality, finishing profile, defensive structure, and total-goal line movement.
5. Load the selected strategy file and apply its entry, stake, and forbidden-condition rules.
6. For formal recommendations, append the evidence snapshot to the 数据表 and the decision to the 推荐表 using `/Users/gangzi/Documents/交易大师/scripts/football_tables_tool.mjs`. Skip workbook writes only when the user explicitly says not to record.

## Recommendation Matrix

For each relevant market, produce:

- `推荐`: actionable pick, 观望, or 不下注.
- `理由`: 2-4 concrete signals.
- `可接受价格`: the minimum odds or line range where the recommendation remains valid.
- `主要风险`: the most likely way the trade fails.
- `信心等级`: 关键, 极高, 非常高, 高, 中, 低, or 不下注.

## Staking Defaults

Use the strategy file's staking rules first. If the strategy file does not specify a number:

- `高`: 5%-8% of current bankroll.
- `中`: 3%-5%.
- `低`: 1%-2%.
- `观望/不下注`: 0%.

Keep same-match exposure within the selected strategy's cap. If multiple markets are valid, allocate across them instead of treating each as independent risk.

## Report Shape

Start with the actionable conclusion, then explain:

```text
使用策略：
主推：
备选：
不建议：
可接受盘口/赔率：
建议仓位：
基本面：
最新新闻：
盘口走势：
风险：
复盘标签：
```

Keep recommendations concise enough to be copied into a betting review later.
