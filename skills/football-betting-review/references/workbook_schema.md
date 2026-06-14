# Workbook Schema

Use one row per bet ticket or intentionally tracked betting unit. Keep a parlay/串关 as one row unless the user explicitly asks to split legs.

## Normalized Fields

- `bet_date`: bet placement date/time if available.
- `recommendation_id`: id from the 推荐表 when the actual bet came from a recorded recommendation.
- `match`: competition or match label.
- `match_date`: kickoff date/time if available.
- `stage`: 小组赛, 淘汰赛, 联赛, 杯赛, 友谊赛, or other visible stage.
- `fixture`: team matchup, such as `巴西 vs 摩洛哥`.
- `market`: 亚盘/让球, 大小球, 胜平负, 波胆, 串关, or platform market text.
- `bet_kind`: one of `赛前`, `滚球`, or `串关`.
- `selection`: exact betting direction, such as `摩洛哥 +0.5/1`, `小2/2.5`, `巴西 2-1`.
- `strategy_tag`: strategy label from the recommendation skill when available.
- `stake`: actual RMB stake.
- `odds`: decimal odds.
- `result`: one of `未结算`, `赢`, `输`, `走水`, `半赢`, `半输`, `取消`.
- `pre_match_notes`: concise reason before the match.
- `post_match_review`: concise review after settlement.
- `notes`: ticket id, screenshot ambiguity, live minute, emotional/discipline context.

## Bet Kind Rules

- Use `串关` for parlays/accumulators even if all legs were selected before kickoff.
- Use `滚球` for in-play bets.
- Use `赛前` for single bets placed before kickoff.

## Strategy Tags

Use these preferred tags:

- `S01-稳健主仓`
- `S02-反热门受让小球`
- `S03-滚球纠偏`
- `S04-小仓高赔`
- `纪律问题-追单`
- `未知策略`

Do not invent a precise strategy tag when the screenshot alone cannot support it; use `未知策略` or leave blank.
