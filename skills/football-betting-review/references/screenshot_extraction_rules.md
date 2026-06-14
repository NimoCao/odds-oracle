# Screenshot Extraction Rules

Extract only visible information. Mark unclear fields in `notes`.

## Ticket Fields

Capture:

- Ticket/order id.
- Fixture and competition.
- Bet placement time and kickoff time.
- Market and selection.
- Handicap or total line.
- Odds.
- Stake.
- Result score.
- Settlement result and profit/loss if visible.
- Whether the ticket is single, live, or parlay.

## Result Normalization

- `赢 +X` or `已结算 赢`: result `赢`.
- `输 -X` or `已结算 输`: result `输`.
- `和 0.00`: result `走水`.
- `赢一半`: result `半赢`.
- `输一半`: result `半输`.
- `投注成功` without settlement: result `未结算`, unless a final score and settlement are visible.

## Market Normalization

- `全场让球`, `滚 全场让球`: market `亚盘/让球`.
- `全场大小`: market `大小球`.
- `全场独赢`: market `胜平负`.
- `全场波胆`: market `波胆`.
- `串关`, `2串1`, `3串1`: market `串关`; bet kind `串关`.

## Ambiguity

If a parlay leg is visible but the overall ticket odds are shown, keep the ticket as one record and summarize legs in `selection`.

If both stake and settlement are visible, store `stake` only as the cash amount risked. Do not store payout as stake.
