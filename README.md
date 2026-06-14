# odds-oracle

足球赔率抓取、赛前建议记录和投注复盘模板工具。

这个仓库只放公开安全的代码、文档和空模板。真实下注记录、赔率快照、预测写回、复盘文件、截图预览和备份文件统一放在 `data/`，并且默认被 `.gitignore` 排除。

## What is included

- `scripts/scrape_titan007_worldcup_odds.py`：抓取 Titan007 2026 世界杯赛程和赔率快照。
- `scripts/build_titan007_worldcup_odds_workbook.mjs`：把抓取到的 CSV/JSON 生成赔率工作簿和空的比赛预测表。
- `scripts/build_world_cup_review_template.mjs`：生成空的世界杯投注复盘模板。
- `scripts/update_prediction_staking_logic.mjs`：为比赛预测表补充/刷新资金分配和风控公式。
- `templates/`：可公开提交的空模板。
- `data/`：本地私有数据目录，不提交真实内容。

## Data privacy

Do not commit:

- 真实投注明细、盈亏、金额、赔率、复盘结论
- 已写入赛前建议的预测表
- Titan007 抓取快照、CSV、JSON、预览图片
- 工作簿备份文件
- 带具体下注建议或个人判断记录的私有脚本

这些文件都应放在 `data/` 下。公开仓库只保留空模板和可复用脚本。

## Usage

生成空复盘模板：

```bash
npm run build:review-template
```

抓取 Titan007 赔率数据：

```bash
npm run scrape:titan007
```

基于 `data/titan007_worldcup_2026_odds/` 里的抓取结果生成工作簿：

```bash
npm run build:odds-workbook
```

刷新比赛预测表的资金分配和风控公式：

```bash
npm run update:staking-logic
```

The workbook builders use `@oai/artifact-tool`, which is available in the Codex workspace runtime used to create this project.

