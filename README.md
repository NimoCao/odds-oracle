# odds-oracle

通用足球投注推荐与复盘工作台。

用户只需要感知两件事：两个 Codex skill，三张 output 表。模板和脚本只是支撑它们工作的底层文件。

## Skills

- `football-betting-recommender`
  - 赛前推荐、滚球推荐、亚盘/让球、大小球、胜平负、串关、波胆、对冲、止盈、补仓判断。
  - 写入 `足球数据表.xlsx` 和 `足球推荐表.xlsx`。
  - 源文件在 `skills/football-betting-recommender/`。

- `football-betting-review`
  - 投注截图记录、文字票据入账、结算更新、策略复盘、ROI/纪律统计。
  - 写入 `足球复盘表.xlsx`。
  - 源文件在 `skills/football-betting-review/`。

`~/.codex/skills/` 下保留同名发现目录，目录内的 `SKILL.md`、`agents`、`references` 链接到本仓库的 `skills/` 源文件，方便 Codex 自动发现。

## Output Tables

用户最终使用的三张表都在 `outputs/football_betting/`：

- `足球数据表.xlsx`
- `足球推荐表.xlsx`
- `足球复盘表.xlsx`

这些 output 表是日常使用入口，会随着推荐和复盘持续更新。
仓库里提交的是空白初始 output 表；真实下注数据写入后，提交前需要自行确认是否包含隐私内容。

## Templates

空表模板放在 `templates/`：

- `足球数据表模板.xlsx`
- `足球推荐表模板.xlsx`
- `足球复盘表模板.xlsx`

模板只用于初始化，不在日常流程里直接编辑。

## Repository Layout

```text
skills/
  football-betting-recommender/
  football-betting-review/
scripts/
  football_tables_tool.mjs
templates/
  足球数据表模板.xlsx
  足球推荐表模板.xlsx
  足球复盘表模板.xlsx
outputs/football_betting/
  足球数据表.xlsx
  足球推荐表.xlsx
  足球复盘表.xlsx
```

## Tool

统一写表和初始化工具：

```bash
node scripts/football_tables_tool.mjs init
```

也可以通过 npm：

```bash
npm run init:football-tables
```

两个 skill 写表时都必须调用 `scripts/football_tables_tool.mjs`，不要绕过它直接改工作簿。
