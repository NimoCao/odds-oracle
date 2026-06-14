import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const templateDir = path.join(repoRoot, "templates");
const outputDir = path.join(repoRoot, "outputs", "football_betting");
const rowCount = 300;

const TABLES = {
  data: {
    label: "数据表",
    sheet: "数据明细",
    template: path.join(templateDir, "足球数据表模板.xlsx"),
    output: path.join(outputDir, "足球数据表.xlsx"),
  },
  recommendation: {
    label: "推荐表",
    sheet: "推荐明细",
    template: path.join(templateDir, "足球推荐表模板.xlsx"),
    output: path.join(outputDir, "足球推荐表.xlsx"),
  },
  review: {
    label: "复盘表",
    sheet: "投注明细",
    template: path.join(templateDir, "足球复盘表模板.xlsx"),
    output: path.join(outputDir, "足球复盘表.xlsx"),
  },
};

const palette = {
  navy: "#17324D",
  teal: "#0F766E",
  amber: "#B45309",
  green: "#15803D",
  red: "#B91C1C",
  slate: "#475569",
  paleBlue: "#EAF3F8",
  paleGreen: "#EAF7EE",
  paleAmber: "#FFF7E6",
  paleRed: "#FDECEC",
  grid: "#D9E2EC",
  soft: "#F8FAFC",
  white: "#FFFFFF",
};

function argValue(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function hasFlag(name) {
  return process.argv.includes(name);
}

function colName(index) {
  let name = "";
  let n = index;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function rangeFor(colCount, rows) {
  return `A1:${colName(colCount)}${rows}`;
}

function styleTitle(sheet, range, title) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[title]];
  r.format.fill = palette.navy;
  r.format.font = { color: palette.white, bold: true, size: 16 };
  r.format.horizontalAlignment = "left";
  r.format.rowHeightPx = 40;
}

function styleHeader(range) {
  range.format.fill = palette.teal;
  range.format.font = { color: palette.white, bold: true };
  range.format.horizontalAlignment = "center";
  range.format.wrapText = true;
  range.format.rowHeightPx = 34;
}

function styleBlock(range, fill = palette.white) {
  range.format.fill = fill;
  range.format.borders = { preset: "all", style: "thin", color: palette.grid };
}

function setColumnWidths(sheet, widths, totalRows = 340) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}1:${col}${totalRows}`).format.columnWidthPx = width;
  }
}

function moneyFormat(sheet, range) {
  sheet.getRange(range).format.numberFormat = '"¥"#,##0.00;[Red]-"¥"#,##0.00;"¥"0.00';
}

function percentFormat(sheet, range) {
  sheet.getRange(range).format.numberFormat = '0.00%;[Red]-0.00%;0.00%';
}

function numberFormat(sheet, range, format) {
  sheet.getRange(range).format.numberFormat = format;
}

function writeGuide(workbook, rows) {
  const guide = workbook.worksheets.add("字段说明");
  styleTitle(guide, "A1:D1", "字段说明");
  guide.getRange("A3:D3").values = [["字段", "是否必填", "说明", "示例"]];
  styleHeader(guide.getRange("A3:D3"));
  guide.getRange(`A4:D${rows.length + 3}`).values = rows;
  styleBlock(guide.getRange(`A4:D${rows.length + 3}`));
  guide.getRange(`A4:D${rows.length + 3}`).format.wrapText = true;
  setColumnWidths(guide, { A: 150, B: 80, C: 420, D: 260 }, 80);
  guide.freezePanes.freezeRows(3);
}

function addParams(workbook) {
  const params = workbook.worksheets.add("参数");
  styleTitle(params, "A1:E1", "资金与选项参数");
  params.getRange("A3:B10").values = [
    ["起始本金", 1000],
    ["推荐默认仓位%", 0.03],
    ["推荐单场上限%", 0.1],
    ["滚球单场上限%", 0.05],
    ["彩票仓单场上限%", 0.03],
    ["高信心仓位%", 0.08],
    ["中信心仓位%", 0.05],
    ["低信心仓位%", 0.02],
  ];
  styleBlock(params.getRange("A3:B10"), palette.paleBlue);
  params.getRange("A3:A10").format.font = { bold: true, color: palette.navy };
  moneyFormat(params, "B3:B3");
  percentFormat(params, "B4:B10");
  params.getRange("D3:H3").values = [["策略标签", "场景", "投注类型", "投注结果", "建议动作"]];
  styleHeader(params.getRange("D3:H3"));
  params.getRange("D4:H12").values = [
    ["S01-稳健主仓", "赛前", "赛前", "未结算", "建议执行"],
    ["S02-反热门受让小球", "滚球", "滚球", "赢", "观望"],
    ["S03-滚球纠偏", "赛后", "串关", "输", "不下注"],
    ["S04-小仓高赔", "复盘", "", "走水", "对冲"],
    ["纪律问题-追单", "", "", "半赢", "止盈"],
    ["未知策略", "", "", "半输", ""],
    ["", "", "", "取消", ""],
    ["", "", "", "", ""],
    ["", "", "", "", ""],
  ];
  styleBlock(params.getRange("D4:H12"));
  setColumnWidths(params, { A: 150, B: 110, D: 170, E: 90, F: 90, G: 90, H: 100 }, 40);
  return params;
}

function buildDataWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(TABLES.data.sheet);
  const headers = [
    "id",
    "记录时间",
    "数据来源",
    "比赛ID",
    "联赛/赛事",
    "阶段/轮次",
    "开赛时间",
    "主队",
    "客队",
    "中立场",
    "主队近况",
    "客队近况",
    "阵容新闻",
    "伤停",
    "天气/场地",
    "欧赔",
    "亚盘",
    "大小球",
    "xG/数据摘要",
    "数据时间戳",
    "备注",
  ];
  styleTitle(sheet, "A1:U1", "足球数据表");
  sheet.getRange("A2:U2").values = [headers];
  styleHeader(sheet.getRange("A2:U2"));
  const formulas = [];
  for (let row = 3; row < 3 + rowCount; row += 1) {
    formulas.push([`=IF(OR(D${row}<>"",H${row}<>"",I${row}<>""),ROW()-2,"")`]);
  }
  sheet.getRange(`A3:A${2 + rowCount}`).formulas = formulas;
  styleBlock(sheet.getRange(`A3:U${2 + rowCount}`));
  sheet.getRange(`A3:U${2 + rowCount}`).format.wrapText = true;
  sheet.getRange(`A3:U${2 + rowCount}`).format.rowHeightPx = 34;
  numberFormat(sheet, `B3:B${2 + rowCount}`, "yyyy-mm-dd hh:mm");
  numberFormat(sheet, `G3:G${2 + rowCount}`, "yyyy-mm-dd hh:mm");
  numberFormat(sheet, `T3:T${2 + rowCount}`, "yyyy-mm-dd hh:mm");
  setColumnWidths(sheet, {
    A: 54, B: 140, C: 130, D: 110, E: 130, F: 110, G: 150, H: 110, I: 110, J: 74,
    K: 220, L: 220, M: 240, N: 180, O: 140, P: 130, Q: 150, R: 130, S: 260, T: 150, U: 220,
  });
  sheet.freezePanes.freezeRows(2);
  writeGuide(workbook, [
    ["比赛ID", "建议", "同一比赛的稳定编号；没有时可留空。", "EPL-2026-06-14-ARS-CHE"],
    ["欧赔/亚盘/大小球", "建议", "记录当前盘口快照，供推荐表引用。", "亚盘 主队 -0.5 @1.91"],
    ["xG/数据摘要", "可选", "近10场、实时xG、射门、控球、伤停等关键数据摘要。", "近10场xG 1.6/0.9，主队边路优势"],
  ]);
  return workbook;
}

function buildRecommendationWorkbook() {
  const workbook = Workbook.create();
  addParams(workbook);
  const sheet = workbook.worksheets.add(TABLES.recommendation.sheet);
  const headers = [
    "id",
    "推荐时间",
    "场景",
    "策略标签",
    "比赛",
    "开赛时间/当前分钟",
    "市场/玩法",
    "投注内容",
    "可接受盘口/赔率",
    "信心等级",
    "建议动作",
    "建议仓位%",
    "建议金额",
    "主要依据",
    "主要风险",
    "不下注/失效条件",
    "数据来源",
    "复盘状态",
    "备注",
  ];
  styleTitle(sheet, "A1:S1", "足球推荐表");
  sheet.getRange("A2:S2").values = [headers];
  styleHeader(sheet.getRange("A2:S2"));
  const idFormulas = [];
  const amountFormulas = [];
  for (let row = 3; row < 3 + rowCount; row += 1) {
    idFormulas.push([`=IF(E${row}="","",ROW()-2)`]);
    amountFormulas.push([`=IF(OR(E${row}="",K${row}="不下注",K${row}="观望",L${row}=""),"",'参数'!$B$3*L${row})`]);
  }
  sheet.getRange(`A3:A${2 + rowCount}`).formulas = idFormulas;
  sheet.getRange(`M3:M${2 + rowCount}`).formulas = amountFormulas;
  styleBlock(sheet.getRange(`A3:S${2 + rowCount}`));
  sheet.getRange(`A3:S${2 + rowCount}`).format.wrapText = true;
  sheet.getRange(`A3:S${2 + rowCount}`).format.rowHeightPx = 40;
  numberFormat(sheet, `B3:B${2 + rowCount}`, "yyyy-mm-dd hh:mm");
  percentFormat(sheet, `L3:L${2 + rowCount}`);
  moneyFormat(sheet, `M3:M${2 + rowCount}`);
  sheet.getRange(`K3:K${2 + rowCount}`).conditionalFormats.addCustom("=$K3=\"建议执行\"", {
    fill: palette.paleGreen,
    font: { color: palette.green, bold: true },
  });
  sheet.getRange(`K3:K${2 + rowCount}`).conditionalFormats.addCustom("=$K3=\"不下注\"", {
    fill: palette.paleRed,
    font: { color: palette.red, bold: true },
  });
  setColumnWidths(sheet, {
    A: 54, B: 140, C: 80, D: 170, E: 190, F: 150, G: 100, H: 180, I: 160, J: 90,
    K: 90, L: 90, M: 95, N: 260, O: 220, P: 240, Q: 200, R: 90, S: 220,
  });
  sheet.freezePanes.freezeRows(2);
  const strategy = workbook.worksheets.add("策略索引");
  styleTitle(strategy, "A1:E1", "策略索引");
  strategy.getRange("A3:E3").values = [["策略标签", "定位", "默认仓位", "适用场景", "禁止条件"]];
  styleHeader(strategy.getRange("A3:E3"));
  strategy.getRange("A4:E7").values = [
    ["S01-稳健主仓", "主仓、单关、保护盘", "3%-8%", "赛前亚盘/大小球/受让", "证据不足、价格已过、同场暴露超限"],
    ["S02-反热门受让小球", "强队热度过高时找受让和小球", "2%-6%", "弱队不被打穿、小比分脚本", "弱队防线伤停严重、强队全主力且价格合理"],
    ["S03-滚球纠偏", "比分错杀时的临场修正", "1%-5%", "滚球亚盘/大小球/对冲止盈", "没有实时数据、只是补亏、有效期已过"],
    ["S04-小仓高赔", "波胆、比分、小额串关", "0.5%-2%", "高赔娱乐仓", "不能做主仓、不能连续加码"],
  ];
  styleBlock(strategy.getRange("A4:E7"));
  strategy.getRange("A4:E7").format.wrapText = true;
  setColumnWidths(strategy, { A: 170, B: 190, C: 90, D: 240, E: 320 }, 40);
  writeGuide(workbook, [
    ["场景", "必填", "赛前、滚球、赛后或复盘。", "赛前"],
    ["策略标签", "必填", "从策略索引中选择，便于复盘统计。", "S01-稳健主仓"],
    ["建议动作", "必填", "建议执行、观望、不下注、对冲、止盈。", "建议执行"],
    ["建议仓位%", "建议", "按本金比例填写，建议金额会自动计算。", "0.05"],
  ]);
  return workbook;
}

function buildReviewWorkbook() {
  const workbook = Workbook.create();
  addParams(workbook);
  const sheet = workbook.worksheets.add(TABLES.review.sheet);
  const headers = [
    "id",
    "投注日期",
    "推荐ID",
    "比赛",
    "比赛日期",
    "阶段",
    "对阵",
    "市场/玩法",
    "投注类型",
    "投注内容",
    "策略标签",
    "投注前本金",
    "投注金额",
    "投注赔率",
    "资金占比",
    "潜在返还",
    "投注结果",
    "结算金额",
    "投注收益",
    "ROI",
    "滚动本金",
    "赛前依据",
    "赛后复盘",
    "备注",
  ];
  styleTitle(sheet, "A1:X1", "足球复盘表");
  sheet.getRange("A2:X2").values = [headers];
  styleHeader(sheet.getRange("A2:X2"));
  const formulaRows = [];
  for (let row = 3; row < 3 + rowCount; row += 1) {
    formulaRows.push([
      `=IF(D${row}="","",ROW()-2)`,
      "", "", "", "", "", "", "", "", "", "",
      `=IF(M${row}="","",'参数'!$B$3+IF(ROW()=3,0,SUM($S$3:INDEX($S:$S,ROW()-1))))`,
      "", "",
      `=IF(OR(M${row}="",L${row}=""),"",M${row}/L${row})`,
      `=IF(OR(M${row}="",N${row}=""),"",M${row}*N${row})`,
      "",
      `=IF(OR(M${row}="",Q${row}="",Q${row}="未结算"),"",SWITCH(Q${row},"赢",M${row}*N${row},"输",0,"走水",M${row},"取消",M${row},"半赢",M${row}*(1+(N${row}-1)/2),"半输",M${row}*0.5,""))`,
      `=IF(R${row}="","",R${row}-M${row})`,
      `=IF(OR(M${row}="",S${row}=""),"",S${row}/M${row})`,
      `=IF(S${row}="","",'参数'!$B$3+SUM($S$3:S${row}))`,
      "", "", "",
    ]);
  }
  sheet.getRange(`A3:X${2 + rowCount}`).formulas = formulaRows;
  styleBlock(sheet.getRange(`A3:X${2 + rowCount}`));
  sheet.getRange(`A3:X${2 + rowCount}`).format.wrapText = true;
  sheet.getRange(`A3:X${2 + rowCount}`).format.rowHeightPx = 34;
  numberFormat(sheet, `B3:B${2 + rowCount}`, "yyyy-mm-dd hh:mm");
  numberFormat(sheet, `E3:E${2 + rowCount}`, "yyyy-mm-dd hh:mm");
  moneyFormat(sheet, `L3:M${2 + rowCount}`);
  numberFormat(sheet, `N3:N${2 + rowCount}`, "0.000");
  percentFormat(sheet, `O3:O${2 + rowCount}`);
  moneyFormat(sheet, `P3:S${2 + rowCount}`);
  percentFormat(sheet, `T3:T${2 + rowCount}`);
  moneyFormat(sheet, `U3:U${2 + rowCount}`);
  sheet.getRange(`S3:S${2 + rowCount}`).conditionalFormats.addCellIs({
    operator: "greaterThan",
    formula: 0,
    format: { fill: palette.paleGreen, font: { color: palette.green, bold: true } },
  });
  sheet.getRange(`S3:S${2 + rowCount}`).conditionalFormats.addCellIs({
    operator: "lessThan",
    formula: 0,
    format: { fill: palette.paleRed, font: { color: palette.red, bold: true } },
  });
  setColumnWidths(sheet, {
    A: 54, B: 140, C: 80, D: 150, E: 140, F: 90, G: 160, H: 100, I: 90, J: 180, K: 170,
    L: 95, M: 95, N: 80, O: 80, P: 95, Q: 90, R: 95, S: 95, T: 80, U: 95, V: 240, W: 240, X: 220,
  });
  sheet.freezePanes.freezeRows(2);
  const summary = workbook.worksheets.add("策略统计");
  styleTitle(summary, "A1:H1", "策略统计");
  summary.getRange("A3:H3").values = [["策略标签", "投注数", "已结算", "胜单", "总投注", "总收益", "ROI", "备注"]];
  styleHeader(summary.getRange("A3:H3"));
  const tags = ["S01-稳健主仓", "S02-反热门受让小球", "S03-滚球纠偏", "S04-小仓高赔", "纪律问题-追单", "未知策略"];
  const rows = tags.map((tag, index) => {
    const row = 4 + index;
    return [
      tag,
      `=COUNTIF('投注明细'!$K$3:$K$302,A${row})`,
      `=COUNTIFS('投注明细'!$K$3:$K$302,A${row},'投注明细'!$Q$3:$Q$302,"<>未结算",'投注明细'!$Q$3:$Q$302,"<>")`,
      `=COUNTIFS('投注明细'!$K$3:$K$302,A${row},'投注明细'!$Q$3:$Q$302,"赢")`,
      `=SUMIF('投注明细'!$K$3:$K$302,A${row},'投注明细'!$M$3:$M$302)`,
      `=SUMIF('投注明细'!$K$3:$K$302,A${row},'投注明细'!$S$3:$S$302)`,
      `=IF(E${row}=0,"",F${row}/E${row})`,
      "",
    ];
  });
  summary.getRange("A4:H9").formulas = rows;
  summary.getRange("A4:A9").values = tags.map((tag) => [tag]);
  styleBlock(summary.getRange("A4:H9"));
  moneyFormat(summary, "E4:F9");
  percentFormat(summary, "G4:G9");
  setColumnWidths(summary, { A: 190, B: 80, C: 80, D: 80, E: 100, F: 100, G: 80, H: 240 }, 40);
  writeGuide(workbook, [
    ["推荐ID", "可选", "来自推荐表的 id，方便推荐与实际下注串联。", "12"],
    ["投注类型", "必填", "赛前、滚球或串关。", "赛前"],
    ["策略标签", "建议", "来自推荐 skill 的策略标签。", "S03-滚球纠偏"],
    ["投注结果", "必填", "未结算、赢、输、走水、半赢、半输、取消。", "赢"],
  ]);
  return workbook;
}

async function saveWorkbook(workbook, target) {
  await fs.mkdir(path.dirname(target), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(target);
  return target;
}

async function copyIfNeeded(source, target, force) {
  await fs.mkdir(path.dirname(target), { recursive: true });
  if (!force) {
    try {
      await fs.access(target);
      return { target, copied: false };
    } catch {
      // continue
    }
  }
  await fs.copyFile(source, target);
  return { target, copied: true };
}

async function initTables() {
  const force = hasFlag("--force");
  const builders = {
    data: buildDataWorkbook,
    recommendation: buildRecommendationWorkbook,
    review: buildReviewWorkbook,
  };
  const results = {};
  for (const [key, table] of Object.entries(TABLES)) {
    const workbook = builders[key]();
    await saveWorkbook(workbook, table.template);
    const copied = await copyIfNeeded(table.template, table.output, force);
    results[key] = { template: table.template, output: table.output, outputUpdated: copied.copied };
  }
  return results;
}

function loadRecords() {
  const recordsJson = argValue("--records-json");
  if (recordsJson) {
    const parsed = JSON.parse(recordsJson);
    return Array.isArray(parsed) ? parsed : [parsed];
  }
  throw new Error("Provide --records-json with a JSON object or array.");
}

function firstPresent(raw, aliases) {
  for (const alias of aliases) {
    if (raw[alias] !== undefined && raw[alias] !== null && raw[alias] !== "") return raw[alias];
  }
  return "";
}

const ALIASES = {
  data: {
    "记录时间": ["record_time", "记录时间", "timestamp", "time"],
    "数据来源": ["source", "数据来源", "来源"],
    "比赛ID": ["match_id", "比赛ID", "id"],
    "联赛/赛事": ["competition", "联赛/赛事", "赛事", "league"],
    "阶段/轮次": ["stage", "阶段/轮次", "阶段", "round"],
    "开赛时间": ["kickoff", "开赛时间", "match_time"],
    "主队": ["home_team", "主队", "home"],
    "客队": ["away_team", "客队", "away"],
    "中立场": ["neutral", "中立场"],
    "主队近况": ["home_form", "主队近况"],
    "客队近况": ["away_form", "客队近况"],
    "阵容新闻": ["lineup_news", "阵容新闻"],
    "伤停": ["injuries", "伤停"],
    "天气/场地": ["weather", "天气/场地"],
    "欧赔": ["euro_odds", "欧赔"],
    "亚盘": ["asian_handicap", "亚盘"],
    "大小球": ["total_goals", "大小球"],
    "xG/数据摘要": ["stats_summary", "xG/数据摘要", "数据摘要"],
    "数据时间戳": ["data_timestamp", "数据时间戳"],
    "备注": ["notes", "备注"],
  },
  recommendation: {
    "推荐时间": ["recommend_time", "推荐时间", "time"],
    "场景": ["scenario", "场景"],
    "策略标签": ["strategy_tag", "策略标签", "strategy"],
    "比赛": ["fixture", "比赛", "match"],
    "开赛时间/当前分钟": ["match_time", "minute", "开赛时间/当前分钟"],
    "市场/玩法": ["market", "市场/玩法", "玩法"],
    "投注内容": ["selection", "投注内容", "pick"],
    "可接受盘口/赔率": ["acceptable_price", "可接受盘口/赔率"],
    "信心等级": ["confidence", "信心等级"],
    "建议动作": ["action", "建议动作"],
    "建议仓位%": ["stake_pct", "建议仓位%", "stake_percent"],
    "建议金额": ["suggested_amount", "建议金额"],
    "主要依据": ["rationale", "主要依据"],
    "主要风险": ["risk", "主要风险"],
    "不下注/失效条件": ["no_bet_conditions", "不下注/失效条件"],
    "数据来源": ["sources", "数据来源"],
    "复盘状态": ["review_status", "复盘状态"],
    "备注": ["notes", "备注"],
  },
  review: {
    "投注日期": ["bet_date", "投注日期", "下注日期"],
    "推荐ID": ["recommendation_id", "推荐ID"],
    "比赛": ["match", "比赛", "赛事"],
    "比赛日期": ["match_date", "比赛日期", "开赛时间"],
    "阶段": ["stage", "阶段"],
    "对阵": ["fixture", "对阵", "teams"],
    "市场/玩法": ["market", "市场/玩法", "玩法"],
    "投注类型": ["bet_kind", "投注类型", "bet_type"],
    "投注内容": ["selection", "投注内容", "pick"],
    "策略标签": ["strategy_tag", "策略标签", "strategy"],
    "投注金额": ["stake", "投注金额", "stake"],
    "投注赔率": ["odds", "投注赔率", "odds"],
    "投注结果": ["result", "投注结果", "result"],
    "赛前依据": ["pre_match_notes", "赛前依据"],
    "赛后复盘": ["post_match_review", "赛后复盘"],
    "备注": ["notes", "备注"],
  },
};

const FORMULA_HEADERS = {
  data: new Set(["id"]),
  recommendation: new Set(["id", "建议金额"]),
  review: new Set(["id", "投注前本金", "资金占比", "潜在返还", "结算金额", "投注收益", "ROI", "滚动本金"]),
};

const EMPTY_ROW_KEYS = {
  data: ["比赛ID", "主队", "客队"],
  recommendation: ["比赛"],
  review: ["比赛", "对阵"],
};

function normalizeRecord(tableKey, raw) {
  const aliases = ALIASES[tableKey];
  const normalized = {};
  for (const [header, names] of Object.entries(aliases)) {
    normalized[header] = firstPresent(raw, names);
  }
  return normalized;
}

async function ensureOutput(tableKey) {
  const table = TABLES[tableKey];
  try {
    await fs.access(table.output);
  } catch {
    await initTables();
  }
  return table.output;
}

async function appendRecords(tableKey) {
  const table = TABLES[tableKey];
  const workbookPath = argValue("--workbook") ?? await ensureOutput(tableKey);
  const records = loadRecords().map((record) => normalizeRecord(tableKey, record));
  const blob = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(blob);
  const sheet = workbook.worksheets.getItem(table.sheet);
  const headerValues = sheet.getRange("A2:AZ2").values[0].filter((value) => value !== null && value !== "");
  const dataRange = sheet.getRangeByIndexes(2, 0, rowCount, headerValues.length);
  const rows = dataRange.values;
  const keyIndexes = EMPTY_ROW_KEYS[tableKey]
    .map((header) => headerValues.indexOf(header))
    .filter((index) => index >= 0);
  const appendStart = rows.findIndex((row) => keyIndexes.every((index) => row[index] === "" || row[index] === null));
  if (appendStart < 0) throw new Error(`${table.label} has no empty rows left.`);
  const startRow = 3 + appendStart;
  records.forEach((record, recordIndex) => {
    headerValues.forEach((header, colIndex) => {
      if (FORMULA_HEADERS[tableKey].has(header)) return;
      const value = record[header];
      if (value === undefined || value === null || value === "") return;
      sheet.getRangeByIndexes(startRow - 1 + recordIndex, colIndex, 1, 1).values = [[value]];
    });
  });
  await saveWorkbook(workbook, workbookPath);
  return {
    table: table.label,
    workbook: workbookPath,
    appendedRows: records.map((_, index) => startRow + index),
    count: records.length,
  };
}

async function main() {
  const action = process.argv[2] ?? "help";
  let result;
  if (action === "init") {
    result = await initTables();
  } else if (action === "append") {
    const tableKey = argValue("--table");
    if (!TABLES[tableKey]) throw new Error("--table must be one of: data, recommendation, review");
    result = await appendRecords(tableKey);
  } else {
    result = {
      usage: [
        "node scripts/football_tables_tool.mjs init [--force]",
        "node scripts/football_tables_tool.mjs append --table data|recommendation|review --records-json '[{...}]'",
      ],
      tables: TABLES,
    };
  }
  console.log(JSON.stringify(result, null, 2,));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
