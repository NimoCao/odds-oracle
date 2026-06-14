import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function argValue(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const defaultBaseDir = path.join(repoRoot, "data/titan007_worldcup_2026_odds");
const baseDir = argValue("--base-dir") ?? (process.argv[2]?.startsWith("--") ? defaultBaseDir : process.argv[2]) ?? defaultBaseDir;
const dataOutputPath = argValue("--data-output") ?? argValue("--output") ?? `${baseDir}/titan007_worldcup_2026_odds.xlsx`;
const predictionOutputPath = argValue("--prediction-output") ?? `${baseDir}/比赛预测.xlsx`;

const palette = {
  navy: "#17324D",
  teal: "#0F766E",
  amber: "#B45309",
  paleBlue: "#EAF3F8",
  paleGreen: "#EAF7EE",
  paleAmber: "#FFF7E6",
  paleRose: "#FCE7F3",
  paleViolet: "#EEF2FF",
  grid: "#D9E2EC",
  white: "#FFFFFF",
};

const selectedCompanies = [
  { id: "8", label: "365", fill: "#FEF3C7" },
];

const preMatchTradingColumns = [
  "基本面摘要",
  "最新新闻",
  "士气预测",
  "盘口走势",
  "胜平负建议",
  "亚盘建议",
  "大小球建议",
  "信心等级",
  "赛前交易报告",
];

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;

  for (; i < text.length; i += 1) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch !== "\r") {
      cell += ch;
    }
  }
  if (cell !== "" || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

function makeBjtText(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):\d{2}(?:\.\d+)?\+08:00$/);
  if (!match) return value;
  return `${match[1]}-${match[2]}-${match[3]} ${match[4]}:${match[5]} BJT`;
}

function normalizeBjtColumn(rows) {
  if (!rows.length) return rows;
  const idx = rows[0].indexOf("match_time_bjt");
  if (idx < 0) return rows;
  for (let i = 1; i < rows.length; i += 1) {
    rows[i][idx] = makeBjtText(rows[i][idx]);
  }
  return rows;
}

function rowsToObjects(rows) {
  if (!rows.length) return [];
  const headers = rows[0];
  return rows
    .slice(1)
    .filter((row) => row.some((value) => String(value ?? "").trim() !== ""))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
}

function buildOddsIndex(rows) {
  const index = new Map();
  for (const row of rows) {
    index.set(`${row.match_id}:${row.company_id}`, row);
  }
  return index;
}

function fmtTriple(a, b, c) {
  if (![a, b, c].some((value) => String(value ?? "").trim() !== "")) return "";
  return [a, b, c].map((value) => String(value ?? "").trim() || "-").join("/");
}

function fmtHandicap(home, line, away) {
  if (![home, line, away].some((value) => String(value ?? "").trim() !== "")) return "";
  return [home, line, away].map((value) => String(value ?? "").trim() || "-").join(" ");
}

function fmtBaseline(row) {
  if (!row) return "";
  const asian = fmtHandicap(row.latest_asian_home, row.latest_asian_handicap, row.latest_asian_away);
  const goals = fmtHandicap(row.latest_over_odds, row.latest_goals_line, row.latest_under_odds);
  return [asian ? `亚 ${asian}` : "", goals ? `大小 ${goals}` : ""].filter(Boolean).join("\n");
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

function rangeFor(rows) {
  const rowCount = rows.length;
  const colCount = rows[0]?.length ?? 1;
  return `A1:${colName(colCount)}${rowCount}`;
}

function setValues(sheet, range, values) {
  sheet.getRange(range).values = values;
}

function setFormulas(sheet, range, formulas) {
  sheet.getRange(range).formulas = formulas;
}

function styleTitle(sheet, range, title) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[title]];
  r.format.fill = palette.navy;
  r.format.font = { color: palette.white, bold: true, size: 16 };
  r.format.horizontalAlignment = "left";
  r.format.rowHeightPx = 38;
}

function styleHeader(range) {
  range.format.fill = palette.teal;
  range.format.font = { color: palette.white, bold: true };
  range.format.horizontalAlignment = "center";
  range.format.wrapText = true;
  range.format.rowHeightPx = 34;
}

function styleGroupHeader(sheet, range, fill) {
  const r = sheet.getRange(range);
  r.merge();
  r.format.fill = fill;
  r.format.font = { color: palette.navy, bold: true };
  r.format.horizontalAlignment = "center";
  r.format.borders = { preset: "all", style: "thin", color: palette.grid };
}

function styleBlock(range, fill = "#F8FAFC") {
  range.format.fill = fill;
  range.format.borders = { preset: "all", style: "thin", color: palette.grid };
}

function setColumnWidths(sheet, widths, totalRows = 2500) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}1:${col}${totalRows}`).format.columnWidthPx = width;
  }
}

function writeTable(sheet, rows, options = {}) {
  if (!rows.length) return;
  const fullRange = rangeFor(rows);
  setValues(sheet, fullRange, rows);
  styleHeader(sheet.getRange(`A1:${colName(rows[0].length)}1`));
  sheet.getRange(fullRange).format.borders = { preset: "all", style: "thin", color: palette.grid };
  sheet.getRange(`A2:${colName(rows[0].length)}${rows.length}`).format.wrapText = Boolean(options.wrap);
  sheet.getRange(`A2:${colName(rows[0].length)}${rows.length}`).format.rowHeightPx = options.rowHeightPx ?? 24;
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

function writePredictionSheet(sheet, matchRecords, oddsWideRecords) {
  const oddsByMatchCompany = buildOddsIndex(oddsWideRecords);
  const width = 3 + selectedCompanies.length * 3 + preMatchTradingColumns.length;
  const lastCol = colName(width);

  styleTitle(sheet, `A1:${lastCol}1`, "比赛预测");
  const noteRange = sheet.getRange(`A2:${lastCol}2`);
  noteRange.merge();
  noteRange.values = [["即时赔率。欧赔格式为胜/平/负；亚盘格式为主队水位 盘口 客队水位；大小球格式为大球水位 盘口 小球水位。右侧赛前交易字段用于写回基本面、新闻、士气、盘口走势和下注建议。"]];
  noteRange.format.fill = palette.paleAmber;
  noteRange.format.font = { color: palette.amber, bold: true };
  noteRange.format.wrapText = true;
  noteRange.format.rowHeightPx = 44;

  const groupHeader = ["基础信息", "", ""];
  const metricHeader = ["小组", "时间(BJT)", "比赛"];
  for (const company of selectedCompanies) {
    groupHeader.push(company.label, "", "");
    metricHeader.push("欧赔", "亚盘", "大小球");
  }
  groupHeader.push("赛前交易", ...Array(preMatchTradingColumns.length - 1).fill(""));
  metricHeader.push(...preMatchTradingColumns);

  const dataRows = matchRecords.map((match) => {
    const companyRows = selectedCompanies.map((company) => oddsByMatchCompany.get(`${match.match_id}:${company.id}`));
    const row = [
      match.group,
      match.match_time_bjt,
      `${match.home_team} vs ${match.away_team}`,
    ];
    for (const odds of companyRows) {
      row.push(
        odds ? fmtTriple(odds.latest_euro_home_win, odds.latest_euro_draw, odds.latest_euro_away_win) : "",
        odds ? fmtHandicap(odds.latest_asian_home, odds.latest_asian_handicap, odds.latest_asian_away) : "",
        odds ? fmtHandicap(odds.latest_over_odds, odds.latest_goals_line, odds.latest_under_odds) : "",
      );
    }
    row.push(...preMatchTradingColumns.map(() => ""));
    return row;
  });

  const rows = [groupHeader, metricHeader, ...dataRows];
  setValues(sheet, `A3:${lastCol}${rows.length + 2}`, rows);
  styleGroupHeader(sheet, "A3:C3", palette.paleBlue);
  selectedCompanies.forEach((company, index) => {
    const start = 4 + index * 3;
    styleGroupHeader(sheet, `${colName(start)}3:${colName(start + 2)}3`, company.fill);
  });
  const tradingStart = 4 + selectedCompanies.length * 3;
  styleGroupHeader(sheet, `${colName(tradingStart)}3:${lastCol}3`, palette.paleGreen);
  styleHeader(sheet.getRange(`A4:${lastCol}4`));
  sheet.getRange(`A3:${lastCol}${rows.length + 2}`).format.borders = { preset: "all", style: "thin", color: palette.grid };
  sheet.getRange(`A5:${lastCol}${rows.length + 2}`).format.wrapText = true;
  sheet.getRange(`A5:${lastCol}${rows.length + 2}`).format.rowHeightPx = 58;
  sheet.getRange(`A5:C${rows.length + 2}`).format.fill = "#F8FAFC";
  sheet.getRange(`${colName(tradingStart)}5:${lastCol}${rows.length + 2}`).format.fill = "#FAFFF5";
  const widths = {
    A: 58,
    B: 175,
    C: 250,
    D: 165,
    E: 170,
    F: 155,
  };
  const tradingWidths = [190, 210, 170, 190, 150, 150, 150, 95, 260];
  preMatchTradingColumns.forEach((_, index) => {
    widths[colName(tradingStart + index)] = tradingWidths[index] ?? 160;
  });
  setColumnWidths(sheet, widths, 100);
}

function writeTradeRecommendations(sheet) {
  const lastCol = "S";
  styleTitle(sheet, `A1:${lastCol}1`, "交易推荐");

  const noteRange = sheet.getRange(`A2:${lastCol}2`);
  noteRange.merge();
  noteRange.values = [["用于记录赛前交易 skill 推荐执行的每笔交易。同一场可记录多笔执行单，建议金额按信心等级给出，并用关键/普通场次的单场累计上限做风控检查。"]];
  noteRange.format.fill = palette.paleAmber;
  noteRange.format.font = { color: palette.amber, bold: true };
  noteRange.format.wrapText = true;
  noteRange.format.rowHeightPx = 38;

  setValues(sheet, "A4:B11", [
    ["起始资金", 1000],
    ["关键场次上限%", 0.3],
    ["一般场次上限%", 0.1],
    ["赛前总暴露上限%", 0.6],
    ["关键信号建议%", 0.3],
    ["高信号建议%", 0.08],
    ["中信号建议%", 0.05],
    ["低信号建议%", 0.02],
  ]);
  styleBlock(sheet.getRange("A4:B11"), palette.paleBlue);
  sheet.getRange("A4:A11").format.font = { bold: true, color: palette.navy };
  moneyFormat(sheet, "B4:B4");
  percentFormat(sheet, "B5:B11");

  setValues(sheet, "D4:E9", [
    ["推荐概览", ""],
    ["总建议金额", ""],
    ["总资金占比", ""],
    ["推荐笔数", ""],
    ["建议执行笔数", ""],
    ["风控异常笔数", ""],
  ]);
  sheet.getRange("D4:E4").merge();
  sheet.getRange("D4:E4").format.fill = palette.navy;
  sheet.getRange("D4:E4").format.font = { color: palette.white, bold: true };
  styleBlock(sheet.getRange("D5:E9"), palette.white);
  sheet.getRange("D5:D9").format.fill = palette.paleBlue;
  sheet.getRange("D5:D9").format.font = { bold: true, color: palette.navy };
  setFormulas(sheet, "E5:E9", [
    ["=SUM($K$13:$K$212)"],
    ["=IF($B$4=0,\"\",E5/$B$4)"],
    ["=COUNTIF($E$13:$E$212,\"<>\")"],
    ["=COUNTIF($J$13:$J$212,\"建议执行\")"],
    ["=COUNTIF($N$13:$N$212,\"超*\")"],
  ]);
  moneyFormat(sheet, "E5:E5");
  percentFormat(sheet, "E6:E6");
  numberFormat(sheet, "E7:E9", "0");

  setValues(sheet, "G4:S10", [
    ["资金分配规则", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["关键/极高/非常高信心场次，赛前同场累计最高 30% 本金；普通场次同场累计最高 10%。", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["默认建议：关键 30%，高 8%，中 5%，低 2%；参数区可手动调整。", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["同一场可写多笔建议执行，例如大小球和让球/亚盘；风控看同场累计。", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["赛前总暴露上限默认 60% 本金；不下注/观望的建议金额自动为 0。", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["若同场多单超过上限，压低金额，而不是仅因同场已有主推就强制改为观望。", "", "", "", "", "", "", "", "", "", "", "", ""],
  ]);
  sheet.getRange("G4:S4").merge();
  sheet.getRange("G5:S10").merge(true);
  sheet.getRange("G4:S10").format.fill = palette.paleGreen;
  sheet.getRange("G4:G4").format.font = { bold: true, color: palette.navy };
  sheet.getRange("G4:S10").format.wrapText = true;
  sheet.getRange("G4:S10").format.borders = { preset: "all", style: "thin", color: palette.grid };

  const headers = [[
    "id",
    "推荐日期",
    "比赛时间(BJT)",
    "小组/阶段",
    "比赛",
    "市场/玩法",
    "投注内容",
    "可接受盘口/赔率",
    "信心等级",
    "建议动作",
    "建议金额",
    "资金占比",
    "单场累计",
    "风控状态",
    "交易依据",
    "主要风险",
    "更新来源",
    "复盘状态",
    "备注",
  ]];
  setValues(sheet, "A12:S12", headers);
  styleHeader(sheet.getRange("A12:S12"));

  const formulaRows = [];
  const criticalConfidenceCheck = (row) => `OR(COUNTIFS($E$13:$E$212,E${row},$J$13:$J$212,"建议执行",$I$13:$I$212,"关键")>0,COUNTIFS($E$13:$E$212,E${row},$J$13:$J$212,"建议执行",$I$13:$I$212,"极高")>0,COUNTIFS($E$13:$E$212,E${row},$J$13:$J$212,"建议执行",$I$13:$I$212,"非常高")>0)`;
  const matchCapFormula = (row) => `IF(${criticalConfidenceCheck(row)},$B$4*$B$5,$B$4*$B$6)`;
  for (let row = 13; row <= 212; row += 1) {
    formulaRows.push([
      `=IF(E${row}="","",ROW()-12)`,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      `=IF(E${row}="","",IF(OR($J${row}="不下注",$J${row}="观望",$I${row}="不下注"),0,MIN(${matchCapFormula(row)},$B$4*SWITCH($I${row},"关键",$B$8,"极高",$B$8,"非常高",$B$8,"高",$B$9,"中",$B$10,"低",$B$11,0))))`,
      `=IF(K${row}="","",K${row}/$B$4)`,
      `=IF(E${row}="","",SUMIF($E$13:$E$212,E${row},$K$13:$K$212))`,
      `=IF(E${row}="","",IF(K${row}>${matchCapFormula(row)},"超单笔",IF(M${row}>${matchCapFormula(row)},"超单场",IF(SUM($K$13:$K$212)>$B$4*$B$7,"超总暴露","OK"))))`,
      "",
      "",
      "",
      "",
      "",
    ]);
  }
  setFormulas(sheet, "A13:S212", formulaRows);

  styleBlock(sheet.getRange("A13:S212"), palette.white);
  sheet.getRange("A13:S212").format.wrapText = true;
  sheet.getRange("A13:S212").format.rowHeightPx = 38;
  sheet.getRange("A12:S212").format.borders = { preset: "all", style: "thin", color: palette.grid };
  setColumnWidths(sheet, {
    A: 90,
    B: 105,
    C: 145,
    D: 90,
    E: 180,
    F: 110,
    G: 180,
    H: 145,
    I: 85,
    J: 90,
    K: 90,
    L: 82,
    M: 90,
    N: 105,
    O: 240,
    P: 210,
    Q: 120,
    R: 95,
    S: 180,
  }, 230);
  numberFormat(sheet, "B13:C212", "yyyy-mm-dd");
  moneyFormat(sheet, "K13:K212");
  percentFormat(sheet, "L13:L212");
  moneyFormat(sheet, "M13:M212");

  sheet.getRange("N13:N212").conditionalFormats.addCustom("=$N13=\"OK\"", {
    fill: palette.paleGreen,
    font: { color: palette.teal, bold: true },
  });
  sheet.getRange("N13:N212").conditionalFormats.addCustom("=AND($N13<>\"\",$N13<>\"OK\")", {
    fill: palette.paleRose,
    font: { color: "#BE123C", bold: true },
  });
  sheet.getRange("I13:I212").conditionalFormats.addCustom("=$I13=\"高\"", {
    fill: palette.paleAmber,
    font: { color: palette.amber, bold: true },
  });
  sheet.getRange("I13:I212").conditionalFormats.addCustom("=OR($I13=\"关键\",$I13=\"极高\",$I13=\"非常高\")", {
    fill: palette.paleRose,
    font: { color: "#BE123C", bold: true },
  });
}

async function saveRenderedImage(rendered, path) {
  if (typeof rendered?.save === "function") {
    await rendered.save(path);
    return;
  }
  const candidate = rendered?.blob ?? rendered?.data ?? rendered?.buffer ?? rendered;
  if (typeof candidate?.arrayBuffer === "function") {
    await fs.writeFile(path, Buffer.from(await candidate.arrayBuffer()));
    return;
  }
  if (candidate instanceof Uint8Array || Buffer.isBuffer(candidate)) {
    await fs.writeFile(path, Buffer.from(candidate));
    return;
  }
  if (typeof candidate === "string") {
    const data = candidate.startsWith("data:")
      ? candidate.slice(candidate.indexOf(",") + 1)
      : candidate;
    await fs.writeFile(path, Buffer.from(data, "base64"));
  }
}

const snapshot = JSON.parse(await fs.readFile(`${baseDir}/odds_snapshot.json`, "utf8"));
const matchesCsv = normalizeBjtColumn(parseCsv(await fs.readFile(`${baseDir}/matches.csv`, "utf8")));
const oddsLongCsv = normalizeBjtColumn(parseCsv(await fs.readFile(`${baseDir}/odds_long.csv`, "utf8")));
const oddsWideCsv = normalizeBjtColumn(parseCsv(await fs.readFile(`${baseDir}/odds_wide.csv`, "utf8")));
const errorsCsv = parseCsv(await fs.readFile(`${baseDir}/errors.csv`, "utf8"));
const matchRecords = rowsToObjects(matchesCsv);
const oddsWideRecords = rowsToObjects(oddsWideCsv);

const dataWorkbook = Workbook.create();
const summary = dataWorkbook.worksheets.add("摘要");
const matches = dataWorkbook.worksheets.add("赛程");
const oddsLong = dataWorkbook.worksheets.add("赔率长表");
const oddsWide = dataWorkbook.worksheets.add("赔率宽表");
const notes = dataWorkbook.worksheets.add("字段说明");

const predictionWorkbook = Workbook.create();
const prediction = predictionWorkbook.worksheets.add("比赛预测");
const tradeRecommendations = predictionWorkbook.worksheets.add("交易推荐");

const groups = new Map();
for (const row of snapshot.matches) {
  const group = row.group;
  groups.set(group, (groups.get(group) ?? 0) + 1);
}
const companyCounts = new Map();
for (const row of snapshot.odds) {
  const key = `${row.company_id} ${row.company_name}`;
  companyCounts.set(key, (companyCounts.get(key) ?? 0) + 1);
}
const companyRows = [...companyCounts.entries()]
  .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
  .slice(0, 20);

styleTitle(summary, "A1:F1", "Titan007 2026世界杯赔率抓取摘要");
setValues(summary, "A3:B10", [
  ["源站", snapshot.source_home_url],
  ["抓取开始", makeBjtText(snapshot.scrape_started_at)],
  ["抓取结束", makeBjtText(snapshot.scrape_finished_at)],
  ["时区", snapshot.timezone],
  ["比赛数", snapshot.matches_count],
  ["比赛-公司行数", snapshot.odds_company_rows],
  ["比赛-公司-阶段行数", snapshot.odds_period_rows],
  ["错误数", snapshot.errors_count],
]);
summary.getRange("A3:A10").format.fill = palette.paleBlue;
summary.getRange("A3:A10").format.font = { bold: true, color: palette.navy };
summary.getRange("A3:B10").format.borders = { preset: "all", style: "thin", color: palette.grid };
setValues(summary, "D3:F3", [["小组", "比赛数", "说明"]]);
styleHeader(summary.getRange("D3:F3"));
setValues(
  summary,
  `D4:F${3 + groups.size}`,
  [...groups.entries()].map(([group, count]) => [group, count, "当前首页固定赛程"])
);
summary.getRange(`D4:F${3 + groups.size}`).format.borders = { preset: "all", style: "thin", color: palette.grid };

const companyStart = 18;
setValues(summary, `A${companyStart}:B${companyStart}`, [["公司", "覆盖场次数"]]);
styleHeader(summary.getRange(`A${companyStart}:B${companyStart}`));
setValues(summary, `A${companyStart + 1}:B${companyStart + companyRows.length}`, companyRows);
summary.getRange(`A${companyStart + 1}:B${companyStart + companyRows.length}`).format.borders = { preset: "all", style: "thin", color: palette.grid };

setValues(summary, "D18:F25", [
  ["口径说明", "", ""],
  ["本数据表来自 Titan007 2026 首页和 zq.titan007.com/analysis/odds/{id}.htm。", "", ""],
  ["首页当前固定列出 A-L 组 72 场小组赛；淘汰赛对阵尚未固定，未纳入。", "", ""],
  ["比赛预测已剥离到独立文件：比赛预测.xlsx。", "", ""],
  ["赔率长表每家公司分开盘和即时两行；赔率宽表每家公司一行。", "", ""],
  ["euro_to_asian_* 是站点给出的欧赔转亚盘字段，asian_* 是实际亚盘字段。", "", ""],
  ["CSV 和 JSON 原始快照同目录保留，便于脚本化分析。", "", ""],
  ["抓取脚本在 scripts/scrape_titan007_worldcup_odds.py。", "", ""],
]);
summary.getRange("D18:F18").merge();
summary.getRange("D19:F25").merge(true);
summary.getRange("D18:F25").format.fill = palette.paleAmber;
summary.getRange("D18:D18").format.font = { bold: true, color: palette.amber };
summary.getRange("D18:F25").format.wrapText = true;
summary.getRange("D18:F25").format.borders = { preset: "all", style: "thin", color: palette.grid };
setColumnWidths(summary, { A: 190, B: 360, D: 120, E: 90, F: 260 }, 80);

writePredictionSheet(prediction, matchRecords, oddsWideRecords);
writeTradeRecommendations(tradeRecommendations);

writeTable(matches, matchesCsv, { wrap: true, rowHeightPx: 28 });
setColumnWidths(matches, {
  A: 86,
  B: 80,
  C: 70,
  D: 165,
  E: 135,
  F: 150,
  G: 260,
  H: 260,
  I: 260,
  J: 290,
}, 120);

writeTable(oddsLong, oddsLongCsv, { wrap: false, rowHeightPx: 22 });
setColumnWidths(oddsLong, {
  A: 86,
  C: 70,
  D: 165,
  E: 130,
  F: 150,
  K: 80,
  L: 110,
  P: 80,
  Q: 70,
  R: 86,
  S: 72,
  T: 86,
  V: 130,
  Y: 90,
  Z: 130,
  AC: 80,
  AD: 90,
  AE: 80,
}, 2200);

writeTable(oddsWide, oddsWideCsv, { wrap: false, rowHeightPx: 22 });
setColumnWidths(oddsWide, {
  A: 86,
  C: 70,
  D: 165,
  E: 130,
  F: 150,
  K: 80,
  L: 110,
  P: 90,
  Q: 80,
  R: 90,
  S: 90,
  T: 130,
  V: 90,
  W: 90,
  X: 130,
  AE: 90,
  AF: 80,
  AG: 90,
  AH: 90,
  AI: 130,
  AK: 90,
  AL: 90,
  AM: 130,
}, 1200);

const noteRows = [
  ["字段", "说明"],
  ...Object.entries(snapshot.field_notes),
  ["match_time_bjt", "比赛时间，按 Asia/Shanghai / 北京时间标注。"],
  ["status_flags", "站点隐藏字段原样保留，用逗号分隔。"],
  ["jc_odds", "站点 iframeAJCOdds 隐藏值，原样保留。"],
  ["source_url", "本行赔率来自的 analysis/odds 接口。"],
  ["比赛预测.xlsx", "独立预测表：按比赛展示 365 即时赔率，并预留赛前交易写回字段。"],
  ["errors.csv", `本次抓取错误数为 ${snapshot.errors_count}，只保留表头表示无错误。`],
];
writeTable(notes, noteRows, { wrap: true, rowHeightPx: 34 });
notes.getRange(`A2:A${noteRows.length}`).format.fill = palette.paleGreen;
notes.getRange(`A2:A${noteRows.length}`).format.font = { bold: true, color: palette.navy };
setColumnWidths(notes, { A: 220, B: 620 }, 80);

if (errorsCsv.length > 1) {
  const errorsSheet = dataWorkbook.worksheets.add("错误");
  writeTable(errorsSheet, errorsCsv, { wrap: true, rowHeightPx: 34 });
  setColumnWidths(errorsSheet, { A: 90, B: 120, C: 120, D: 300, E: 420 }, 80);
}

const summaryCheck = await dataWorkbook.inspect({
  kind: "table",
  range: "摘要!A1:F27",
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 6,
});
console.log(summaryCheck.ndjson);

const matchViewLastCol = colName(3 + selectedCompanies.length * 3 + preMatchTradingColumns.length);
const predictionCheck = await predictionWorkbook.inspect({
  kind: "table",
  range: `比赛预测!A1:${matchViewLastCol}12`,
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 15,
});
console.log(predictionCheck.ndjson);

const tradeRecommendationsCheck = await predictionWorkbook.inspect({
  kind: "table",
  range: "交易推荐!A1:S18",
  include: "values,formulas",
  tableMaxRows: 18,
  tableMaxCols: 19,
});
console.log(tradeRecommendationsCheck.ndjson);

const matchesCheck = await dataWorkbook.inspect({
  kind: "table",
  range: "赛程!A1:J8",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 10,
});
console.log(matchesCheck.ndjson);

const dataFormulaErrors = await dataWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "data workbook formula error scan",
});
console.log(dataFormulaErrors.ndjson);

const predictionFormulaErrors = await predictionWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "prediction workbook formula error scan",
});
console.log(predictionFormulaErrors.ndjson);

await fs.mkdir(`${baseDir}/previews`, { recursive: true });
await saveRenderedImage(
  await dataWorkbook.render({ sheetName: "摘要", range: "A1:F27", scale: 2 }),
  `${baseDir}/previews/summary.png`
);
await saveRenderedImage(
  await predictionWorkbook.render({ sheetName: "比赛预测", range: `A1:${matchViewLastCol}22`, scale: 1 }),
  `${baseDir}/previews/prediction.png`
);
await saveRenderedImage(
  await predictionWorkbook.render({ sheetName: "交易推荐", range: "A1:S30", scale: 1 }),
  `${baseDir}/previews/trade_recommendations.png`
);
await saveRenderedImage(
  await dataWorkbook.render({ sheetName: "赛程", range: "A1:J20", scale: 1 }),
  `${baseDir}/previews/matches.png`
);
await saveRenderedImage(
  await dataWorkbook.render({ sheetName: "赔率长表", range: "A1:AE28", scale: 1 }),
  `${baseDir}/previews/odds_long.png`
);
await saveRenderedImage(
  await dataWorkbook.render({ sheetName: "赔率宽表", range: "A1:AM24", scale: 1 }),
  `${baseDir}/previews/odds_wide.png`
);
await saveRenderedImage(
  await dataWorkbook.render({ sheetName: "字段说明", range: "A1:B12", scale: 2 }),
  `${baseDir}/previews/notes.png`
);

const dataOutput = await SpreadsheetFile.exportXlsx(dataWorkbook);
await dataOutput.save(dataOutputPath);
console.log(`SAVED ${dataOutputPath}`);

const predictionOutput = await SpreadsheetFile.exportXlsx(predictionWorkbook);
await predictionOutput.save(predictionOutputPath);
console.log(`SAVED ${predictionOutputPath}`);
