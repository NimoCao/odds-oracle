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
const outputPath = argValue("--output") ?? path.join(repoRoot, "templates/世界杯足球策略复盘模板.xlsx");
const outputDir = path.dirname(outputPath);
const previewDir = argValue("--preview-dir") ?? path.join(repoRoot, "data/world_cup_strategy_review/previews");
const rowCount = 200;

const workbook = Workbook.create();

const dashboard = workbook.worksheets.add("资金总览");
const bets = workbook.worksheets.add("投注明细");
const guide = workbook.worksheets.add("字段说明");
const params = workbook.worksheets.add("参数");

const palette = {
  navy: "#17324D",
  teal: "#0F766E",
  amber: "#D97706",
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
  r.format.font = { color: palette.white, bold: true, size: 18 };
  r.format.horizontalAlignment = "left";
  r.format.rowHeightPx = 42;
}

function styleHeader(range) {
  range.format.fill = palette.teal;
  range.format.font = { color: palette.white, bold: true };
  range.format.horizontalAlignment = "center";
  range.format.wrapText = true;
  range.format.rowHeightPx = 34;
}

function styleBlock(range, fill = palette.soft) {
  range.format.fill = fill;
  range.format.borders = { preset: "all", style: "thin", color: palette.grid };
}

function setColumnWidths(sheet, widths, totalRows = 240) {
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

// 参数
styleTitle(params, "A1:E1", "资金与风控参数");
setValues(params, "A3:B8", [
  ["起始本金", 1000],
  ["单笔建议上限", 50],
  ["单笔常规单位", 20],
  ["高信号单位", 30],
  ["最大连续亏损警戒", 3],
  ["备注", "参数可手动调整，明细表公式会引用起始本金。"],
]);
styleBlock(params.getRange("A3:B8"), palette.paleBlue);
params.getRange("A3:A8").format.font = { bold: true, color: palette.navy };
moneyFormat(params, "B3:B6");
numberFormat(params, "B7:B7", "0");
params.getRange("A10:E10").merge();
params.getRange("A10:E10").values = [["选项清单：建议保持这些固定词，便于后续透视统计。"]];
params.getRange("A10:E10").format.fill = palette.navy;
params.getRange("A10:E10").format.font = { color: palette.white, bold: true };

const optionRows = [
  ["场次分类", "市场/玩法", "投注结果", "策略标签", "投注类型"],
  ["小组赛", "胜平负", "未结算", "赛前模型", "赛前"],
  ["1/8决赛", "让球胜平负", "赢", "盘口异动", "滚球"],
  ["1/4决赛", "亚洲让球", "输", "阵容消息", "串关"],
  ["半决赛", "大小球", "走水", "体能赛程", ""],
  ["季军赛", "角球", "半赢", "对冲", ""],
  ["决赛", "比分", "半输", "价值投注", ""],
  ["其他", "半全场", "取消", "情绪修正", ""],
  ["", "球员进球", "", "其他", ""],
  ["", "冠军/晋级", "", "", ""],
  ["", "其他", "", "", ""],
];
setValues(params, "A11:E21", optionRows);
styleHeader(params.getRange("A11:E11"));
styleBlock(params.getRange("A12:E21"), palette.white);
setColumnWidths(params, { A: 130, B: 150, C: 110, D: 130, E: 110 }, 40);

// 投注明细
styleTitle(bets, "A1:W1", "单笔投注复盘明细");
const headers = [[
  "id",
  "投注日期",
  "投注场次",
  "比赛日期",
  "场次分类",
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
]];
setValues(bets, "A2:W2", headers);
styleHeader(bets.getRange("A2:W2"));

const formulaRows = [];
for (let row = 3; row < 3 + rowCount; row += 1) {
  formulaRows.push([
    `=IF(C${row}="","",ROW()-2)`,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    `=IF(L${row}="","",'参数'!$B$3+IF(ROW()=3,0,SUM($R$3:INDEX($R:$R,ROW()-1))))`,
    "",
    "",
    `=IF(OR(L${row}="",K${row}=""),"",L${row}/K${row})`,
    `=IF(OR(L${row}="",M${row}=""),"",L${row}*M${row})`,
    "",
    `=IF(OR(L${row}="",P${row}="",P${row}="未结算"),"",SWITCH(P${row},"赢",L${row}*M${row},"输",0,"走水",L${row},"取消",L${row},"半赢",L${row}*(1+(M${row}-1)/2),"半输",L${row}*0.5,""))`,
    `=IF(Q${row}="","",Q${row}-L${row})`,
    `=IF(OR(L${row}="",R${row}=""),"",R${row}/L${row})`,
    `=IF(R${row}="","",'参数'!$B$3+SUM($R$3:R${row}))`,
    "",
    "",
    "",
  ]);
}
setFormulas(bets, `A3:W${2 + rowCount}`, formulaRows);

styleBlock(bets.getRange(`A3:W${2 + rowCount}`), palette.white);
bets.getRange(`A3:W${2 + rowCount}`).format.wrapText = true;
bets.getRange(`A3:W${2 + rowCount}`).format.rowHeightPx = 30;
bets.getRange("A2:W202").format.borders = { preset: "all", style: "thin", color: palette.grid };
setColumnWidths(bets, {
  A: 54,
  B: 98,
  C: 130,
  D: 98,
  E: 94,
  F: 140,
  G: 110,
  H: 84,
  I: 170,
  J: 110,
  K: 100,
  L: 92,
  M: 78,
  N: 82,
  O: 96,
  P: 86,
  Q: 96,
  R: 96,
  S: 82,
  T: 100,
  U: 220,
  V: 220,
  W: 160,
});
moneyFormat(bets, "K3:L202");
numberFormat(bets, "M3:M202", "0.000");
percentFormat(bets, "N3:N202");
moneyFormat(bets, "O3:R202");
percentFormat(bets, "S3:S202");
moneyFormat(bets, "T3:T202");
numberFormat(bets, "B3:B202", "yyyy-mm-dd");
numberFormat(bets, "D3:D202", "yyyy-mm-dd");

bets.getRange("R3:R202").conditionalFormats.addCellIs({
  operator: "greaterThan",
  formula: 0,
  format: { fill: palette.paleGreen, font: { color: palette.green, bold: true } },
});
bets.getRange("R3:R202").conditionalFormats.addCellIs({
  operator: "lessThan",
  formula: 0,
  format: { fill: palette.paleRed, font: { color: palette.red, bold: true } },
});
bets.getRange("L3:L202").conditionalFormats.addCustom("=AND($L3<>\"\",$L3>'参数'!$B$4)", {
  fill: palette.paleAmber,
  font: { color: palette.amber, bold: true },
});
bets.getRange("P3:P202").conditionalFormats.addCustom("=$P3=\"赢\"", {
  fill: palette.paleGreen,
  font: { color: palette.green, bold: true },
});
bets.getRange("P3:P202").conditionalFormats.addCustom("=$P3=\"输\"", {
  fill: palette.paleRed,
  font: { color: palette.red, bold: true },
});

setValues(bets, "A204:W208", [
  ["使用说明", "每一行只记录一笔投注；同一场比赛多个玩法拆成多行。", ...Array(21).fill("")],
  ["结算口径", "赢=金额*赔率；输=0；走水/取消=退回本金；半赢=本金+半额盈利；半输=退回半额本金。", ...Array(21).fill("")],
  ["纪律提醒", "黄色高亮表示投注金额超过参数页的单笔建议上限，复盘时优先检查是否冲动加仓。", ...Array(21).fill("")],
  ["字段规范", "场次分类、玩法、投注类型、结果、策略标签建议使用参数页固定词，后续才好按阶段和策略统计。", ...Array(21).fill("")],
  ["本金设置", "默认起始本金为1000元，可在参数页 B3 调整。", ...Array(21).fill("")],
]);
bets.getRange("A204:A208").format.fill = palette.navy;
bets.getRange("A204:A208").format.font = { color: palette.white, bold: true };
bets.getRange("B204:W208").merge(true);
bets.getRange("B204:W208").format.fill = palette.paleBlue;
bets.getRange("B204:W208").format.wrapText = true;
bets.getRange("A204:W208").format.borders = { preset: "all", style: "thin", color: palette.grid };

// 资金总览
styleTitle(dashboard, "A1:N1", "世界杯足球策略复盘总览");
setValues(dashboard, "A3:B11", [
  ["起始本金", ""],
  ["当前本金", ""],
  ["总投注金额", ""],
  ["已结算笔数", ""],
  ["胜率", ""],
  ["净收益", ""],
  ["总 ROI", ""],
  ["最大单笔亏损", ""],
  ["最大单笔盈利", ""],
]);
setFormulas(dashboard, "B3:B11", [
  ["='参数'!B3"],
  ["='参数'!B3+SUM('投注明细'!R3:R202)"],
  ["=SUM('投注明细'!L3:L202)"],
  ["=COUNTIF('投注明细'!P3:P202,\"<>未结算\")-COUNTBLANK('投注明细'!P3:P202)"],
  ["=IF(B6=0,\"\",COUNTIF('投注明细'!P3:P202,\"赢\")/B6)"],
  ["=SUM('投注明细'!R3:R202)"],
  ["=IF(B5=0,\"\",B8/B5)"],
  ["=IFERROR(MIN('投注明细'!R3:R202),\"\")"],
  ["=IFERROR(MAX('投注明细'!R3:R202),\"\")"],
]);
styleBlock(dashboard.getRange("A3:B11"), palette.white);
dashboard.getRange("A3:A11").format.fill = palette.paleBlue;
dashboard.getRange("A3:A11").format.font = { bold: true, color: palette.navy };
moneyFormat(dashboard, "B3:B5");
numberFormat(dashboard, "B6:B6", "0");
percentFormat(dashboard, "B7:B7");
moneyFormat(dashboard, "B8:B8");
percentFormat(dashboard, "B9:B9");
moneyFormat(dashboard, "B10:B11");

setValues(dashboard, "D3:H3", [["按场次分类统计", "", "", "", ""]]);
dashboard.getRange("D3:H3").merge();
dashboard.getRange("D3:H3").format.fill = palette.navy;
dashboard.getRange("D3:H3").format.font = { color: palette.white, bold: true };
setValues(dashboard, "D4:H4", [["场次分类", "笔数", "投注金额", "净收益", "ROI"]]);
styleHeader(dashboard.getRange("D4:H4"));
setValues(dashboard, "D5:D11", [["小组赛"], ["1/8决赛"], ["1/4决赛"], ["半决赛"], ["季军赛"], ["决赛"], ["其他"]]);
setFormulas(dashboard, "E5:H11", [
  ["=COUNTIF('投注明细'!E:E,D5)", "=SUMIF('投注明细'!E:E,D5,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D5,'投注明细'!R:R)", "=IF(F5=0,\"\",G5/F5)"],
  ["=COUNTIF('投注明细'!E:E,D6)", "=SUMIF('投注明细'!E:E,D6,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D6,'投注明细'!R:R)", "=IF(F6=0,\"\",G6/F6)"],
  ["=COUNTIF('投注明细'!E:E,D7)", "=SUMIF('投注明细'!E:E,D7,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D7,'投注明细'!R:R)", "=IF(F7=0,\"\",G7/F7)"],
  ["=COUNTIF('投注明细'!E:E,D8)", "=SUMIF('投注明细'!E:E,D8,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D8,'投注明细'!R:R)", "=IF(F8=0,\"\",G8/F8)"],
  ["=COUNTIF('投注明细'!E:E,D9)", "=SUMIF('投注明细'!E:E,D9,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D9,'投注明细'!R:R)", "=IF(F9=0,\"\",G9/F9)"],
  ["=COUNTIF('投注明细'!E:E,D10)", "=SUMIF('投注明细'!E:E,D10,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D10,'投注明细'!R:R)", "=IF(F10=0,\"\",G10/F10)"],
  ["=COUNTIF('投注明细'!E:E,D11)", "=SUMIF('投注明细'!E:E,D11,'投注明细'!L:L)", "=SUMIF('投注明细'!E:E,D11,'投注明细'!R:R)", "=IF(F11=0,\"\",G11/F11)"],
]);
styleBlock(dashboard.getRange("D5:H11"), palette.white);
moneyFormat(dashboard, "F5:G11");
percentFormat(dashboard, "H5:H11");

setValues(dashboard, "J3:N3", [["按投注类型统计", "", "", "", ""]]);
dashboard.getRange("J3:N3").merge();
dashboard.getRange("J3:N3").format.fill = palette.navy;
dashboard.getRange("J3:N3").format.font = { color: palette.white, bold: true };
setValues(dashboard, "J4:N4", [["投注类型", "笔数", "投注金额", "净收益", "ROI"]]);
styleHeader(dashboard.getRange("J4:N4"));
setValues(dashboard, "J5:J7", [["赛前"], ["滚球"], ["串关"]]);
setFormulas(dashboard, "K5:N7", [
  ["=COUNTIF('投注明细'!H:H,J5)", "=SUMIF('投注明细'!H:H,J5,'投注明细'!L:L)", "=SUMIF('投注明细'!H:H,J5,'投注明细'!R:R)", "=IF(L5=0,\"\",M5/L5)"],
  ["=COUNTIF('投注明细'!H:H,J6)", "=SUMIF('投注明细'!H:H,J6,'投注明细'!L:L)", "=SUMIF('投注明细'!H:H,J6,'投注明细'!R:R)", "=IF(L6=0,\"\",M6/L6)"],
  ["=COUNTIF('投注明细'!H:H,J7)", "=SUMIF('投注明细'!H:H,J7,'投注明细'!L:L)", "=SUMIF('投注明细'!H:H,J7,'投注明细'!R:R)", "=IF(L7=0,\"\",M7/L7)"],
]);
styleBlock(dashboard.getRange("J5:N7"), palette.white);
moneyFormat(dashboard, "L5:M7");
percentFormat(dashboard, "N5:N7");

setValues(dashboard, "A14:H14", [["近十笔结算记录", "", "", "", "", "", "", ""]]);
dashboard.getRange("A14:H14").merge();
dashboard.getRange("A14:H14").format.fill = palette.navy;
dashboard.getRange("A14:H14").format.font = { color: palette.white, bold: true };
setValues(dashboard, "A15:H15", [["id", "投注场次", "投注类型", "投注内容", "投注金额", "赔率", "结果", "收益"]]);
styleHeader(dashboard.getRange("A15:H15"));
setFormulas(dashboard, "A16:H25", [
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A16)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A17)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A18)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A19)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A20)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A21)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A22)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A23)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A24)),\"\")"],
  ["=IFERROR(INDEX(FILTER('投注明细'!A3:A202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!C3:C202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!H3:H202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!I3:I202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!L3:L202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!M3:M202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!P3:P202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")", "=IFERROR(INDEX(FILTER('投注明细'!R3:R202,'投注明细'!P3:P202<>\"未结算\",'投注明细'!P3:P202<>\"\"),ROWS($A$16:A25)),\"\")"],
]);
styleBlock(dashboard.getRange("A16:H25"), palette.white);
moneyFormat(dashboard, "E16:E25");
numberFormat(dashboard, "F16:F25", "0.000");
moneyFormat(dashboard, "H16:H25");
setColumnWidths(dashboard, { A: 90, B: 150, C: 100, D: 180, E: 100, F: 75, G: 75, H: 95, J: 100, K: 80, L: 100, M: 100, N: 82 }, 40);

dashboard.getRange("A28:H30").values = [
  ["复盘口径", "这份表按单笔投注记录，并用“投注类型”区分赛前、滚球和串关；同场多玩法不合并，方便之后找出具体策略问题。", "", "", "", "", "", ""],
  ["资金纪律", "默认本金1000元，参数页单笔建议上限为50元；超出会在明细页黄色提示。", "", "", "", "", "", ""],
  ["结果维护", "未结算先填“未结算”，比赛结束后改成赢/输/走水/半赢/半输/取消。", "", "", "", "", "", ""],
];
dashboard.getRange("A28:A30").format.fill = palette.navy;
dashboard.getRange("A28:A30").format.font = { color: palette.white, bold: true };
dashboard.getRange("B28:H30").merge(true);
dashboard.getRange("B28:H30").format.fill = palette.paleBlue;
dashboard.getRange("B28:H30").format.wrapText = true;
dashboard.getRange("A28:H30").format.borders = { preset: "all", style: "thin", color: palette.grid };

// 字段说明
styleTitle(guide, "A1:D1", "字段说明");
setValues(guide, "A3:D3", [["字段", "是否必填", "说明", "示例"]]);
styleHeader(guide.getRange("A3:D3"));
setValues(guide, "A4:D26", [
  ["id", "自动", "单笔投注编号，由明细行自动生成。", "1"],
  ["投注日期", "建议填", "实际下注日期。", "2026-06-12"],
  ["投注场次", "必填", "比赛或盘口名称；同场多单可重复。", "A组第1轮"],
  ["比赛日期", "建议填", "比赛开赛日期。", "2026-06-13"],
  ["场次分类", "必填", "小组赛、淘汰赛阶段等。", "小组赛"],
  ["对阵", "必填", "双方球队。", "巴西 vs 德国"],
  ["市场/玩法", "必填", "胜平负、亚洲让球、大小球等。", "大小球"],
  ["投注类型", "必填", "区分下注发生场景：赛前、滚球、串关。", "滚球"],
  ["投注内容", "必填", "具体方向。", "大2.5球"],
  ["策略标签", "建议填", "用于复盘策略来源。", "阵容消息"],
  ["投注前本金", "自动", "该笔下注前的滚动本金。", "1000"],
  ["投注金额", "必填", "该笔投注本金。", "20"],
  ["投注赔率", "必填", "十进制赔率。", "1.850"],
  ["资金占比", "自动", "投注金额 / 投注前本金。", "2.00%"],
  ["潜在返还", "自动", "投注金额 * 投注赔率。", "37"],
  ["投注结果", "必填", "未结算、赢、输、走水、半赢、半输、取消。", "赢"],
  ["结算金额", "自动", "按投注结果自动计算返还金额。", "37"],
  ["投注收益", "自动", "结算金额 - 投注金额。", "17"],
  ["ROI", "自动", "投注收益 / 投注金额。", "85.00%"],
  ["滚动本金", "自动", "起始本金 + 累计已结算收益。", "1017"],
  ["赛前依据", "建议填", "下注前的判断依据，越具体越能复盘。", "主力前锋复出，盘口升水"],
  ["赛后复盘", "建议填", "赛后验证判断是否成立。", "节奏符合预期但射门质量偏低"],
  ["备注", "选填", "临场信息、对冲、情绪状态等。", "临场追单，需复盘"],
]);
styleBlock(guide.getRange("A4:D26"), palette.white);
guide.getRange("A4:A26").format.fill = palette.paleBlue;
guide.getRange("A4:A26").format.font = { bold: true, color: palette.navy };
guide.getRange("A3:D26").format.wrapText = true;
guide.getRange("A3:D26").format.borders = { preset: "all", style: "thin", color: palette.grid };
setColumnWidths(guide, { A: 120, B: 80, C: 390, D: 180 }, 40);
guide.getRange("A4:D26").format.rowHeightPx = 34;

// Compact visual checks and export.
const dashboardCheck = await workbook.inspect({
  kind: "table",
  range: "资金总览!A1:N30",
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 14,
});
console.log(dashboardCheck.ndjson);

const betsCheck = await workbook.inspect({
  kind: "table",
  range: "投注明细!A1:W8",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 23,
});
console.log(betsCheck.ndjson);

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(formulaErrors.ndjson);

await fs.mkdir(previewDir, { recursive: true });
const dashboardPreview = await workbook.render({ sheetName: "资金总览", range: "A1:N30", scale: 1.5 });
await saveRenderedImage(dashboardPreview, `${previewDir}/dashboard.png`);
const betsPreview = await workbook.render({ sheetName: "投注明细", range: "A1:W28", scale: 1 });
await saveRenderedImage(betsPreview, `${previewDir}/bets.png`);
const guidePreview = await workbook.render({ sheetName: "字段说明", range: "A1:D26", scale: 2 });
await saveRenderedImage(guidePreview, `${previewDir}/guide.png`);
const paramsPreview = await workbook.render({ sheetName: "参数", range: "A1:E21", scale: 2 });
await saveRenderedImage(paramsPreview, `${previewDir}/params.png`);

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`SAVED ${outputPath}`);
