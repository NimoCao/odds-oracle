import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

function argValue(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const baseDir = argValue("--base-dir") ?? path.join(repoRoot, "data/titan007_worldcup_2026_odds");
const workbookPath = argValue("--workbook") ?? path.join(baseDir, "比赛预测.xlsx");
const previewPath = argValue("--preview") ?? path.join(baseDir, "previews/trade_recommendations_staking_logic.png");

const palette = {
  navy: "#17324D",
  teal: "#0F766E",
  amber: "#B45309",
  paleBlue: "#EAF3F8",
  paleGreen: "#EAF7EE",
  paleAmber: "#FFF7E6",
  paleRose: "#FCE7F3",
  grid: "#D9E2EC",
};

function timestampForBackup() {
  return new Intl.DateTimeFormat("sv-SE", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).format(new Date()).replace(/[-: ]/g, "");
}

function styleBlock(range, fill = "#F8FAFC") {
  range.format.fill = fill;
  range.format.borders = { preset: "all", style: "thin", color: palette.grid };
}

function moneyFormat(sheet, range) {
  sheet.getRange(range).format.numberFormat = '"¥"#,##0.00;[Red]-"¥"#,##0.00;"¥"0.00';
}

function percentFormat(sheet, range) {
  sheet.getRange(range).format.numberFormat = '0.00%;[Red]-0.00%;0.00%';
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
  }
}

const backupPath = workbookPath.replace(/\.xlsx$/, `_before_staking_logic_${timestampForBackup()}.xlsx`);
await fs.copyFile(workbookPath, backupPath);

const blob = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(blob);
const trades = workbook.worksheets.getItem("交易推荐");

const note = trades.getRange("A2:S2");
note.merge();
note.values = [["用于记录赛前交易 skill 推荐执行的每笔交易。同一场可记录多笔执行单，建议金额按信心等级给出，并用关键/普通场次的单场累计上限做风控检查。"]];
note.format.fill = palette.paleAmber;
note.format.font = { color: palette.amber, bold: true };
note.format.wrapText = true;
note.format.rowHeightPx = 38;

trades.getRange("A4:B11").values = [
  ["起始资金", 1000],
  ["关键场次上限%", 0.3],
  ["一般场次上限%", 0.1],
  ["赛前总暴露上限%", 0.6],
  ["关键信号建议%", 0.3],
  ["高信号建议%", 0.08],
  ["中信号建议%", 0.05],
  ["低信号建议%", 0.02],
];
styleBlock(trades.getRange("A4:B11"), palette.paleBlue);
trades.getRange("A4:A11").format.font = { bold: true, color: palette.navy };
moneyFormat(trades, "B4:B4");
percentFormat(trades, "B5:B11");

trades.getRange("G4").values = [["资金分配规则"]];
trades.getRange("G5").values = [["关键/极高/非常高信心场次，赛前同场累计最高 30% 本金；普通场次同场累计最高 10%。"]];
trades.getRange("G6").values = [["默认建议：关键 30%，高 8%，中 5%，低 2%；参数区可手动调整。"]];
trades.getRange("G7").values = [["同一场可写多笔建议执行，例如大小球和让球/亚盘；风控看同场累计。"]];
trades.getRange("G8").values = [["赛前总暴露上限默认 60% 本金；不下注/观望的建议金额自动为 0。"]];
trades.getRange("G9").values = [["若同场多单超过上限，压低金额，而不是仅因同场已有主推就强制改为观望。"]];
trades.getRange("G10").values = [[""]];
trades.getRange("G4:S10").format.fill = palette.paleGreen;
trades.getRange("G4:G4").format.font = { bold: true, color: palette.navy };
trades.getRange("G4:S10").format.wrapText = true;
trades.getRange("G4:S10").format.borders = { preset: "all", style: "thin", color: palette.grid };
trades.getRange("A1:A230").format.columnWidthPx = 90;

const criticalConfidenceCheck = (row) => `OR(COUNTIFS($E$13:$E$212,E${row},$J$13:$J$212,"建议执行",$I$13:$I$212,"关键")>0,COUNTIFS($E$13:$E$212,E${row},$J$13:$J$212,"建议执行",$I$13:$I$212,"极高")>0,COUNTIFS($E$13:$E$212,E${row},$J$13:$J$212,"建议执行",$I$13:$I$212,"非常高")>0)`;
const matchCapFormula = (row) => `IF(${criticalConfidenceCheck(row)},$B$4*$B$5,$B$4*$B$6)`;

const idFormulas = [];
const amountFormulas = [];
const percentFormulas = [];
const matchTotalFormulas = [];
const riskFormulas = [];
for (let row = 13; row <= 212; row += 1) {
  idFormulas.push([`=IF(E${row}="","",ROW()-12)`]);
  amountFormulas.push([`=IF(E${row}="","",IF(OR($J${row}="不下注",$J${row}="观望",$I${row}="不下注"),0,MIN(${matchCapFormula(row)},$B$4*SWITCH($I${row},"关键",$B$8,"极高",$B$8,"非常高",$B$8,"高",$B$9,"中",$B$10,"低",$B$11,0))))`]);
  percentFormulas.push([`=IF(K${row}="","",K${row}/$B$4)`]);
  matchTotalFormulas.push([`=IF(E${row}="","",SUMIF($E$13:$E$212,E${row},$K$13:$K$212))`]);
  riskFormulas.push([`=IF(E${row}="","",IF(K${row}>${matchCapFormula(row)},"超单笔",IF(M${row}>${matchCapFormula(row)},"超单场",IF(SUM($K$13:$K$212)>$B$4*$B$7,"超总暴露","OK"))))`]);
}
trades.getRange("A13:A212").formulas = idFormulas;
trades.getRange("K13:K212").formulas = amountFormulas;
trades.getRange("L13:L212").formulas = percentFormulas;
trades.getRange("M13:M212").formulas = matchTotalFormulas;
trades.getRange("N13:N212").formulas = riskFormulas;
moneyFormat(trades, "K13:K212");
percentFormat(trades, "L13:L212");
moneyFormat(trades, "M13:M212");

const remarkRange = trades.getRange("S13:S212");
remarkRange.values = remarkRange.values.map(([value]) => {
  if (typeof value !== "string") return [value];
  const replacements = {
    高: "高信心默认8%本金",
    中: "中信心默认5%本金",
    低: "低信心默认2%本金",
  };
  return [
    value.replace(/默认(高|中|低)信心金额\d+(?:\.\d+)?/g, (_, level) => replacements[level]),
  ];
});

trades.getRange("N13:N212").conditionalFormats.addCustom("=$N13=\"OK\"", {
  fill: palette.paleGreen,
  font: { color: palette.teal, bold: true },
});
trades.getRange("N13:N212").conditionalFormats.addCustom("=AND($N13<>\"\",$N13<>\"OK\")", {
  fill: palette.paleRose,
  font: { color: "#BE123C", bold: true },
});
trades.getRange("I13:I212").conditionalFormats.addCustom("=OR($I13=\"关键\",$I13=\"极高\",$I13=\"非常高\")", {
  fill: palette.paleRose,
  font: { color: "#BE123C", bold: true },
});

const tradesCheck = await workbook.inspect({
  kind: "table",
  range: "交易推荐!A1:S18",
  include: "values,formulas",
  tableMaxRows: 18,
  tableMaxCols: 19,
});
console.log(tradesCheck.ndjson);

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 200 },
  summary: "formula error scan after staking logic update",
});
console.log(formulaErrors.ndjson);

await fs.mkdir(path.dirname(previewPath), { recursive: true });
await saveRenderedImage(
  await workbook.render({ sheetName: "交易推荐", range: "A1:S30", scale: 1 }),
  previewPath
);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(workbookPath);

console.log(`BACKUP ${backupPath}`);
console.log(`SAVED ${workbookPath}`);
console.log(`PREVIEW ${previewPath}`);
