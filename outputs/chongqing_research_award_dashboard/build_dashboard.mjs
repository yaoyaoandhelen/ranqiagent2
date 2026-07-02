import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/chongqing_research_award_dashboard";
const payload = JSON.parse(await fs.readFile(`${outputDir}/data.json`, "utf8"));
const rows = payload.rows;
const maxRow = rows.length + 1;

const workbook = Workbook.create();
const readme = workbook.worksheets.add("说明");
const raw = workbook.worksheets.add("原始数据");
const dash = workbook.worksheets.add("动态看板");
const lookup = workbook.worksheets.add("图表数据");

const colors = {
  navy: "#17324D",
  blue: "#2F6B9A",
  teal: "#278D8D",
  gold: "#C99A2E",
  red: "#B8554E",
  gray: "#6B7280",
  light: "#F5F7FA",
  paleBlue: "#EAF2F8",
  paleGold: "#FBF4E3",
  border: "#D8DEE8",
  white: "#FFFFFF",
};

function colName(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function rangeFor(sheet, row, col, values) {
  const height = values.length;
  const width = values[0].length;
  return sheet.getRange(`${colName(col)}${row}:${colName(col + width - 1)}${row + height - 1}`);
}

function writeBlock(sheet, row, col, values, asFormulas = false) {
  const range = rangeFor(sheet, row, col, values);
  if (asFormulas) range.formulas = values;
  else range.values = values;
  return range;
}

function styleTitle(range) {
  range.format.fill = { color: colors.paleBlue };
  range.format.font = { color: colors.navy, bold: true, size: 16 };
  range.format.borders = { preset: "outside", style: "thin", color: colors.navy };
}

function styleHeader(range) {
  range.format.fill = { color: colors.paleBlue };
  range.format.font = { bold: true, color: colors.navy };
  range.format.borders = { preset: "outside", style: "thin", color: colors.border };
}

function styleTable(range) {
  range.format.borders = { preset: "inside", style: "thin", color: colors.border };
}

function styleCard(range, fill = colors.light) {
  range.format.fill = { color: fill };
  range.format.borders = { preset: "outside", style: "thin", color: colors.border };
}

for (const sheet of [readme, raw, dash, lookup]) {
  sheet.showGridLines = false;
}

readme.getRange("A1:F1").merge();
readme.getRange("A1").values = [["重庆社科科学院发展研究奖 - 评审结果动态分析"]];
styleTitle(readme.getRange("A1:F1"));
writeBlock(readme, 3, 1, [
  ["工作簿结构", "用途"],
  ["原始数据", "合并三个获奖结果表，保留来源奖项、评分差异、分数区间与临界标记。"],
  ["动态看板", "通过奖项筛选联动 KPI、分数分布、评分分歧和评分均值对比图。"],
  ["图表数据", "看板图表的公式驱动数据区，便于审计和后续扩展。"],
]);
styleHeader(readme.getRange("A3:B3"));
styleTable(readme.getRange("A3:B6"));
writeBlock(readme, 8, 1, [
  ["关注点", "看板解释"],
  ["名额结构", "观察一二三等奖数量是否符合评奖安排。"],
  ["分数区间", "识别获奖作品集中在高分段还是临界分段。"],
  ["评分分歧", "大会评分相对小组评分的上调/下调情况，用于发现复核重点。"],
  ["临界作品", "各奖项分数较靠后的作品，可作为复核和公示风险关注对象。"],
  ["重复选题", "同名成果多次出现时提示选题聚集度，便于系统后续做去重或主题归并。"],
]);
styleHeader(readme.getRange("A8:B8"));
styleTable(readme.getRange("A8:B13"));
readme.getRange("A:A").format.columnWidth = 18;
readme.getRange("B:B").format.columnWidth = 72;

const rawHeaders = ["编号", "成果名称", "奖项", "奖项内排名", "小组评分", "大会评分", "大会-小组差", "综合最终得分", "分数区间", "是否临界"];
const rawValues = rows.map((r) => rawHeaders.map((h) => r[h]));
writeBlock(raw, 1, 1, [rawHeaders]);
writeBlock(raw, 2, 1, rawValues);
styleHeader(raw.getRange("A1:J1"));
styleTable(raw.getRange(`A1:J${maxRow}`));
raw.freezePanes.freezeRows(1);
raw.getRange("A:A").format.columnWidth = 14;
raw.getRange("B:B").format.columnWidth = 52;
raw.getRange("C:D").format.columnWidth = 12;
raw.getRange("E:H").format.columnWidth = 13;
raw.getRange("I:J").format.columnWidth = 12;
raw.getRange(`E2:H${maxRow}`).format.numberFormat = "0.00";

lookup.getRange("A1").values = [["奖项选项"]];
writeBlock(lookup, 2, 1, [["全部"], ["一等奖"], ["二等奖"], ["三等奖"]]);
lookup.getRange("C1").values = [["奖项数量"]];
writeBlock(lookup, 2, 3, [["奖项", "数量"], ["一等奖", null], ["二等奖", null], ["三等奖", null]]);
writeBlock(lookup, 3, 4, [
  [`=COUNTIF('原始数据'!$C$2:$C$${maxRow},C3)`],
  [`=COUNTIF('原始数据'!$C$2:$C$${maxRow},C4)`],
  [`=COUNTIF('原始数据'!$C$2:$C$${maxRow},C5)`],
], true);
styleHeader(lookup.getRange("A1:A1"));
styleHeader(lookup.getRange("C2:D2"));
lookup.getRange("A:D").format.columnWidth = 14;

writeBlock(dash, 1, 1, [["重庆社科科学院发展研究奖评审结果动态看板", null, null, null, null, null, null, null, null, null, null, null, null, null]]);
styleTitle(dash.getRange("A1:N1"));
dash.getRange("A1:N1").format.rowHeight = 30;
dash.getRange("A3").values = [["奖项筛选"]];
dash.getRange("B3").values = [["全部"]];
dash.getRange("B3").dataValidation = { rule: { type: "list", formula1: "'图表数据'!$A$2:$A$5" } };
styleHeader(dash.getRange("A3:A3"));
styleCard(dash.getRange("B3:B3"), colors.paleGold);

writeBlock(dash, 5, 1, [
  ["获奖成果数", "平均最终分", "最高分", "最低分", "大会-小组均差"],
]);
styleHeader(dash.getRange("A5:E5"));
writeBlock(dash, 6, 1, [[
  `=IF($B$3="全部",COUNTA('原始数据'!$A$2:$A$${maxRow}),COUNTIF('原始数据'!$C$2:$C$${maxRow},$B$3))`,
  `=IF($B$3="全部",AVERAGE('原始数据'!$H$2:$H$${maxRow}),AVERAGEIF('原始数据'!$C$2:$C$${maxRow},$B$3,'原始数据'!$H$2:$H$${maxRow}))`,
  `=IF($B$3="全部",MAX('原始数据'!$H$2:$H$${maxRow}),MAXIFS('原始数据'!$H$2:$H$${maxRow},'原始数据'!$C$2:$C$${maxRow},$B$3))`,
  `=IF($B$3="全部",MIN('原始数据'!$H$2:$H$${maxRow}),MINIFS('原始数据'!$H$2:$H$${maxRow},'原始数据'!$C$2:$C$${maxRow},$B$3))`,
  `=IF($B$3="全部",AVERAGE('原始数据'!$G$2:$G$${maxRow}),AVERAGEIF('原始数据'!$C$2:$C$${maxRow},$B$3,'原始数据'!$G$2:$G$${maxRow}))`,
]], true);
styleCard(dash.getRange("A6:E6"), colors.light);
dash.getRange("A6:E6").format.font = { bold: true, size: 14, color: colors.navy };
dash.getRange("B6:E6").format.numberFormat = "0.00";

writeBlock(dash, 8, 1, [["分数区间", "数量"]]);
writeBlock(dash, 9, 1, [["94分及以上"], ["92-94分"], ["90-92分"], ["88-90分"], ["88分以下"]]);
writeBlock(dash, 9, 2, [
  [`=IF($B$3="全部",COUNTIF('原始数据'!$I$2:$I$${maxRow},A9),COUNTIFS('原始数据'!$I$2:$I$${maxRow},A9,'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIF('原始数据'!$I$2:$I$${maxRow},A10),COUNTIFS('原始数据'!$I$2:$I$${maxRow},A10,'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIF('原始数据'!$I$2:$I$${maxRow},A11),COUNTIFS('原始数据'!$I$2:$I$${maxRow},A11,'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIF('原始数据'!$I$2:$I$${maxRow},A12),COUNTIFS('原始数据'!$I$2:$I$${maxRow},A12,'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIF('原始数据'!$I$2:$I$${maxRow},A13),COUNTIFS('原始数据'!$I$2:$I$${maxRow},A13,'原始数据'!$C$2:$C$${maxRow},$B$3))`],
], true);
styleHeader(dash.getRange("A8:B8"));
styleTable(dash.getRange("A8:B13"));

writeBlock(dash, 8, 4, [["评分项", "均值"]]);
writeBlock(dash, 9, 4, [["小组评分"], ["大会评分"], ["综合最终得分"]]);
writeBlock(dash, 9, 5, [
  [`=IF($B$3="全部",AVERAGE('原始数据'!$E$2:$E$${maxRow}),AVERAGEIF('原始数据'!$C$2:$C$${maxRow},$B$3,'原始数据'!$E$2:$E$${maxRow}))`],
  [`=IF($B$3="全部",AVERAGE('原始数据'!$F$2:$F$${maxRow}),AVERAGEIF('原始数据'!$C$2:$C$${maxRow},$B$3,'原始数据'!$F$2:$F$${maxRow}))`],
  [`=IF($B$3="全部",AVERAGE('原始数据'!$H$2:$H$${maxRow}),AVERAGEIF('原始数据'!$C$2:$C$${maxRow},$B$3,'原始数据'!$H$2:$H$${maxRow}))`],
], true);
styleHeader(dash.getRange("D8:E8"));
styleTable(dash.getRange("D8:E11"));
dash.getRange("E9:E11").format.numberFormat = "0.00";

writeBlock(dash, 15, 1, [["大会-小组差异", "数量"]]);
writeBlock(dash, 16, 1, [["大会低1分以上"], ["大会低0-1分"], ["大会高0-1分"], ["大会高1分以上"]]);
writeBlock(dash, 16, 2, [
  [`=IF($B$3="全部",COUNTIF('原始数据'!$G$2:$G$${maxRow},"<-1"),COUNTIFS('原始数据'!$G$2:$G$${maxRow},"<-1",'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIFS('原始数据'!$G$2:$G$${maxRow},">=-1",'原始数据'!$G$2:$G$${maxRow},"<0"),COUNTIFS('原始数据'!$G$2:$G$${maxRow},">=-1",'原始数据'!$G$2:$G$${maxRow},"<0",'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIFS('原始数据'!$G$2:$G$${maxRow},">=0",'原始数据'!$G$2:$G$${maxRow},"<1"),COUNTIFS('原始数据'!$G$2:$G$${maxRow},">=0",'原始数据'!$G$2:$G$${maxRow},"<1",'原始数据'!$C$2:$C$${maxRow},$B$3))`],
  [`=IF($B$3="全部",COUNTIF('原始数据'!$G$2:$G$${maxRow},">=1"),COUNTIFS('原始数据'!$G$2:$G$${maxRow},">=1",'原始数据'!$C$2:$C$${maxRow},$B$3))`],
], true);
styleHeader(dash.getRange("A15:B15"));
styleTable(dash.getRange("A15:B19"));

writeBlock(dash, 15, 4, [["奖项", "数量"]]);
writeBlock(dash, 16, 4, [["一等奖"], ["二等奖"], ["三等奖"]]);
writeBlock(dash, 16, 5, [
  [`=COUNTIF('原始数据'!$C$2:$C$${maxRow},D16)`],
  [`=COUNTIF('原始数据'!$C$2:$C$${maxRow},D17)`],
  [`=COUNTIF('原始数据'!$C$2:$C$${maxRow},D18)`],
], true);
styleHeader(dash.getRange("D15:E15"));
styleTable(dash.getRange("D15:E18"));

const criticalRows = rows
  .filter((r) => r["是否临界"])
  .sort((a, b) => a["综合最终得分"] - b["综合最终得分"])
  .slice(0, 12)
  .map((r) => [r["编号"], r["成果名称"], r["奖项"], r["综合最终得分"], r["大会-小组差"]]);
writeBlock(dash, 22, 1, [["临界作品关注清单", "", "", "", ""]]);
dash.getRange("A22:E22").merge();
styleHeader(dash.getRange("A22:E22"));
writeBlock(dash, 23, 1, [["编号", "成果名称", "奖项", "最终分", "大会-小组差"], ...criticalRows]);
styleHeader(dash.getRange("A23:E23"));
styleTable(dash.getRange(`A23:E${23 + criticalRows.length}`));
dash.getRange(`D24:E${23 + criticalRows.length}`).format.numberFormat = "0.00";

const byTitle = new Map();
for (const r of rows) {
  const item = byTitle.get(r["成果名称"]) ?? { count: 0, total: 0, awards: new Set() };
  item.count += 1;
  item.total += r["综合最终得分"];
  item.awards.add(r["奖项"]);
  byTitle.set(r["成果名称"], item);
}
const duplicateRows = [...byTitle.entries()]
  .filter(([, v]) => v.count > 1)
  .sort((a, b) => b[1].count - a[1].count || b[1].total / b[1].count - a[1].total / a[1].count)
  .slice(0, 10)
  .map(([title, v]) => [title, v.count, [...v.awards].join("、"), v.total / v.count]);
writeBlock(dash, 22, 7, [["重复选题聚集度", "", "", ""]]);
dash.getRange("G22:J22").merge();
styleHeader(dash.getRange("G22:J22"));
writeBlock(dash, 23, 7, [["成果名称", "出现次数", "涉及奖项", "平均最终分"], ...duplicateRows]);
styleHeader(dash.getRange("G23:J23"));
styleTable(dash.getRange(`G23:J${23 + duplicateRows.length}`));
dash.getRange(`J24:J${23 + duplicateRows.length}`).format.numberFormat = "0.00";

for (const c of ["A", "D", "G"]) dash.getRange(`${c}:${c}`).format.columnWidth = 19;
dash.getRange("B:B").format.columnWidth = 14;
dash.getRange("C:C").format.columnWidth = 12;
dash.getRange("E:E").format.columnWidth = 14;
dash.getRange("H:H").format.columnWidth = 11;
dash.getRange("I:I").format.columnWidth = 18;
dash.getRange("J:J").format.columnWidth = 13;
dash.getRange("F:F").format.columnWidth = 4;
dash.getRange("K:N").format.columnWidth = 13;
dash.getRange("A1:N30").format.wrapText = true;
dash.getRange("A5:E6").format.wrapText = false;
dash.getRange("A1:N1").format.wrapText = false;
styleTitle(dash.getRange("A1:N1"));
dash.freezePanes.freezeRows(3);

const scoreChart = dash.charts.add("bar", dash.getRange("A8:B13"), "Auto");
scoreChart.title.text = "筛选奖项的分数区间分布";
scoreChart.setPosition(dash.getRange("G4:N12"));
scoreChart.width = 620;
scoreChart.height = 260;
scoreChart.hasLegend = false;
scoreChart.xAxis = { textStyle: { fontSize: 10 } };
scoreChart.yAxis = { majorGridlines: { fill: colors.border, style: "solid", width: 1 } };

const avgChart = dash.charts.add("bar", dash.getRange("D8:E11"), "Auto");
avgChart.title.text = "小组评分、大会评分与最终分均值";
avgChart.setPosition(dash.getRange("G13:N21"));
avgChart.width = 620;
avgChart.height = 245;
avgChart.hasLegend = false;
avgChart.yAxis = { numberFormatCode: "0.00", majorGridlines: { fill: colors.border, style: "solid", width: 1 } };

const diffChart = dash.charts.add("bar", dash.getRange("A15:B19"), "Auto");
diffChart.title.text = "大会评分相对小组评分差异分布";
diffChart.setPosition(dash.getRange("K22:N31"));
diffChart.width = 360;
diffChart.height = 260;
diffChart.hasLegend = false;

const awardChart = dash.charts.add("bar", dash.getRange("D15:E18"), "Auto");
awardChart.title.text = "各奖项成果数量结构";
awardChart.setPosition(dash.getRange("K31:N38"));
awardChart.width = 360;
awardChart.height = 190;
awardChart.hasLegend = false;

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errorScan.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({ sheetName: "动态看板", range: "A1:N40", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/dashboard_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/社科院发展研究奖_动态分析看板.xlsx`);
