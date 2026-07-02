import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "outputs/chongqing_research_award_dashboard";
const input = await FileBlob.load(`${outputDir}/社科院发展研究奖_动态分析看板.xlsx`);
const workbook = await SpreadsheetFile.importXlsx(input);

const key = await workbook.inspect({
  kind: "table",
  range: "动态看板!A1:N20",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 14,
  maxChars: 5000,
});
console.log(key.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of ["说明", "原始数据", "动态看板", "图表数据"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/verify_${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
